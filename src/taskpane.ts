/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

Office.onReady(info => {
  if (info.host === Office.HostType.Excel) {
    Excel.run(async context => {
      // hot reload
      const range = context.workbook.getSelectedRange();
      await context.sync()
      const binding = context.workbook.bindings.add(range, "Range", "HotReload");
      binding.onDataChanged.add(run);
      await context.sync();
      
      // initial load
      run({ binding });
      await context.sync();
    });
    
    // Load UI
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("copy").onclick = copy;
  }
});

function copy() {
  const node = document.getElementById("container");
  // eslint-disable-next-line no-undef
  const selection = window.getSelection();
  const range = document.createRange();
  range.selectNodeContents(node);
  selection.removeAllRanges();
  selection.addRange(range);
  document.execCommand("copy");
  selection.empty();
}

export async function run(event: Excel.BindingDataChangedEventArgs) {
  try {
    Excel.run(async context => {
      const range = event.binding.getRange();
      range.load('values');
      const props = range.getCellProperties({
        format: {
          font: {
            bold: true,
            color: true,
            italic: true,
            strikethrough: true,
            subscript: true,
            superscript: true,
            underline: true
          }
        }
      });

      await context.sync();

      document.getElementById("container").innerText = rangeToLatex(range, props);
    });
  } catch (error) {
    console.error(error);
    // "Please select a single range"
    //if (error instanceof OfficeExtension.Error) {
    //  console.log("Debug info: " + JSON.stringify(error.debugInfo));
    //}
  }
}

function rangeToLatex(range: Excel.Range, props: OfficeExtension.ClientResult<Excel.CellProperties[][]>) {
  // Convert single-row ranges (1D array) to standard 2D
  if (typeof range.values[0] === 'string') {
    range.values = [range.values];
  }

  let packages = {};
  const rows = range.values.map((row, i) => {
    return row.map((value, j) => {
      if (value === '') return '';
      value = value.toString(); // values can be strings, numbers or bools

      const formatted = applyFormat(value, props.value[i][j].format, packages);
      packages = formatted.packages;
      return formatted.value;
    }).join(' & ') + '\\\\';
  });

  return [
    ...Object.values(packages),
    "\\begin{tabular}",
    ...rows,
    "\\end{tabular}"
  ].join('\n');
}

// Only applies if format applied to entire cell, not individual letters
function applyFormat(value: string, format, packages) {
  if (format.font.bold) value = `\\textbf{${value}}`

  if (format.font.color !== "#000000") {
    if (!packages.color) packages.color = "\\usepackage{xcolor}";
    value = `\\textcolor{${format.font.color}}{${value}}`;
  }

  if (format.font.italic) value = `\\textit{${value}}`

  // apply font?

  // don't change font size

  if (format.font.strikethrough) {
    if (!packages.strikethrough) packages.strikethrough = "\\usepackage{cancel}";
    value = `\\cancel{${value}}`;
  }

  if (format.font.subscript) {
    if (!packages.subscript) packages.subscript = "\\usepackage{fixltx2e}";
    value = `\\textsubscript{${value}}`;
  }

  if (format.font.superscript) value = `\\textsuperscript{${value}}`

  if (format.font.underline === "Single") value = `\\underline{${value}}`

  return { value, packages };
}