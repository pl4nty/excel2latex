/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

Office.onReady(info => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});

export async function run() {
  try {
    Excel.run(async context => {
      const range = context.workbook.getSelectedRange();
      range.load(['values']);
      const cellProps = range.getCellProperties({
        format: {
          font: {
            color: true
          }
        }
      });
      await context.sync();

      const grid = range.values.map((row, i) => {
        if (typeof row === 'string') {
          return [row, cellProps.value[i][0].format]
        } else {
          return row.map((value, j) => [value, cellProps.value[i][j].format])
        }
      });

      document.getElementById("container").innerHTML = toLatex(grid);
    });
    
  } catch (error) {
    console.error(error);
  }
}

function toLatex(range) {
  let packages = {
    color: '',
    strikethrough: '',
    subscript: ''
  };
  let body = range.map(row => {
    return row.map(cell => {
      cell[0] = cell[0].toString(); // cells can be strings, numbers or bools
      if (cell[0] === '') cell[0]=' ';
      let [p,text] = applyFormat(packages, cell);
      packages = p;
      return text;
    }).join(' & ')+'\\\\\n';
  });
  body = `\\begin{tabular}\n${body}\\end{tabular}`;
  return body;
}

// Only applies if format applied to entire cell, not just letters
function applyFormat(packages, [value, format]) {
  if (format.font.bold) value = `\\textbf{${value}}`
  if (format.font.color !== "#000000") {
    if (!packages.color) packages.color = "\\usepackage{xcolor}\n";
    value = `\\textcolor{${format.font.color}}{${value}}`;
  }
  if (format.font.italic) value = `\\textit{${value}}`
  // apply font?
  // don't change font size
  if (format.font.strikethrough) {
    if (!packages.strikethrough) packages.strikethrough = "\\usepackage{cancel}\n";
    value = `\\cancel{${value}}`;
  }
  if (format.font.subscript) {
    if (!packages.subscript) packages.subscript = "\\usepackage{fixltx2e}\n";
    value = `\\textsubscript{${value}}`;
  }
  if (format.font.superscript) value = `\\textsuperscript{${value}}`
  if (format.font.underline) value = `\\underline{${value}}`
  return [packages, value];
}

document.getElementById("copy").addEventListener('click', () => {
  const node = document.getElementById("container");
  const selection = window.getSelection();
  const range = document.createRange();
  range.selectNodeContents(node);
  selection.removeAllRanges();
  selection.addRange(range);
  document.execCommand("copy");
  selection.empty();
});