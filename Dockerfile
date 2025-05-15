FROM node:22.15.1 as build
WORKDIR /app

COPY package*.json ./
RUN npm ci

COPY . .
RUN npm run build

FROM joseluisq/static-web-server:2.36.1

COPY --from=build /app/dist /public
