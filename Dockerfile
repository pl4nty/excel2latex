FROM node:24.11.0 as build
WORKDIR /app

COPY package*.json ./
RUN npm ci

COPY . .
RUN npm run build

FROM joseluisq/static-web-server:2.39.0

COPY --from=build /app/dist /public
