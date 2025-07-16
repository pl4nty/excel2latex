FROM node:22.17.1 as build
WORKDIR /app

COPY package*.json ./
RUN npm ci

COPY . .
RUN npm run build

FROM joseluisq/static-web-server:2.37.0

COPY --from=build /app/dist /public
