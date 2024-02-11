FROM node:21.6.1 as build
WORKDIR /app

COPY package*.json ./
RUN npm ci

COPY . .
RUN npm run build

FROM joseluisq/static-web-server:2.26.0

COPY --from=build /app/dist /public
