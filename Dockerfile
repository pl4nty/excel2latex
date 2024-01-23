FROM node:21.6.0 as build
WORKDIR /app

COPY package*.json ./
RUN npm ci

COPY . .
RUN npm run build

FROM joseluisq/static-web-server:2.25.0

COPY --from=build /app/dist /public
