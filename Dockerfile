FROM node:21.7.0 as build
WORKDIR /app

COPY package*.json ./
RUN npm ci

COPY . .
RUN npm run build

FROM joseluisq/static-web-server:2.27.0

COPY --from=build /app/dist /public
