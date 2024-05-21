FROM node:21.7.3 as build
WORKDIR /app

COPY package*.json ./
RUN npm ci

COPY . .
RUN npm run build

FROM joseluisq/static-web-server:2.31.1

COPY --from=build /app/dist /public
