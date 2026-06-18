# --- build stage ---
FROM node:20-bookworm-slim AS build
WORKDIR /app
COPY zscaler.crt /usr/local/share/ca-certificates/zscaler.crt
RUN apt-get update \
  && apt-get install -y --no-install-recommends ca-certificates \
  && update-ca-certificates \
  && rm -rf /var/lib/apt/lists/*
ENV NODE_EXTRA_CA_CERTS=/usr/local/share/ca-certificates/zscaler.crt
ENV NPM_CONFIG_CAFILE=/usr/local/share/ca-certificates/zscaler.crt
COPY frontend/package.json frontend/package-lock.json* ./
RUN npm ci --include=dev --no-audit --no-fund
COPY frontend/ .
RUN npm run build

# --- serve stage ---
FROM nginx:alpine
COPY frontend/nginx.conf /etc/nginx/conf.d/default.conf
COPY --from=build /app/dist /usr/share/nginx/html
EXPOSE 80
