# Runtime image. No npm install — the app has zero runtime dependencies.
# Test-only tools (jsdom, openpyxl) stay in CI and are not copied in.
FROM node:20-alpine

LABEL org.opencontainers.image.title="Open365" \
      org.opencontainers.image.description="Self-hosted open-source office suite" \
      org.opencontainers.image.licenses="AGPL-3.0-or-later"

WORKDIR /app

COPY package.json server.js LICENSE ./
COPY public ./public

RUN mkdir -p /data /app/data \
  && chown -R node:node /data /app

USER node

ENV NODE_ENV=production \
    HOST=0.0.0.0 \
    PORT=4173 \
    DATA_DIR=/data

EXPOSE 4173
VOLUME ["/data"]

HEALTHCHECK --interval=30s --timeout=5s --start-period=15s --retries=3 \
  CMD node -e "fetch('http://127.0.0.1:'+(process.env.PORT||4173)+'/healthz').then(r=>process.exit(r.ok?0:1)).catch(()=>process.exit(1))"

CMD ["node", "server.js"]
