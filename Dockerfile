FROM node:22-slim AS build

WORKDIR /app

COPY package.json package-lock.json tsconfig.json ./
RUN npm ci

COPY src ./src
RUN npm run build

FROM node:22-slim AS runtime

WORKDIR /app
ENV NODE_ENV=production
ENV PORT=3000
ENV POWERBI_MCP_CACHE_DIR=/data/powerbi-mcp-claude

COPY package.json package-lock.json ./
RUN npm ci --omit=dev \
  && npm cache clean --force \
  && mkdir -p /data/powerbi-mcp-claude \
  && chown -R node:node /app /data/powerbi-mcp-claude

COPY --from=build --chown=node:node /app/dist ./dist

EXPOSE 3000
VOLUME ["/data/powerbi-mcp-claude"]

USER node

CMD ["node", "dist/src/httpServer.js"]
