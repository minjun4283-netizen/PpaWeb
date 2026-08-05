# Builds the frontend static site and the backend into one small image that
# serves both from a single Node process — simplest thing to run on an
# intranet server that doesn't have a reverse proxy set up yet.

FROM node:20-slim AS frontend-build
WORKDIR /app/frontend
COPY frontend/package*.json ./
RUN npm install
COPY frontend/ ./
RUN npm run build

FROM node:20-slim AS backend-build
WORKDIR /app/backend
COPY backend/package*.json ./
RUN npm install
COPY backend/ ./
RUN npm run build

FROM node:20-slim
WORKDIR /app
ENV NODE_ENV=production
COPY backend/package*.json ./
RUN npm install --omit=dev
COPY --from=backend-build /app/backend/dist ./dist
COPY --from=frontend-build /app/frontend/dist ./dist/public

VOLUME ["/app/data"]
ENV DB_PATH=/app/data/ppaweb.db
ENV PORT=4000
EXPOSE 4000

CMD ["node", "dist/index.js"]
