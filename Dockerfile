FROM node:22-bookworm-slim
WORKDIR /app
ENV NODE_ENV=production
COPY package*.json ./
RUN npm ci --omit=dev
COPY . .
RUN mkdir -p /app/data
ENV PORT=3050 HOST=0.0.0.0 PM_DATA_ROOT=/app/data
EXPOSE 3050
CMD ["node","server.js"]
