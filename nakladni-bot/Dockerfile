# Node 24 запускає .ts напряму, тож збирати TypeScript не треба.
FROM node:24-alpine

WORKDIR /app
ENV NODE_ENV=production

# Спершу лише package*.json: поки залежності не змінюються, Docker бере цей шар із кешу.
COPY package.json package-lock.json ./
RUN npm ci --omit=dev && npm cache clean --force

COPY src ./src
# Стартовий прайс. У docker-compose папку data/ підміняє volume, щоб оновлення прайсу не губилися.
COPY data ./data
RUN chown -R node:node /app/data

# Не запускаємо від root.
USER node

CMD ["node", "src/bot.ts"]
