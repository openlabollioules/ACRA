FROM node:16-alpine

# Set working directory
WORKDIR /app

# Copy package.json and package-lock.json
COPY frontend/package*.json ./

# Install dependencies
RUN npm install

# Copy the public directory
COPY frontend/public ./public

# Copy the source files
COPY frontend/src ./src

EXPOSE 3000

CMD ["npm", "start"]