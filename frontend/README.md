# ACRA - PowerPoint Management Application

This React application provides a user interface for creating, managing, and viewing PowerPoint presentations with an integrated chatbot.

## Features

- PowerPoint creation and editing using pptxgenjs
- File listing and management 
- Integrated ChatBot (via iframe)

## Project Structure

The application is structured into three main areas:
1. PPTX Viewer/Editor (top left)
2. File Manager (bottom left)
3. ChatBot (right side)

## Development

### Prerequisites

- Node.js 16+
- Docker and Docker Compose

### Local Development

```bash
# Install dependencies
npm install

# Start the development server
npm start
```

### Docker Development

```bash
# Build and start all services
docker-compose up --build
```

## Technologies Used

- React
- pptxgenjs for PowerPoint manipulation
- Docker for containerization 