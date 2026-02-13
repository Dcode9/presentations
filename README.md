# AI Presentation Generator

Generate PowerPoint (.pptx) presentations using AI — powered by **Cerebras** for content and **Pollinations** for images.

## How It Works

A `.pptx` file is just a ZIP archive containing XML files and media (images). This app:

1. Takes your topic and sends it to **Cerebras AI** to generate structured slide content (titles, bullets, colors, animations)
2. Generates images for each slide using **Pollinations AI**
3. Assembles everything into valid Office Open XML (OOXML) structure
4. Packages it as a `.pptx` file you can open in PowerPoint, Google Slides, or LibreOffice

## Quick Start

```bash
npm install
npm start
```

Then open http://localhost:3000 in your browser.

You'll need a [Cerebras API key](https://cloud.cerebras.ai/) to generate presentations.

## Project Structure

```
├── server.js                 # Express server with API endpoints
├── lib/
│   ├── cerebras.js           # Cerebras API integration
│   ├── image-gen.js          # Pollinations image generation
│   ├── pptx-builder.js       # PPTX ZIP assembly
│   └── pptx-templates.js     # Office Open XML templates
├── public/
│   └── index.html            # Web frontend with live progress
└── test/
    └── test-pptx-templates.js # Template validation tests
```

## API Endpoints

- `POST /api/generate` — Start a presentation generation job
  - Body: `{ "topic": "...", "apiKey": "csk-..." }`
  - Returns: `{ "jobId": "..." }`
- `GET /api/progress/:jobId` — SSE stream of real-time generation logs
- `GET /api/download/:jobId` — Download the generated `.pptx` file

## Features

- AI-generated slide content with proper structure
- AI-generated images for each slide
- Entrance animations (fade in, fly in, wipe, appear)
- Slide transitions
- Custom color schemes per slide
- Real-time progress logging (see everything the AI does)
- Valid PPTX output compatible with PowerPoint and other apps