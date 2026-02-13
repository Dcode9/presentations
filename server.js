/**
 * AI Presentation Generator - Server
 *
 * Express server that orchestrates:
 * 1. Cerebras API for slide content generation
 * 2. Pollinations API for image generation
 * 3. PPTX (ZIP) assembly
 *
 * Uses Server-Sent Events (SSE) for real-time progress logging.
 */

const express = require('express');
const path = require('path');
const { generateSlideContent } = require('./lib/cerebras');
const { generateImage } = require('./lib/image-gen');
const { buildPptx } = require('./lib/pptx-builder');

const app = express();
const PORT = process.env.PORT || 3000;

app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

/**
 * POST /api/generate
 *
 * Body: { topic: string, apiKey: string }
 *
 * Returns the PPTX file as a downloadable ZIP.
 * Streams progress via SSE-style logging in a separate endpoint.
 */

// In-memory store for generation jobs (for progress tracking)
const jobs = new Map();

/**
 * POST /api/generate - Start a presentation generation job
 */
app.post('/api/generate', (req, res) => {
  const { topic, apiKey } = req.body;

  if (!topic || !topic.trim()) {
    return res.status(400).json({ error: 'Topic is required' });
  }
  if (!apiKey || !apiKey.trim()) {
    return res.status(400).json({ error: 'Cerebras API key is required' });
  }

  const jobId = Date.now().toString(36) + Math.random().toString(36).substring(2, 8);
  const job = {
    id: jobId,
    topic: topic.trim(),
    status: 'started',
    logs: [],
    result: null,
    error: null,
  };
  jobs.set(jobId, job);

  // Start generation in background
  runGeneration(job, apiKey.trim()).catch((err) => {
    job.status = 'error';
    job.error = err.message;
    job.logs.push(`❌ Error: ${err.message}`);
  });

  res.json({ jobId });
});

/**
 * GET /api/progress/:jobId - SSE endpoint for real-time logs
 */
app.get('/api/progress/:jobId', (req, res) => {
  const job = jobs.get(req.params.jobId);
  if (!job) {
    return res.status(404).json({ error: 'Job not found' });
  }

  res.writeHead(200, {
    'Content-Type': 'text/event-stream',
    'Cache-Control': 'no-cache',
    Connection: 'keep-alive',
  });

  let lastIndex = 0;
  const interval = setInterval(() => {
    // Send any new log entries
    while (lastIndex < job.logs.length) {
      const logEntry = job.logs[lastIndex];
      res.write(`data: ${JSON.stringify({ type: 'log', message: logEntry })}\n\n`);
      lastIndex++;
    }

    // Check if job is complete
    if (job.status === 'done') {
      res.write(`data: ${JSON.stringify({ type: 'done', jobId: job.id })}\n\n`);
      clearInterval(interval);
      res.end();
    } else if (job.status === 'error') {
      res.write(`data: ${JSON.stringify({ type: 'error', message: job.error })}\n\n`);
      clearInterval(interval);
      res.end();
    }
  }, 200);

  req.on('close', () => {
    clearInterval(interval);
  });
});

/**
 * GET /api/download/:jobId - Download the generated PPTX
 */
app.get('/api/download/:jobId', (req, res) => {
  const job = jobs.get(req.params.jobId);
  if (!job) {
    return res.status(404).json({ error: 'Job not found' });
  }
  if (job.status !== 'done' || !job.result) {
    return res.status(400).json({ error: 'PPTX not ready yet' });
  }

  const filename = job.topic.replace(/[^a-zA-Z0-9 ]/g, '').substring(0, 50).trim().replace(/\s+/g, '_') || 'presentation';

  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.presentationml.presentation');
  res.setHeader('Content-Disposition', `attachment; filename="${filename}.pptx"`);
  res.send(job.result);
});

/**
 * Run the full generation pipeline.
 */
async function runGeneration(job, apiKey) {
  const log = (msg) => {
    const timestamp = new Date().toISOString().substring(11, 23);
    const entry = `[${timestamp}] ${msg}`;
    job.logs.push(entry);
  };

  try {
    log('🚀 Starting presentation generation...');
    log(`📋 Topic: "${job.topic}"`);

    // Step 1: Generate slide content via Cerebras
    log('');
    log('═══════════════════════════════════════');
    log('  STEP 1: Generate Slide Content (Cerebras API)');
    log('═══════════════════════════════════════');

    const slideData = await generateSlideContent(apiKey, job.topic, log);
    const slides = slideData.slides;

    log('');
    log('📊 Generated Slides:');
    slides.forEach((s, i) => {
      log(`   Slide ${i + 1}: "${s.title}" (${s.bullets.length} bullets, bg: #${s.bgColor}, anim: ${s.animation})`);
    });

    // Step 2: Generate images via Pollinations
    log('');
    log('═══════════════════════════════════════');
    log('  STEP 2: Generate Images (Pollinations API)');
    log('═══════════════════════════════════════');

    const images = [];
    for (let i = 0; i < slides.length; i++) {
      const slide = slides[i];
      if (slide.imagePrompt) {
        log(`\n🖼️  Slide ${i + 1} image:`);
        log(`   Prompt: "${slide.imagePrompt}"`);
        try {
          const buffer = await generateImage(slide.imagePrompt, {
            width: 1024,
            height: 768,
            onLog: log,
          });
          images.push({ buffer, filename: `image${i + 1}.png` });
        } catch (err) {
          log(`⚠️ Image generation failed for slide ${i + 1}: ${err.message}`);
          log('   Continuing without image for this slide...');
          images.push(null);
        }
      } else {
        images.push(null);
      }
    }

    // Step 3: Build PPTX
    log('');
    log('═══════════════════════════════════════');
    log('  STEP 3: Assemble PPTX File');
    log('═══════════════════════════════════════');

    const pptxStream = buildPptx({
      title: job.topic,
      slides,
      images,
      onLog: log,
    });

    // Collect the archive into a buffer
    const chunks = [];
    await new Promise((resolve, reject) => {
      pptxStream.on('data', (chunk) => chunks.push(chunk));
      pptxStream.on('end', resolve);
      pptxStream.on('error', reject);
    });

    const pptxBuffer = Buffer.concat(chunks);
    job.result = pptxBuffer;

    log('');
    log('═══════════════════════════════════════');
    log('  ✅ GENERATION COMPLETE');
    log('═══════════════════════════════════════');
    log(`📁 PPTX file size: ${(pptxBuffer.length / 1024).toFixed(1)} KB`);
    log(`📊 Total slides: ${slides.length}`);
    log(`🖼️  Images included: ${images.filter(Boolean).length}`);
    log('');
    log('💾 Ready for download!');

    job.status = 'done';

    // Clean up old jobs after 30 minutes
    setTimeout(() => {
      jobs.delete(job.id);
    }, 30 * 60 * 1000);

  } catch (err) {
    log(`❌ Fatal error: ${err.message}`);
    job.status = 'error';
    job.error = err.message;
    throw err;
  }
}

app.listen(PORT, () => {
  console.log(`🚀 AI Presentation Generator running at http://localhost:${PORT}`);
  console.log(`   Open in your browser to start creating presentations!`);
});

module.exports = app;
