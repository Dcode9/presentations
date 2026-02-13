/**
 * Image generation using Pollinations API (enter.pollinations.ai).
 *
 * Uses the gen.pollinations.ai gateway with the imagen-2 model.
 * API docs: https://github.com/pollinations/pollinations/blob/main/APIDOCS.md
 * API keys: https://enter.pollinations.ai
 */

const fetch = require('node-fetch');

const POLLINATIONS_BASE = 'https://gen.pollinations.ai/image';
const POLLINATIONS_MODEL = 'imagen-2';

/**
 * Generate an image from a text prompt using Pollinations API with imagen-2.
 *
 * @param {string} prompt - The image description prompt
 * @param {object} opts
 * @param {number} opts.width - Image width (default 1024)
 * @param {number} opts.height - Image height (default 768)
 * @param {function} opts.onLog - Logging callback
 * @returns {Promise<Buffer>} PNG image data as Buffer
 */
async function generateImage(prompt, opts = {}) {
  const { width = 1024, height = 768, onLog = () => {} } = opts;

  const apiKey = process.env.POLLINATIONS_API_KEY;

  const encodedPrompt = encodeURIComponent(prompt);
  let url = `${POLLINATIONS_BASE}/${encodedPrompt}?model=${POLLINATIONS_MODEL}&width=${width}&height=${height}&nologo=true`;

  const headers = {};
  if (apiKey) {
    headers['Authorization'] = `Bearer ${apiKey}`;
    onLog(`🎨 Generating image (imagen-2, authenticated): "${prompt.substring(0, 80)}..."`);
  } else {
    onLog(`🎨 Generating image (imagen-2, no API key): "${prompt.substring(0, 80)}..."`);
  }
  onLog(`   URL: ${url.substring(0, 150)}...`);
  onLog(`   Model: ${POLLINATIONS_MODEL}`);

  const response = await fetch(url, { headers, timeout: 120000 });

  if (!response.ok) {
    throw new Error(`Pollinations API error (${response.status}): ${response.statusText}`);
  }

  const buffer = await response.buffer();
  onLog(`✅ Image generated: ${(buffer.length / 1024).toFixed(1)} KB`);

  return buffer;
}

module.exports = { generateImage };
