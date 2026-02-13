/**
 * Cerebras API integration for generating slide content.
 *
 * Uses the Cerebras Cloud API (OpenAI-compatible) to generate structured
 * slide data from a user's topic/prompt.
 */

const fetch = require('node-fetch');

const CEREBRAS_API_URL = 'https://api.cerebras.ai/v1/chat/completions';

const SYSTEM_PROMPT = `You are an expert presentation designer and content creator. Your job is to generate structured slide content for PowerPoint presentations.

IMPORTANT RULES:
1. You MUST respond with ONLY valid JSON. No markdown, no explanation, no code blocks.
2. The JSON must be an object with a "slides" array.
3. Each slide object must have these exact fields:
   - "title": string - A clear, concise slide title (max 80 chars)
   - "bullets": array of strings - 3-6 bullet points per slide (each max 150 chars)
   - "imagePrompt": string - A detailed prompt for generating an illustration for this slide. Describe the image content, style, and mood. Be specific about what visual would complement the content.
   - "bgColor": string - Background color as 6-char hex (e.g., "FFFFFF" for white, "1A1A2E" for dark)
   - "titleColor": string - Title text color as 6-char hex
   - "bodyColor": string - Body text color as 6-char hex
   - "animation": string - One of: "fadeIn", "flyIn", "wipe", "appear"

4. Generate 5-10 slides for a complete presentation.
5. The first slide should be a title slide with the presentation title and key themes as bullets.
6. The last slide should be a summary/conclusion or Q&A slide.
7. Choose colors that work well together - ensure text is readable against the background.
8. Vary animations across slides for visual interest.
9. Make bullet points informative and concise - they should convey real content, not filler.
10. Image prompts should describe professional, relevant visuals. Include style hints like "flat design", "photograph", "illustration", "infographic style", etc.

EXAMPLE OUTPUT FORMAT:
{
  "slides": [
    {
      "title": "Introduction to Machine Learning",
      "bullets": ["Understanding AI fundamentals", "Key algorithms and approaches", "Real-world applications"],
      "imagePrompt": "A futuristic digital brain with neural network connections glowing in blue, flat design illustration style, dark background",
      "bgColor": "1A1A2E",
      "titleColor": "E94560",
      "bodyColor": "EAEAEA",
      "animation": "fadeIn"
    }
  ]
}`;

/**
 * Call Cerebras API to generate slide content
 * @param {string} apiKey - Cerebras API key
 * @param {string} topic - User's presentation topic
 * @param {function} onLog - Callback for logging progress
 * @returns {Promise<object>} Parsed slide data
 */
async function generateSlideContent(apiKey, topic, onLog = () => {}) {
  onLog('🧠 Calling Cerebras API to generate slide content...');
  onLog(`📝 Topic: "${topic}"`);

  const requestBody = {
    model: 'llama-4-scout-17b-16e-instruct',
    messages: [
      { role: 'system', content: SYSTEM_PROMPT },
      {
        role: 'user',
        content: `Create a professional PowerPoint presentation about: "${topic}"\n\nGenerate detailed slide content with appropriate colors, animations, and image prompts. Remember to respond with ONLY valid JSON.`,
      },
    ],
    temperature: 0.7,
    max_completion_tokens: 4096,
  };

  onLog('📡 Request payload prepared, sending to Cerebras...');
  onLog(`   Model: ${requestBody.model}`);

  const response = await fetch(CEREBRAS_API_URL, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      Authorization: `Bearer ${apiKey}`,
    },
    body: JSON.stringify(requestBody),
  });

  if (!response.ok) {
    const errorText = await response.text();
    throw new Error(`Cerebras API error (${response.status}): ${errorText}`);
  }

  const data = await response.json();
  onLog('✅ Cerebras API response received');

  const content = data.choices[0].message.content;
  onLog(`📄 Raw response length: ${content.length} chars`);
  onLog('📄 Raw response preview:');
  onLog(content.substring(0, 500) + (content.length > 500 ? '...' : ''));

  // Parse JSON from response - handle potential markdown code blocks
  let jsonStr = content.trim();
  // Remove markdown code block wrapper if present
  if (jsonStr.startsWith('```')) {
    jsonStr = jsonStr.replace(/^```(?:json)?\s*\n?/, '').replace(/\n?```\s*$/, '');
  }

  let slideData;
  try {
    slideData = JSON.parse(jsonStr);
  } catch (e) {
    onLog('⚠️ Failed to parse JSON directly, attempting to extract JSON object...');
    const match = jsonStr.match(/\{[\s\S]*\}/);
    if (match) {
      slideData = JSON.parse(match[0]);
    } else {
      throw new Error('Could not parse slide data from AI response');
    }
  }

  if (!slideData.slides || !Array.isArray(slideData.slides)) {
    throw new Error('AI response does not contain a "slides" array');
  }

  onLog(`✅ Parsed ${slideData.slides.length} slides successfully`);

  // Validate and sanitize each slide
  slideData.slides = slideData.slides.map((slide, i) => ({
    title: String(slide.title || `Slide ${i + 1}`),
    bullets: Array.isArray(slide.bullets) ? slide.bullets.map(String) : [],
    imagePrompt: String(slide.imagePrompt || ''),
    bgColor: validHex(slide.bgColor, 'FFFFFF'),
    titleColor: validHex(slide.titleColor, '000000'),
    bodyColor: validHex(slide.bodyColor, '333333'),
    animation: ['fadeIn', 'flyIn', 'wipe', 'appear'].includes(slide.animation) ? slide.animation : 'fadeIn',
  }));

  return slideData;
}

function validHex(val, fallback) {
  if (typeof val === 'string' && /^[0-9A-Fa-f]{6}$/.test(val)) {
    return val.toUpperCase();
  }
  return fallback;
}

module.exports = { generateSlideContent, SYSTEM_PROMPT };
