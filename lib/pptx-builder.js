/**
 * PPTX Builder
 *
 * Assembles all XML files and images into a valid .pptx (ZIP) file.
 */

const archiver = require('archiver');
const templates = require('./pptx-templates');

/**
 * Build a PPTX file as a readable stream.
 *
 * @param {object} opts
 * @param {string} opts.title - Presentation title
 * @param {Array} opts.slides - Array of slide data objects
 * @param {Array} opts.images - Array of { buffer: Buffer, filename: string } for each slide image
 * @param {function} opts.onLog - Logging callback
 * @returns {archiver.Archiver} ZIP stream that can be piped
 */
function buildPptx(opts) {
  const { title, slides, images = [], onLog = () => {} } = opts;
  const slideCount = slides.length;

  onLog(`📦 Building PPTX with ${slideCount} slides...`);

  const archive = archiver('zip', { zlib: { level: 9 } });

  archive.on('warning', (err) => {
    onLog(`⚠️ Archiver warning: ${err.message}`);
  });

  archive.on('error', (err) => {
    throw err;
  });

  // Collect image entries for content types
  const imageEntries = images
    .filter(img => img && img.filename)
    .map(img => {
      const ext = img.filename.split('.').pop() || 'png';
      return { filename: img.filename, ext };
    });

  // [Content_Types].xml
  const ctXml = templates.contentTypesXml(slideCount, imageEntries);
  archive.append(ctXml, { name: '[Content_Types].xml' });
  onLog('  ✓ [Content_Types].xml');

  // _rels/.rels
  archive.append(templates.rootRelsXml(), { name: '_rels/.rels' });
  onLog('  ✓ _rels/.rels');

  // ppt/presentation.xml
  archive.append(templates.presentationXml(slideCount), { name: 'ppt/presentation.xml' });
  onLog('  ✓ ppt/presentation.xml');

  // ppt/_rels/presentation.xml.rels
  archive.append(templates.presentationRelsXml(slideCount), { name: 'ppt/_rels/presentation.xml.rels' });
  onLog('  ✓ ppt/_rels/presentation.xml.rels');

  // Theme
  archive.append(templates.theme1Xml(), { name: 'ppt/theme/theme1.xml' });
  onLog('  ✓ ppt/theme/theme1.xml');

  // Slide Master
  archive.append(templates.slideMaster1Xml(), { name: 'ppt/slideMasters/slideMaster1.xml' });
  archive.append(templates.slideMaster1RelsXml(), { name: 'ppt/slideMasters/_rels/slideMaster1.xml.rels' });
  onLog('  ✓ ppt/slideMasters/slideMaster1.xml');

  // Slide Layout
  archive.append(templates.slideLayout1Xml(), { name: 'ppt/slideLayouts/slideLayout1.xml' });
  archive.append(templates.slideLayout1RelsXml(), { name: 'ppt/slideLayouts/_rels/slideLayout1.xml.rels' });
  onLog('  ✓ ppt/slideLayouts/slideLayout1.xml');

  // Slides
  for (let i = 0; i < slideCount; i++) {
    const slide = slides[i];
    const imageEntry = images[i];
    const hasImage = imageEntry && imageEntry.buffer && imageEntry.buffer.length > 0;
    const imageRId = hasImage ? 'rId2' : null;

    const slideContent = templates.slideXml({
      slideIndex: i,
      title: slide.title,
      bullets: slide.bullets || [],
      imageRId,
      bgColor: slide.bgColor || 'FFFFFF',
      titleColor: slide.titleColor || '000000',
      bodyColor: slide.bodyColor || '333333',
      animation: slide.animation || 'fadeIn',
    });

    archive.append(slideContent, { name: `ppt/slides/slide${i + 1}.xml` });
    onLog(`  ✓ ppt/slides/slide${i + 1}.xml - "${slide.title}"`);

    // Slide relationship
    const imageFilename = hasImage ? imageEntry.filename : null;
    archive.append(
      templates.slideRelsXml(i, imageFilename),
      { name: `ppt/slides/_rels/slide${i + 1}.xml.rels` }
    );

    // Add image to media folder
    if (hasImage) {
      archive.append(imageEntry.buffer, { name: `ppt/media/${imageEntry.filename}` });
      onLog(`  ✓ ppt/media/${imageEntry.filename}`);
    }
  }

  // docProps
  archive.append(templates.coreXml(title), { name: 'docProps/core.xml' });
  archive.append(templates.appXml(slideCount), { name: 'docProps/app.xml' });
  onLog('  ✓ docProps/core.xml & app.xml');

  onLog('📦 PPTX assembly complete, finalizing ZIP...');
  archive.finalize();

  return archive;
}

module.exports = { buildPptx };
