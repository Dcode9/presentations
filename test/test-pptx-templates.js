/**
 * Tests for PPTX template generation.
 *
 * Validates that generated XML is well-formed and contains required elements
 * for a valid .pptx file structure.
 */

const assert = require('assert');
const templates = require('../lib/pptx-templates');
const { buildPptx } = require('../lib/pptx-builder');

let passed = 0;
let failed = 0;

function test(name, fn) {
  try {
    fn();
    console.log(`  ✅ ${name}`);
    passed++;
  } catch (err) {
    console.log(`  ❌ ${name}`);
    console.log(`     ${err.message}`);
    failed++;
  }
}

console.log('\n🧪 PPTX Template Tests\n');

// --- Content Types ---
console.log('Content Types:');

test('contentTypesXml includes slide overrides', () => {
  const xml = templates.contentTypesXml(3, [{ filename: 'img.png', ext: 'png' }]);
  assert(xml.includes('slide1.xml'), 'should include slide1');
  assert(xml.includes('slide2.xml'), 'should include slide2');
  assert(xml.includes('slide3.xml'), 'should include slide3');
  assert(xml.includes('Extension="png"'), 'should include png extension');
  assert(xml.includes('ContentType="image/png"'), 'should include image/png');
});

test('contentTypesXml handles jpeg images', () => {
  const xml = templates.contentTypesXml(1, [{ filename: 'img.jpg', ext: 'jpg' }]);
  assert(xml.includes('Extension="jpg"'), 'should include jpg extension');
  assert(xml.includes('image/jpeg'), 'should include image/jpeg');
});

test('contentTypesXml handles no images', () => {
  const xml = templates.contentTypesXml(2, []);
  assert(xml.includes('slide1.xml'), 'should include slide1');
  assert(xml.includes('slide2.xml'), 'should include slide2');
  assert(!xml.includes('Extension="png"'), 'should not include png when no images');
});

// --- Root Relationships ---
console.log('\nRoot Relationships:');

test('rootRelsXml has presentation relationship', () => {
  const xml = templates.rootRelsXml();
  assert(xml.includes('officeDocument'), 'should reference officeDocument');
  assert(xml.includes('ppt/presentation.xml'), 'should target presentation.xml');
});

// --- Presentation ---
console.log('\nPresentation:');

test('presentationXml includes correct number of slide IDs', () => {
  const xml = templates.presentationXml(5);
  for (let i = 0; i < 5; i++) {
    assert(xml.includes(`id="${256 + i}"`), `should include slide id ${256 + i}`);
    assert(xml.includes(`r:id="rId${i + 2}"`), `should include rId${i + 2}`);
  }
});

test('presentationXml includes slide dimensions', () => {
  const xml = templates.presentationXml(1);
  assert(xml.includes('cx="12192000"'), 'should include widescreen width');
  assert(xml.includes('cy="6858000"'), 'should include widescreen height');
});

// --- Presentation Relationships ---
console.log('\nPresentation Relationships:');

test('presentationRelsXml includes slide relationships', () => {
  const xml = templates.presentationRelsXml(3);
  assert(xml.includes('slides/slide1.xml'), 'should reference slide1');
  assert(xml.includes('slides/slide2.xml'), 'should reference slide2');
  assert(xml.includes('slides/slide3.xml'), 'should reference slide3');
  assert(xml.includes('slideMaster'), 'should reference slideMaster');
  assert(xml.includes('theme/theme1.xml'), 'should reference theme');
});

// --- Theme ---
console.log('\nTheme:');

test('theme1Xml has valid structure', () => {
  const xml = templates.theme1Xml();
  assert(xml.includes('a:theme'), 'should have theme element');
  assert(xml.includes('a:clrScheme'), 'should have color scheme');
  assert(xml.includes('a:fontScheme'), 'should have font scheme');
  assert(xml.includes('Calibri'), 'should reference Calibri font');
});

// --- Slide Master ---
console.log('\nSlide Master:');

test('slideMaster1Xml has valid structure', () => {
  const xml = templates.slideMaster1Xml();
  assert(xml.includes('p:sldMaster'), 'should have slideMaster element');
  assert(xml.includes('p:clrMap'), 'should have color mapping');
  assert(xml.includes('p:sldLayoutIdLst'), 'should reference slide layouts');
});

// --- Slide Layout ---
console.log('\nSlide Layout:');

test('slideLayout1Xml is blank layout', () => {
  const xml = templates.slideLayout1Xml();
  assert(xml.includes('type="blank"'), 'should be blank type');
  assert(xml.includes('p:sldLayout'), 'should have slideLayout element');
});

// --- Slide XML ---
console.log('\nSlide XML:');

test('slideXml generates title and bullets', () => {
  const xml = templates.slideXml({
    slideIndex: 0,
    title: 'Test Title',
    bullets: ['Point 1', 'Point 2', 'Point 3'],
    bgColor: 'FFFFFF',
    titleColor: '000000',
    bodyColor: '333333',
    animation: 'fadeIn',
  });
  assert(xml.includes('Test Title'), 'should include title');
  assert(xml.includes('Point 1'), 'should include bullet 1');
  assert(xml.includes('Point 2'), 'should include bullet 2');
  assert(xml.includes('Point 3'), 'should include bullet 3');
  assert(xml.includes('val="FFFFFF"'), 'should include bg color');
  assert(xml.includes('p:timing'), 'should include animations');
  assert(xml.includes('p:transition'), 'should include transitions');
});

test('slideXml includes image when imageRId provided', () => {
  const xml = templates.slideXml({
    slideIndex: 0,
    title: 'Image Slide',
    bullets: ['Content'],
    imageRId: 'rId2',
    animation: 'fadeIn',
  });
  assert(xml.includes('p:pic'), 'should include picture element');
  assert(xml.includes('r:embed="rId2"'), 'should reference image rId');
});

test('slideXml does not include image when imageRId is null', () => {
  const xml = templates.slideXml({
    slideIndex: 0,
    title: 'No Image',
    bullets: ['Content'],
    imageRId: null,
    animation: 'fadeIn',
  });
  assert(!xml.includes('p:pic'), 'should not include picture element');
});

test('slideXml supports different animation types', () => {
  for (const anim of ['fadeIn', 'flyIn', 'wipe', 'appear']) {
    const xml = templates.slideXml({
      slideIndex: 0,
      title: 'Anim Test',
      bullets: ['Test'],
      animation: anim,
    });
    assert(xml.includes('p:timing'), `${anim}: should include timing`);
    assert(xml.includes('presetClass="entr"'), `${anim}: should be entrance animation`);
  }
});

// --- Slide Relationships ---
console.log('\nSlide Relationships:');

test('slideRelsXml references slideLayout', () => {
  const xml = templates.slideRelsXml(0, null);
  assert(xml.includes('slideLayout'), 'should reference slideLayout');
});

test('slideRelsXml includes image relationship when provided', () => {
  const xml = templates.slideRelsXml(0, 'image1.png');
  assert(xml.includes('image1.png'), 'should reference image file');
  assert(xml.includes('relationships/image'), 'should have image type');
});

// --- XML Escaping ---
console.log('\nXML Escaping:');

test('escapeXml handles special characters', () => {
  assert.strictEqual(templates.escapeXml('a & b'), 'a &amp; b');
  assert.strictEqual(templates.escapeXml('<tag>'), '&lt;tag&gt;');
  assert.strictEqual(templates.escapeXml('"quoted"'), '&quot;quoted&quot;');
  assert.strictEqual(templates.escapeXml("it's"), 'it&apos;s');
});

test('escapeXml handles empty/null input', () => {
  assert.strictEqual(templates.escapeXml(''), '');
  assert.strictEqual(templates.escapeXml(null), '');
  assert.strictEqual(templates.escapeXml(undefined), '');
});

// --- DocProps ---
console.log('\nDocProps:');

test('coreXml includes title', () => {
  const xml = templates.coreXml('My Presentation');
  assert(xml.includes('My Presentation'), 'should include title');
  assert(xml.includes('dc:title'), 'should have dc:title element');
});

test('appXml includes slide count', () => {
  const xml = templates.appXml(5);
  assert(xml.includes('<Slides>5</Slides>'), 'should include slide count');
});

// --- PPTX Builder ---
console.log('\nPPTX Builder:');

test('buildPptx returns an archiver stream', () => {
  const slides = [
    { title: 'Slide 1', bullets: ['Bullet A'], bgColor: 'FFFFFF', titleColor: '000000', bodyColor: '333333', animation: 'fadeIn' },
    { title: 'Slide 2', bullets: ['Bullet B'], bgColor: '1A1A2E', titleColor: 'FFFFFF', bodyColor: 'CCCCCC', animation: 'flyIn' },
  ];

  const archive = buildPptx({
    title: 'Test Presentation',
    slides,
    images: [null, null],
    onLog: () => {},
  });

  assert(archive, 'should return an archive object');
  assert(typeof archive.pipe === 'function', 'archive should be pipeable');

  // Collect and verify it produces output
  const chunks = [];
  archive.on('data', (chunk) => chunks.push(chunk));
  archive.on('end', () => {
    const buf = Buffer.concat(chunks);
    // ZIP files start with PK (0x504B)
    assert(buf[0] === 0x50 && buf[1] === 0x4B, 'output should be a valid ZIP');
    assert(buf.length > 1000, 'ZIP should have meaningful content');
  });
});

// --- Summary ---
setTimeout(() => {
  console.log(`\n${'─'.repeat(40)}`);
  console.log(`Results: ${passed} passed, ${failed} failed`);
  if (failed > 0) {
    process.exit(1);
  }
}, 1000);
