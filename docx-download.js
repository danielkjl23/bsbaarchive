/* ============================================================
   REAL .docx GENERATOR — replaces the old downloadChapterAsWord()
   ============================================================
   WHY: The old version saved an HTML file with a ".doc" extension.
   Desktop Word has a legacy hack that opens HTML-as-.doc and
   understands "mso-columns" — that's why it worked (but only
   single-column-looking, ish) on your PC. Mobile Word does NOT
   support that hack at all, so nothing showed up.

   This version builds an actual OOXML .docx package (a real zip
   with document.xml / styles.xml / relationship files inside) using
   JSZip, entirely in the browser. Real .docx section properties
   (<w:cols w:num="2"/>) are honored by desktop AND mobile Word.

   SETUP REQUIRED (do this in chapter.html):
   1. Add this line in <head> (or right before your existing
      inline <script> block, before this code runs):

        <script src="https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js"></script>

   2. Replace your existing `function downloadChapterAsWord() { ... }`
      block (and nothing else) with everything below. The FAB entry
      `{ id: 'download', label: 'Download this chapter', onSelect: downloadChapterAsWord }`
      does not need to change — same function name, it's just async now.
   ============================================================ */

function escapeXml(str) {
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

// Word highlight colors are a fixed enum — "cyan" is the closest
// built-in match to the #00e5ff swatch used on-screen.
const DOCX_HIGHLIGHT_NAME = 'cyan';

// Walks a DOM node, tracking inherited bold/italic/underline/highlight
// state, and returns an array of <w:r> run XML strings.
function nodeToDocxRuns(node, fmt) {
  const runs = [];

  function pushTextRun(text, f) {
    if (text === '') return;
    const rprParts = [];
    if (f.bold) rprParts.push('<w:b/>');
    if (f.italic) rprParts.push('<w:i/>');
    if (f.underline) rprParts.push('<w:u w:val="single"/>');
    if (f.highlight) rprParts.push(`<w:highlight w:val="${DOCX_HIGHLIGHT_NAME}"/>`);
    const rpr = rprParts.length ? `<w:rPr>${rprParts.join('')}</w:rPr>` : '';
    runs.push(`<w:r>${rpr}<w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r>`);
  }

  function walk(n, f) {
    if (n.nodeType === 3) {
      // Text node — collapse internal newlines/tabs like a browser would.
      const text = n.textContent.replace(/\s+/g, ' ');
      pushTextRun(text, f);
      return;
    }
    if (n.nodeType !== 1) return;

    if (n.tagName === 'BR') {
      runs.push('<w:r><w:br/></w:r>');
      return;
    }

    const next = { ...f };
    const cls = n.classList || { contains: () => false };
    const tag = n.tagName;

    if (tag === 'STRONG' || tag === 'B' || cls.contains('user-bold')) next.bold = true;
    if (tag === 'EM' || tag === 'I') next.italic = true;
    if (tag === 'U' || cls.contains('user-underline')) next.underline = true;
    if (cls.contains('user-highlight')) next.highlight = true;

    n.childNodes.forEach((child) => walk(child, next));
  }

  walk(node, fmt);
  return runs;
}

function paragraphXml(runsXml, { align = 'both', styleRpr = '' } = {}) {
  const jc = align ? `<w:jc w:val="${align}"/>` : '';
  return `<w:p><w:pPr>${jc}</w:pPr>${runsXml.join('')}</w:p>`;
}

// Converts the chapter's rendered HTML (content.innerHTML) into the
// paragraph XML that goes inside <w:body>. Handles p, h2/h3, and
// ul/ol > li (as manually-prefixed bullets/numbers — no native
// numbering.xml, keeps this dependency-free and reliably portable).
function htmlToDocxBodyXml(containerEl) {
  const paragraphs = [];

  Array.from(containerEl.children).forEach((el) => {
    const tag = el.tagName;

    if (tag === 'H2' || tag === 'H3') {
      const runs = nodeToDocxRuns(el, {});
      // Force heading runs bold + larger, on top of whatever
      // highlight/underline marks the user added.
      const boldedRuns = runs.map((r) =>
        r.includes('<w:rPr>')
          ? r.replace('<w:rPr>', '<w:rPr><w:b/><w:sz w:val="24"/><w:szCs w:val="24"/>')
          : r.replace('<w:r>', '<w:r><w:rPr><w:b/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr>')
      );
      paragraphs.push(paragraphXml(boldedRuns, { align: 'left' }));
      return;
    }

    if (tag === 'UL' || tag === 'OL') {
      let i = 0;
      Array.from(el.children).forEach((li) => {
        if (li.tagName !== 'LI') return;
        i += 1;
        const prefix = tag === 'OL' ? `${i}. ` : '\u2022 ';
        const runs = nodeToDocxRuns(li, {});
        paragraphs.push(paragraphXml([`<w:r><w:t xml:space="preserve">${prefix}</w:t></w:r>`, ...runs]));
      });
      return;
    }

    if (tag === 'P') {
      const runs = nodeToDocxRuns(el, {});
      paragraphs.push(paragraphXml(runs));
      return;
    }

    // Fallback for any other block-level tag — treat as a plain paragraph.
    const runs = nodeToDocxRuns(el, {});
    if (runs.length) paragraphs.push(paragraphXml(runs));
  });

  return paragraphs.join('');
}

async function buildDocxBlob(bodyParagraphsXml, titleText) {
  const zip = new JSZip();

  zip.file(
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>`
  );

  zip.file(
    '_rels/.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`
  );

  zip.file(
    'word/_rels/document.xml.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`
  );

  zip.file(
    'word/styles.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault>
      <w:rPr>
        <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>
        <w:sz w:val="16"/>
        <w:szCs w:val="16"/>
      </w:rPr>
    </w:rPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
  </w:style>
</w:styles>`
  );

  const titleParagraph = paragraphXml(
    [`<w:r><w:rPr><w:b/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr><w:t xml:space="preserve">${escapeXml(titleText)}</w:t></w:r>`],
    { align: 'left' }
  );

  const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    ${titleParagraph}
    ${bodyParagraphsXml}
    <w:sectPr>
      <w:cols w:num="2" w:space="480"/>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>
    </w:sectPr>
  </w:body>
</w:document>`;

  zip.file('word/document.xml', documentXml);

  return zip.generateAsync({
    type: 'blob',
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  });
}

async function downloadChapterAsWord() {
  if (typeof JSZip === 'undefined') {
    alert('Download isn\u2019t available right now \u2014 please refresh the page and try again.');
    return;
  }

  const bodyXml = htmlToDocxBodyXml(content);
  const titleText = `Chapter ${num} \u2014 ${moduleTitle}`;

  let blob;
  try {
    blob = await buildDocxBlob(bodyXml, titleText);
  } catch (err) {
    alert('Sorry, something went wrong building the Word file.');
    return;
  }

  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `${moduleTitle.replace(/\s+/g, '-')}-Chapter-${num}.docx`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}
