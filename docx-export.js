// ============================================================
// Shared .docx export utilities.
// Used by: chapter-picker.js (multi-chapter "Select Chapter
// Download" popup on the module pages). Loaded via:
//   <script src="https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js"></script>
//   <script src="docx-export.js"></script>
// (JSZip must load before this file.)
// ============================================================

function docxEscapeXml(str) {
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
function docxNodeToRuns(node, fmt) {
  const runs = [];

  function pushTextRun(text, f) {
    if (text === '') return;
    const rprParts = [];
    if (f.bold) rprParts.push('<w:b/>');
    if (f.italic) rprParts.push('<w:i/>');
    if (f.underline) rprParts.push('<w:u w:val="single"/>');
    if (f.highlight) rprParts.push(`<w:highlight w:val="${DOCX_HIGHLIGHT_NAME}"/>`);
    const rpr = rprParts.length ? `<w:rPr>${rprParts.join('')}</w:rPr>` : '';
    runs.push(`<w:r>${rpr}<w:t xml:space="preserve">${docxEscapeXml(text)}</w:t></w:r>`);
  }

  function walk(n, f) {
    if (n.nodeType === 3) {
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

function docxParagraphXml(runsXml, { align = 'both' } = {}) {
  const jc = align ? `<w:jc w:val="${align}"/>` : '';
  return `<w:p><w:pPr>${jc}</w:pPr>${runsXml.join('')}</w:p>`;
}

// Bold, larger-size heading paragraph — used for "CHAPTER N".
function docxHeadingParagraphXml(text, { size = 28 } = {}) {
  return `<w:p><w:pPr><w:jc w:val="left"/></w:pPr><w:r><w:rPr><w:b/><w:sz w:val="${size}"/><w:szCs w:val="${size}"/></w:rPr><w:t xml:space="preserve">${docxEscapeXml(
    text
  )}</w:t></w:r></w:p>`;
}

function docxPageBreakXml() {
  return '<w:p><w:r><w:br w:type="page"/></w:r></w:p>';
}

// Converts a chapter's HTML string (from window.CHAPTER_CONTENT, or a
// saved localStorage working copy) into an array of <w:p> paragraph
// XML strings. Handles p, h2/h3, and ul/ol > li (as manually-prefixed
// bullets/numbers — no native numbering.xml, keeps this dependency-free).
function docxHtmlStringToParagraphs(htmlString) {
  const container = document.createElement('div');
  container.innerHTML = htmlString || '';
  const paragraphs = [];

  Array.from(container.children).forEach((el) => {
    const tag = el.tagName;

    if (tag === 'H2' || tag === 'H3') {
      const runs = docxNodeToRuns(el, {});
      const boldedRuns = runs.map((r) =>
        r.includes('<w:rPr>')
          ? r.replace('<w:rPr>', '<w:rPr><w:b/><w:sz w:val="24"/><w:szCs w:val="24"/>')
          : r.replace('<w:r>', '<w:r><w:rPr><w:b/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr>')
      );
      paragraphs.push(docxParagraphXml(boldedRuns, { align: 'left' }));
      return;
    }

    if (tag === 'UL' || tag === 'OL') {
      let i = 0;
      Array.from(el.children).forEach((li) => {
        if (li.tagName !== 'LI') return;
        i += 1;
        const prefix = tag === 'OL' ? `${i}. ` : '\u2022 ';
        const runs = docxNodeToRuns(li, {});
        paragraphs.push(docxParagraphXml([`<w:r><w:t xml:space="preserve">${prefix}</w:t></w:r>`, ...runs]));
      });
      return;
    }

    if (tag === 'P') {
      const runs = docxNodeToRuns(el, {});
      paragraphs.push(docxParagraphXml(runs));
      return;
    }

    const runs = docxNodeToRuns(el, {});
    if (runs.length) paragraphs.push(docxParagraphXml(runs));
  });

  return paragraphs;
}

// Wraps finished body content XML (paragraphs already concatenated)
// into a complete, real OOXML .docx package (zipped with JSZip) with
// a two-column section — the same layout used by the single-chapter
// download in chapter.html.
async function docxBuildBlobFromBody(bodyContentXml) {
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

  const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    ${bodyContentXml}
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

// chapters: [{ num, html }, ...] — already sorted the way they should
// appear. Each chapter gets a bold "CHAPTER N" heading; a page break is
// inserted before every chapter except the first. A chapter with no
// content yet just shows its heading with nothing underneath.
async function docxBuildMultiChapterBlob(chapters) {
  let bodyParts = [];

  chapters.forEach((ch, idx) => {
    if (idx > 0) bodyParts.push(docxPageBreakXml());
    bodyParts.push(docxHeadingParagraphXml(`CHAPTER ${ch.num}`));
    bodyParts = bodyParts.concat(docxHtmlStringToParagraphs(ch.html));
  });

  return docxBuildBlobFromBody(bodyParts.join(''));
}

function docxTriggerDownload(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

window.DocxExport = {
  htmlStringToParagraphs: docxHtmlStringToParagraphs,
  headingParagraphXml: docxHeadingParagraphXml,
  buildBlobFromBody: docxBuildBlobFromBody,
  buildMultiChapterBlob: docxBuildMultiChapterBlob,
  triggerDownload: docxTriggerDownload,
};
