// CanvSyllabus: a .docx or .pdf syllabus in, clean Canvas-safe HTML out.
//
// Everything here runs in the browser tab. The file is read with the
// FileReader API, converted in memory, and never sent anywhere. The page's
// Content Security Policy (connect-src 'none') means this script could not
// upload it even if it tried.
//
// Pipeline:
//   .docx -> mammoth -> rough HTML --\
//                                     +-> clean() -> one strict DOMPurify pass -> preview + clipboard
//   .pdf  -> pdf.js text -> lines ---/
//
// clean() is the single place that decides what Canvas gets. Both converters
// feed it, and its last step is a DOMPurify pass with the Canvas-safe
// whitelist, so nothing reaches the clipboard without going through that gate.

import * as pdfjsLib from "./vendor/pdf.min.mjs?v=0.1.1";

const VERSION = "0.1.1";

pdfjsLib.GlobalWorkerOptions.workerSrc = "./vendor/pdf.worker.min.mjs?v=0.1.1";

// What Canvas gets. A deliberately narrow subset of what the Rich Content
// Editor accepts: structure, emphasis, links, and tables, with no styles,
// classes, ids, or images.
const ALLOWED_TAGS = [
  "h2", "h3", "h4", "h5", "h6", "p", "br",
  "ul", "ol", "li",
  "strong", "em", "u", "s", "sub", "sup", "a",
  "table", "caption", "thead", "tbody", "tr", "th", "td",
];
// border is set only on <table>, by this script, so schedule tables have
// visible grid lines in Canvas. scope is set only on header cells.
const ALLOWED_ATTR = ["href", "colspan", "rowspan", "scope", "border"];
const SAFE_URL = /^(?:https?:|mailto:)/i;
// DOMPurify checks every allowed attribute's value against the URL rule
// unless told otherwise, which would throw away colspan="2". These carry
// numbers and keywords, never URLs.
const NOT_URLS = ["colspan", "rowspan", "scope", "border"];

const MAX_BYTES = 50 * 1024 * 1024;

// ---------- .docx ----------

// Canvas uses the page title as the H1, and its editor's heading menu starts
// at Heading 2, so every Word heading moves down one level.
const STYLE_MAP = [
  "p[style-name='Title'] => h2:fresh",
  "p[style-name='Subtitle'] => p:fresh",
  "p[style-name='Heading 1'] => h2:fresh",
  "p[style-name='Heading 2'] => h3:fresh",
  "p[style-name='Heading 3'] => h4:fresh",
  "p[style-name='Heading 4'] => h5:fresh",
  "p[style-name='Heading 5'] => h6:fresh",
  "p[style-name='Heading 6'] => h6:fresh",
  // Word's built-in list styles. Lists made with the toolbar buttons carry
  // their own level, and the deeper-level rules (mammoth's own defaults)
  // come first so those keep it. A paragraph styled "List Bullet 2" from the
  // Styles gallery is different: Word gives it a separate list at level one
  // and shows the depth through the style, which mammoth doesn't read, so
  // without the style rules those items come out flat.
  "p:unordered-list(2) => ul|ol > li > ul > li:fresh",
  "p:unordered-list(3) => ul|ol > li > ul|ol > li > ul > li:fresh",
  "p:unordered-list(4) => ul|ol > li > ul|ol > li > ul|ol > li > ul > li:fresh",
  "p:ordered-list(2) => ul|ol > li > ol > li:fresh",
  "p:ordered-list(3) => ul|ol > li > ul|ol > li > ol > li:fresh",
  "p:ordered-list(4) => ul|ol > li > ul|ol > li > ul|ol > li > ol > li:fresh",
  "p[style-name='List Bullet 2'] => ul|ol > li > ul > li:fresh",
  "p[style-name='List Bullet 3'] => ul|ol > li > ul|ol > li > ul > li:fresh",
  "p[style-name='List Number 2'] => ul|ol > li > ol > li:fresh",
  "p[style-name='List Number 3'] => ul|ol > li > ul|ol > li > ol > li:fresh",
  "p[style-name='List Bullet'] => ul > li:fresh",
  "p[style-name='List Number'] => ol > li:fresh",
  "u => u",
  "strike => s",
  "comment-reference => !",
];

async function docxToRough(buffer) {
  const result = await window.mammoth.convertToHtml(
    { arrayBuffer: buffer },
    {
      styleMap: STYLE_MAP,
      // Don't build base64 copies of images we're about to throw away.
      // clean() counts the empty <img> tags and removes them.
      convertImage: window.mammoth.images.imgElement(() => Promise.resolve({ src: "" })),
    }
  );
  return { html: result.value, notes: [] };
}

// ---------- .pdf ----------

const BULLET = /^[•●▪■◦○◆►▶‣⁃–—‐‒\-\*✓✔]\s+/;
const NUMBERED = /^\d{1,2}[.)]\s+/;
const PAGE_NUMBER = /^(page\s*)?\d{1,4}(\s*(of|\/)\s*\d{1,4})?$/i;

function escapeText(s) {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

async function pdfToRough(buffer) {
  const task = pdfjsLib.getDocument({
    data: new Uint8Array(buffer),
    isEvalSupported: false,
    disableFontFace: true,
    useSystemFonts: false,
    useWorkerFetch: false,
    // Fonts that aren't embedded produce warnings about font data this page
    // deliberately doesn't load. Text extraction doesn't need it.
    verbosity: pdfjsLib.VerbosityLevel.ERRORS,
  });
  const pdf = await task.promise;

  const pages = [];
  let charCount = 0;
  for (let n = 1; n <= pdf.numPages; n++) {
    const page = await pdf.getPage(n);
    const content = await page.getTextContent();
    // Font names (the only reliable bold signal) are available once the
    // page's operators have been parsed. If that fails, carry on without bold.
    const fontNames = {};
    try {
      await page.getOperatorList();
      for (const item of content.items) {
        if (!item.fontName || item.fontName in fontNames) continue;
        try {
          fontNames[item.fontName] = page.commonObjs.get(item.fontName)?.name || "";
        } catch {
          fontNames[item.fontName] = "";
        }
      }
    } catch {
      /* bold detection is a nicety */
    }
    const items = [];
    for (const it of content.items) {
      if (typeof it.str !== "string" || it.str === "") continue;
      const t = it.transform;
      const size = Math.hypot(t[2], t[3]) || it.height || 0;
      if (!size) continue;
      const fname = fontNames[it.fontName] || content.styles?.[it.fontName]?.fontFamily || "";
      items.push({
        str: it.str,
        x: t[4],
        y: t[5],
        w: it.width,
        size,
        bold: /bold|black|heavy|semibold|demi/i.test(fname),
      });
      charCount += it.str.trim().length;
    }
    pages.push(linesFromItems(items));
    page.cleanup();
  }
  await task.destroy();

  if (charCount < 20 * Math.max(1, pages.length)) {
    const err = new Error("scanned");
    err.userMessage =
      "This PDF doesn't contain any text I can read. It's most likely a scan, which is a picture of a document. If you have the Word file it came from, use that instead.";
    throw err;
  }

  const notes = [];
  const removed = dropRunningHeadersAndFooters(pages);
  if (removed) notes.push(`Removed ${removed} line${removed === 1 ? "" : "s"} that repeated at the top or bottom of each page, such as running headers and page numbers.`);
  return { html: blocksToHtml(pages.flat()), notes };
}

function linesFromItems(items) {
  // Top of the page first, then left to right.
  items.sort((a, b) => (Math.abs(b.y - a.y) > Math.min(a.size, b.size) * 0.5 ? b.y - a.y : a.x - b.x));
  const lines = [];
  for (const it of items) {
    const last = lines[lines.length - 1];
    if (last && Math.abs(last.y - it.y) < Math.min(last.size, it.size) * 0.5) {
      last.items.push(it);
      last.size = Math.max(last.size, it.size);
    } else {
      lines.push({ y: it.y, size: it.size, items: [it] });
    }
  }
  for (const line of lines) {
    line.items.sort((a, b) => a.x - b.x);
    let text = "";
    let html = "";
    let prev = null;
    let runBold = null;
    let run = "";
    const flush = () => {
      if (!run) return;
      html += runBold ? `<strong>${escapeText(run)}</strong>` : escapeText(run);
      run = "";
    };
    for (const it of line.items) {
      let piece = it.str;
      if (prev) {
        const gap = it.x - (prev.x + prev.w);
        if (gap > it.size * 0.15 && !/\s$/.test(text) && !/^\s/.test(piece)) piece = " " + piece;
      }
      const bold = it.bold && it.str.trim() !== "";
      if (it.str.trim() !== "" && bold !== runBold) {
        // A space that opens a new run belongs outside the <strong>.
        const lead = piece.match(/^\s*/)[0];
        run += lead;
        piece = piece.slice(lead.length);
        flush();
        runBold = bold;
      }
      run += piece;
      text += piece;
      prev = it;
    }
    flush();
    line.text = text.replace(/\s+/g, " ").trim();
    line.html = html.trim();
    line.x = line.items[0].x;
    line.right = Math.max(...line.items.map((i) => i.x + i.w));
    const nonSpace = line.items.filter((i) => i.str.trim() !== "");
    line.allBold = nonSpace.length > 0 && nonSpace.every((i) => i.bold);
  }
  return lines.filter((l) => l.text !== "");
}

function dropRunningHeadersAndFooters(pages) {
  const key = (l) => l.text.toLowerCase().replace(/\d+/g, "#").replace(/\s+/g, " ");
  const edge = (lines) => [...lines.slice(0, 2), ...lines.slice(-2)];
  const counts = new Map();
  for (const lines of pages) {
    for (const k of new Set(edge(lines).map(key))) counts.set(k, (counts.get(k) || 0) + 1);
  }
  const threshold = Math.max(2, Math.ceil(pages.length * 0.5));
  let removed = 0;
  pages.forEach((lines, i) => {
    const drop = new Set(
      edge(lines).filter((l) => PAGE_NUMBER.test(l.text) || (pages.length > 1 && counts.get(key(l)) >= threshold))
    );
    if (drop.size) {
      pages[i] = lines.filter((l) => !drop.has(l));
      removed += drop.size;
    }
  });
  return removed;
}

function blocksToHtml(lines) {
  if (!lines.length) return "";
  // Body text size: the size most of the characters are set in.
  const bySize = new Map();
  for (const l of lines) {
    const s = Math.round(l.size * 2) / 2;
    bySize.set(s, (bySize.get(s) || 0) + l.text.length);
  }
  const body = [...bySize.entries()].sort((a, b) => b[1] - a[1])[0][0];
  const isHeading = (l) => l.size >= body * 1.15 && l.text.length <= 120;
  const headingSizes = [...new Set(lines.filter(isHeading).map((l) => Math.round(l.size * 2) / 2))].sort((a, b) => b - a);
  const level = (l) => Math.min(4, 2 + headingSizes.indexOf(Math.round(l.size * 2) / 2));
  const maxRight = Math.max(...lines.map((l) => l.right));

  const blocks = [];
  let cur = null;
  let prev = null;
  for (const l of lines) {
    const gap = prev ? prev.y - l.y : 0;
    const looseGap = !prev || gap < 0 || gap > prev.size * 1.75;
    const prevEndedShort = prev && prev.right < maxRight * 0.8;
    let kind;
    let html = l.html;
    if (isHeading(l)) kind = "h" + level(l);
    else if (BULLET.test(l.text)) kind = "ul";
    else if (NUMBERED.test(l.text)) kind = "ol";
    else kind = "p";

    if (kind === "ul" || kind === "ol") {
      html = stripMarker(html, kind === "ul" ? BULLET : NUMBERED);
      cur = { kind, html, x: l.x, lastSize: l.size };
      blocks.push(cur);
    } else if (
      cur &&
      !looseGap &&
      cur.kind === kind &&
      kind[0] === "h" &&
      Math.abs(cur.lastSize - l.size) < 0.5
    ) {
      cur.html += " " + html;
    } else if (
      cur &&
      kind === "p" &&
      !looseGap &&
      !l.allBold &&
      !cur.boldLine &&
      (cur.kind === "p" || ((cur.kind === "ul" || cur.kind === "ol") && l.x > cur.x + 1)) &&
      Math.abs(cur.lastSize - l.size) < 0.75 &&
      !prevEndedShort
    ) {
      cur.html = joinLines(cur.html, html);
      cur.lastSize = l.size;
    } else {
      cur = { kind, html, x: l.x, lastSize: l.size, boldLine: kind === "p" && l.allBold };
      blocks.push(cur);
    }
    prev = l;
  }

  let out = "";
  let list = null;
  for (const b of blocks) {
    if (b.kind === "ul" || b.kind === "ol") {
      if (list !== b.kind) {
        if (list) out += `</${list}>`;
        out += `<${b.kind}>`;
        list = b.kind;
      }
      out += `<li>${b.html}</li>`;
      continue;
    }
    if (list) {
      out += `</${list}>`;
      list = null;
    }
    out += `<${b.kind}>${b.html}</${b.kind}>`;
  }
  if (list) out += `</${list}>`;
  return out;
}

function stripMarker(html, re) {
  // The marker is plain text at the start of the line, possibly inside a
  // leading <strong>. Remove it from whichever text comes first.
  return html.replace(/^(<strong>)?([^<]*)/, (m, tag, text) => (tag || "") + text.replace(re, ""));
}

function joinLines(a, b) {
  // Rejoin a word the PDF hyphenated across a line break.
  if (/[a-z]-$/.test(a) && /^[a-z]/.test(b)) return a.slice(0, -1) + b;
  return a + " " + b;
}

// ---------- shared cleanup ----------

const BLOCKS = new Set(["H2", "H3", "H4", "H5", "H6", "P", "UL", "OL", "LI", "TABLE", "CAPTION", "THEAD", "TBODY", "TR", "TH", "TD"]);

function clean(rough, options) {
  const notes = [];
  // First pass: drop anything dangerous or irrelevant (scripts, styles,
  // classes, ids, Word cruft), but keep a few tags this function still needs
  // to look at before deciding.
  const loose = window.DOMPurify.sanitize(rough, {
    ALLOWED_TAGS: [...ALLOWED_TAGS, "h1", "b", "i", "img", "span", "div"],
    ALLOWED_ATTR: ["href", "colspan", "rowspan"],
    ALLOWED_URI_REGEXP: SAFE_URL,
    ADD_URI_SAFE_ATTR: NOT_URLS,
    KEEP_CONTENT: true,
  });
  const doc = new DOMParser().parseFromString(`<body>${loose}</body>`, "text/html");
  const root = doc.body;

  rename(root, "h1", "h2");
  rename(root, "b", "strong");
  rename(root, "i", "em");

  const images = root.querySelectorAll("img");
  if (images.length) {
    notes.push(
      `Removed ${images.length} image${images.length === 1 ? "" : "s"}. Canvas handles images better when you upload them to Course Files and insert them with the editor's image button.`
    );
    images.forEach((img) => img.remove());
  }

  root.querySelectorAll("span, div").forEach(unwrap);

  // Links: only real web and email links survive. In-document links (Word
  // bookmarks and footnote arrows) point nowhere once they're in Canvas.
  root.querySelectorAll("a").forEach((a) => {
    const href = a.getAttribute("href") || "";
    if (!SAFE_URL.test(href)) {
      if (a.textContent.trim() === "↑") a.remove();
      else unwrap(a);
    }
  });

  // Word wraps every table cell and list item in its own paragraph.
  root.querySelectorAll("td, th, li").forEach((cell) => {
    const ps = [...cell.children].filter((c) => c.tagName === "P");
    if (ps.length === 1 && cell.children.length === 1) unwrap(ps[0]);
  });

  removeEmpty(root);

  manualLists(root);

  if (options.fakeHeadings) {
    const promoted = promoteBoldLines(root);
    if (promoted) {
      notes.push(
        `Turned ${promoted} short bold line${promoted === 1 ? "" : "s"} into ${promoted === 1 ? "a heading" : "headings"}. If ${promoted === 1 ? "that was" : "any of those were"} wrong, untick "Treat short bold lines as headings."`
      );
    }
  }

  // Headings are already bold. Bold inside a heading is noise.
  root.querySelectorAll("h2, h3, h4, h5, h6").forEach((h) => h.querySelectorAll("strong").forEach(unwrap));

  tidyTables(root);
  autolink(root);
  removeEmpty(root);

  const strict = window.DOMPurify.sanitize(root.innerHTML, {
    ALLOWED_TAGS,
    ALLOWED_ATTR,
    ALLOWED_URI_REGEXP: SAFE_URL,
    ADD_URI_SAFE_ATTR: NOT_URLS,
    KEEP_CONTENT: true,
  });
  return { html: format(strict), notes };
}

function rename(root, from, to) {
  root.querySelectorAll(from).forEach((el) => {
    const n = el.ownerDocument.createElement(to);
    for (const a of el.attributes) n.setAttribute(a.name, a.value);
    while (el.firstChild) n.appendChild(el.firstChild);
    el.replaceWith(n);
  });
}

function unwrap(el) {
  el.replaceWith(...el.childNodes);
}

function isBlank(el) {
  return el.textContent.replace(/[\s ​]/g, "") === "" && !el.querySelector("td, th, li");
}

function removeEmpty(root) {
  // Inline wrappers with nothing in them, then empty blocks.
  root.querySelectorAll("strong, em, u, s, sub, sup, a").forEach((el) => {
    if (isBlank(el)) unwrap(el);
  });
  root.querySelectorAll("p, h2, h3, h4, h5, h6, li, ul, ol").forEach((el) => {
    if (isBlank(el)) el.remove();
  });
  // Stray line breaks at the edges of a paragraph, and runs of them inside.
  root.querySelectorAll("p, li, td, th").forEach((el) => {
    while (el.firstChild && (el.firstChild.nodeName === "BR" || (el.firstChild.nodeType === 3 && !el.firstChild.textContent.trim()))) el.firstChild.remove();
    while (el.lastChild && (el.lastChild.nodeName === "BR" || (el.lastChild.nodeType === 3 && !el.lastChild.textContent.trim()))) el.lastChild.remove();
    el.querySelectorAll("br + br").forEach((br) => br.remove());
  });
}

// Paragraphs typed with a bullet character or "1." instead of Word's list
// feature become real lists.
function manualLists(root) {
  const kinds = [
    { re: BULLET, tag: "ul", min: 1 },
    { re: NUMBERED, tag: "ol", min: 2 },
  ];
  for (const { re, tag, min } of kinds) {
    const ps = [...root.querySelectorAll("p")].filter((p) => p.parentNode === root && re.test(p.textContent.trim()));
    const used = new Set();
    for (const p of ps) {
      if (used.has(p)) continue;
      const run = [p];
      let next = p.nextElementSibling;
      while (next && next.tagName === "P" && re.test(next.textContent.trim())) {
        run.push(next);
        next = next.nextElementSibling;
      }
      run.forEach((x) => used.add(x));
      if (run.length < min) continue;
      const list = root.ownerDocument.createElement(tag);
      p.before(list);
      for (const item of run) {
        stripLeadingText(item, re);
        const li = root.ownerDocument.createElement("li");
        while (item.firstChild) li.appendChild(item.firstChild);
        list.appendChild(li);
        item.remove();
      }
    }
  }
}

function stripLeadingText(el, re) {
  const walker = el.ownerDocument.createTreeWalker(el, NodeFilter.SHOW_TEXT);
  let node;
  while ((node = walker.nextNode())) {
    if (!node.textContent.trim()) continue;
    node.textContent = node.textContent.replace(/^\s+/, "").replace(re, "");
    return;
  }
}

// A paragraph that is short, bold all the way through, and not a sentence is
// almost always a heading someone made by hand.
function promoteBoldLines(root) {
  const level = root.querySelector("h2") ? "h3" : "h2";
  let count = 0;
  for (const p of root.querySelectorAll("p")) {
    if (p.parentNode !== root) continue;
    const text = p.textContent.replace(/\s+/g, " ").trim();
    if (!text || text.length > 80) continue;
    if (/[.!?]$/.test(text) && text.split(" ").length > 6) continue;
    if (p.querySelector("a")) continue;
    if (!allTextBold(p)) continue;
    const h = root.ownerDocument.createElement(level);
    while (p.firstChild) h.appendChild(p.firstChild);
    p.replaceWith(h);
    count++;
  }
  return count;
}

function allTextBold(el) {
  const walker = el.ownerDocument.createTreeWalker(el, NodeFilter.SHOW_TEXT);
  let node;
  let any = false;
  while ((node = walker.nextNode())) {
    if (!node.textContent.trim()) continue;
    any = true;
    if (!node.parentElement.closest("strong")) return false;
  }
  return any;
}

function tidyTables(root) {
  for (const table of root.querySelectorAll("table")) {
    table.setAttribute("border", "1");
    const doc = root.ownerDocument;
    // Make sure every row lives in a thead or tbody.
    let tbody = table.querySelector(":scope > tbody");
    for (const tr of [...table.querySelectorAll(":scope > tr")]) {
      if (!tbody) {
        tbody = doc.createElement("tbody");
        table.appendChild(tbody);
      }
      tbody.appendChild(tr);
    }
    // If there's no header row but the first row is bold all the way across,
    // it's a header row. Canvas's accessibility checker wants those marked.
    let thead = table.querySelector(":scope > thead");
    if (!thead && tbody && tbody.rows.length > 1) {
      const first = tbody.rows[0];
      const cells = [...first.cells];
      if (cells.length && cells.every((c) => c.textContent.trim() === "" || allTextBold(c))) {
        thead = doc.createElement("thead");
        table.insertBefore(thead, tbody);
        thead.appendChild(first);
      }
    }
    if (thead) {
      for (const td of [...thead.querySelectorAll("td")]) {
        const th = doc.createElement("th");
        for (const a of td.attributes) th.setAttribute(a.name, a.value);
        while (td.firstChild) th.appendChild(td.firstChild);
        td.replaceWith(th);
      }
      thead.querySelectorAll("th").forEach((th) => {
        th.setAttribute("scope", "col");
        th.querySelectorAll("strong").forEach(unwrap);
      });
    }
  }
}

const LINKABLE = /\b(https?:\/\/[^\s<>"]+|www\.[^\s<>"]+\.[^\s<>"]+|[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,})/g;

// Bare web addresses and email addresses become real links.
function autolink(root) {
  const doc = root.ownerDocument;
  const walker = doc.createTreeWalker(root, NodeFilter.SHOW_TEXT);
  const targets = [];
  let node;
  while ((node = walker.nextNode())) {
    if (node.parentElement.closest("a")) continue;
    LINKABLE.lastIndex = 0;
    if (LINKABLE.test(node.textContent)) targets.push(node);
  }
  for (const node of targets) {
    const text = node.textContent;
    const frag = doc.createDocumentFragment();
    let last = 0;
    LINKABLE.lastIndex = 0;
    let m;
    while ((m = LINKABLE.exec(text))) {
      let found = m[0];
      // Sentence punctuation after an address isn't part of it.
      const trail = found.match(/[.,;:!?)\]'"]+$/);
      if (trail) found = found.slice(0, -trail[0].length);
      const start = m.index;
      frag.append(text.slice(last, start));
      const a = doc.createElement("a");
      a.href = found.includes("@") && !found.includes("/") ? "mailto:" + found : /^www\./i.test(found) ? "https://" + found : found;
      a.textContent = found;
      frag.append(a);
      last = start + found.length;
      LINKABLE.lastIndex = last;
    }
    frag.append(text.slice(last));
    node.replaceWith(frag);
  }
}

// Readable HTML for the Show HTML panel and the HTML editor: one block per
// line, nested blocks indented.
function format(html) {
  const doc = new DOMParser().parseFromString(`<body>${html}</body>`, "text/html");
  const lines = [];
  const walk = (parent, depth) => {
    let inline = "";
    const pad = "  ".repeat(depth);
    const flushInline = () => {
      const t = inline.replace(/\s+/g, " ").trim();
      if (t) lines.push(pad + t);
      inline = "";
    };
    for (const node of parent.childNodes) {
      if (node.nodeType === 1 && BLOCKS.has(node.tagName)) {
        flushInline();
        const hasBlockChild = [...node.children].some((c) => BLOCKS.has(c.tagName));
        const open = node.outerHTML.slice(0, node.outerHTML.indexOf(">") + 1);
        const close = `</${node.tagName.toLowerCase()}>`;
        if (hasBlockChild) {
          lines.push(pad + open);
          walk(node, depth + 1);
          lines.push(pad + close);
        } else {
          lines.push(pad + open + node.innerHTML.replace(/\s+/g, " ").trim() + close);
        }
      } else if (node.nodeType === 1) {
        inline += node.outerHTML;
      } else if (node.nodeType === 3) {
        inline += node.textContent.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
      }
    }
    flushInline();
  };
  walk(doc.body, 0);
  return lines.join("\n");
}

// ---------- clipboard ----------

// One copy, two flavors. The visual editor takes the text/html flavor and
// gets formatted content. The HTML editor (</>) is a plain-text box, so the
// browser hands it the text/plain flavor, which is the same HTML as code.
// The copy is always built from the cleaned string, never by selecting the
// preview, because a selection copy drags this page's styles along with it.
async function copyBoth(html) {
  if (window.ClipboardItem && navigator.clipboard && navigator.clipboard.write) {
    try {
      await navigator.clipboard.write([
        new ClipboardItem({
          "text/html": new Blob([html], { type: "text/html" }),
          "text/plain": new Blob([html], { type: "text/plain" }),
        }),
      ]);
      return true;
    } catch {
      /* fall through to the older route */
    }
  }
  return legacyCopy({ "text/html": html, "text/plain": html });
}

async function copyPlain(text) {
  if (navigator.clipboard && navigator.clipboard.writeText) {
    try {
      await navigator.clipboard.writeText(text);
      return true;
    } catch {
      /* fall through */
    }
  }
  return legacyCopy({ "text/plain": text });
}

function legacyCopy(flavors) {
  let ok = false;
  const onCopy = (e) => {
    for (const [type, data] of Object.entries(flavors)) e.clipboardData.setData(type, data);
    e.preventDefault();
    ok = true;
  };
  document.addEventListener("copy", onCopy);
  try {
    document.execCommand("copy");
  } catch {
    ok = false;
  } finally {
    document.removeEventListener("copy", onCopy);
  }
  return ok;
}

// ---------- preview ----------

// Neutral styling close to how Canvas shows a syllabus, so the preview shows
// what lands in Canvas rather than what looks nice on this page.
const PREVIEW_CSS = `
  html { background: #fff; color: #2d3b45; }
  body { margin: 0; padding: 18px 22px 28px; font: 16px/1.5 "Lato", "Helvetica Neue", Helvetica, Arial, sans-serif; }
  h2, h3, h4, h5, h6 { line-height: 1.25; margin: 1.2em 0 0.45em; font-weight: 700; }
  h2 { font-size: 1.6em; } h3 { font-size: 1.3em; } h4 { font-size: 1.12em; } h5, h6 { font-size: 1em; }
  body > :first-child { margin-top: 0; }
  p { margin: 0 0 0.75em; }
  ul, ol { margin: 0 0 0.75em; padding-left: 1.6em; }
  a { color: #0374b5; pointer-events: none; }
  table { border-collapse: collapse; margin: 0 0 1em; }
  th, td { border: 1px solid #c7cdd1; padding: 4px 8px; text-align: left; vertical-align: top; }
`;

function showPreview(frame, html) {
  frame.srcdoc = `<!doctype html><html lang="en"><head><meta charset="utf-8"><style>${PREVIEW_CSS}</style></head><body>${html}</body></html>`;
  frame.onload = () => {
    try {
      const h = frame.contentDocument.documentElement.scrollHeight;
      frame.style.height = Math.min(Math.max(h + 4, 200), 900) + "px";
    } catch {
      /* sandboxed without same-origin: keep the default height */
    }
  };
}

// ---------- page wiring ----------

const $ = (id) => document.getElementById(id);

const state = { rough: null, notes: [], kind: null, html: "" };

function setStatus(msg, isError) {
  const el = $("cs-status");
  el.textContent = msg;
  el.classList.toggle("cs-status--error", !!isError);
}

function render() {
  if (!state.rough) return;
  const { html, notes } = clean(state.rough, { fakeHeadings: $("cs-fake-headings").checked });
  state.html = html;
  const all = [...state.notes, ...notes];
  if (state.kind === "pdf") {
    all.unshift(
      "PDFs store where words sit on a page, not what they are, so this conversion is rougher than a Word file. Check the headings and lists in the preview. Tables come through as lines of text."
    );
  }
  const list = $("cs-notes");
  list.replaceChildren(...all.map((n) => Object.assign(document.createElement("li"), { textContent: n })));
  list.hidden = all.length === 0;
  $("cs-html").value = html;
  showPreview($("cs-preview"), html);
  $("cs-result").hidden = false;
  $("cs-copy-status").textContent = "";
}

async function handleFile(file) {
  $("cs-result").hidden = true;
  state.rough = null;
  if (!file) return;
  const name = file.name.toLowerCase();
  if (file.size > MAX_BYTES) return setStatus("That file is over 50 MB, which is a lot for a syllabus. Try saving a copy without the images.", true);
  if (name.endsWith(".doc")) return setStatus("That's the older Word format (.doc). Open it in Word, choose Save As, pick Word Document (.docx), and try that file.", true);
  if (name.endsWith(".pages")) return setStatus("That's a Pages file. In Pages, choose File, Export To, Word, and try the .docx it makes.", true);
  const kind = name.endsWith(".docx") ? "docx" : name.endsWith(".pdf") ? "pdf" : null;
  if (!kind) return setStatus("I can read .docx and .pdf files. If your syllabus is a Google Doc, choose File, Download, Microsoft Word (.docx) first.", true);

  setStatus(`Reading ${file.name}…`);
  try {
    const buffer = await file.arrayBuffer();
    const head = new Uint8Array(buffer.slice(0, 5));
    const magic = String.fromCharCode(...head);
    if (kind === "docx" && !magic.startsWith("PK")) throw Object.assign(new Error("bad docx"), { userMessage: "That file is named .docx but isn't a Word document inside. Try opening it in Word and saving it again." });
    if (kind === "pdf" && magic !== "%PDF-") throw Object.assign(new Error("bad pdf"), { userMessage: "That file is named .pdf but isn't a PDF inside." });
    const rough = kind === "docx" ? await docxToRough(buffer) : await pdfToRough(buffer);
    state.rough = rough.html;
    state.notes = rough.notes;
    state.kind = kind;
    render();
    setStatus(`Converted ${file.name}. Check the preview, then copy.`);
  } catch (err) {
    console.error(err);
    if (err && err.name === "PasswordException") return setStatus("That PDF is password-protected. Save an unprotected copy and try again.", true);
    setStatus(err.userMessage || "Something went wrong reading that file. If it opens fine on your computer, please open an issue on GitHub and say which program made it.", true);
  }
}

function init() {
  const input = $("cs-file");
  const drop = $("cs-drop");
  input.addEventListener("change", () => handleFile(input.files[0]));
  ["dragenter", "dragover"].forEach((t) =>
    drop.addEventListener(t, (e) => {
      e.preventDefault();
      drop.classList.add("cs-drop--over");
    })
  );
  ["dragleave", "drop"].forEach((t) => drop.addEventListener(t, () => drop.classList.remove("cs-drop--over")));
  drop.addEventListener("drop", (e) => {
    e.preventDefault();
    const file = e.dataTransfer.files[0];
    if (file) handleFile(file);
  });
  // A file dropped anywhere else on the page shouldn't navigate away from it.
  window.addEventListener("dragover", (e) => e.preventDefault());
  window.addEventListener("drop", (e) => e.preventDefault());

  $("cs-fake-headings").addEventListener("change", render);

  $("cs-copy").addEventListener("click", async () => {
    const ok = await copyBoth(state.html);
    $("cs-copy-status").textContent = ok
      ? "Copied. In Canvas, click into the syllabus editor and paste."
      : "Your browser blocked the copy. Open Show the HTML below and copy it from there.";
  });
  $("cs-copy-code").addEventListener("click", async () => {
    const ok = await copyPlain(state.html);
    $("cs-copy-code-status").textContent = ok ? "Copied the code." : "Your browser blocked the copy. Click in the box, select all, and copy.";
  });

  document.querySelectorAll(".tc-version").forEach((el) => (el.textContent = VERSION));
  console.info(`CanvSyllabus ${VERSION}`);
}

init();
