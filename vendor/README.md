# Vendored libraries

Every library CanvSyllabus uses is copied into this folder and served from the
same GitHub Pages site as the tool. Nothing loads from a CDN, so the page makes
no requests to anyone else, and the Content Security Policy in `index.html`
enforces that.

| File | Library | Version | License |
| --- | --- | --- | --- |
| `mammoth.browser.min.js` | mammoth (.docx to HTML) | 1.13.0 | BSD-2-Clause |
| `pdf.min.mjs`, `pdf.worker.min.mjs` | pdf.js (`pdfjs-dist`) | 6.2.108 | Apache-2.0 |
| `purify.min.js` | DOMPurify | 3.4.14 | MPL-2.0 or Apache-2.0 |

License texts for these and for mammoth's bundled dependencies (JSZip, pako,
underscore, @xmldom/xmldom, base64-js, lop, duck) are in `licenses/`.
dingbat-to-unicode (BSD-2-Clause, Michael Williamson) is also bundled inside
the mammoth build.

## How the mammoth build was made

mammoth publishes its browser build only to npm, so this one is bundled from
the tagged source on GitHub (`mwilliamson/mammoth.js`, tag `1.13.0`) with
esbuild 0.28, using mammoth's own `browser` field mappings:

```
esbuild entry.js --bundle --platform=browser --format=iife \
  --global-name=mammoth --minify --target=es2018 \
  --alias:option=./shims/option.js \
  --alias:xmlbuilder=./shims/xmlbuilder.js
```

Two dependencies are replaced with small shims, kept in `mammoth-build/`. `option` is a tiny some/none
type, reimplemented for the four methods mammoth's style-map parser calls.
`xmlbuilder` is only used to write a style map back into a .docx, which
CanvSyllabus never does, so the shim throws if anything ever calls it.

pdf.js and DOMPurify are their official prebuilt files, unmodified. The pdf.js
files come from its `legacy/build` folder on purpose. The modern build calls
JavaScript features so new that browsers a version or two behind fail on
them, and faculty run every browser imaginable.
