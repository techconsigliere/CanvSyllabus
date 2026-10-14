# CanvSyllabus

A Canvas syllabus converter for teachers. Drop in a Word (.docx) or PDF
syllabus, check the preview, and copy clean, Canvas-safe HTML straight into
your course's Syllabus page.

**Use it:** https://techconsigliere.github.io/CanvSyllabus/

Runs entirely in your browser. The file is never uploaded, there's no account,
and the page's Content Security Policy blocks it from making network requests
to anyone. Readable source in `canvsyllabus.js`; vendored libraries and their
licenses in `vendor/`.

## Building

The version lives in `VERSION`. After changing it, or `design/tc.css`, run:

```
python3 build.py            # stamp the version, inline the design system
python3 build.py --check    # before you push
```

`build.py` writes the version into the page, the script, and the `?v=` on
every script URL, so a new version always means fresh JavaScript after a hard
refresh.

License: MIT.
