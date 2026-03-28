# ExamJsonReader

Browser-based JSON renderer and proofreading tool for structured exam-paper data.

This project is designed for editors who need to inspect, validate, and adjust JSON that represents exam content such as passages, questions, options, and tokenized text.

## Scope

Included:
- Browser UI source
- Static build pipeline
- JSON render / edit logic

## Core Features

- Three-column proofreading workflow
  - file list
  - rendered preview
  - editable JSON source
- Live sync between rendered preview and JSON source
- Token-level editing in the render panel
- Visual marking for `i = 0` tokens
- Editable `contentType` in the render panel
- Editable `sectionType` in the render panel
- Backward-compatible support for:
  - legacy section inference from `contentType`
  - new `{ sectionType, section }` structure
- Section folding for explicit sections
- Section splitting and section deletion from the render panel
- Export options:
  - `另存为` JSON
  - `导出txt` rendered plain text

## Expected Data Shapes

The renderer supports both:

### Legacy shape

```json
{
  "data": [
    [
      { "contentType": "body", "content": [] }
    ],
    [
      { "contentType": "mcQuestion", "content": [] },
      { "contentType": "mcOption", "content": [] }
    ]
  ]
}
```

### New sectioned shape

```json
{
  "data": [
    {
      "sectionType": "阅读理解",
      "section": [
        [
          { "contentType": "body", "content": [] }
        ],
        [
          { "contentType": "mcQuestion", "content": [] },
          { "contentType": "mcOption", "content": [] }
        ]
      ]
    }
  ]
}
```

## Local Development

### Start a local server

```bash
npm run dev
```

Then open:

```text
http://localhost:5173
```

### Build deployable static files

```bash
npm run build
```

Build output:

```text
dist/
```

## Deployment

The built `dist/` folder can be deployed to a static web server on an intranet.

## Project Structure

```text
index.html          App shell
src/main.js         Rendering, editing, export logic
src/styles.css      UI styling
scripts/build.py    Static build script
dist/               Generated output (ignored in git)
```

## Notes

- This tool assumes exam-paper JSON can contain token arrays such as:

```json
[{ "w": "word", "i": 123, "bf": "base" }]
```

- Legacy files can be progressively upgraded in the UI by adding `sectionType`.
- Save-location reuse depends on browser support for the File System Access API.

## License

No license file has been added yet. Add one before public reuse if needed.
