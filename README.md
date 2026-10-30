# pdf-editor

Framework-agnostic PDF editor as a single custom element. Angular, React, Vue, and plain HTML all use the same tag.

```html
<pdf-editor src="report.pdf"></pdf-editor>
```

This step is the shell and the viewer: open a PDF, zoom, pan, and move between pages. Annotation, shapes, editing, forms, and saving come in later steps. See [ROADMAP.md](ROADMAP.md).

## Try it

```bash
npm install
npm run dev
```

Open `http://localhost:5173/demo/index.html`. The demo loads a two-page sample. Use the menu to open another PDF, or drop a file on the page.

## Use the tag

Install the package, then import it once so the element is registered.

```ts
import 'pdf-editor';
```

### Attributes, methods, and events

| Name | Kind | Description |
| --- | --- | --- |
| `src` | attribute / property | PDF URL. Relative URLs resolve against the host page. |
| `load(source)` | method | Open a URL string, `File`, `ArrayBuffer`, or `Uint8Array`. |
| `page` | read-only | Current page, starting at 1. |
| `pageCount` | read-only | Pages in the open document, or `0`. |
| `pdf-load` | event | Fired after a document opens. `detail` is `{ pages }`. |
| `pdf-error` | event | Fired when a file cannot be opened or drawn. `detail` is `{ message }`. |

Events bubble and are composed, so listeners on the tag see them from outside the shadow root.

Page previous and next are in the menu, and the Left and Right arrow keys do the same when focus is inside the editor and not in a text field.

### Angular

```ts
import { CUSTOM_ELEMENTS_SCHEMA, Component } from '@angular/core';
import 'pdf-editor';

@Component({
  selector: 'app-viewer',
  schemas: [CUSTOM_ELEMENTS_SCHEMA],
  template: `<pdf-editor src="assets/report.pdf"></pdf-editor>`,
})
export class ViewerComponent {}
```

Give the element a height. It fills its parent.

### React

```tsx
import { useEffect, useRef } from 'react';
import 'pdf-editor';
import type { PdfEditor, PdfLoadDetail } from 'pdf-editor';

export function Viewer() {
  const ref = useRef<PdfEditor>(null);

  useEffect(() => {
    const element = ref.current;
    if (!element) return;
    const onLoad = (event: Event) => {
      const detail = (event as CustomEvent<PdfLoadDetail>).detail;
      console.log(detail.pages);
    };
    element.addEventListener('pdf-load', onLoad);
    return () => element.removeEventListener('pdf-load', onLoad);
  }, []);

  return <pdf-editor ref={ref} src="report.pdf" style={{ display: 'block', height: '100%' }} />;
}
```

React does not forward custom events through `onPdfLoad`. Listen on the element. For TypeScript, extend JSX:

```ts
import type { PdfEditor } from 'pdf-editor';

declare module 'react' {
  namespace JSX {
    interface IntrinsicElements {
      'pdf-editor': React.DetailedHTMLProps<React.HTMLAttributes<PdfEditor>, PdfEditor> & {
        src?: string;
      };
    }
  }
}
```

### Plain HTML

```html
<script type="module">
  import 'pdf-editor';
</script>
<pdf-editor src="report.pdf" style="display:block;height:100vh"></pdf-editor>
```

The package build copies the PDF.js worker, fonts, and wasm into `dist/pdfjs`, next to `pdf-editor.js`. Serve that folder with the script so the viewer can load them.

## Build

```bash
npm run build
```

Output is `dist/pdf-editor.js`, its type declarations, and `dist/pdfjs/`.
