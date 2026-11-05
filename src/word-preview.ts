import { renderAsync } from 'docx-preview';

/** True when the bytes are a zip package, which is how .docx files are stored. */
export function isZipDocument(bytes: Uint8Array): boolean {
  return bytes.length > 3 && bytes[0] === 0x50 && bytes[1] === 0x4b;
}

/**
 * Renders a .docx into `body` as one `section.docx` per page.
 * Styles are written into the same element so they stay inside the shadow root.
 */
export async function renderDocxPages(bytes: Uint8Array, body: HTMLElement): Promise<number> {
  body.replaceChildren();
  const blob = new Blob([bytes.slice()], {
    type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  });
  await renderAsync(blob, body, body, {
    className: 'docx',
    inWrapper: true,
    breakPages: true,
    ignoreLastRenderedPageBreak: false,
    useBase64URL: true,
    renderHeaders: true,
    renderFooters: true,
    renderFootnotes: true,
    renderEndnotes: true,
    renderComments: false,
  });
  return body.querySelectorAll('section.docx').length;
}
