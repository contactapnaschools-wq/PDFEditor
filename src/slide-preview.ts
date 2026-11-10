import { init } from 'pptx-preview';

const SLIDE_WIDTH = 960;

/**
 * Renders a .pptx into `body` as one `.preview-slide` per slide.
 * The caller shows a single slide at a time.
 */
export async function renderSlidePages(bytes: Uint8Array, body: HTMLElement): Promise<number> {
  body.replaceChildren();
  const mount = document.createElement('div');
  body.append(mount);
  const previewer = init(mount, { width: SLIDE_WIDTH, mode: 'list' });
  const copy = bytes.slice();
  await previewer.preview(copy.buffer);
  const slides = body.querySelectorAll('.pptx-preview-slide-wrapper');
  slides.forEach((slide) => slide.classList.add('preview-slide'));
  return slides.length;
}
