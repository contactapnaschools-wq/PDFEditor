import '../src/index';

const editor = document.querySelector('#editor');

editor?.addEventListener('pdf-load', (event) => {
  const detail = (event as CustomEvent<{ pages: number }>).detail;
  console.log('pdf-load', detail);
});

editor?.addEventListener('pdf-error', (event) => {
  const detail = (event as CustomEvent<{ message: string }>).detail;
  console.log('pdf-error', detail);
});
