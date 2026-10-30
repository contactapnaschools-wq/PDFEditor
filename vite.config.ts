import { cpSync, createReadStream, existsSync, mkdirSync, statSync } from 'node:fs';
import { extname, join, resolve, sep } from 'node:path';
import { fileURLToPath } from 'node:url';
import { defineConfig, type Plugin } from 'vite';

const rootDir = fileURLToPath(new URL('.', import.meta.url));

function contentType(file: string): string {
  switch (extname(file)) {
    case '.mjs':
    case '.js':
      return 'text/javascript';
    case '.wasm':
      return 'application/wasm';
    case '.ttf':
      return 'font/ttf';
    default:
      return 'application/octet-stream';
  }
}

/** Serves PDF.js worker, fonts, cmaps, and wasm in dev, and copies them beside the library build. */
function pdfjsAssets(): Plugin {
  const vendor = resolve(rootDir, 'node_modules/pdfjs-dist');
  const vendorPrefix = vendor.endsWith(sep) ? vendor : vendor + sep;

  const safeFile = (relativePath: string): string | null => {
    const resolved = resolve(vendor, relativePath);
    if (resolved !== vendor && !resolved.startsWith(vendorPrefix)) return null;
    if (!existsSync(resolved) || !statSync(resolved).isFile()) return null;
    return resolved;
  };

  const copyAll = (destRoot: string) => {
    mkdirSync(destRoot, { recursive: true });
    cpSync(join(vendor, 'build/pdf.worker.min.mjs'), join(destRoot, 'pdf.worker.min.mjs'));
    cpSync(join(vendor, 'cmaps'), join(destRoot, 'cmaps'), { recursive: true });
    cpSync(join(vendor, 'standard_fonts'), join(destRoot, 'standard_fonts'), { recursive: true });
    cpSync(join(vendor, 'wasm'), join(destRoot, 'wasm'), { recursive: true });
  };

  return {
    name: 'pdfjs-assets',
    configureServer(server) {
      server.middlewares.use((req, res, next) => {
        const raw = req.url?.split('?')[0] ?? '';
        if (!raw.startsWith('/pdfjs/')) {
          next();
          return;
        }

        let relative = '';
        if (raw === '/pdfjs/pdf.worker.min.mjs') relative = 'build/pdf.worker.min.mjs';
        else if (raw.startsWith('/pdfjs/cmaps/')) relative = `cmaps/${decodeURIComponent(raw.slice('/pdfjs/cmaps/'.length))}`;
        else if (raw.startsWith('/pdfjs/standard_fonts/')) {
          relative = `standard_fonts/${decodeURIComponent(raw.slice('/pdfjs/standard_fonts/'.length))}`;
        } else if (raw.startsWith('/pdfjs/wasm/')) relative = `wasm/${decodeURIComponent(raw.slice('/pdfjs/wasm/'.length))}`;
        else {
          next();
          return;
        }

        const file = safeFile(relative);
        if (!file) {
          next();
          return;
        }
        res.setHeader('Content-Type', contentType(file));
        createReadStream(file).pipe(res);
      });
    },
    writeBundle(options) {
      const dir = options.dir ?? resolve(rootDir, 'dist');
      copyAll(resolve(dir, 'pdfjs'));
    },
  };
}

export default defineConfig({
  plugins: [pdfjsAssets()],
  server: {
    port: 5173,
    strictPort: true,
  },
  optimizeDeps: {
    exclude: ['pdfjs-dist'],
  },
  build: {
    lib: {
      entry: resolve(rootDir, 'src/index.ts'),
      formats: ['es'],
      fileName: 'pdf-editor',
    },
    sourcemap: true,
    rollupOptions: {
      output: {
        inlineDynamicImports: true,
      },
    },
  },
});
