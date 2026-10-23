import tailwindcss from '@tailwindcss/vite';
import react from '@vitejs/plugin-react';
import fs from 'fs';
import path from 'path';
import {defineConfig, Plugin} from 'vite';

function embeddedTemplateSyncPlugin(): Plugin {
  return {
    name: 'embedded-template-sync',
    configureServer(server) {
      server.middlewares.use('/__sync_embedded_template', (req, res) => {
        if (req.method !== 'POST') {
          res.statusCode = 405;
          res.end('Method Not Allowed');
          return;
        }
        let body = '';
        req.on('data', (chunk) => {
          body += chunk;
        });
        req.on('end', () => {
          try {
            const parsed = JSON.parse(body);
            if (parsed && typeof parsed.base64 === 'string' && parsed.base64.length > 0) {
              const targetPath = path.resolve(__dirname, 'src/utils/embeddedCustomTemplate.json');
              const existingRaw = fs.existsSync(targetPath)
                ? fs.readFileSync(targetPath, 'utf-8')
                : '{}';
              let existing: any = {};
              try {
                existing = JSON.parse(existingRaw);
              } catch {
                existing = {};
              }
              const nextContent = JSON.stringify(
                {
                  base64: parsed.base64,
                  fileName: parsed.fileName || existing.fileName || 'შაბლონი.xlsx',
                  instructions: Array.isArray(parsed.instructions)
                    ? parsed.instructions
                    : existing.instructions || [],
                },
                null,
                2
              );
              if (existingRaw.trim() !== nextContent.trim()) {
                fs.writeFileSync(targetPath, nextContent, 'utf-8');
              }
            }
            res.statusCode = 200;
            res.setHeader('Content-Type', 'application/json');
            res.end(JSON.stringify({ok: true}));
          } catch (err: any) {
            res.statusCode = 400;
            res.end(JSON.stringify({ok: false, error: err?.message}));
          }
        });
      });
    },
  };
}

export default defineConfig(() => {
  return {
    base: './',
    plugins: [react(), tailwindcss(), embeddedTemplateSyncPlugin()],
    resolve: {
      alias: {
        '@': path.resolve(__dirname, '.'),
      },
    },
    server: {
      // HMR is disabled in AI Studio via DISABLE_HMR env var.
      // Do not modify—file watching is disabled to prevent flickering during agent edits.
      hmr: process.env.DISABLE_HMR !== 'true',
      // Disable file watching when DISABLE_HMR is true to save CPU during agent edits.
      watch: process.env.DISABLE_HMR === 'true' ? null : {},
    },
  };
});
