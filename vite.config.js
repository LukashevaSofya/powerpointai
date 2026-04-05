import { defineConfig } from 'vite';
import fs from 'fs';
import { homedir } from 'os';
import { join } from 'path';

/**
 * Важное изменение: Надстройки Office ВСЕГДА требуют HTTPS.
 * Если npx office-addin-dev-certs не установлен, сервер не запустится,
 * что лучше, чем запуск по HTTP, который не загрузится в PowerPoint.
 */
let httpsConfig = false;
try {
  const certDir = join(homedir(), '.office-addin-dev-certs');
  const keyPath = join(certDir, 'localhost.key');
  const certPath = join(certDir, 'localhost.crt');
  
  if (fs.existsSync(keyPath) && fs.existsSync(certPath)) {
    httpsConfig = {
      key: fs.readFileSync(keyPath),
      cert: fs.readFileSync(certPath)
    };
    console.log('✅ Office dev certs loaded (HTTPS)');
  } else {
    console.error('❌ SSL certs not found! Run: npx office-addin-dev-certs install');
  }
} catch (e) {
  console.error('Error loading certs:', e);
}

export default defineConfig({
  root: '.',
  publicDir: 'public',
  server: {
    port: 3001,
    https: httpsConfig,
    // Включаем CORS для WebView2
    headers: {
      'Access-Control-Allow-Origin': '*'
    }
  },
  build: {
    outDir: 'dist',
    assetsDir: 'assets',
    sourcemap: true,
    rollupOptions: {
      input: {
        main: './index.html'
      }
    }
  }
});