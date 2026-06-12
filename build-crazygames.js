/* Builds the CrazyGames submission variant from index.html:
   - injects the CrazyGames SDK v3 script tag (activates the Portal facade)
   - blanks SHARE_URL (portals forbid external links)
   - strips the PWA manifest + service worker (portal hosts the files)
   Output: dist/crazygames/index.html                                   */
const fs = require('fs');
const path = require('path');

const SRC = path.join(__dirname, 'index.html');
const OUT_DIR = path.join(__dirname, 'dist', 'crazygames');
const OUT = path.join(OUT_DIR, 'index.html');

let html = fs.readFileSync(SRC, 'utf8');
const checks = [];

// 1. inject the SDK before the first <style> in <head>
const sdkTag = '<script src="https://sdk.crazygames.com/crazygames-sdk-v3.js"></script>';
if (!html.includes(sdkTag)) {
  const anchor = '<link rel="manifest" href="./manifest.webmanifest">';
  if (!html.includes(anchor)) throw new Error('manifest anchor not found');
  html = html.replace(anchor, sdkTag);   // replaces manifest link (also strips PWA)
  checks.push('SDK tag injected, manifest link removed');
}

// 2. strip apple-touch-icon (portal build needs no install surface)
html = html.replace('<link rel="apple-touch-icon" href="./icon-192.png">\n', '');

// 3. blank the share URL
const shareLine = "const SHARE_URL = 'https://skimmilk12.github.io/buzzer-beater/';";
if (!html.includes(shareLine)) throw new Error('SHARE_URL line not found');
html = html.replace(shareLine, "const SHARE_URL = '';");
checks.push('SHARE_URL blanked');

// 4. remove service worker registration
const swBlock = `if('serviceWorker' in navigator && location.protocol === 'https:'){
  navigator.serviceWorker.register('./sw.js').catch(()=>{});
}`;
if (!html.includes(swBlock)) throw new Error('SW registration block not found');
html = html.replace(swBlock, '');
checks.push('service worker registration removed');

fs.mkdirSync(OUT_DIR, { recursive: true });
fs.writeFileSync(OUT, html);
console.log('built ' + OUT + ' (' + html.length + ' bytes)');
checks.forEach(c => console.log('  ✓ ' + c));
