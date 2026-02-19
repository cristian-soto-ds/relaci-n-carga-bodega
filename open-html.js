/**
 * Abre index.html en el navegador por defecto (sin servidor).
 */
const { execSync } = require('child_process');
const path = require('path');

const htmlPath = path.join(process.cwd(), 'index.html');

if (process.platform === 'win32') {
  execSync(`start "" "${htmlPath}"`, { stdio: 'inherit' });
} else if (process.platform === 'darwin') {
  execSync(`open "${htmlPath}"`, { stdio: 'inherit' });
} else {
  execSync(`xdg-open "${htmlPath}"`, { stdio: 'inherit' });
}
