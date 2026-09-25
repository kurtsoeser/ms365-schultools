import fs from 'node:fs';

const p = 'landing/index.html';
let s = fs.readFileSync(p, 'utf8');

const map = {
  'dashboard.png': 'dashboard.png?v=light1',
  'werkzeuge.png': 'werkzeuge.png?v=light1',
  'einrichtung.png': 'einrichtung.png?v=light1',
  'kursteams.png': 'kursteams.png?v=light1',
  'gruppen.png': 'gruppen.png?v=light1',
  'schularbeiten.png': 'schularbeiten.png?v=light1',
  'projektwochen.png': 'projektwochen.png?v=light1',
  'aufraeumen.png': 'aufraeumen.png?v=light1',
  'hilfe.png': 'hilfe.png?v=light1'
};

for (const [from, to] of Object.entries(map)) {
  s = s.split(`assets/screens/${from}`).join(`assets/screens/${to}`);
}

fs.writeFileSync(p, s);
const n = (s.match(/screens\/[a-z0-9-]+\.png\?v=light1/g) || []).length;
console.log('updated', n, 'urls');
