import fs from 'node:fs';
import path from 'node:path';

const root = process.cwd();
const configPath = path.join(root, 'ms365-config.js');

if (!fs.existsSync(configPath)) {
    console.error('check-tracked-secrets: ms365-config.js fehlt.');
    process.exit(1);
}

const content = fs.readFileSync(configPath, 'utf8');
const functionKeyMatch = content.match(/functionKey:\s*['"]([^'"]*)['"]/);

if (functionKeyMatch && functionKeyMatch[1].trim().length > 0) {
    console.error('');
    console.error('FEHLER: ms365-config.js enthält einen Azure Function Key.');
    console.error('Bitte den Key in ms365-config.local.js auslagern (Vorlage: ms365-config.local.example.js).');
    console.error('In ms365-config.js muss functionKey leer bleiben: functionKey: \'\'');
    console.error('');
    process.exit(1);
}

console.log('check-tracked-secrets: OK (keine Secrets in ms365-config.js)');
