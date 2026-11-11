#!/usr/bin/env node
import { copyFileSync, existsSync, mkdirSync } from 'node:fs';
import { homedir, platform } from 'node:os';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const __dirname = dirname(fileURLToPath(import.meta.url));
const publishSrc = resolve(__dirname, '../publish.xml');

function resolveWpsJsAddonsDir() {
  const os = platform();
  if (os === 'win32') {
    const appData = process.env.APPDATA || join(homedir(), 'AppData', 'Roaming');
    return join(appData, 'kingsoft', 'wps', 'jsaddons');
  }
  if (os === 'darwin') {
    const containerDir = join(
      homedir(),
      'Library',
      'Containers',
      'com.kingsoft.wpsoffice.mac',
      'Data',
      'Library',
      'Application Support',
      'Kingsoft',
      'wps',
      'jsaddons',
    );
    return containerDir;
  }
  return join(homedir(), '.local', 'share', 'Kingsoft', 'wps', 'jsaddons');
}

const targetDir = resolveWpsJsAddonsDir();
mkdirSync(targetDir, { recursive: true });
const targetPublish = join(targetDir, 'publish.xml');

if (!existsSync(targetPublish)) {
  copyFileSync(publishSrc, targetPublish);
  console.log(`✓ Installed WPS Presentation add-in registry to: ${targetPublish}`);
} else {
  const backup = join(targetDir, 'decken.publish.xml');
  copyFileSync(publishSrc, backup);
  console.log(`ℹ Existing publish.xml detected at ${targetPublish}.`);
  console.log(`  Wrote Decken entry template to ${backup} — merge <jspluginonline name="decken-wpp-addin" .../> into publish.xml.`);
}
