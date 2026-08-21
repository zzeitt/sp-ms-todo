#!/usr/bin/env node

const fs = require('fs');
const path = require('path');
const AdmZip = require('adm-zip');
const { build } = require('esbuild');

const ROOT_DIR = path.join(__dirname, '..');
const SRC_DIR = path.join(ROOT_DIR, 'src');
const DIST_DIR = path.join(ROOT_DIR, 'dist');

async function buildPlugin() {
  console.log('Building ms-todo-provider...');

  if (fs.existsSync(DIST_DIR)) {
    fs.rmSync(DIST_DIR, { recursive: true });
  }
  fs.mkdirSync(DIST_DIR, { recursive: true });

  await build({
    entryPoints: [path.join(SRC_DIR, 'plugin.ts')],
    bundle: true,
    outfile: path.join(DIST_DIR, 'plugin.js'),
    platform: 'browser',
    target: 'es2020',
    format: 'iife',
    define: {
      'process.env.NODE_ENV': '"production"',
      'process.env.MS_TODO_CLIENT_ID': JSON.stringify(process.env.MS_TODO_CLIENT_ID || ''),
    },
    logLevel: 'info',
    minify: true,
    sourcemap: false,
  });

  fs.copyFileSync(path.join(SRC_DIR, 'manifest.json'), path.join(DIST_DIR, 'manifest.json'));

  const iconSrc = path.join(SRC_DIR, 'icon.svg');
  if (fs.existsSync(iconSrc)) {
    fs.copyFileSync(iconSrc, path.join(DIST_DIR, 'icon.svg'));
    console.log('Copied icon.svg');
  }

  // SP 的「从本地加载」只接受 .zip，且 manifest.json / plugin.js / icon.svg
  // 必须在压缩包根目录（不能套子文件夹）。打包成根级 zip 供直接导入。
  const zip = new AdmZip();
  for (const fileName of ['manifest.json', 'plugin.js', 'icon.svg']) {
    const filePath = path.join(DIST_DIR, fileName);
    if (fs.existsSync(filePath)) {
      zip.addLocalFile(filePath);
    }
  }
  const zipPath = path.join(DIST_DIR, 'ms-todo-provider.zip');
  zip.writeZip(zipPath);
  console.log('Wrote ' + path.basename(zipPath));

  console.log('Build complete!');
}

buildPlugin().catch((err) => {
  console.error('Build failed:', err);
  process.exit(1);
});
