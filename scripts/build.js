#!/usr/bin/env node

const fs = require('fs');
const path = require('path');
const AdmZip = require('adm-zip');
const { build } = require('esbuild');

const ROOT_DIR = path.join(__dirname, '..');
const SRC_DIR = path.join(ROOT_DIR, 'src');
const DIST_DIR = path.join(ROOT_DIR, 'dist');

// 单一版本来源：package.json（与 src/manifest.json 由 bump 脚本保持一致）。注入到
// plugin.js，供独立视图展示/打印构建版本，方便在移动端确认是否跑的是最新构建。
const PLUGIN_VERSION = JSON.parse(
  fs.readFileSync(path.join(ROOT_DIR, 'package.json'), 'utf8'),
).version;

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
      'process.env.PLUGIN_VERSION': JSON.stringify(PLUGIN_VERSION),
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

  const indexSrc = path.join(SRC_DIR, 'index.html');
  if (fs.existsSync(indexSrc)) {
    fs.copyFileSync(indexSrc, path.join(DIST_DIR, 'index.html'));
    console.log('Copied index.html');
  }

  // SP 的「从本地加载」只接受 .zip，且 manifest.json / plugin.js / icon.svg / index.html
  // 必须在压缩包根目录（不能套子文件夹）。打包成根级 zip 供直接导入。
  const zip = new AdmZip();
  for (const fileName of ['manifest.json', 'plugin.js', 'icon.svg', 'index.html']) {
    const filePath = path.join(DIST_DIR, fileName);
    if (fs.existsSync(filePath)) {
      zip.addLocalFile(filePath);
    }
  }
  const zipPath = path.join(DIST_DIR, 'ms-todo-provider.zip');
  zip.writeZip(zipPath);
  console.log('Wrote ' + path.basename(zipPath));

  // 额外产出一个带版本号文件名的 zip，方便经坚果云等同步盘手工传到移动端：
  // 固定文件名在同步盘/文件选择器里会被当成"旧文件"命中本地缓存（旧包），
  // 而带版本号的新文件名必然触发重新下载，保证手机端拿到的一定是最新构建。
  const versionedZipPath = path.join(DIST_DIR, `ms-todo-provider-${PLUGIN_VERSION}.zip`);
  fs.copyFileSync(zipPath, versionedZipPath);
  console.log('Wrote ' + path.basename(versionedZipPath));

  console.log('Build complete!');
}

buildPlugin().catch((err) => {
  console.error('Build failed:', err);
  process.exit(1);
});
