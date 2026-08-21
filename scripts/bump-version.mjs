#!/usr/bin/env node
// Bump the plugin version in both package.json and src/manifest.json.
//
// 版本单一来源是 package.json 的 `version`；src/manifest.json 是 SP 运行时读到的
// 版本号，必须与之同步。此脚本一次性更新两者，避免手改漏一处。发版流程沿用
// zot-tool：在 PR 里 bump 版本 → 合并到 main 后由 auto-release workflow 打 tag 并发
// GitHub Release。
//
// 用法：
//   node scripts/bump-version.mjs patch|minor|major
//   node scripts/bump-version.mjs 1.2.3
// 或：npm run bump:patch / bump:minor / bump:major

import { readFileSync, writeFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';

const ROOT = join(dirname(fileURLToPath(import.meta.url)), '..');
const PKG_PATH = join(ROOT, 'package.json');
const MANIFEST_PATH = join(ROOT, 'src', 'manifest.json');

function semverBump(current, level) {
  const m = /^(\d+)\.(\d+)\.(\d+)$/.exec(current);
  if (!m) throw new Error(`invalid semver: ${current}`);
  let [, major, minor, patch] = m.map(Number);
  if (level === 'major') {
    major += 1;
    minor = 0;
    patch = 0;
  } else if (level === 'minor') {
    minor += 1;
    patch = 0;
  } else if (level === 'patch') {
    patch += 1;
  } else {
    throw new Error(`unknown bump level: ${level}`);
  }
  return `${major}.${minor}.${patch}`;
}

const arg = process.argv[2];
if (!arg) {
  console.error('Usage: node scripts/bump-version.mjs <patch|minor|major|X.Y.Z>');
  process.exit(1);
}

const pkg = JSON.parse(readFileSync(PKG_PATH, 'utf8'));
const current = pkg.version;
const next = /^(patch|minor|major)$/.test(arg) ? semverBump(current, arg) : arg;

if (!/^\d+\.\d+\.\d+$/.test(next)) {
  console.error(`Invalid target version: ${next} (expected X.Y.Z or patch/minor/major)`);
  process.exit(1);
}

// package.json：整文件重写（本就是 2 空格缩进，语义不变）。
pkg.version = next;
writeFileSync(PKG_PATH, JSON.stringify(pkg, null, 2) + '\n', 'utf8');

// manifest.json：仅替换 version 行，保留原有排版。
const manifestText = readFileSync(MANIFEST_PATH, 'utf8');
const updatedManifest = manifestText.replace(
  /("version"\s*:\s*)"[^"]*"/,
  `$1"${next}"`,
);
if (updatedManifest === manifestText) {
  console.error('manifest.json has no "version" field — aborting without writing.');
  process.exit(1);
}
writeFileSync(MANIFEST_PATH, updatedManifest, 'utf8');

console.log(`Bumped version ${current} -> ${next}`);
console.log('Updated: package.json, src/manifest.json');
console.log('');
console.log('Next: commit in a PR and merge to main — the auto-release workflow');
console.log(`will tag v${next} and publish a GitHub Release.`);
