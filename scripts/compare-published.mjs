import fs from 'node:fs';
import path from 'node:path';
import { createHash } from 'node:crypto';
import { initAuth } from '../node_modules/@google/clasp/build/src/auth/auth.js';

export function compareFiles(remoteFiles, root) {
  const normalize = (name, text) => name.endsWith('.json')
    ? JSON.stringify(JSON.parse(text)) : text.replace(/\r\n?/g, '\n').replace(/\n+$/, '');
  const local = new Map();
  function walk(dir) {
    for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
      const full = path.join(dir, entry.name);
      if (entry.isDirectory()) walk(full);
      else if (/\.(js|gs|html|json)$/.test(entry.name)) {
        const name = path.relative(root, full).split(path.sep).join('/').replace(/\.gs$/, '.js');
        local.set(name, normalize(name, fs.readFileSync(full, 'utf8')));
      }
    }
  }
  walk(root);
  const remote = new Map();
  for (const file of remoteFiles) {
    const ext = { SERVER_JS: '.js', HTML: '.html', JSON: '.json' }[file.type];
    if (!ext) throw new Error('Unsupported remote file type');
    const name = file.name + ext;
    remote.set(name, normalize(name, file.source ?? ''));
  }
  return [...new Set([...local.keys(), ...remote.keys()])].sort().map(name => ({
    name,
    status: !local.has(name) ? 'published only' : !remote.has(name) ? 'repository only'
      : local.get(name) === remote.get(name) ? 'identical' : 'different',
    localHash: local.has(name) ? createHash('sha256').update(local.get(name)).digest('hex') : null,
    remoteHash: remote.has(name) ? createHash('sha256').update(remote.get(name)).digest('hex') : null,
  }));
}

async function main() {
  const authPath = process.argv[2];
  const deploymentId = process.env.DEPLOYMENT_ID;
  if (!authPath || !deploymentId) throw new Error('Auth file and GAS_DEPLOYMENT_ID required');
  const config = JSON.parse(fs.readFileSync('.clasp.json', 'utf8'));
  const { credentials } = await initAuth({ authFilePath: authPath });
  if (!credentials) throw new Error('No OAuth credentials');
  const base = `https://script.googleapis.com/v1/projects/${encodeURIComponent(config.scriptId)}`;
  const { data: deployment } = await credentials.request({ url: `${base}/deployments/${encodeURIComponent(deploymentId)}` });
  const version = deployment.deploymentConfig?.versionNumber;
  if (!Number.isInteger(version) || version < 1) throw new Error('Deployment has no immutable version');
  const { data: content } = await credentials.request({ url: `${base}/content`, params: { versionNumber: version } });
  if (!Array.isArray(content.files) || !content.files.length) throw new Error('No published files returned');
  const rows = compareFiles(content.files, config.rootDir || '.');
  const changed = rows.filter(r => r.status !== 'identical');
  const report = [
    `Published version: ${version}`,
    `Repository commit: ${process.env.GITHUB_SHA || 'local'}`,
    `Compared files: ${rows.length}; identical: ${rows.length - changed.length}; differences: ${changed.length}`,
    'Comparison ignores line endings, trailing final newlines, and JSON formatting. No Google data or code changed.',
    ...changed.map(r => `${r.status}: ${r.name}`),
  ].join('\n');
  console.log(report);
  if (process.env.GITHUB_STEP_SUMMARY) fs.appendFileSync(process.env.GITHUB_STEP_SUMMARY, '```text\n' + report + '\n```\n');
}
if (process.argv[1] && path.resolve(process.argv[1]) === path.resolve(new URL(import.meta.url).pathname)) {
  main().catch(error => {
    // Never serialize OAuth/API error objects: they can contain request headers.
    console.error(`Comparison failed (HTTP ${error.response?.status || 'n/a'}); check credentials, deployment ID and permissions.`);
    process.exitCode = 1;
  });
}
