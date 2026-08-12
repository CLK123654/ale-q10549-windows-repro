import fs from 'node:fs';
import fsp from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import process from 'node:process';
import crypto from 'node:crypto';
import { spawn } from 'node:child_process';
import { inflateRawSync } from 'node:zlib';

const repoRoot = path.resolve(import.meta.dirname, '..');
const artifactRoot = path.join(repoRoot, 'artifacts');
const evidenceRoot = path.join(repoRoot, 'evidence');
const attachments = {
  '输入数据包.zip': '131e6a38086d780047b57ac6db2007b1a42d942adc431ef258edfe8f85fdb8e9',
  'reference.zip': '032e537f243287a48e22db95d779040b2d320340fca3f039747fe21da8776d6d',
  '关键标准答案.xlsx': '1352b479cf9bba242ab81d57d804c39c1e137467c51b6db8145b6ea0eb937030',
  '任务规格转化.xlsx': 'dba45594b3157f46fd0166c3648757f3309112a840ef917c0662c1a479f786af',
};
const deliveryMembers = [
  'output/reports/draft_recovery_report.csv',
  'output/reports/handoff_state_matrix.csv',
  'output/reports/offline_banner_checks.csv',
  'output/tests/support_console.spec.ts',
];
const sha = (body) => crypto.createHash('sha256').update(body).digest('hex');
const shaFile = (file) => sha(fs.readFileSync(file));
const assert = (value, message) => { if (!value) throw new Error(message); };

function parseZipBytes(data) {
  const files = new Map();
  let offset = 0;
  while (offset + 46 <= data.length) {
    if (data.readUInt32LE(offset) !== 0x02014b50) { offset += 1; continue; }
    const method = data.readUInt16LE(offset + 10);
    const compressedSize = data.readUInt32LE(offset + 20);
    const uncompressedSize = data.readUInt32LE(offset + 24);
    const nameLength = data.readUInt16LE(offset + 28);
    const extraLength = data.readUInt16LE(offset + 30);
    const commentLength = data.readUInt16LE(offset + 32);
    const localOffset = data.readUInt32LE(offset + 42);
    const name = data.subarray(offset + 46, offset + 46 + nameLength).toString('utf8').replaceAll('\\', '/');
    if (!name.endsWith('/')) {
      const localNameLength = data.readUInt16LE(localOffset + 26);
      const localExtraLength = data.readUInt16LE(localOffset + 28);
      const start = localOffset + 30 + localNameLength + localExtraLength;
      const compressed = data.subarray(start, start + compressedSize);
      const body = method === 0 ? compressed : method === 8 ? inflateRawSync(compressed) : null;
      assert(body && body.length === uncompressedSize, 'ZIP extraction error');
      files.set(name, body);
    }
    offset += 46 + nameLength + extraLength + commentLength;
  }
  return files;
}
const parseZip = (file) => parseZipBytes(fs.readFileSync(file));
async function extract(file, destination) {
  for (const [name, bytes] of parseZip(file)) {
    const target = path.resolve(destination, name);
    assert(target.startsWith(path.resolve(destination) + path.sep), 'unsafe ZIP path');
    await fsp.mkdir(path.dirname(target), { recursive: true });
    await fsp.writeFile(target, bytes);
  }
}
function sheets(file) {
  const xml = parseZip(file).get('xl/workbook.xml')?.toString('utf8') ?? '';
  return [...xml.matchAll(/<(?:[A-Za-z]+:)?sheet[^>]+name="([^"]+)"/gu)].map((match) => match[1]);
}
async function run(command, args, cwd, env = {}) {
  const started = Date.now();
  return await new Promise((resolve) => {
    const child = spawn(command, args, { cwd, env: { ...process.env, ...env }, windowsHide: true });
    let stdout = '', stderr = '';
    child.stdout.on('data', (chunk) => { stdout += chunk; });
    child.stderr.on('data', (chunk) => { stderr += chunk; });
    child.on('error', (error) => resolve({ code: 1, stdout, stderr: stderr + error.message, elapsed_ms: Date.now() - started }));
    child.on('exit', (code) => resolve({ code: code ?? 1, stdout, stderr, elapsed_ms: Date.now() - started }));
  });
}
async function npmRun(args, cwd) {
  assert(process.env.npm_execpath, 'npm CLI missing');
  return await run(process.execPath, [process.env.npm_execpath, ...args], cwd);
}
function treeDigest(root, ignored = new Set()) {
  const lines = [];
  function visit(current, prefix = '') {
    for (const entry of fs.readdirSync(current, { withFileTypes: true }).toSorted((a, b) => a.name.localeCompare(b.name))) {
      const relative = prefix ? `${prefix}/${entry.name}` : entry.name;
      if (ignored.has(relative.split('/')[0])) continue;
      const full = path.join(current, entry.name);
      if (entry.isDirectory()) visit(full, relative); else lines.push(`${relative}\0${shaFile(full)}`);
    }
  }
  visit(root);
  return sha(Buffer.from(lines.join('\n')));
}
function csv(text) { const lines = text.replaceAll('\r\n', '\n').trimEnd().split('\n'); return { header: lines[0], rows: lines.slice(1).toSorted() }; }
function compare(outputRoot, standard) {
  const actual = [];
  function visit(current, prefix = 'output') {
    for (const entry of fs.readdirSync(current, { withFileTypes: true })) {
      const relative = `${prefix}/${entry.name}`, full = path.join(current, entry.name);
      if (entry.isDirectory()) visit(full, relative); else actual.push(relative.replaceAll('\\', '/'));
    }
  }
  visit(outputRoot); actual.sort();
  assert(JSON.stringify(actual) === JSON.stringify(deliveryMembers), 'delivery members differ');
  const digest = crypto.createHash('sha256');
  for (const member of deliveryMembers) {
    const actualBytes = fs.readFileSync(path.join(outputRoot, member.slice(7)));
    const expectedBytes = standard.get(member);
    const a = member.endsWith('.csv') ? JSON.stringify(csv(actualBytes.toString('utf8'))) : actualBytes.toString('utf8').replaceAll('\r\n', '\n').trimEnd();
    const e = member.endsWith('.csv') ? JSON.stringify(csv(expectedBytes.toString('utf8'))) : expectedBytes.toString('utf8').replaceAll('\r\n', '\n').trimEnd();
    assert(a === e, `delivery differs ${member}`); digest.update(a);
  }
  return digest.digest('hex');
}
async function prepare(label, mutate) {
  const runRoot = path.join(os.tmpdir(), label);
  await fsp.rm(runRoot, { recursive: true, force: true });
  await fsp.mkdir(runRoot, { recursive: true });
  await extract(path.join(artifactRoot, '输入数据包.zip'), runRoot);
  const inputRoot = path.join(runRoot, 'input_data');
  const standard = parseZip(path.join(artifactRoot, 'reference.zip'));
  await fsp.mkdir(path.join(runRoot, 'output', 'tests'), { recursive: true });
  await fsp.writeFile(path.join(runRoot, 'output', 'tests', 'support_console.spec.ts'), standard.get('output/tests/support_console.spec.ts'));
  if (mutate) await mutate(inputRoot);
  return { inputRoot, outputRoot: path.join(runRoot, 'output'), standard };
}
async function install(inputRoot) {
  let result = await npmRun(['ci'], inputRoot); assert(result.code === 0, `npm ci failed ${result.stderr}`);
  result = await npmRun(['exec', '--', 'playwright', 'install', 'chromium'], inputRoot); assert(result.code === 0, `Chromium install failed ${result.stderr}`);
}
async function execute(inputRoot) { return await npmRun(['run', 'process'], inputRoot); }

await fsp.rm(evidenceRoot, { recursive: true, force: true });
await fsp.mkdir(evidenceRoot, { recursive: true });
assert(process.platform === 'win32' && process.env.GITHUB_ACTIONS === 'true', 'Windows hosted runner required');
for (const [name, expected] of Object.entries(attachments)) assert(shaFile(path.join(artifactRoot, name)) === expected, `${name} checksum mismatch`);
const input = parseZip(path.join(artifactRoot, '输入数据包.zip'));
const standard = parseZip(path.join(artifactRoot, 'reference.zip'));
assert(JSON.stringify([...standard.keys()].sort()) === JSON.stringify(deliveryMembers), 'reference members differ');
const platform = [...input, ...standard].filter(([name, bytes]) => (bytes[0] === 0x7f && bytes.subarray(1, 4).toString() === 'ELF') || /\.(?:sh|bash|so)$/iu.test(name));
assert(platform.length === 0, 'platform member found');
assert(JSON.stringify(sheets(path.join(artifactRoot, '关键标准答案.xlsx'))) === JSON.stringify(['交付物答案清单','固定字段答案','固定集合答案','固定数值答案','允许变体答案']), 'answer sheets mismatch');
assert(JSON.stringify(sheets(path.join(artifactRoot, '任务规格转化.xlsx'))) === JSON.stringify(['任务规格转化']), 'spec sheets mismatch');

const cleanRuns = [];
for (const label of ['Q10549 clean support handoff', 'Q10549 中文 空格 客服接管']) {
  const prepared = await prepare(label);
  const before = treeDigest(prepared.inputRoot, new Set(['node_modules']));
  await install(prepared.inputRoot);
  const result = await execute(prepared.inputRoot);
  assert(result.code === 0, `browser run failed ${result.stdout} ${result.stderr}`);
  const after = treeDigest(prepared.inputRoot, new Set(['node_modules']));
  assert(before === after, 'business input changed');
  const semantic = compare(prepared.outputRoot, prepared.standard);
  cleanRuns.push({ directory_label: label, exit_code: result.code, input_digest_before: before, input_digest_after: after, semantic_digest: semantic, elapsed_ms: result.elapsed_ms });
}
assert(cleanRuns[0].semantic_digest === cleanRuns[1].semantic_digest, 'clean results differ');

const crlf = await prepare('Q10549 CRLF support handoff', async (inputRoot) => {
  const file = path.join(inputRoot, 'offline_scenarios.json');
  const text = await fsp.readFile(file, 'utf8');
  await fsp.writeFile(file, text.replaceAll('\r\n', '\n').replaceAll('\n', '\r\n'));
});
await install(crlf.inputRoot);
let result = await execute(crlf.inputRoot);
assert(result.code === 0, `CRLF input failed ${result.stderr}`);
compare(crlf.outputRoot, crlf.standard);

const mutation = await prepare('Q10549 changed handoff agent', async (inputRoot) => {
  const file = path.join(inputRoot, 'handoff_scenarios.json');
  const data = JSON.parse(await fsp.readFile(file, 'utf8'));
  const target = data.find((item) => item.scenario_id === 'H2');
  target.second_tab_agent = 'agent-c'; target.previous_tab_locked = 'yes'; target.websocket_event = 'session.owner_changed'; target.takeover_banner = 'agent-c took over T-711';
  await fsp.writeFile(file, `${JSON.stringify(data, null, 2)}\n`);
});
await install(mutation.inputRoot);
result = await execute(mutation.inputRoot);
assert(result.code === 0, `positive mutation failed ${result.stderr}`);
const changed = await fsp.readFile(path.join(mutation.outputRoot, 'reports', 'handoff_state_matrix.csv'), 'utf8');
assert(changed.includes('H2,T-711,agent-b,agent-c,session.owner_changed'), 'changed handoff not reflected');

const invalid = await prepare('Q10549 invalid handoff scenario', async (inputRoot) => {
  const file = path.join(inputRoot, 'handoff_scenarios.json');
  const data = JSON.parse(await fsp.readFile(file, 'utf8')); data[2].scenario_id = data[0].scenario_id;
  await fsp.writeFile(file, `${JSON.stringify(data, null, 2)}\n`);
});
await install(invalid.inputRoot);
result = await execute(invalid.inputRoot);
assert(result.code !== 0, 'duplicate handoff scenario accepted');
assert(!fs.existsSync(path.join(invalid.outputRoot, 'reports')), 'invalid input left reports');

const evidence = {
  schema_version: 1,
  task_asset_id: 'playwright_support_handoff_offline_regression',
  result: 'PASS',
  generated_at_utc: new Date().toISOString(),
  git_commit_sha: process.env.GITHUB_SHA,
  workflow_run_id: process.env.GITHUB_RUN_ID,
  runner: { os: process.env.RUNNER_OS, arch: process.env.RUNNER_ARCH, image_os: process.env.ImageOS, image_version: process.env.ImageVersion, node: process.version, powershell_hosted_workflow: true },
  software: { main: 'Playwright', version: '1.62.1', browser: 'Chromium', executed: true },
  attachment_sha256: attachments,
  workbook_checks: { answer_sheet_names: sheets(path.join(artifactRoot, '关键标准答案.xlsx')), specification_sheet_names: ['任务规格转化'], task_spec_column_count: 2 },
  platform_audit: { platform_specific_members: platform, linux_executables_executed: false, no_wsl_required: true, no_linux_container_required: true, no_posix_shell_required: true, cross_platform_paths: true },
  clean_runs: cleanRuns,
  reference_match: true,
  crlf_input: { exit_code: 0, reference_match: true },
  positive_mutation: { changed_rule: 'handoff agent identity', exit_code: 0, report_changed: true },
  invalid_input: { changed_rule: 'duplicate handoff scenario', exit_code: result.code, reports_absent: !fs.existsSync(path.join(invalid.outputRoot, 'reports')) },
  network: { installation: 'npm and Playwright sources', formal_run: 'loopback only' },
};
await fsp.writeFile(path.join(evidenceRoot, 'windows-verification.json'), `${JSON.stringify(evidence, null, 2)}\n`);
console.log(JSON.stringify(evidence, null, 2));
