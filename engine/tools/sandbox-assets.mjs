// Build and serve an isolated development host. Font binaries are read-only;
// only compiled JavaScript is emitted under engine/target/.
import { spawnSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import { readFileSync, readdirSync, existsSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');
export const sha256 = bytes => createHash('sha256').update(bytes).digest('hex');

export function createSandboxAssets({symbolFallback=false}={}) {
  const host = readdirSync(resolve(root,'web/host')).filter(name => name.endsWith('.ts')).sort();
  const tsc = resolve(root,'../apps/windows/node_modules/typescript/bin/tsc');
  if (!existsSync(tsc)) throw new Error('Existing local TypeScript compiler is missing; no production install was attempted.');
  const compiled = spawnSync(process.execPath,[tsc,'--target','es2022','--module','es2022','--moduleResolution','bundler',
    '--strict','--lib','es2022,dom,dom.iterable','--rootDir','web','--outDir','target/sandbox',
    ...host.map(name => `web/host/${name}`),'web/sandbox/main.ts'],{cwd:root,encoding:'utf8',timeout:30000});
  if (compiled.error) throw compiled.error;
  if (compiled.status !== 0) throw new Error(`Sandbox host compilation failed:\n${compiled.stdout}\n${compiled.stderr}`);
  const wasm = readFileSync(resolve(root,'web/pkg/visualtex_engine_wasm_bg.wasm'));
  const bindings = readFileSync(resolve(root,'web/pkg/visualtex_engine_wasm.js'));
  const fonts = [
    {id:'math',file:process.env.VISUALTEX_TEST_MATH_FONT ?? resolve(root,'../apps/windows/src-windows/artifacts/office-fonts/latinmodern-math.otf'),
      url:'math.otf',mime:'font/otf',hash:'6075562b771f8b82f0c179e363389684f2dd09de30038269e2628e504bd7be0f'},
    {id:'text',file:process.env.VISUALTEX_TEST_TEXT_FONT ?? 'C:/Windows/Fonts/NotoSansSC-VF.ttf',
      url:'text.ttf',mime:'font/ttf',hash:'763146584cf0710223441356b4395e279021b0806c196614377a7a0174ae074a'},
    ...(symbolFallback ? [{id:'mathFallback',file:resolve(root,'target/test-resources/stix-two-2.13/STIXTwoMath-Regular.otf'),
      url:'math-fallback.otf',mime:'font/otf',hash:'f2076b9f1676438439dd41e23676f5ab99056e83d6b8f8c27841591ef2ccfa72'}] : []),
  ].map(font => {
    const bytes = readFileSync(font.file);
    if (sha256(bytes) !== font.hash) throw new Error(`The fixed ${font.id} font fixture has changed. No family-name fallback is allowed.`);
    return {...font,bytes};
  });
  const assets = new Map([
    ['/assets/core.js',{type:'text/javascript; charset=utf-8',bytes:bindings}],
    ['/assets/core.wasm',{type:'application/wasm',bytes:wasm}],
    ['/sandbox/style.css',{type:'text/css; charset=utf-8',bytes:readFileSync(resolve(root,'web/sandbox/style.css'))}],
    ['/sandbox/main.js',{type:'text/javascript; charset=utf-8',bytes:readFileSync(resolve(root,'target/sandbox/sandbox/main.js'))}],
  ]);
  for (const name of host) assets.set(`/host/${name.replace(/\.ts$/u,'.js')}`,{
    type:'text/javascript; charset=utf-8',bytes:readFileSync(resolve(root,'target/sandbox/host',name.replace(/\.ts$/u,'.js'))),
  });
  for (const font of fonts) assets.set(`/assets/${font.url}`,{type:font.mime,bytes:font.bytes});
  const fingerprint = createHash('sha256');
  for (const [path,resource] of assets) {fingerprint.update(path);fingerprint.update(resource.bytes);}
  const buildId = `sandbox-${fingerprint.digest('hex').slice(0,24)}`;
  const manifest = {schemaVersion:1,engineProtocol:1,buildId,assets:[
    {id:'core',kind:'wasm',url:'core.wasm',sha256:sha256(wasm),byteLength:wasm.length},
    ...fonts.map(font => ({id:font.id,kind:'font',url:font.url,sha256:font.hash,byteLength:font.bytes.length})),
  ]};
  assets.set('/manifest.json',{type:'application/json; charset=utf-8',bytes:Buffer.from(JSON.stringify(manifest))});
  const html = readFileSync(resolve(root,'web/sandbox/index.html'),'utf8').replace('__ENGINE_BUILD_ID__',buildId);
  return {assets,html,manifest,root};
}
