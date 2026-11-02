// setup.js — одноразовая подготовка новой машины.
// Скачивает зависимости, Electron, Pandoc и наполняет локальный кэш (.cache/),
// который нужен для офлайн-сборок и не попадает в git.
const { execSync } = require("child_process");
const path = require("path");
const fs = require("fs");
const os = require("os");


const root = __dirname;
const env = {
  ...process.env,
  ELECTRON_BUILDER_CACHE: path.join(root, ".cache", "builder"),
  ELECTRON_CACHE: path.join(root, ".cache", "electron"),
};
const run = (cmd) => execSync(cmd, { stdio: "inherit", cwd: root, env });

const arch = process.arch === "arm64" ? "arm64" : "x64";
console.log(`[setup] arch=${arch}, cache=${env.ELECTRON_BUILDER_CACHE}`);

// ─── Pandoc ──────────────────────────────────────────────────────────────────
// pandoc пишется на Haskell, отдельной JS-библиотеки-конвертера не существует;
// npm-пакеты вида "node-pandoc" — лишь обёртки над тем же exe. Поэтому берём
// официальный статический бинарник и кладём его в bin/ — electron-builder
// упакует его в сборку (asarUnpacked), а index.js найдёт рядом с приложением.
const PANDOC_VERSION = "3.6.3";


// Синхронное скачивание через временный runner-файл: GitHub отдаёт 302 на
// objects.githubusercontent.com, поэтому редиректы обрабатываются вручную;
// блокирующий запуск дочернего процесса нужен потому, что setup.js — обычный
// скрипт без top-level await.
function downloadFileSync(url, dest) {
  const runner = path.join(os.tmpdir(), `mdv-dl-${process.pid}-${Date.now()}.js`);
  const code = `
    const https=require('https'),fs=require('fs');
    const url=${JSON.stringify(url)},dest=${JSON.stringify(dest)};
    const get=(u,d)=>{ if(d>5){console.error('too many redirects');process.exit(2);}
      https.get(u,{headers:{'User-Agent':'mdv-setup'}},res=>{
        if([301,302,303,307,308].includes(res.statusCode)&&res.headers.location){res.resume();return get(res.headers.location,d+1);}
        if(res.statusCode!==200){res.resume();console.error('HTTP '+res.statusCode);process.exit(3);}
        const out=fs.createWriteStream(dest+'.part');
        res.pipe(out); out.on('finish',()=>out.close(()=>{fs.renameSync(dest+'.part',dest);process.exit(0);})); out.on('error',e=>{console.error(e.message);process.exit(4);});
      }).on('error',e=>{console.error(e.message);process.exit(5);});
    };
    get(url,0);`;
  try {
    fs.writeFileSync(runner, code);
    execSync(`node "${runner}"`, { stdio: "inherit" });
  } finally {
    try { fs.rmSync(runner, { force: true }); } catch (_) {}
  }
}

function ensurePandocSync() {
  const binDir = path.join(root, "bin");
  const exeName = process.platform === "win32" ? "pandoc.exe" : "pandoc";
  const target = path.join(binDir, exeName);
  if (fs.existsSync(target)) {
    console.log(`[setup] pandoc уже есть: ${target}`);
    return;
  }
  // URL под текущую ОС/архитектуру (официальные релизные архивы pandoc).
  let asset;
  if (process.platform === "win32") {
    asset = arch === "arm64"
      ? `pandoc-${PANDOC_VERSION}-windows-arm64.zip`
      : `pandoc-${PANDOC_VERSION}-windows-x86_64.zip`;
  } else if (process.platform === "darwin") {
    asset = `pandoc-${PANDOC_VERSION}-universal-macOS.zip`;
  } else {
    asset = arch === "arm64"
      ? `pandoc-${PANDOC_VERSION}-linux-arm64.tar.gz`
      : `pandoc-${PANDOC_VERSION}-linux-amd64.tar.gz`;
  }
  const url = `https://github.com/jgm/pandoc/releases/download/${PANDOC_VERSION}/${asset}`;
  const cacheDir = path.join(root, ".cache", "pandoc");
  const cacheArc = path.join(cacheDir, asset);
  fs.mkdirSync(cacheDir, { recursive: true });
  if (!fs.existsSync(cacheArc)) {
    console.log(`[setup] скачиваю pandoc ${PANDOC_VERSION} (${asset})...`);
    downloadFileSync(url, cacheArc);
  } else {
    console.log(`[setup] использую кэшированный архив: ${cacheArc}`);
  }
  // Распаковка во временный каталог.
  const tmpDir = path.join(cacheDir, "unpacked");
  fs.rmSync(tmpDir, { recursive: true, force: true });
  fs.mkdirSync(tmpDir, { recursive: true });
  if (asset.endsWith(".zip")) {
    if (process.platform === "win32") {
      execSync(`powershell -NoProfile -ExecutionPolicy Bypass -Command "Expand-Archive -LiteralPath '${cacheArc}' -DestinationPath '${tmpDir}' -Force"`, { stdio: "inherit" });
    } else {
      execSync(`unzip -q -o '${cacheArc}' -d '${tmpDir}'`, { stdio: "inherit" });
    }
  } else {
    execSync(`tar -xzf '${cacheArc}' -C '${tmpDir}'`, { stdio: "inherit" });
  }
  // В архиве папка вида pandoc-3.6.3/(bin/)pandoc(.exe)
  const found = [];
  (function walk(dir) {
    for (const e of fs.readdirSync(dir, { withFileTypes: true })) {
      const p = path.join(dir, e.name);
      if (e.isDirectory()) walk(p);
      else if (e.name === exeName) found.push(p);
    }
  })(tmpDir);
  if (!found.length) throw new Error(`В архиве не найден ${exeName}`);
  fs.mkdirSync(binDir, { recursive: true });
  fs.copyFileSync(found[0], target);
  try { fs.chmodSync(target, 0o755); } catch (_) { /* на Windows no-op */ }
  console.log(`[setup] ✓ pandoc установлен в ${target}`);
}

// 1. Зависимости + бинарник Electron (postinstall положит zip в .cache/electron)
console.log("[setup] шаг 1/3: npm install");
run("npm install");

// 2. Pandoc в bin/ (нужен для экспорта .docx; кэшируется в .cache/pandoc)
console.log("[setup] шаг 2/3: pandoc");
try {
  ensurePandocSync();
} catch (err) {
  console.warn(`[setup] не удалось подготовить pandoc (${err.message}).`);
  console.warn("[setup] экспорт .docx будет использовать системный pandoc или встроенный конвертер.");
}

// 3. Прогревочная сборка: тянет nsis, nsis-resources и 7zip-win-<arch> в .cache/builder.
//    Повторяем до 3 раз — на случай нестабильного канала (бывает ECONNRESET).
console.log(`[setup] шаг 3/3: прогревочная сборка portable --${arch}`);
const build = `npx electron-builder --win portable --${arch}`;
for (let attempt = 1; ; attempt++) {
  try {
    run(build);
    break;
  } catch (err) {
    if (attempt >= 3) {
      console.error("[setup] 3 попытки не удались. Проверьте интернет или прокси.");
      throw err;
    }
    console.log(`[setup] попытка ${attempt} не удалась, повторяю через 5 сек...`);
    execSync("ping -n 6 127.0.0.1 >nul");
  }
}

console.log("");
console.log("[setup] ✓ готово. .cache/ наполнен, в git не попадёт.");
console.log("[setup] Дальше используйте:");
console.log("         npm start          — запустить приложение");
console.log("         npm run dist       — собрать portable-исполняемый файл");