// setup.js — одноразовая подготовка новой машины:
// зависимости + локальный кэш (.cache/) для офлайн-сборки.
const { execSync } = require("child_process");
const path = require("path");

const root = __dirname;
const env = {
  ...process.env,
  ELECTRON_BUILDER_CACHE: path.join(root, ".cache", "builder"),
  ELECTRON_CACHE: path.join(root, ".cache", "electron"),
};
const run = (cmd) => execSync(cmd, { stdio: "inherit", cwd: root, env });

const arch = process.arch === "arm64" ? "arm64" : "x64";
console.log(`[setup] arch=${arch}, cache=${env.ELECTRON_BUILDER_CACHE}`);

// 1. Зависимости + бинарник Electron (postinstall положит zip в .cache/electron)
run("npm install");

// 2. Прогревочная сборка: тянет nsis, nsis-resources и 7zip-win-<arch> в .cache/builder.
//    Ретраи — на случай вашего нестабильного канала (ECONNRESET).
const build = `npx electron-builder --win portable --${arch}`;
for (let attempt = 1; ; attempt++) {
  try {
    run(build);
    break;
  } catch (err) {
    if (attempt >= 3) throw err;
    console.log(`[setup] попытка ${attempt} не удалась, повторяю...`);
  }
}

console.log("[setup] готово. .cache/ в git не попадёт.");