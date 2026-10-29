// setup.js — одноразовая подготовка новой машины.
// Скачивает зависимости, Electron и наполняет локальный кэш (.cache/),
// который нужен для офлайн-сборок и не попадает в git.
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
console.log("[setup] шаг 1/2: npm install");
run("npm install");

// 2. Прогревочная сборка: тянет nsis, nsis-resources и 7zip-win-<arch> в .cache/builder.
//    Повторяем до 3 раз — на случай нестабильного канала (бывает ECONNRESET).
console.log(`[setup] шаг 2/2: прогревочная сборка portable --${arch}`);
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