const form = document.querySelector("#adminForm");
const enabled = document.querySelector("#enabled");
const newPassword = document.querySelector("#newPassword");
const confirmPassword = document.querySelector("#confirmPassword");
const status = document.querySelector("#status");
const downloadLink = document.querySelector("#downloadLink");
let currentPolicy = null;
let downloadUrl = null;
let edited = false;
form.addEventListener("input", () => { edited = true; });

function bytesToBase64(bytes) {
  let binary = "";
  for (const byte of bytes) binary += String.fromCharCode(byte);
  return btoa(binary);
}

function validPolicy(data) {
  return data && data.version === 1 && typeof data.enabled === "boolean" &&
    Number.isInteger(data.iterations) && data.iterations >= 100000 && data.iterations <= 1000000 &&
    typeof data.salt === "string" && typeof data.hash === "string";
}

async function loadPolicy() {
  if (location.protocol === "file:") {
    status.textContent = "本地配置生成模式。请输入新密码；生成文件后替换 public/scheduler-access.json 并重新部署。";
    return;
  }
  try {
    const response = await fetch(`/scheduler-access.json?v=${Date.now()}`, { cache: "no-store" });
    if (!response.ok) throw new Error(`HTTP ${response.status}`);
    const data = await response.json();
    if (!validPolicy(data)) throw new Error("配置格式不正确");
    currentPolicy = data;
    if (!edited) enabled.checked = data.enabled;
    status.textContent = `已读取当前部署的配置：${data.enabled ? "允许排课" : "暂停排课"}。修改后请下载并重新部署。`;
  } catch (error) {
    status.textContent = `无法读取当前配置（${error.message}）。仍可输入新密码生成配置；留空不能保留原密码。`;
  }
}

form.addEventListener("submit", async (event) => {
  event.preventDefault();
  if (newPassword.value !== confirmPassword.value) {
    status.textContent = "两次输入的新密码不一致。";
    return;
  }
  if (enabled.checked && !newPassword.value && !currentPolicy?.enabled) {
    status.textContent = "首次启用或重新启用时，请设置新密码。";
    return;
  }
  try {
    status.textContent = "正在生成配置……";
    let { salt, hash, iterations } = currentPolicy ?? {
      salt: bytesToBase64(crypto.getRandomValues(new Uint8Array(16))),
      hash: bytesToBase64(crypto.getRandomValues(new Uint8Array(32))),
      iterations: 310000,
    };
    if (newPassword.value) {
      const saltBytes = crypto.getRandomValues(new Uint8Array(16));
      const key = await crypto.subtle.importKey("raw", new TextEncoder().encode(newPassword.value), "PBKDF2", false, ["deriveBits"]);
      const hashBytes = new Uint8Array(await crypto.subtle.deriveBits(
        { name: "PBKDF2", salt: saltBytes, iterations: 310000, hash: "SHA-256" }, key, 256));
      salt = bytesToBase64(saltBytes);
      hash = bytesToBase64(hashBytes);
      iterations = 310000;
    }
    const policy = { version: 1, enabled: enabled.checked, iterations, salt, hash, updatedUtc: new Date().toISOString() };
    if (downloadUrl) URL.revokeObjectURL(downloadUrl);
    const objectUrl = URL.createObjectURL(new Blob([JSON.stringify(policy, null, 2) + "\n"], { type: "application/json" }));
    downloadUrl = objectUrl;
    downloadLink.href = objectUrl;
    downloadLink.hidden = false;
    downloadLink.click();
    currentPolicy = policy;
    newPassword.value = "";
    confirmPassword.value = "";
    status.textContent = "配置已生成。如果没有自动下载，请点击下方下载链接。用文件替换网页工程/public/scheduler-access.json，再重新构建并部署。";
  } catch (error) {
    status.textContent = `生成失败：${error.message}`;
  }
});

loadPolicy();
