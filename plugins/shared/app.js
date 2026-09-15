import { getHost } from "/hosts.js";

const status = document.querySelector("#status");
const progress = document.querySelector("#progress");
const buttons = [...document.querySelectorAll("button")];
let host;

function busy(value) {
  buttons.forEach((button) => { button.disabled = value; });
}

function download(base64, filename) {
  const binary = atob(base64);
  const bytes = Uint8Array.from(binary, (char) => char.charCodeAt(0));
  const url = URL.createObjectURL(new Blob([bytes], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }));
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  link.click();
  URL.revokeObjectURL(url);
}

async function convert(operation, target) {
  busy(true);
  progress.value = 20;
  status.textContent = "正在读取工作簿……";
  try {
    const request = { ...(await host.payload()), operation };
    if (operation === "floating-native") {
      const selected = await host.selection();
      Object.assign(request, selected, { target });
    }
    status.textContent = "正在转换图片……";
    progress.value = 70;
    const response = await fetch("/api/convert", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(request) });
    const result = await response.json();
    if (!response.ok) throw new Error(result.error || "转换失败");
    download(result.workbookBase64, result.filename);
    progress.value = 100;
    status.textContent = `完成：成功 ${result.converted} 张，已生成 ${result.filename}`;
  } catch (error) {
    progress.value = 0;
    status.textContent = `失败：${error.message}`;
  } finally {
    busy(false);
  }
}

document.querySelector("#to-floating").addEventListener("click", () => convert("compatible"));
document.querySelector("#to-excel").addEventListener("click", () => convert("floating-native", "excel"));
document.querySelector("#to-wps").addEventListener("click", () => convert("floating-native", "wps"));

async function initialize() {
  try {
    if (typeof Office !== "undefined" && Office.onReady) await Office.onReady();
    host = await getHost();
    document.querySelector("#host").textContent = `当前宿主：${host.name}`;
  } catch (error) {
    document.querySelector("#host").textContent = error.message;
    buttons.forEach((button) => { button.disabled = true; });
  }
}

initialize();
