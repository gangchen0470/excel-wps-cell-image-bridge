const input = document.querySelector("#workbook");
const status = document.querySelector("#status");
const progress = document.querySelector("#progress");
const inspect = document.querySelector("#inspect");
const repair = document.querySelector("#repair");
const save = document.querySelector("#save");
const list = document.querySelector("#images");
let file = null;
let result = null;
let busy = false;

function refresh() {
  input.disabled = busy;
  inspect.disabled = busy || !file;
  repair.disabled = busy || !file;
  save.disabled = busy || !result;
  progress.hidden = !busy;
}

input.addEventListener("change", () => {
  file = input.files[0] || null;
  result = null;
  list.replaceChildren();
  if (file && (!/\.xlsx$/i.test(file.name) || file.size > 25 * 1024 * 1024)) {
    status.textContent = "请选择不超过 25 MB 的 .xlsx 文件。";
    file = null;
  } else {
    status.textContent = file ? "文件已就绪，可以检测或直接修复。" : "选择文件后即可检测或修复。";
  }
  document.querySelector("#filename").textContent = file?.name || "尚未选择有效文件";
  refresh();
});

async function base64(file) {
  const bytes = new Uint8Array(await file.arrayBuffer());
  let binary = "";
  for (let i = 0; i < bytes.length; i += 0x8000) {
    binary += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
  }
  return btoa(binary);
}

async function run(operation) {
  if (!file || busy) return;
  busy = true;
  result = null;
  list.replaceChildren();
  refresh();
  status.textContent = operation === "inspect" ? "正在检测图片……" : "正在生成兼容副本……";
  try {
    const response = await fetch("/api/convert", {
      method: "POST", headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ operation, sourceName: file.name, workbookBase64: await base64(file) }),
    });
    const data = await response.json();
    if (!response.ok) throw new Error(data.error || "处理失败");
    for (const image of data.images) {
      const item = document.createElement("li");
      item.textContent = `${image.sheet_name} · ${image.cell}`;
      list.append(item);
    }
    if (operation === "inspect") {
      status.textContent = data.count ? `检测到 ${data.count} 张 WPS 图片，可点击“一键修复”。` : "没有检测到 WPS 单元格图片。";
    } else {
      result = data;
      status.textContent = `已修复 ${data.converted} 张图片。请点击“另存为兼容版”保存结果。`;
    }
  } catch (error) {
    status.textContent = `处理失败：${error.message}`;
  } finally {
    busy = false;
    refresh();
  }
}

inspect.addEventListener("click", () => run("inspect"));
repair.addEventListener("click", () => run("compatible"));
save.addEventListener("click", () => {
  if (!result || busy) return;
  const bytes = Uint8Array.from(atob(result.workbookBase64), char => char.charCodeAt(0));
  const url = URL.createObjectURL(new Blob([bytes], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }));
  const link = document.createElement("a");
  link.href = url;
  link.download = result.filename;
  document.body.append(link);
  link.click();
  link.remove();
  setTimeout(() => URL.revokeObjectURL(url), 60000);
  status.textContent = `已发起下载：${result.filename}。请在浏览器下载列表中确认保存。`;
});
refresh();
