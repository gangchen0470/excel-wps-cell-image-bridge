function excelReady() {
  return typeof Office !== "undefined" && typeof Excel !== "undefined";
}

async function excelSelection() {
  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const range = context.workbook.getSelectedRange();
    sheet.load("name");
    range.load("address");
    await context.sync();
    return { sheet: sheet.name, range: range.address.replace(/^.*!/, "").replace(/\$/g, "") };
  });
}

function excelBytes() {
  return new Promise((resolve, reject) => {
    Office.context.document.getFileAsync(Office.FileType.Compressed, { sliceSize: 4 * 1024 * 1024 }, (result) => {
      if (result.status !== Office.AsyncResultStatus.Succeeded) return reject(new Error(result.error.message));
      const file = result.value;
      const chunks = [];
      let index = 0;
      const next = () => file.getSliceAsync(index, (sliceResult) => {
        if (sliceResult.status !== Office.AsyncResultStatus.Succeeded) {
          file.closeAsync();
          return reject(new Error(sliceResult.error.message));
        }
        chunks.push(new Uint8Array(sliceResult.value.data));
        index += 1;
        if (index < file.sliceCount) return next();
        file.closeAsync();
        const size = chunks.reduce((total, chunk) => total + chunk.length, 0);
        const bytes = new Uint8Array(size);
        let offset = 0;
        for (const chunk of chunks) { bytes.set(chunk, offset); offset += chunk.length; }
        resolve(bytes);
      });
      next();
    });
  });
}

function bytesToBase64(bytes) {
  let binary = "";
  const step = 0x8000;
  for (let i = 0; i < bytes.length; i += step) binary += String.fromCharCode(...bytes.subarray(i, i + step));
  return btoa(binary);
}

const excelHost = {
  name: "Microsoft Excel",
  selection: excelSelection,
  async payload() {
    const url = Office.context.document.url || "workbook.xlsx";
    return { workbookBase64: bytesToBase64(await excelBytes()), sourceName: decodeURIComponent(url.split(/[\\/]/).pop()) };
  },
};

const wpsHost = {
  name: "WPS 表格",
  async selection() {
    const sheet = Application.ActiveSheet || Application.ActiveWorkbook.ActiveSheet;
    const selection = Application.Selection;
    const address = typeof selection.Address === "function" ? selection.Address() : selection.Address;
    return { sheet: sheet.Name, range: String(address).replace(/^.*!/, "").replace(/\$/g, "") };
  },
  async payload() {
    Application.ActiveWorkbook.Save();
    return { sourcePath: Application.ActiveWorkbook.FullName };
  },
};

export async function getHost() {
  if (typeof Application !== "undefined" && Application.ActiveWorkbook) return wpsHost;
  if (excelReady()) return excelHost;
  throw new Error("未检测到 Excel 或 WPS 加载项宿主");
}
