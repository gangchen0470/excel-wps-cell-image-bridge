using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Xml.Linq;


namespace CellImageBridgeVsto {
    public partial class ThisAddIn {
        private dynamic app;
        private bool busy;
        private static readonly Regex Formula = new Regex(@"^\s*=?\s*(?:_xlfn\.)?DISPIMG\s*\(\s*""([^""]+)""\s*[,;]\s*1\s*\)\s*$", RegexOptions.IgnoreCase);
        private class Picture { public byte[] Bytes; public int Width; public int Height; public string Extension; }
        private class Job { public dynamic Cell; public dynamic Area; public Picture Picture; public string Formula; public dynamic Shape; public bool Cleared; }

        private static void Trace(string message) {
            try { File.AppendAllText(Path.Combine(Path.GetTempPath(), "CellImageBridge.log"), DateTime.Now.ToString("O") + " " + message + Environment.NewLine); } catch { }
        }
        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject() {
            Trace("CreateRibbonExtensibilityObject");
            return new RibbonController(this);
        }

        public string GetCustomUI(string ribbonID) {
            Trace("GetCustomUI " + ribbonID);
            return @"<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui'><ribbon><tabs><tab id='CellImageBridgeTab' label='图片修复'><group id='CellImageBridgeGroup' label='WPS 单元格图片'><button id='DetectImages' label='检测 WPS 图片' size='large' imageMso='FindDialog' onAction='Detect'/><button id='RepairImages' label='一键转换' size='large' imageMso='PictureInsertFromFile' onAction='Repair'/><button id='SaveImages' label='另存为兼容版' size='large' imageMso='FileSaveAs' onAction='SaveCopy'/></group></tab></tabs></ribbon></customUI>";
        }
        private static void Notify(string text) { MessageBox.Show(text, "WPS 图片修复", MessageBoxButtons.OK, MessageBoxIcon.Information); }
        public void Detect(object control) { try { Notify("检测到 " + DetectActiveWorkbook() + " 张可转换的 WPS 图片。"); } catch (Exception e) { Notify(e.Message); } }
        public void Repair(object control) { try { int count = RepairActiveWorkbook(); Notify(count == 0 ? "未发现需要转换的 WPS 图片。" : "已转换 " + count + " 张图片，结果已显示在当前工作簿。\n尚未保存，请检查后保存或另存为兼容版。"); } catch (Exception e) { Notify("转换未完成：" + e.Message); } }
        public void SaveCopy(object control) {
            try {
                dynamic wb = Workbook();
                object chosen = app.GetSaveAsFilename(Path.GetFileNameWithoutExtension((string)wb.Name) + "_兼容版.xlsx", "Excel 工作簿 (*.xlsx), *.xlsx");
                if (chosen is bool) return;
                string path = Path.GetFullPath(Convert.ToString(chosen));
                if (!path.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase)) throw new Exception("请使用 .xlsx 扩展名。");
                if (File.Exists(path)) throw new Exception("为保留已有文件，请选择一个新的文件名。");
                wb.SaveCopyAs(path);
                Notify("兼容副本已保存：\n" + path);
            } catch (Exception e) { Notify(e.Message); }
        }
        private dynamic Workbook() {
            if (app == null || app.ActiveWorkbook == null) throw new Exception("请先在 Excel 中打开 WPS 工作簿。");
            dynamic wb = app.ActiveWorkbook;
            if (!((string)wb.Name).EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) || String.IsNullOrEmpty((string)wb.Path))
                throw new Exception("请打开已保存的 .xlsx 工作簿。");
            if (!File.Exists((string)wb.FullName)) throw new Exception("此版本需要本地工作簿，请先将云端文件下载到本地。");
            return wb;
        }
        private static XElement ReadXml(ZipArchive zip, string name) {
            var entry = zip.GetEntry(name);
            if (entry == null) throw new Exception("原文件缺少 " + name + "。请重新打开 WPS 原始文件，转换前不要在 Excel 覆盖保存。");
            using (var stream = entry.Open()) return XElement.Load(stream);
        }
        private static Dictionary<string, Picture> LoadPictures(string path) {
            var result = new Dictionary<string, Picture>();
            using (var file = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var zip = new ZipArchive(file, ZipArchiveMode.Read)) {
                var root = ReadXml(zip, "xl/cellimages.xml");
                var rels = ReadXml(zip, "xl/_rels/cellimages.xml.rels");
                XNamespace xdr = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
                XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
                XNamespace r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
                foreach (var pic in root.Descendants(xdr + "pic")) {
                    var prop = pic.Descendants(xdr + "cNvPr").FirstOrDefault();
                    var blip = pic.Descendants(a + "blip").FirstOrDefault();
                    if (prop == null || blip == null) continue;
                    string id = (string)prop.Attribute("name");
                    string rid = (string)blip.Attribute(r + "embed");
                    var rel = rels.Elements().FirstOrDefault(x => (string)x.Attribute("Id") == rid);
                    if (String.IsNullOrEmpty(id) || rel == null || (string)rel.Attribute("TargetMode") == "External") continue;
                    if ((string)rel.Attribute("Type") != "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image") continue;
                    var uri = new Uri(new Uri("http://package/xl/cellimages.xml"), (string)rel.Attribute("Target"));
                    var entry = zip.GetEntry(Uri.UnescapeDataString(uri.AbsolutePath).TrimStart('/'));
                    if (entry == null) continue;
                    using (var stream = entry.Open()) using (var buffer = new MemoryStream()) {
                        stream.CopyTo(buffer);
                        buffer.Position = 0;
                        using (var image = Image.FromStream(buffer)) {
                            if (result.ContainsKey(id)) throw new Exception("发现重复图片 ID：" + id);
                            result.Add(id, new Picture { Bytes = buffer.ToArray(), Width = image.Width, Height = image.Height, Extension = Path.GetExtension(entry.Name) });
                        }
                    }
                }
            }
            return result;
        }
        private List<Job> Plan(dynamic wb) {
            var jobs = new List<Job>();
            foreach (dynamic sheet in wb.Worksheets) {
                dynamic cells;
                try { cells = sheet.UsedRange.SpecialCells(-4123); } // xlCellTypeFormulas
                catch (COMException) { continue; }
                foreach (dynamic cell in cells.Cells) {
                    string formula = Convert.ToString(cell.Formula);
                    if (formula.IndexOf("DISPIMG", StringComparison.OrdinalIgnoreCase) < 0) continue;
                    var match = Formula.Match(formula);
                    if (!match.Success) throw new Exception("暂不支持的图片公式：" + sheet.Name + "!" + cell.Address);
                    if ((bool)sheet.ProtectContents) throw new Exception("请先取消工作表保护：" + sheet.Name);
                    dynamic area = cell.MergeArea;
                    if ((double)area.Width <= 3 || (double)area.Height <= 3) throw new Exception("单元格区域太小或已隐藏：" + sheet.Name + "!" + cell.Address);
                    jobs.Add(new Job { Cell = cell, Area = area, Formula = formula });
                }
            }
            if (jobs.Count == 0) return jobs;
            var pictures = LoadPictures((string)wb.FullName);
            foreach (var job in jobs) {
                string id = Formula.Match(job.Formula).Groups[1].Value;
                Picture picture;
                if (!pictures.TryGetValue(id, out picture)) throw new Exception("原文件中找不到图片：" + id + "。请使用尚未被 Excel 覆盖保存的 WPS 文件。");
                job.Picture = picture;
            }
            return jobs;
        }
        public int DetectActiveWorkbook() { return Plan((object)Workbook()).Count; }
        public int RepairActiveWorkbook() {
            if (busy) throw new Exception("正在转换，请稍候。");
            dynamic wb = Workbook();
            List<Job> jobs = Plan((object)wb);
            if (jobs.Count == 0) return 0;
            busy = true;
            bool events = app.EnableEvents, screen = app.ScreenUpdating;
            string folder = Path.Combine(Path.GetTempPath(), "CellImageBridge-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(folder);
            try {
                app.EnableEvents = false;
                app.ScreenUpdating = false;
                for (int i = 0; i < jobs.Count; i++) {
                    var job = jobs[i];
                    string imagePath = Path.Combine(folder, i + job.Picture.Extension);
                    File.WriteAllBytes(imagePath, job.Picture.Bytes);
                    double width = job.Area.Width, height = job.Area.Height;
                    double scale = Math.Min((width - 3) / job.Picture.Width, (height - 3) / job.Picture.Height);
                    double drawW = job.Picture.Width * scale, drawH = job.Picture.Height * scale;
                    // Excel uses points: 2 pixels at 96 DPI = 1.5 points per side.
                    job.Shape = job.Cell.Worksheet.Shapes.AddPicture(imagePath, 0, -1,
                        (double)job.Area.Left + (width - drawW) / 2, (double)job.Area.Top + (height - drawH) / 2, drawW, drawH);
                    job.Shape.LockAspectRatio = -1;
                    job.Shape.Placement = 1; // xlMoveAndSize
                    job.Shape.AlternativeText = "WPS 图片：" + job.Cell.Address;
                }
                foreach (var job in jobs) { job.Cell.Formula = ""; job.Cleared = true; }
                return jobs.Count;
            } catch (Exception original) {
                var errors = new List<string>();
                foreach (var job in jobs) {
                    try { if (job.Cleared) job.Cell.Formula = job.Formula; } catch (Exception e) { errors.Add(e.Message); }
                    try { if (job.Shape != null) job.Shape.Delete(); } catch (Exception e) { errors.Add(e.Message); }
                }
                if (errors.Count > 0) throw new Exception(original.Message + "\n部分回滚失败，请不要保存并重新打开原文件。\n" + String.Join("\n", errors));
                throw;
            } finally {
                app.EnableEvents = events;
                app.ScreenUpdating = screen;
                busy = false;
                try { Directory.Delete(folder, true); } catch (IOException) { }
            }
        }
        private void ThisAddIn_Startup(object sender, EventArgs e) {
            app = this.Application;
            Trace("VSTO startup");
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e) {
            app = null;
            Trace("VSTO shutdown");
        }

        private void InternalStartup() {
            this.Startup += new EventHandler(ThisAddIn_Startup);
            this.Shutdown += new EventHandler(ThisAddIn_Shutdown);
        }
    }

    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    public sealed class RibbonController : Microsoft.Office.Core.IRibbonExtensibility {
        private readonly ThisAddIn addIn;

        public RibbonController(ThisAddIn addIn) {
            this.addIn = addIn;
        }

        public string GetCustomUI(string ribbonID) { return addIn.GetCustomUI(ribbonID); }
        public void Detect(object control) { addIn.Detect(control); }
        public void Repair(object control) { addIn.Repair(control); }
        public void SaveCopy(object control) { addIn.SaveCopy(control); }
    }
}


