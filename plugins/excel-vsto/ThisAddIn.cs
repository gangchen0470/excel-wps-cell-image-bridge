using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Net;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Xml.Linq;
using Microsoft.Win32;


namespace CellImageBridgeVsto {
    public partial class ThisAddIn {
        private dynamic app;
        private bool busy;
        private const string CurrentVersion = "1.0.13";
        private const string GitHubUpdateManifestUrl = "https://api.github.com/repos/gangchen0470/excel-wps-cell-image-bridge/contents/update.json?ref=main";
        private const string GiteeUpdateManifestUrl = "https://gitee.com/chengang0470/excel-wps-cell-image-bridge/raw/master/update.json";
        private static readonly Regex Formula = new Regex(@"^\s*=?\s*(?:_xlfn\.)?DISPIMG\s*\(\s*""([^""]+)""\s*[,;]\s*1\s*\)\s*$", RegexOptions.IgnoreCase);
        private class Picture { public byte[] Bytes; public int Width; public int Height; public string Extension; }
        private class PackageJob { public string Sheet; public string Cell; public Picture Picture; }
        private class Job { public dynamic Cell; public dynamic Area; public Picture Picture; public string Formula; public string Source; public dynamic Shape; public bool Cleared; }

        private static void Trace(string message) {
            try { File.AppendAllText(Path.Combine(Path.GetTempPath(), "CellImageBridge.log"), DateTime.Now.ToString("O") + " " + message + Environment.NewLine); } catch { }
        }
        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject() {
            Trace("CreateRibbonExtensibilityObject");
            return new RibbonController(this);
        }

        public string GetCustomUI(string ribbonID) {
            Trace("GetCustomUI " + ribbonID);
            return @"<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui'><ribbon><tabs><tab id='CellImageBridgeTab' label='图片修复'><group id='CellImageBridgeGroup' label='Excel / WPS 单元格图片'><button id='DetectImages' label='检测单元格图片' size='large' imageMso='FindDialog' onAction='Detect'/><button id='RepairImages' label='一键转换' size='large' imageMso='PictureInsertFromFile' onAction='Repair'/><button id='SaveImages' label='修复并另存兼容版' size='large' imageMso='FileSaveAs' onAction='SaveCopy'/></group><group id='CellImageBridgeUpdateGroup' label='插件'><button id='CheckUpdate' label='检查更新' size='large' imageMso='RefreshAll' onAction='CheckUpdate'/></group></tab></tabs></ribbon></customUI>";
        }
        private static void Notify(string text) { MessageBox.Show(text, "Excel / WPS 图片修复", MessageBoxButtons.OK, MessageBoxIcon.Information); }
        public void Detect(object control) { try { var jobs = Plan((object)Workbook()); int wps = jobs.Count(x => x.Source == "wps"); Trace("Detect: total=" + jobs.Count + ", wps=" + wps); Notify("共检测到 " + jobs.Count + " 张单元格图片。\nExcel：" + (jobs.Count - wps) + " 张\nWPS：" + wps + " 张"); } catch (Exception e) { Trace("Detect failed: " + e); Notify(e.Message); } }
        public void Repair(object control) { try { int count = RepairActiveWorkbook(); Trace("Repair: converted=" + count); Notify(count == 0 ? "未发现需要转换的 Excel 或 WPS 单元格图片。" : "已转换 " + count + " 张图片。\n尚未保存，请检查后使用“修复并另存兼容版”。"); } catch (Exception e) { Trace("Repair failed: " + e); Notify("转换未完成：" + e.Message); } }
        public void CheckUpdate(object control) {
            ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls12;
            Exception lastError = null;
            foreach (var source in new[] { new { Name = "GitHub", Manifest = GitHubUpdateManifestUrl, DownloadField = "githubDownloadUrl" }, new { Name = "Gitee", Manifest = GiteeUpdateManifestUrl, DownloadField = "giteeDownloadUrl" } }) {
                try {
                    var request = (HttpWebRequest)WebRequest.Create(source.Manifest);
                    request.Timeout = 5000;
                    request.ReadWriteTimeout = 5000;
                    request.UserAgent = "CellImageBridgeVsto/" + CurrentVersion;
                    if (source.Name == "GitHub") request.Accept = "application/vnd.github.raw+json";
                    string json;
                    using (var response = request.GetResponse())
                    using (var reader = new StreamReader(response.GetResponseStream())) json = reader.ReadToEnd();
                    string versionText = JsonField(json, "version");
                    string downloadUrl = JsonField(json, source.DownloadField);
                    if (String.IsNullOrEmpty(downloadUrl)) downloadUrl = JsonField(json, "downloadUrl");
                    Uri download;
                    string expectedHash = JsonField(json, "sha256");
                    if (String.IsNullOrEmpty(versionText) || !Uri.TryCreate(downloadUrl, UriKind.Absolute, out download) || download.Scheme != Uri.UriSchemeHttps ||
                        String.IsNullOrEmpty(expectedHash) || !Regex.IsMatch(expectedHash, "^[0-9a-fA-F]{64}$"))
                        throw new Exception("更新信息格式无效。");
                    var latest = new Version(versionText);
                    var current = new Version(CurrentVersion);
                    Trace("CheckUpdate: source=" + source.Name + ", latest=" + latest);
                    if (latest <= current) { Notify("当前已是最新版本：" + CurrentVersion + "（" + source.Name + "）"); return; }
                    if (MessageBox.Show("发现新版本 " + latest + "（当前 " + current + "）。\n更新源：" + source.Name + "\n将下载、校验并准备新版。Excel 可以继续使用；新版会在下次启动 Excel 时生效。\n\n现在开始吗？", "插件在线更新", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes) return;
                    string packageFolder = DownloadUpdate(download, latest, expectedHash);
                    StageUpdate(packageFolder, latest);
                    Notify("新版 " + latest + " 已下载、校验并准备完成。\n\n当前 Excel 可以继续使用；下次启动 Excel 时会自动使用新版。");
                    return;
                } catch (Exception e) {
                    lastError = e;
                    Trace("CheckUpdate " + source.Name + " failed: " + e);
                }
            }
            Notify("检查更新失败：GitHub 和 Gitee 均不可用。\n" + lastError.Message + "\n\n可手动打开：\nhttps://gitee.com/chengang0470/excel-wps-cell-image-bridge");
        }
        private static string JsonField(string json, string field) {
            var match = Regex.Match(json, "\\\"" + Regex.Escape(field) + "\\\"\\s*:\\s*\\\"([^\\\"]+)\\\"");
            return match.Success ? match.Groups[1].Value : null;
        }
        private static string DownloadUpdate(Uri url, Version version, string expectedHash) {
            string folder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "CellImageBridgeVsto", "updates", version + "-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(folder);
            string archive = Path.Combine(folder, "CellImageBridgeVsto.zip");
            var request = (HttpWebRequest)WebRequest.Create(url);
            request.Timeout = 15000;
            request.ReadWriteTimeout = 15000;
            request.UserAgent = "CellImageBridgeVsto/" + CurrentVersion;
            using (var response = request.GetResponse())
            using (var input = response.GetResponseStream())
            using (var output = new FileStream(archive, FileMode.Create, FileAccess.Write)) {
                byte[] buffer = new byte[81920];
                int read; long total = 0;
                while ((read = input.Read(buffer, 0, buffer.Length)) > 0) {
                    total += read;
                    if (total > 50L * 1024 * 1024) throw new Exception("更新包超过 50 MB 限制。");
                    output.Write(buffer, 0, read);
                }
            }
            string actualHash;
            using (var file = File.OpenRead(archive))
            using (var sha = SHA256.Create()) actualHash = BitConverter.ToString(sha.ComputeHash(file)).Replace("-", "");
            if (!String.Equals(actualHash, expectedHash, StringComparison.OrdinalIgnoreCase)) throw new Exception("更新包校验失败，请稍后重试。");
            string extracted = Path.Combine(folder, "package");
            ZipFile.ExtractToDirectory(archive, extracted);
            if (!File.Exists(Path.Combine(extracted, "CellImageBridgeVsto.vsto")) || !File.Exists(Path.Combine(extracted, "Install-ExcelAddin.ps1")))
                throw new Exception("更新包内容不完整。");
            return extracted;
        }
        private static void StageUpdate(string packageFolder, Version version) {
            string installRoot = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "CellImageBridgeVsto", "installed", version.ToString());
            string staging = installRoot + ".staging-" + Guid.NewGuid().ToString("N");
            CopyDirectory(packageFolder, staging);
            string stagedManifest = Path.Combine(staging, "CellImageBridgeVsto.vsto");
            if (!File.Exists(stagedManifest)) throw new Exception("新版部署清单不存在。");
            if (Directory.Exists(installRoot)) Directory.Delete(installRoot, true);
            Directory.Move(staging, installRoot);
            string manifest = new Uri(Path.Combine(installRoot, "CellImageBridgeVsto.vsto")).AbsoluteUri + "|vstolocal";
            bool updated = false;
            foreach (string keyName in new[] {
                @"Software\Microsoft\Office\Excel\Addins\CellImageBridgeVsto",
                @"Software\WOW6432Node\Microsoft\Office\Excel\Addins\CellImageBridgeVsto"
            }) {
                using (RegistryKey key = Registry.CurrentUser.OpenSubKey(keyName, true)) {
                    if (key == null) continue;
                    key.SetValue("Manifest", manifest, RegistryValueKind.String);
                    key.SetValue("LoadBehavior", 3, RegistryValueKind.DWord);
                    updated = true;
                }
            }
            if (!updated) throw new Exception("未找到插件注册信息，请使用发布包中的 Install-ExcelAddin.cmd 安装。");
            File.WriteAllText(Path.Combine(installRoot, "update-ready.txt"), "version=" + version + Environment.NewLine + "prepared=" + DateTime.Now.ToString("O"));
        }
        private static void CopyDirectory(string source, string destination) {
            Directory.CreateDirectory(destination);
            foreach (string file in Directory.GetFiles(source)) File.Copy(file, Path.Combine(destination, Path.GetFileName(file)), true);
            foreach (string directory in Directory.GetDirectories(source)) CopyDirectory(directory, Path.Combine(destination, Path.GetFileName(directory)));
        }
        public void SaveCopy(object control) {
            try {
                dynamic wb = Workbook();
                int converted = RepairActiveWorkbook();
                object chosen = app.GetSaveAsFilename(Path.GetFileNameWithoutExtension((string)wb.Name) + "_兼容版.xlsx", "Excel 工作簿 (*.xlsx), *.xlsx");
                if (chosen is bool) return;
                string path = Path.GetFullPath(Convert.ToString(chosen));
                if (!path.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase)) throw new Exception("请使用 .xlsx 扩展名。");
                if (File.Exists(path)) throw new Exception("为保留已有文件，请选择一个新的文件名。");
                wb.SaveCopyAs(path);
                Notify("已转换 " + converted + " 张图片并保存兼容副本：\n" + path + "\n该文件可供 WPS 和旧版 Excel 查看。");
            } catch (Exception e) { Notify(e.Message); }
        }
        private dynamic Workbook() {
            if (app == null || app.ActiveWorkbook == null) throw new Exception("请先在 Excel 中打开工作簿。");
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
        private static string ResolvePart(string source, string target) {
            return Uri.UnescapeDataString(new Uri(new Uri("http://package/" + source), target).AbsolutePath).TrimStart('/');
        }
        private static Dictionary<string, string> Relations(ZipArchive zip, string source) {
            int slash = source.LastIndexOf('/');
            string name = source.Substring(0, slash + 1) + "_rels/" + source.Substring(slash + 1) + ".rels";
            var result = new Dictionary<string, string>();
            var entry = zip.GetEntry(name); if (entry == null) return result;
            using (var stream = entry.Open()) foreach (var rel in XElement.Load(stream).Elements()) {
                string id = (string)rel.Attribute("Id"), target = (string)rel.Attribute("Target");
                if (!String.IsNullOrEmpty(id) && !String.IsNullOrEmpty(target) && (string)rel.Attribute("TargetMode") != "External") result[id] = ResolvePart(source, target);
            }
            return result;
        }
        private static Picture ReadPicture(ZipArchive zip, string path) {
            var entry = zip.GetEntry(path); if (entry == null) throw new Exception("找不到图片文件：" + path);
            using (var input = entry.Open()) using (var buffer = new MemoryStream()) { input.CopyTo(buffer); buffer.Position = 0;
                using (var image = Image.FromStream(buffer)) return new Picture { Bytes = buffer.ToArray(), Width = image.Width, Height = image.Height, Extension = Path.GetExtension(entry.Name) };
            }
        }
        private static List<PackageJob> LoadExcelPictures(string path) {
            var found = new List<PackageJob>();
            using (var file = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var zip = new ZipArchive(file, ZipArchiveMode.Read)) {
                string[] required = { "xl/metadata.xml", "xl/richData/rdrichvalue.xml", "xl/richData/richValueRel.xml", "xl/workbook.xml" };
                if (required.Any(x => zip.GetEntry(x) == null)) return found;
                XNamespace main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
                XNamespace rich = "http://schemas.microsoft.com/office/spreadsheetml/2017/richdata", xrel = "http://schemas.microsoft.com/office/spreadsheetml/2022/richvaluerel";
                var metadata = ReadXml(zip, "xl/metadata.xml"); var indexes = new List<int>();
                var valueMetadata = metadata.Descendants(main + "valueMetadata").FirstOrDefault(); if (valueMetadata == null) return found;
                foreach (var bk in valueMetadata.Elements(main + "bk")) { int v = -1; var rc = bk.Element(main + "rc"); if (rc != null) Int32.TryParse((string)rc.Attribute("v"), out v); indexes.Add(v); }
                var ids = ReadXml(zip, "xl/richData/richValueRel.xml").Descendants(xrel + "rel").Select(x => (string)x.Attribute(r + "id") ?? "").ToList();
                var rels = Relations(zip, "xl/richData/richValueRel.xml"); var media = new List<string>();
                foreach (var rv in ReadXml(zip, "xl/richData/rdrichvalue.xml").Descendants(rich + "rv")) { int n = -1; var v = rv.Elements(rich + "v").FirstOrDefault(); if (v != null) Int32.TryParse(v.Value, out n); media.Add(n >= 0 && n < ids.Count && rels.ContainsKey(ids[n]) ? rels[ids[n]] : null); }
                var workbookRels = Relations(zip, "xl/workbook.xml");
                foreach (var sheet in ReadXml(zip, "xl/workbook.xml").Descendants(main + "sheet")) {
                    string sheetName = (string)sheet.Attribute("name"), id = (string)sheet.Attribute(r + "id"); if (String.IsNullOrEmpty(id) || !workbookRels.ContainsKey(id)) continue;
                    var entry = zip.GetEntry(workbookRels[id]); if (entry == null) continue; XElement xml; using (var stream = entry.Open()) xml = XElement.Load(stream);
                    foreach (var cell in xml.Descendants(main + "c")) { int vm; if (!Int32.TryParse((string)cell.Attribute("vm"), out vm) || vm <= 0 || vm > indexes.Count) continue; int ri = indexes[vm - 1];
                        if (ri >= 0 && ri < media.Count && !String.IsNullOrEmpty(media[ri])) found.Add(new PackageJob { Sheet = sheetName, Cell = (string)cell.Attribute("r"), Picture = ReadPicture(zip, media[ri]) });
                    }
                }
            }
            return found;
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
                    jobs.Add(new Job { Cell = cell, Area = area, Formula = formula, Source = "wps" });
                }
            }
            if (jobs.Count > 0) {
                var pictures = LoadPictures((string)wb.FullName);
                foreach (var job in jobs) {
                    string id = Formula.Match(job.Formula).Groups[1].Value; Picture picture;
                    if (!pictures.TryGetValue(id, out picture)) throw new Exception("原文件中找不到图片：" + id + "。请使用尚未被 Excel 覆盖保存的 WPS 文件。");
                    job.Picture = picture;
                }
            }
            foreach (var item in LoadExcelPictures((string)wb.FullName)) {
                dynamic sheet = wb.Worksheets[item.Sheet];
                if ((bool)sheet.ProtectContents) throw new Exception("请先取消工作表保护：" + item.Sheet);
                dynamic originalCell = sheet.Range[item.Cell], area = originalCell.MergeArea;
                dynamic cell = area.Cells[1, 1];
                if ((double)area.Width <= 3 || (double)area.Height <= 3) throw new Exception("单元格区域太小或已隐藏：" + item.Sheet + "!" + item.Cell);
                jobs.Add(new Job { Cell = cell, Area = area, Picture = item.Picture, Source = "excel" });
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
                    // Clear the complete cell or merged area before adding the floating picture.
                    // Clearing an Excel in-cell image after AddPicture can remove the new shape too.
                    job.Area.ClearContents();
                    job.Cleared = true;
                    // Excel uses points: 2 pixels at 96 DPI = 1.5 points per side.
                    job.Shape = job.Cell.Worksheet.Shapes.AddPicture(imagePath, 0, -1,
                        (double)job.Area.Left + (width - drawW) / 2, (double)job.Area.Top + (height - drawH) / 2, drawW, drawH);
                    job.Shape.LockAspectRatio = -1;
                    job.Shape.Placement = 1; // xlMoveAndSize
                    job.Shape.AlternativeText = (job.Source == "excel" ? "Excel" : "WPS") + " 单元格图片：" + job.Cell.Address;
                }
                return jobs.Count;
            } catch (Exception original) {
                var errors = new List<string>();
                foreach (var job in jobs) {
                    try { if (job.Cleared && job.Source == "wps") job.Cell.Formula = job.Formula; } catch (Exception e) { errors.Add(e.Message); }
                    try { if (job.Shape != null) job.Shape.Delete(); } catch (Exception e) { errors.Add(e.Message); }
                }
                Trace("Conversion failed: " + original);
                if (errors.Count > 0) throw new Exception(original.Message + "\n部分回滚失败，请不要保存并重新打开原文件。\n" + String.Join("\n", errors));
                if (jobs.Any(x => x.Cleared && x.Source == "excel")) throw new Exception(original.Message + "\nExcel 单元格图片已从当前编辑会话清除，请不要保存，关闭并重新打开原文件。", original);
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
        public void CheckUpdate(object control) { addIn.CheckUpdate(control); }
    }
}


