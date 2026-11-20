using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;

internal sealed class ConversionResult
{
    internal int Success;
    internal readonly List<string> Failures = new List<string>();

    internal void Show()
    {
        if (Failures.Count == 0)
        {
            if (Success > 0)
                MessageBox.Show("已转换 " + Success + " 个文件，PDF 与 Word 文件保存在同一文件夹。",
                    Program.MenuTitle, MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        const int maxDetails = 6;
        string message = "已转换 " + Success + " 个文件，" + Failures.Count + " 个失败。\n\n";
        message += string.Join("\n\n", Failures.GetRange(0, Math.Min(maxDetails, Failures.Count)));
        if (Failures.Count > maxDetails)
            message += "\n\n另有 " + (Failures.Count - maxDetails) + " 个文件失败。";
        MessageBox.Show(message, Program.MenuTitle, MessageBoxButtons.OK, MessageBoxIcon.Warning);
    }
}

internal static class WpsPdf
{
    internal static void Convert(List<string> paths, ConversionResult result)
    {
        var uniquePaths = new HashSet<string>(paths, StringComparer.OrdinalIgnoreCase);
        WpsSession wps;
        try
        {
            wps = new WpsSession();
        }
        catch (Exception error)
        {
            foreach (string path in uniquePaths)
                result.Failures.Add(Path.GetFileName(path) + " — " + error.Message);
            return;
        }

        using (wps)
        {
            foreach (string path in uniquePaths)
            {
                try
                {
                    wps.Export(path);
                    result.Success++;
                }
                catch (Exception error)
                {
                    result.Failures.Add(Path.GetFileName(path) + " — " + error.Message);
                }
            }
        }
    }

    private sealed class WpsSession : IDisposable
    {
        private const int PdfFormat = 17; // wdExportFormatPDF
        private const int DoNotSaveChanges = 0;
        private dynamic application;
        private dynamic documents;
        private bool canQuit;
        private int? originalAlerts;
        private int? originalSecurity;

        internal WpsSession()
        {
            var type = Type.GetTypeFromProgID("KWPS.Application");
            if (type == null)
                throw new InvalidOperationException("找不到 WPS 文字。请安装或修复 WPS Office。");

            try
            {
                application = Activator.CreateInstance(type);
                documents = application.Documents;
                canQuit = !(bool)application.Visible && (int)documents.Count == 0;
                originalAlerts = (int)application.DisplayAlerts;
                originalSecurity = (int)application.AutomationSecurity;
                application.DisplayAlerts = 0;
                application.AutomationSecurity = 3; // msoAutomationSecurityForceDisable
            }
            catch
            {
                Dispose();
                throw;
            }
        }

        internal void Export(string path)
        {
            string extension = Path.GetExtension(path);
            if (!string.Equals(extension, ".doc", StringComparison.OrdinalIgnoreCase) &&
                !string.Equals(extension, ".docx", StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("请选择 .doc 或 .docx 文件。");
            if (!File.Exists(path))
                throw new FileNotFoundException("文件不存在或无法访问。");

            string folder = Path.GetDirectoryName(path);
            string stem = Path.GetFileNameWithoutExtension(path);
            string temporaryPdf = Path.Combine(folder, "~wpspdf-" + Guid.NewGuid().ToString("N") + ".pdf");
            dynamic doc = FindOpenDocument(path);
            bool openedForExport = (object)doc == null;
            try
            {
                if (openedForExport)
                {
                    // A placeholder password lets WPS reject password-protected files
                    // without asking for a password during an unattended batch.
                    doc = documents.Open(path, false, true, false, Guid.NewGuid().ToString("N"));
                }
                doc.ExportAsFixedFormat(temporaryPdf, PdfFormat, false);
                if (new FileInfo(temporaryPdf).Length == 0)
                    throw new IOException("WPS 导出了空文件。");
                Publish(temporaryPdf, folder, stem);
            }
            finally
            {
                if ((object)doc != null)
                {
                    if (openedForExport)
                        TryComCleanup(() => doc.Close(DoNotSaveChanges));
                    Marshal.ReleaseComObject((object)doc);
                }
                if (File.Exists(temporaryPdf))
                    File.Delete(temporaryPdf);
            }
        }

        private object FindOpenDocument(string path)
        {
            int count = documents.Count;
            for (int i = 1; i <= count; i++)
            {
                dynamic doc = documents.Item(i);
                bool found = false;
                try
                {
                    found = string.Equals((string)doc.FullName, path, StringComparison.OrdinalIgnoreCase);
                    if (found)
                        return doc;
                }
                finally
                {
                    if (!found)
                        Marshal.ReleaseComObject((object)doc);
                }
            }
            return null;
        }

        private static void Publish(string temporaryPdf, string folder, string stem)
        {
            for (int i = 1; ; i++)
            {
                string name = i == 1 ? stem + ".pdf" : stem + " (" + i + ").pdf";
                try
                {
                    File.Move(temporaryPdf, Path.Combine(folder, name));
                    return;
                }
                catch (IOException error)
                {
                    const int fileExists = 80;
                    const int alreadyExists = 183;
                    int win32Error = error.HResult & 0xffff;
                    if (win32Error != fileExists && win32Error != alreadyExists)
                        throw;
                }
            }
        }

        public void Dispose()
        {
            // Statically dispatched release still works after WPS has shut down.
            // Binding a dynamic invocation against its disconnected RCW can fail.
            if ((object)application == null)
                return;

            if (originalSecurity.HasValue)
                TryComCleanup(() => application.AutomationSecurity = originalSecurity.Value);
            if (originalAlerts.HasValue)
                TryComCleanup(() => application.DisplayAlerts = originalAlerts.Value);
            if (canQuit)
                TryComCleanup(() =>
                {
                    if (!(bool)application.Visible && (int)documents.Count == 0)
                        application.Quit(DoNotSaveChanges);
                });

            if ((object)documents != null)
                Marshal.ReleaseComObject((object)documents);
            Marshal.ReleaseComObject((object)application);
        }

        private static void TryComCleanup(Action cleanup)
        {
            try
            {
                cleanup();
            }
            catch (COMException)
            {
                // WPS can shut down while a batch is running.
            }
        }
    }
}
