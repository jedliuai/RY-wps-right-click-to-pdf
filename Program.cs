using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Windows.Forms;

internal static class Program
{
    internal const string MenuTitle = "转换为 PDF（WPS）";

    [STAThread]
    private static void Main(string[] args)
    {
        if (args.Length == 0)
            return;

        Application.EnableVisualStyles();
        try
        {
            var result = new ConversionResult();
            RunQueue(args, result);
            result.Show();
        }
        catch (Exception error)
        {
            MessageBox.Show(error.Message, MenuTitle, MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private static void RunQueue(string[] paths, ConversionResult result)
    {
        // Explorer launches the command once per file. Atomically publish each
        // request so that only one process opens WPS and drains the queue.
        string folder = Path.Combine(AppDomain.CurrentDomain.BaseDirectory,
            "Queue-" + Process.GetCurrentProcess().SessionId);
        Directory.CreateDirectory(folder);

        string ticket = Path.Combine(folder, Guid.NewGuid().ToString("N"));
        File.WriteAllLines(ticket + ".tmp", paths);
        File.Move(ticket + ".tmp", ticket + ".job");

        using (var worker = new Mutex(false, @"Local\WpsPdfContextMenu.Conversion"))
        {
            do
            {
                if (!TryBecomeWorker(worker))
                    return;

                try
                {
                    // Give Explorer time to start the other processes in a multi-select.
                    Thread.Sleep(600);
                    List<string> batch = TakeQueuedFiles(folder);
                    while (batch.Count > 0)
                    {
                        WpsPdf.Convert(batch, result);
                        batch = TakeQueuedFiles(folder);
                    }
                }
                finally
                {
                    worker.ReleaseMutex();
                }

                // Check again after releasing the mutex. A sender may have queued
                // a job while we were still the worker, just after the final read.
            } while (Directory.GetFiles(folder, "*.job").Length > 0);
        }
    }

    private static bool TryBecomeWorker(Mutex worker)
    {
        try
        {
            return worker.WaitOne(0);
        }
        catch (AbandonedMutexException)
        {
            return true;
        }
    }

    private static List<string> TakeQueuedFiles(string folder)
    {
        var paths = new List<string>();
        foreach (string ticket in Directory.GetFiles(folder, "*.job"))
        {
            paths.AddRange(File.ReadAllLines(ticket));
            File.Delete(ticket);
        }
        return paths;
    }
}
