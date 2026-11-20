param([switch]$Uninstall)

$ErrorActionPreference = 'Stop'
$verbName = 'WpsPdfContextMenu.Convert'
$menuTitle = '转换为 PDF（WPS）'
$verbKeys = '.doc', '.docx' | ForEach-Object {
    "SystemFileAssociations\$_\shell\$verbName"
}
$installFolder = Join-Path ([Environment]::GetFolderPath('LocalApplicationData')) 'WpsPdfContextMenu'
$exePath = Join-Path $installFolder 'WpsPdfContextMenu.exe'

function Notify-Explorer {
    if (-not ('WpsPdfMenu.Shell' -as [type])) {
        Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
namespace WpsPdfMenu {
    public static class Shell {
        [DllImport("shell32.dll")]
        public static extern void SHChangeNotify(int eventId, uint flags, IntPtr item1, IntPtr item2);
    }
}
'@
    }
    [WpsPdfMenu.Shell]::SHChangeNotify(0x08000000, 0, [IntPtr]::Zero, [IntPtr]::Zero)
}

# Match the native Explorer registry view even if a 32-bit PowerShell ran setup.
$framework = 'Framework'
$view = [Microsoft.Win32.RegistryView]::Default
if ([Environment]::Is64BitOperatingSystem) {
    $framework = 'Framework64'
    $view = [Microsoft.Win32.RegistryView]::Registry64
}
$hive = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::CurrentUser, $view)
try {
    $classes = $hive.CreateSubKey('Software\Classes')
    try {
        if ($Uninstall) {
            foreach ($key in $verbKeys) {
                $classes.DeleteSubKeyTree($key, $false)
            }
            Notify-Explorer

            if ([IO.File]::Exists($exePath)) {
                [IO.File]::Delete($exePath)
            }
            Write-Host '已移除 WPS 转 PDF 右键菜单。'
            return
        }

        $wpsType = [Type]::GetTypeFromProgID('KWPS.Application')
        if (-not $wpsType) {
            throw '找不到 WPS 文字。请先安装或修复 WPS Office。'
        }

        $compiler = Join-Path $env:WINDIR "Microsoft.NET\$framework\v4.0.30319\csc.exe"
        [void][IO.Directory]::CreateDirectory($installFolder)
        $buildPath = Join-Path $installFolder ('build-' + [guid]::NewGuid().ToString('N') + '.exe')
        try {
            & $compiler /nologo /target:winexe /codepage:65001 /optimize+ "/out:$buildPath" `
                /reference:Microsoft.CSharp.dll /reference:System.Windows.Forms.dll `
                (Join-Path $PSScriptRoot 'Program.cs') (Join-Path $PSScriptRoot 'WpsPdf.cs')
            if ($LASTEXITCODE -ne 0) {
                throw '编译右键菜单程序失败。'
            }
            if ([IO.File]::Exists($exePath)) {
                [IO.File]::Replace($buildPath, $exePath, [NullString]::Value)
            } else {
                [IO.File]::Move($buildPath, $exePath)
            }
        } finally {
            if ([IO.File]::Exists($buildPath)) {
                [IO.File]::Delete($buildPath)
            }
        }

        foreach ($key in $verbKeys) {
            $verb = $classes.CreateSubKey($key)
            try {
                $verb.SetValue('MUIVerb', $menuTitle)
                $verb.SetValue('MultiSelectModel', 'Player')
                $verb.SetValue('NeverDefault', '')
                $command = $verb.CreateSubKey('command')
                try {
                    $command.SetValue('', '"' + $exePath + '" "%1"')
                } finally {
                    $command.Dispose()
                }
            } finally {
                $verb.Dispose()
            }
        }
        Notify-Explorer
        Write-Host "已安装右键菜单 $menuTitle"
    } finally {
        $classes.Dispose()
    }
} finally {
    $hive.Dispose()
}
