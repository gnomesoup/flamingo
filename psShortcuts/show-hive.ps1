
function Show-Process($Process, [Switch]$Maximize)
{
  $sig = '
    [DllImport("user32.dll")] public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);
    [DllImport("user32.dll")] public static extern int SetForegroundWindow(IntPtr hwnd);
  '
  
  if ($Maximize) { $Mode = 3 } else { $Mode = 4 }
  $type = Add-Type -MemberDefinition $sig -Name WindowAPI -PassThru
  $hwnd = $process.MainWindowHandle
  $null = $type::ShowWindowAsync($hwnd, $Mode)
  $null = $type::SetForegroundWindow($hwnd) 
}

$rvt = Get-Process -Name "HIVE CMS" -ErrorAction Ignore | Where-Object {
    $_.MainWindowTitle -like "HIVE*"
}

if ($rvt) {
    Show-Process($rvt)
}
else {
    Start-Process -FilePath 'C:\Program Files\CTC Software\HIVE CMS\HIVE CMS.exe'
}

