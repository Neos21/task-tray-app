Add-Type -AssemblyName System.Windows.Forms;

# ================================================================================
# Task Tray App
# ================================================================================

# 多重起動していたら終了する (Mutual Exclusion) https://qiita.com/magiclib/items/cc2de9169c781642e52d
$mutex = New-Object Threading.Mutex($False, 'Global\mutex-task-tray-app');
if(!($mutex.WaitOne(0, $False))) {  # `!` = `-Not`
  Write-Host 'This App Is Already Launched, Aborted';
  $mutex.Close();
  Sleep 1;
  Exit;
}

# 定数 : タイマー処理の間隔定義 (ミリ秒) : 3分間隔
$timerIntervalMs = 1000 * 60 * 3;  # デバッグ用 : `1000 * 3` (3秒間隔)
# 定数 : 押下するキー : デバッグ時は `{PGDN}` あたりを使うと分かりやすいかと
$keyToPress      = '{F15}';
# 定数 : タスクトレイアイコン用 Exe ファイル定義 (タイマー On 時・Off 時で用意する)
$iconExeTimerOn  = 'C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe';  # `(Get-Process -id $PID | Select-Object -ExpandProperty Path)`
$iconExeTimerOff = 'C:\Windows\System32\cmd.exe';

# イベントハンドラ内で書き込みたい変数は `script` スコープで宣言する https://github.com/yokra9/RunCat_for_Windows_on_PowerShell/blob/master/RunCatPS/src/runcat.ps1
$script:isTerminalWindowClosed = $True;  # ターミナルウィンドウが閉じているか否か
$script:isTimerEnabled         = $True;  # タイマーが実行中か否か

# デバッグログを出力する https://www.tempo96.com/entry/pwsh-log
function Debug-Log($message) {
  Write-Host "[$((Get-Date).ToString('yyyy-MM-dd HH:mm:ss.fff'))] ${message}";
}

try {
  Debug-Log 'Start App';
  
  # Exe ファイルからアイコンを取り出す
  $iconTimerOn  = [Drawing.Icon]::ExtractAssociatedIcon($iconExeTimerOn );
  $iconTimerOff = [Drawing.Icon]::ExtractAssociatedIcon($iconExeTimerOff);
  
  # メインウィンドウ・タスクバーを非表示にする
  $terminalWindow    = Add-Type -MemberDefinition '[DllImport("user32.dll")] public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);' -Name Win32Functions -PassThru;
  $windowHandle      = (Get-Process -PID $Pid).MainWindowHandle;  # 再表示するために変数に控える https://stackoverflow.com/questions/5847481/showwindowasync-dont-show-hidden-window-sw-show
  $initialWindowMode = if($script:isTerminalWindowClosed) { Write-Output 0; } else { Write-Output 9; };  # https://learn.microsoft.com/ja-jp/windows/win32/api/winuser/nf-winuser-showwindow
  [void]$terminalWindow::ShowWindowAsync($windowHandle, $initialWindowMode);
  
  # コンテキストを定義する
  $applicationContext = New-Object Windows.Forms.ApplicationContext;
  
  # タスクトレイアイコンを表示する
  $notifyIcon         = New-Object Windows.Forms.NotifyIcon;
  $initialIcon        = if($script:isTimerEnabled) { Write-Output $iconTimerOn ; } else { Write-Output $iconTimerOff ; };
  $initialIconText    = if($script:isTimerEnabled) { Write-Output 'Timer Is On'; } else { Write-Output 'Timer Is Off'; };
  $notifyIcon.Icon    = $initialIcon;
  $notifyIcon.Text    = $initialIconText;  # ツールチップ
  $notifyIcon.Visible = $True;
  
  # タイマー処理を定義する
  $timer = New-Object Windows.Forms.Timer;
  $timer.Add_Tick({
    Debug-Log 'Timer Tick';
    $timer.Stop();
    # キーを押下する https://learn.microsoft.com/ja-jp/dotnet/api/system.windows.forms.sendkeys?view=windowsdesktop-6.0 https://onceuponatimeit.hatenablog.com/entry/2016/02/20/004837
    [Windows.Forms.SendKeys]::SendWait($keyToPress);
    # インターバルを再設定してタイマーを再開する
    $timer.Interval = $timerIntervalMs;
    $timer.Start();
  });
  if($script:isTimerEnabled) {
    Debug-Log 'Initial Timer Started';
    $timer.Interval = 1;  # 初回は即時実行する
    $timer.Start();       # http://blog.syo-ko.com/?eid=1542
  }
  else {
    Debug-Log 'Initial Timer Is Disabled';
  }
  
  # アイコンクリック時にタイマーを On・Off 切替する
  $notifyIcon.add_Click({
    if(!($_.Button -eq [Windows.Forms.MouseButtons]::Left)) { return; }
    Debug-Log 'Task Tray Icon Is Clicked';
    if($script:isTimerEnabled) {  # タイマー実行中なら止める
      $timer.Stop();
      Debug-Log 'Timer Stopped';
    }
    else {  # タイマー停止中なら即時実行する
      $timer.Interval = 1;
      $timer.Start();
      Debug-Log 'Timer Restarted';
    }
    # タイマーの On・Off を切り替える
    $script:isTimerEnabled = ! $script:isTimerEnabled;
    # タスクトレイアイコン・ツールチップを変更する
    $icon            = if($script:isTimerEnabled) { Write-Output $iconTimerOn ; } else { Write-Output $iconTimerOff ; };
    $iconText        = if($script:isTimerEnabled) { Write-Output 'Timer Is On'; } else { Write-Output 'Timer Is Off'; };
    $notifyIcon.Icon = $icon;
    $notifyIcon.Text = $iconText;
    # バルーンチップを表示する
    $balloonText = if($script:isTimerEnabled) { Write-Output 'Timer Restarted'; } else { Write-Output 'Timer Stopped'; };
    $notifyIcon.BalloonTipIcon  = [Windows.Forms.ToolTipIcon]::Info;  # https://papanda925.com/?p=1890
    $notifyIcon.BalloonTipText  = $balloonText;
    $notifyIcon.BalloonTipTitle = $balloonText;
    $notifyIcon.ShowBalloonTip(1000);
  });
  
  # ターミナルウィンドウを表示・非表示するメニューを定義する
  $menuItemTerminalWindow      = New-Object Windows.Forms.MenuItem;
  $menuItemTerminalWindow.Text = if($script:isTerminalWindowClosed) { Write-Output 'Show Terminal Window'; } else { Write-Output 'Hide Terminal Window'; };
  $menuItemTerminalWindow.add_Click({
    # ターミナルウィンドウの表示・非表示を切り替える
    $script:isTerminalWindowClosed = ! $script:isTerminalWindowClosed;
    $windowMode = if($script:isTerminalWindowClosed) { Write-Output 0; } else { Write-Output 9; };
    [void]$terminalWindow::ShowWindowAsync($windowHandle, $windowMode);
    # ツールチップを変更する
    $menuItemTerminalWindow.Text = if($script:isTerminalWindowClosed) { Write-Output 'Show Terminal Window'      ; } else { Write-Output 'Hide Terminal Window'     ; };
    $debugMessage                = if($script:isTerminalWindowClosed) { Write-Output 'Terminal Window Was Hidden'; } else { Write-Output 'Terminal Window Was Shown'; };
    Debug-Log $debugMessage;
  });
  
  # アプリ終了メニューを定義する
  $menuItemExit      = New-Object Windows.Forms.MenuItem;
  $menuItemExit.Text = 'Exit';
  $menuItemExit.add_Click({
    Debug-Log 'Exit Menu Is Clicked';
    $applicationContext.ExitThread();
  });
  
  # コンテキストメニューを用意する
  $notifyIcon.ContextMenu = New-Object Windows.Forms.ContextMenu;
  $notifyIcon.contextMenu.MenuItems.AddRange($menuItemTerminalWindow);
  $notifyIcon.contextMenu.MenuItems.AddRange($menuItemExit);
  
  # アプリを起動する
  Debug-Log 'Launch App';
  [void][Windows.Forms.Application]::Run($applicationContext);
  
  # 終了する
  Debug-Log 'Exiting...';
  $notifyIcon.Visible = $False;
}
catch {
  Debug-Log 'Catch : An Error Has Occurred';
  [Windows.Forms.MessageBox]::Show('An Error Has Occurred. Aborted', 'Error');  # https://gist.github.com/esperecyan/648b4831c6b65ea347abdf045451eb93
}
finally {
  Debug-Log 'Finally : Close Mutex';
  $mutex.ReleaseMutex();
  $mutex.Close();
  Debug-Log 'End';
  #Sleep 3;  # デバッグ用 : 終了処理の確認用
}
