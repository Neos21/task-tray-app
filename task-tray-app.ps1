Add-Type -AssemblyName System.Windows.Forms;

# ================================================================================
# Task Tray App
# ================================================================================

# ���d�N�����Ă�����I������ (Mutual Exclusion) https://qiita.com/magiclib/items/cc2de9169c781642e52d
$mutex = New-Object Threading.Mutex($False, 'Global\mutex-task-tray-app');
if(!($mutex.WaitOne(0, $False))) {  # `!` = `-Not`
  Write-Host 'This App Is Already Launched, Aborted';
  $mutex.Close();
  Sleep 1;
  Exit;
}

# �萔 : �^�C�}�[�����̊Ԋu��` (�~���b) : 3���Ԋu
$timerIntervalMs = 1000 * 60 * 3;  # �f�o�b�O�p : `1000 * 3` (3�b�Ԋu)
# �萔 : ��������L�[ : �f�o�b�O���� `{PGDN}` ��������g���ƕ�����₷������
$keyToPress      = '{F15}';
# �萔 : �^�X�N�g���C�A�C�R���p Exe �t�@�C����` (�^�C�}�[ On ���EOff ���ŗp�ӂ���)
$iconExeTimerOn  = 'C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe';  # `(Get-Process -id $Pid | Select-Object -ExpandProperty Path)`
$iconExeTimerOff = 'C:\Windows\System32\cmd.exe';

# �C�x���g�n���h�����ŏ������݂����ϐ��� `script` �X�R�[�v�Ő錾���� https://github.com/yokra9/RunCat_for_Windows_on_PowerShell/blob/master/RunCatPS/src/runcat.ps1
$script:isTerminalWindowClosed = $True;  # �^�[�~�i���E�B���h�E�����Ă��邩�ۂ�
$script:isTimerEnabled         = $True;  # �^�C�}�[�����s�����ۂ�

# �f�o�b�O���O���o�͂��� https://www.tempo96.com/entry/pwsh-log
function Debug-Log($message) {
  Write-Host "[$((Get-Date).ToString('yyyy-MM-dd HH:mm:ss.fff'))] ${message}";
}

try {
  Debug-Log 'Start App';
  
  # Exe �t�@�C������A�C�R�������o��
  $iconTimerOn  = [Drawing.Icon]::ExtractAssociatedIcon($iconExeTimerOn );
  $iconTimerOff = [Drawing.Icon]::ExtractAssociatedIcon($iconExeTimerOff);
  
  # ���C���E�B���h�E�E�^�X�N�o�[���\���ɂ���
  $terminalWindow    = Add-Type -MemberDefinition '[DllImport("user32.dll")] public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);' -Name Win32Functions -PassThru;
  $windowHandle      = (Get-Process -PID $Pid).MainWindowHandle;  # �ĕ\�����邽�߂ɕϐ��ɍT���� https://stackoverflow.com/questions/5847481/showwindowasync-dont-show-hidden-window-sw-show
  $initialWindowMode = if($script:isTerminalWindowClosed) { Write-Output 0; } else { Write-Output 9; };  # https://learn.microsoft.com/ja-jp/windows/win32/api/winuser/nf-winuser-showwindow
  [void]$terminalWindow::ShowWindowAsync($windowHandle, $initialWindowMode);
  
  # �R���e�L�X�g���`����
  $applicationContext = New-Object Windows.Forms.ApplicationContext;
  
  # �^�X�N�g���C�A�C�R����\������
  $notifyIcon         = New-Object Windows.Forms.NotifyIcon;
  $initialIcon        = if($script:isTimerEnabled) { Write-Output $iconTimerOn ; } else { Write-Output $iconTimerOff ; };
  $initialIconText    = if($script:isTimerEnabled) { Write-Output 'Timer Is On'; } else { Write-Output 'Timer Is Off'; };
  $notifyIcon.Icon    = $initialIcon;
  $notifyIcon.Text    = $initialIconText;  # �c�[���`�b�v
  $notifyIcon.Visible = $True;
  
  # �^�C�}�[�������`����
  $timer = New-Object Windows.Forms.Timer;
  $timer.Add_Tick({
    Debug-Log 'Timer Tick';
    $timer.Stop();
    # �L�[���������� https://learn.microsoft.com/ja-jp/dotnet/api/system.windows.forms.sendkeys?view=windowsdesktop-6.0 https://onceuponatimeit.hatenablog.com/entry/2016/02/20/004837
    [Windows.Forms.SendKeys]::SendWait($keyToPress);
    # �C���^�[�o�����Đݒ肵�ă^�C�}�[���ĊJ����
    $timer.Interval = $timerIntervalMs;
    $timer.Start();
  });
  if($script:isTimerEnabled) {
    Debug-Log 'Initial Timer Started';
    $timer.Interval = 1;  # ����͑������s����
    $timer.Start();       # http://blog.syo-ko.com/?eid=1542
  }
  else {
    Debug-Log 'Initial Timer Is Disabled';
  }
  
  # �A�C�R���N���b�N���Ƀ^�C�}�[�� On�EOff �ؑւ���
  $notifyIcon.add_Click({
    if(!($_.Button -eq [Windows.Forms.MouseButtons]::Left)) { return; }
    Debug-Log 'Task Tray Icon Is Clicked';
    if($script:isTimerEnabled) {  # �^�C�}�[���s���Ȃ�~�߂�
      $timer.Stop();
      Debug-Log 'Timer Stopped';
    }
    else {  # �^�C�}�[��~���Ȃ瑦�����s����
      $timer.Interval = 1;
      $timer.Start();
      Debug-Log 'Timer Restarted';
    }
    # �^�C�}�[�� On�EOff ��؂�ւ���
    $script:isTimerEnabled = ! $script:isTimerEnabled;
    # �^�X�N�g���C�A�C�R���E�c�[���`�b�v��ύX����
    $icon            = if($script:isTimerEnabled) { Write-Output $iconTimerOn ; } else { Write-Output $iconTimerOff ; };
    $iconText        = if($script:isTimerEnabled) { Write-Output 'Timer Is On'; } else { Write-Output 'Timer Is Off'; };
    $notifyIcon.Icon = $icon;
    $notifyIcon.Text = $iconText;
    # �o���[���`�b�v��\������
    $balloonText = if($script:isTimerEnabled) { Write-Output 'Timer Restarted'; } else { Write-Output 'Timer Stopped'; };
    $notifyIcon.BalloonTipIcon  = [Windows.Forms.ToolTipIcon]::Info;  # https://papanda925.com/?p=1890
    $notifyIcon.BalloonTipText  = $balloonText;
    $notifyIcon.BalloonTipTitle = $balloonText;
    $notifyIcon.ShowBalloonTip(1000);
  });
  
  # �^�[�~�i���E�B���h�E��\���E��\�����郁�j���[���`����
  $menuItemTerminalWindow      = New-Object Windows.Forms.MenuItem;
  $menuItemTerminalWindow.Text = if($script:isTerminalWindowClosed) { Write-Output 'Show Terminal Window'; } else { Write-Output 'Hide Terminal Window'; };
  $menuItemTerminalWindow.add_Click({
    # �^�[�~�i���E�B���h�E�̕\���E��\����؂�ւ���
    $script:isTerminalWindowClosed = ! $script:isTerminalWindowClosed;
    $windowMode = if($script:isTerminalWindowClosed) { Write-Output 0; } else { Write-Output 9; };
    [void]$terminalWindow::ShowWindowAsync($windowHandle, $windowMode);
    # �c�[���`�b�v��ύX����
    $menuItemTerminalWindow.Text = if($script:isTerminalWindowClosed) { Write-Output 'Show Terminal Window'      ; } else { Write-Output 'Hide Terminal Window'     ; };
    $debugMessage                = if($script:isTerminalWindowClosed) { Write-Output 'Terminal Window Was Hidden'; } else { Write-Output 'Terminal Window Was Shown'; };
    Debug-Log $debugMessage;
  });
  
  # �A�v���I�����j���[���`����
  $menuItemExit      = New-Object Windows.Forms.MenuItem;
  $menuItemExit.Text = 'Exit';
  $menuItemExit.add_Click({
    Debug-Log 'Exit Menu Is Clicked';
    $applicationContext.ExitThread();
  });
  
  # �R���e�L�X�g���j���[��p�ӂ���
  $notifyIcon.ContextMenu = New-Object Windows.Forms.ContextMenu;
  $notifyIcon.contextMenu.MenuItems.AddRange($menuItemTerminalWindow);
  $notifyIcon.contextMenu.MenuItems.AddRange($menuItemExit);
  
  # �A�v�����N������
  Debug-Log 'Launch App';
  [void][Windows.Forms.Application]::Run($applicationContext);
  
  # �I������
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
  #Sleep 3;  # �f�o�b�O�p : �I�������̊m�F�p
}
