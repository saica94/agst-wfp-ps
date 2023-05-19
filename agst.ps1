Set-ExecutionPolicy RemoteSigned -Scope Process -Force
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Windows.Forms, System.Drawing

# 管理者権限の確認
$currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
$bool_admin = $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

$ScriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
Write-Output $ScriptPath
$AssemblyLocation = Join-Path -Path $ScriptPath -ChildPath assembly\
Write-Output $AssemblyLocation
foreach ($Assembly in (Get-ChildItem $AssemblyLocation -Filter *.dll)) {
  [System.Reflection.Assembly]::LoadFrom($Assembly.FullName)
}


if (-not $bool_admin) {
  $msgBoxInput = [System.Windows.Forms.MessageBox]::Show("管理者権限で起動されていない為、実行できません。`r`n管理者権限で実行しますか？", "確認", "YesNo", "Question", "Button2")
  if ($msgBoxInput -eq "Yes") {
    Start-Process "C:\Program Files\PowerShell\7\pwsh.exe" $MyInvocation.MyCommand.Path -Verb runAs
  }
  exit
}

# マウス加速関連
$Mouse_acc_Path = "HKCU:\Control Panel\Mouse"
$Mouse_acc_Key1 = "MouseSpeed"
$Mouse_acc_Key2 = "MouseThreshold1"
$Mouse_acc_Key3 = "MouseThreshold2"
# タスクバー関連
$TaskBar_Path = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
$TaskBar_Size_Key1 = "TaskbarSi"
$TaskIcon_Position_Key1 = "TaskbarAl"
$TaskBar_ClassicMode_Key1 = "Start_ShowClassicMode"
$TaskBar_Search_Disabled_Path = "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Search"
$TaskBar_Search_Disabled_Key1 = "SearchboxTaskbarMode"
$TaskBar_Widged_Key1 = "Taskbarda"
# XBox Game bar 関連
$XBox_Game_bar_Path = "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\GameDVR"
$XBox_Game_bar_Key1 = "AppCaptureEnabled"
# 固定キー関連
$StickyKeys_Path = "HKCU:\Control Panel\Accessibility\StickyKeys"
$StickyKeys_Key1 = "Flags"
# Windows 広告関連
$AdvertisingInfo_Path = "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\AdvertisingInfo"
$AdvertisingInfo_Key1 = "Enabled"
$SyncProviderNotif_Key1 = "ShowSyncProviderNotifications"
$ContentDeliver_Path = "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\ContentDeliveryManager"
$ContentDeliver_Key1 = "SubscribedContent-338388Enabled"
$ContentDeliver_Key2 = "SubscribedContent-310093Enabled"
$ContentDeliver_Key3 = "SubscribedContent-338389Enabled"
$LockScreenOverlay_Key1 = "RotatingLockScreenOverlayEnabled"
$UserProfileEngagement_Path = "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\UserProfileEngagement"
$UserProfileEngagement_Key1 = "ScoobeSystemSettingEnabled"
$Privacy_Path = "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Privacy"
$Privacy_Key1 = "TailoredExperiencesWithDiagnosticDataEnabled"
# ウィンドウのスナップ関連
$Window_snap_Path = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
$Window_snap_key1 = "EnableSnapBar"

function Run_Click(){
  $acc_off = $OpWindow.FindName("acc_off")
  $LARGE = $Opwindow.FindName("LARGE")
  $MEDIUM = $Opwindow.FindName("MEDIUM")
  $Start_win10 = $OpWindow.FindName("start_Win10")
  $xbox_enable = $OpWindow.FindName("xbox_enable")
  $Ads_enable = $OpWindow.FindName("Ads_enable")
  $search_show = $OpWindow.FindName("search_show")
  $widget_show = $OpWindow.FindName("widget_show")
  $snap_enable = $OpWindow.FindName("window_snap_enable")
  $context_win10 = $OpWindow.FindName("context_win10")
  $task_left = $OpWindow.FindName("task_left")
  $static_enable = $OpWindow.FindName("static_enable")
  $pwr_on = $OpWindow.FindName("pwr_on")
  $TextBox = $MainWindow.FindName("Output")
  $chrome = $MainWindow.FindName("Chrome")


  if($chrome.IsChecked) {
    # Google ChromeのDL＆インストール
    $TextBox.AppendText("Google Chrome のダウンロードを行います。`r`n")
    $Path = $env:TEMP
    $Installer = "chrome_installer.exe"
    Invoke-WebRequest "https://dl.google.com/tag/s/appguid%3D%7B8A69D345-D564-463C-AFF1-A69D9E530F96%7D%26browser%3D0%26usagestats%3D1%26appname%3DGoogle%2520Chrome%26needsadmin%3Dprefers%26brand%3DGTPM/update2/installers/ChromeSetup.exe" -OutFile $Path\$Installer
    $TextBox.AppendText("Google Chrome のインストールを行います。`r`n")
    Start-Process -FilePath $Path\$Installer -Args "/silent /install" -Verb RunAs -Wait
    $TextBox.AppendText("インストーラーの削除を行います。`r`n")
    Remove-Item $Path\$Installer
  }

  if($acc_off.IsChecked){
    $TextBox.AppendText("マウス加速を無効化します`r`n")
    Set-ItemProperty $Mouse_acc_Path -name $Mouse_acc_Key1 -Value "0" -Force
    Set-ItemProperty $Mouse_acc_Path -name $Mouse_acc_Key2 -Value "0" -Force
    Set-ItemProperty $Mouse_acc_Path -name $Mouse_acc_Key3 -Value "0" -Force
  } else {
    $TextBox.AppendText("マウス加速を有効化します。`r`n")
    Set-ItemProperty $Mouse_acc_Path -name $Mouse_acc_Key1 -Value "1" -Force
    Set-ItemProperty $Mouse_acc_Path -name $Mouse_acc_Key2 -Value "1" -Force
    Set-ItemProperty $Mouse_acc_Path -name $Mouse_acc_Key3 -Value "1" -Force
  }

  if($LARGE.IsChecked){
    $TextBox.AppendText("タスクバーのサイズを`"大`"に変更します。`r`n")
    Set-ItemProperty $TaskBar_Path -Name $TaskBar_Size_Key1 -Value "2" -Force
  } elseif($MEDIUM.IsChecked){
    $TextBox.AppendText("タスクバーのサイズを`"中`"に変更します。`r`n")
    Set-ItemProperty $TaskBar_Path -Name $TaskBar_Size_Key1 -Value "1" -Force
  } else {
    $TextBox.AppendText("タスクバーのサイズを`"小`"に変更します。`r`n")
    Set-ItemProperty $TaskBar_Path -Name $TaskBar_Size_Key1 -Value "0" -Force
  }

  if($start_win10.IsChecked) {
    $TextBox.AppendText("Windows10仕様のスタートメニューに変更します。`r`n")
    Set-ItemProperty $TaskBar_Path -Name $TaskBar_ClassicMode_Key1 -Value "0" -Force
  } else {
    $TextBox.AppendText("Windows11仕様のスタートメニューに変更します。`r`n")
    Set-ItemProperty $TaskBar_Path -Name $TaskBar_ClassicMode_Key1 -Value "1" -Force
  }

  if($xbox_enable.IsChecked) {
    $TextBox.AppendText("XBOX Game bar を有効化します。`r`n")
    Set-ItemProperty $XBox_Game_bar_Path -Name $Xbox_Game_bar_Key1 -Value "1" -Force
  } else {
    $TextBox.AppendText("XBOX Game bar を無効化します。`r`n")
    Set-ItemProperty $XBox_Game_bar_Path -Name $Xbox_Game_bar_Key1 -Value "0" -Force
  }

  if($Ads_enable.IsChecked) {
    $TextBox.AppendText("Windows 広告を有効にします。`r`n")
    Set-ItemProperty $AdvertisingInfo_Path -Name $AdvertisingInfo_Key1 -Value "0" -Force
    Set-ItemProperty $AdvertisingInfo_Path -Name $SyncProviderNotif_Key1 -Value "0" -Force
    Set-ItemProperty $ContentDeliver_Path -Name $ContentDeliver_Key1 -Value "0" -Force
    Set-ItemProperty $ContentDeliver_Path -Name $ContentDeliver_Key2 -Value "0" -Force
    Set-ItemProperty $ContentDeliver_Path -Name $ContentDeliver_Key3 -Value "0" -Force
    Set-ItemProperty $ContentDeliver_Path -Name $LockScreenOverlay_Key1 -Value "0" -Force
    Set-ItemProperty $UserProfileEngagement_Path -Name $UserProfileEngagement_Key1 -Value "0" -Force
    Set-ItemProperty $Privacy_Path -Name $Privacy_Key1 -Value "0" -Force
  } else {
    $TextBox.AppendText("Windows 広告を無効にします。`r`n")
    Set-ItemProperty $AdvertisingInfo_Path -Name $AdvertisingInfo_Key1 -Value "1" -Force
    Set-ItemProperty $AdvertisingInfo_Path -Name $SyncProviderNotif_Key1 -Value "1" -Force
    Set-ItemProperty $ContentDeliver_Path -Name $ContentDeliver_Key1 -Value "1" -Force
    Set-ItemProperty $ContentDeliver_Path -Name $ContentDeliver_Key2 -Value "1" -Force
    Set-ItemProperty $ContentDeliver_Path -Name $ContentDeliver_Key3 -Value "1" -Force
    Set-ItemProperty $ContentDeliver_Path -Name $LockScreenOverlay_Key1 -Value "1" -Force
    Set-ItemProperty $UserProfileEngagement_Path -Name $UserProfileEngagement_Key1 -Value "1" -Force
    Set-ItemProperty $Privacy_Path -Name $Privacy_Key1 -Value "1" -Force
  }

  if($search_show.IsChecked){
    $TextBox.AppendText("タスクバーの検索アイコンを表示します。`r`n")
    Set-ItemProperty $TaskBar_Search_Disabled_Path -Name $TaskBar_Search_Disabled_Key1 -Value "1" -Force
  } else {
    $TextBox.AppendText("タスクバーの検索アイコンを非表示にします。`r`n")
    Set-ItemProperty $TaskBar_Search_Disabled_Path -Name $TaskBar_Search_Disabled_Key1 -Value "0" -Force
  }

  if($widget_show.IsChecked){
    $TextBox.AppendText("タスクバーのウィジェットアイコンを表示します。`r`n")
    Set-ItemProperty $TaskBar_Path -Name $TaskBar_Widged_Key1 -Value "1" -Force
  } else {
    $TextBox.AppendText("タスクバーのウィジェットアイコンを非表示にします。`r`n")
    Set-ItemProperty $TaskBar_Path -Name $TaskBar_Widged_Key1 -Value "0" -Force
  }

  if($snap_enable.IsChecked){
    $TextBox.AppendText("ウィンドウスナップの一部機能を有効化します。`r`n")
    Set-ItemProperty $Window_snap_Path -Name $Window_snap_key1 -Value "1" -Force
  } else {
    $TextBox.AppendText("ウィンドウスナップの一部機能を無効化します。`r`n")
    Set-ItemProperty $Window_snap_Path -Name $Window_snap_key1 -Value "0" -Force
  }

  if($context_win10.IsChecked){
    $TextBox.AppendText("コンテキストメニューをWindows10仕様に変更します。`r`n")
    reg add “HKCU\Software\Classes\CLSID\{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}\InprocServer32” /f /ve
  } else {
    $TextBox.AppendText("コンテキストメニューをWindows11仕様に変更します。`r`n")
    $context_reg = (Test-Path -LiteralPath "HKCU\Software\Classes\CLSID\{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}")
    if ($context_reg -eq $True) {
      Remove-Item -LiteralPath "HKCU\Software\Classes\CLSID\{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}" -Recurse
    }
  }

    $TextBox.AppendText("タスクバーの設定を反映させる為にexplorer.exeを再起動します。`r`n")
    $nid = (Get-Process explorer).id
    Stop-Process -Id $nid
    Wait-Process -Id $nid
    Start-Process "explorer.exe"

  if($task_left.IsChecked){
    $TextBox.AppendText("タスクアイコンを左揃えにします。`r`n")
    Set-ItemProperty $TaskBar_Path -Name $TaskIcon_Position_Key1 -Value "0" -Force
  } else {
    $TextBox.AppendText("タスクアイコンを中央揃えにします。`r`n")
    Set-ItemProperty $TaskBar_Path -Name $TaskIcon_Position_Key1 -Value "1" -Force
  }

  if($static_enable.IsChecked){
    $TextBox.AppendText("固定キー機能を有効化します。`r`n")
    Set-ITemProperty $StickyKeys_Path -Name $StickyKeys_Key1 -Value "1" -Force
  } else {
    $TextBox.AppendText("固定キー機能を無効化します。`r`n")
      Set-ITemProperty $StickyKeys_Path -Name $StickyKeys_Key1 -Value "0" -Force
  }

  if($pwr_on.IsChecked){
    $TextBox.AppendText("電源プランを変更をします。`r`n")
    $TextBox.AppendText("一部のRyzen CPUではこの処理はスキップされます。`r`n")
    # 電源プランの変更
    $ReturnData = New-Object PSObject | Select-Object CPUName
    $Win32_Processor = Get-WmiObject Win32_Processor
    $ReturnData.CPUName = @($Win32_Processor.Name)[0]
    if ($ReturnData.CPUName -Like "AMD Ryzen * 5**0*") {
      # 休止状態を無効にし、高速スタートアップを自動的に無効化させる
      powercfg /hibernate off
      powercfg -setacvalueindex scheme_balanced sub_buttons pbuttonaction 3
    } else {
      # 電源プランを高パフォーマンスに変更
      powercfg /setactive scheme_min
      powercfg /hibernate off
      powercfg -setacvalueindex scheme_min sub_buttons pbuttonaction 3
    }
  }

  $TextBox.AppendText("全ての処理が完了しました。`r`n")
  $TextBox.AppendText("このウィンドウを閉じても構いません。`r`n")
  $TextBox.AppendText("なお、一部の設定を正しく反映させるためにPCの再起動を行ってください。")

}

# xaml読み込み
try {
  [xml]$MainWindowxaml = ([System.IO.File]::ReadAllText("$ScriptPath\xaml\MainWindow.xaml"))
  $MainWindowReader = New-Object System.Xml.XmlNodeReader $MainWindowxaml
  $MainWindow = [Windows.Markup.XamlReader]::Load($MainWindowReader)
  [xml]$OptionWindowxaml = ([System.IO.File]::ReadAllText("$ScriptPath\xaml\OptionWindow.xaml"))
  $OpReader = New-Object System.Xml.XmlNodeReader $OptionWindowxaml
  $OpWindow = [Windows.Markup.XamlReader]::Load($OpReader)
  [xml]$HelpWindowxaml = ([System.IO.File]::ReadAllText("$ScriptPath\xaml\Help.xaml"))
  $hlpReader = New-Object System.Xml.XmlNodeReader $HelpWindowxaml
  $HelpWindow = [Windows.Markup.XamlReader]::Load($hlpReader)
}
catch {
  Write-Error "Error bulding Xaml data.`n$_"
  exit
}


$MainWindowxaml.SelectNodes("//*[@*[contains(translate(name(.),'n','N'),'Name')]]") | ForEach-Object {
  New-Variable  -Name $_.Name -Value $MainWindow.FindName($_.Name) -Force
}
$OptionWindowxaml.SelectNodes("//*[@*[contains(translate(name(.),'n','N'),'Name')]]") | ForEach-Object {
  New-Variable  -Name $_.Name -Value $OpWindow.FindName($_.Name) -Force
}
$HelpWindowxaml.SelectNodes("//*[@*[contains(translate(name(.),'n','N'),'Name')]]") | ForEach-Object {
  New-Variable  -Name $_.Name -Value $HelpWindow.FindName($_.Name) -Force
}

$Option = $MainWindow.FindName("Options")
$Option.add_Click({ $OpWindow.ShowDialog() })
$Apply = $OpWindow.FindName("Apply")
$Apply.add_Click({ $OpWindow.Hide() })
$AutoRun = $MainWindow.FindName("Run")
$AutoRun.add_Click({ Run_Click })
$Help.Source = "$PSScriptRoot\Image\Help.png"
$hlp.add_Click({ $HelpWindow.ShowDialog()})
$img_Twitter.Source = "$PSScriptRoot\Image\Twitter.png"
$img_github.Source = "$PSScriptRoot\Image\github.png"
$img_Discord.Source = "$PSScriptRoot\Image\discord.png"
$btn_Twitter.add_Click({ (Start-Process "https://twitter.com/daikiy6111") })
$btn_GitHub.add_Click({ (Start-Process "https://github.com/saica94/agst/") })
$btn_ok.add_Click({ $HelpWindow.Hide()})
Write-Output $PSScriptRoot
$MainWindow.ShowDialog()