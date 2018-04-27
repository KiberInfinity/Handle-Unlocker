# ========================================================
#
# 	Script Information
#
#	Title: 			Unlocker GUI for SysInernals Handle
#	Author:			KiberInfinity
#	Created:	      22.04.2018 - 21:34:47
#
# ========================================================

param([switch]$Elevated, [string]$target="")

function Test-Admin
{
	$currentUser = New-Object Security.Principal.WindowsPrincipal $([Security.Principal.WindowsIdentity]::GetCurrent() )
	$currentUser.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
}

if ((Test-Admin) -eq $false)
{
	if ($elevated)
	{
		# tried to elevate, did not work, aborting
	}
	else
	{
		Start-Process powershell.exe -windowstyle hidden -Verb RunAs -ArgumentList("-windowstyle hidden -ExecutionPolicy unrestricted -file ""{0}"" -target ""$($target)"" -elevated " -f ($myinvocation.MyCommand.Definition))
	}

	exit
}

#'running with full privileges'

#region ScriptForm Designer

#region Constructor

[void][System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void][System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")

#endregion

#region Post-Constructor Custom Code

#endregion

#region Form Creation
#Warning: It is recommended that changes inside this region be handled using the ScriptForm Designer.
#When working with the ScriptForm designer this region and any changes within may be overwritten.
#~~< ListForm >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$ListForm = New-Object System.Windows.Forms.Form
$ListForm.ClientSize = New-Object System.Drawing.Size(555, 357)
$ListForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
$ListForm.Text = "Lock handles"
#~~< Panel2 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Panel2 = New-Object System.Windows.Forms.Panel
$Panel2.Dock = [System.Windows.Forms.DockStyle]::Bottom
$Panel2.Location = New-Object System.Drawing.Point(0, 314)
$Panel2.Size = New-Object System.Drawing.Size(555, 43)
$Panel2.TabIndex = 2
$Panel2.Text = ""
#~~< BtnRefresh >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$BtnRefresh = New-Object System.Windows.Forms.Button
$BtnRefresh.Location = New-Object System.Drawing.Point(3, 3)
$BtnRefresh.Size = New-Object System.Drawing.Size(100, 32)
$BtnRefresh.TabIndex = 2
$BtnRefresh.Text = "Refresh"
$BtnRefresh.UseVisualStyleBackColor = $true
$BtnRefresh.add_Click({BtnRefreshClick($BtnRefresh)})
#~~< BtnUnlock >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$BtnUnlock = New-Object System.Windows.Forms.Button
$BtnUnlock.Anchor = ([System.Windows.Forms.AnchorStyles]([System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right))
$BtnUnlock.Location = New-Object System.Drawing.Point(452, 3)
$BtnUnlock.Size = New-Object System.Drawing.Size(100, 32)
$BtnUnlock.TabIndex = 1
$BtnUnlock.Text = "Unlock"
$BtnUnlock.UseVisualStyleBackColor = $true
$BtnUnlock.add_Click({BtnUnlockClick($BtnUnlock)})
#~~< BtnUnlockAll >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$BtnUnlockAll = New-Object System.Windows.Forms.Button
$BtnUnlockAll.Anchor = ([System.Windows.Forms.AnchorStyles]([System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right))
$BtnUnlockAll.Location = New-Object System.Drawing.Point(346, 3)
$BtnUnlockAll.Size = New-Object System.Drawing.Size(100, 32)
$BtnUnlockAll.TabIndex = 0
$BtnUnlockAll.Text = "Unlock All"
$BtnUnlockAll.UseVisualStyleBackColor = $true
$BtnUnlockAll.add_Click({BtnUnlockAllClick($BtnUnlockAll)})
$Panel2.Controls.Add($BtnRefresh)
$Panel2.Controls.Add($BtnUnlock)
$Panel2.Controls.Add($BtnUnlockAll)
#~~< HandlesView >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$HandlesView = New-Object System.Windows.Forms.ListView
$HandlesView.Anchor = ([System.Windows.Forms.AnchorStyles]([System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right))
$HandlesView.FullRowSelect = $true
$HandlesView.GridLines = $true
$HandlesView.Location = New-Object System.Drawing.Point(0, 0)
$HandlesView.ShowItemToolTips = $true
$HandlesView.Size = New-Object System.Drawing.Size(555, 308)
$HandlesView.TabIndex = 1
$HandlesView.Text = "ListView1"
$HandlesView.UseCompatibleStateImageBehavior = $false
$HandlesView.View = [System.Windows.Forms.View]::Details
#~~< ColumnIcon >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$ColumnIcon = New-Object System.Windows.Forms.ColumnHeader
$ColumnIcon.Text = ""
$ColumnIcon.Width = 28
#~~< ColumnName >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$ColumnName = New-Object System.Windows.Forms.ColumnHeader
$ColumnName.Text = "ProcessName"
$ColumnName.Width = 100
#~~< ColumnPID >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$ColumnPID = New-Object System.Windows.Forms.ColumnHeader
$ColumnPID.Text = "PID"
#~~< ColumnHandle >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$ColumnHandle = New-Object System.Windows.Forms.ColumnHeader
$ColumnHandle.Text = "Handle"
$ColumnHandle.Width = 100
#~~< ColumnPath >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$ColumnPath = New-Object System.Windows.Forms.ColumnHeader
$ColumnPath.Text = "Path"
$ColumnPath.Width = 230
$HandlesView.Columns.AddRange([System.Windows.Forms.ColumnHeader[]](@($ColumnIcon, $ColumnName, $ColumnPID, $ColumnHandle, $ColumnPath)))
#~~< components >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$components = New-Object System.ComponentModel.Container
#~~< ProcIcons >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$ProcIcons = New-Object System.Windows.Forms.ImageList($components)
$ProcIcons.ColorDepth = [System.Windows.Forms.ColorDepth]::Depth32Bit
$ProcIcons.ImageSize = New-Object System.Drawing.Size(16, 16)
$ProcIcons.TransparentColor = [System.Drawing.Color]::Transparent
$HandlesView.SmallImageList = $ProcIcons
$HandlesView.add_SelectedIndexChanged({ListView1SelectedIndexChanged($HandlesView)})
$ListForm.Controls.Add($Panel2)
$ListForm.Controls.Add($HandlesView)
$ListForm.add_Shown({ListFormShown($ListForm)})

#endregion

#region Custom Code

#endregion

#region Event Loop

function Main{
	[System.Windows.Forms.Application]::EnableVisualStyles()
	[System.Windows.Forms.Application]::Run($ListForm)
}

#endregion

#endregion

#region Event Handlers

try {
	Add-Type -Namespace 'System' -Name 'WinAPI' -ErrorAction 'Stop' -MemberDefinition '
	public enum ProcessDPIAwareness {
	ProcessDPIUnaware = 0,
	ProcessSystemDPIAware = 1,
	ProcessPerMonitorDPIAware = 2
	}
	[DllImport("shcore.dll")] public static extern int SetProcessDpiAwareness(ProcessDPIAwareness value);
	'
}
catch {
}
[System.WinAPI]::Proc
$processDPIAwareness = [System.WinAPI+ProcessDPIAwareness]::ProcessPerMonitorDPIAware
$processDPIAwarenessResult = [System.WinAPI]::SetProcessDPIAwareness($processDPIAwareness)

# Show an Open Folder Dialog and return the directory selected by the user.
function Read-FolderBrowserDialog([string]$Message, [string]$InitialDirectory, [switch]$NoNewFolderButton)
{
	$browseForFolderOptions = 16
	if ($NoNewFolderButton) { $browseForFolderOptions += 512 }

	$app = New-Object -ComObject Shell.Application
	$folder = $app.BrowseForFolder(0, $Message, $browseForFolderOptions, $InitialDirectory)
	if ($folder) { $selectedDirectory = $folder.Self.Path } else { $selectedDirectory = '' }
	[System.Runtime.Interopservices.Marshal]::ReleaseComObject($app) > $null
	return $selectedDirectory
}

function Get-PSScriptRoot
{
	$ScriptRoot = ""

	Try
	{
		$ScriptRoot = Get-Variable -Name PSScriptRoot -ValueOnly -ErrorAction Stop
	}
	Catch
	{
		$ScriptRoot = Split-Path $script:MyInvocation.MyCommand.Path
	}

	return $ScriptRoot
}

$script_root = Get-PSScriptRoot;
$handle_path = "$($script_root)\handle.exe"

if ($target-eq "")
{
	$target = Read-FolderBrowserDialog("Select folder for unlock")
}

$ListForm.Icon = [Drawing.Icon]::ExtractAssociatedIcon("$($script_root)/icon.ico")

function ListView1SelectedIndexChanged( $object ){

}

function UpdateHandlesList($target_path)
{
	$icos = @{ }

	$ListForm.Text = "Try to unlock: $($target)";
	$HandlesView.Items.Clear();
	$HandlesView.Enabled = 0;
	[System.Windows.Forms.Application]::DoEvents();
	$handle = & $handle_path "$($target_path)"
	foreach ($line in $handle)
	{
		#[System.Windows.Forms.Application]::DoEvents();
		$m = $line -match '^(.+?)\s+pid:\s*(\d+)\s+type:\s*(\S+)\s+([a-f0-9]+):\s*(.+)'
		if ($m)
		{
			#Write-Output $line
			#Write-Output $Matches
			$proc_id = $Matches[2] -as[int];
			#$idx = -1;
			if ($icos[$proc_id]) {
				#$idx = $icos[$proc_id];
			}
			else {
				#$icos[$pid]
				$fv = Get-Process -Id $proc_id;
				if ($fv.Path)
				{
					#
					# [System.Windows.Forms.MessageBox]::Show($fv.Path, $proc_id);
					$ico = [Drawing.Icon]::ExtractAssociatedIcon($fv.Path);
					#$ListForm.Icon = $ico;
					$ProcIcons.Images.Add($proc_id, $ico);
				}
				$icos[$proc_id] = 1;
			}
			$item = New-Object System.Windows.Forms.ListViewItem([System.String[]] ( @(
						"",
						$Matches[1],
						$Matches[2],
						"$($Matches[4]) [$($Matches[3])]",
						$Matches[5]
			) ), -1 )
			if ($ProcIcons.Images.IndexOfKey($proc_id) -ge 0)
			{
				$item.ImageKey = $proc_id;
			}
			$item.Tag = "-c " + $Matches[4] + " -y -p " + $Matches[2]
			$HandlesView.Items.AddRange([System.Windows.Forms.ListViewItem[]] ( @($item) ))
		}
		#$HandlesView.add_SelectedIndexChanged({ ListView1SelectedIndexChanged($HandlesView) })
	}
	$HandlesView.Enabled = 1;
}
function UnlockItem($item) {
	$closed = "Closed.";

	if ($item.SubItems[3].Text -eq $closed)
	{
		return;
	}
	$close_params = $item.Tag;

	$cmd = "& ""$( $handle_path )"" $( $close_params )";
	Invoke-Expression $cmd | Tee-Object -Variable handle;
	$result_type = 0;
	foreach ($line in $handle)
	{
		$m = $line -match 'Error closing handle';
		if ($m)
		{
			$result_type = 2;
		}
		else
		{
			$m = $line -match 'Handle closed\.';
			if ($m)
			{
				$result_type = 1;
			}
		}
	}
	switch($result_type)
	{
		1 {
			$item.SubItems[3].Text = $closed;
			$item.BackColor = 'ButtonFace';
		}
		2 { [System.Windows.Forms.MessageBox]::Show($handle, "Error") ; }
		0 { [System.Windows.Forms.MessageBox]::Show($handle, "Unknown result") ; }
	}
}

function BtnUnlockClick( $object ){
	#Error closing handle
	#Handle closed.
	#[System.Windows.Forms.MessageBox]::Show($HandlesView.SelectedItems.Count.toString(),"test");
	if ($HandlesView.SelectedItems.Count -gt 0) {
		for ($i = 0; $i -lt $HandlesView.SelectedItems.Count; $i++)
		{
			$item = $HandlesView.SelectedItems[$i];
			UnlockItem($item);
		}
		#UpdateHandlesList($target);
	}
}

function BtnUnlockAllClick( $object ){
	if ($HandlesView.Items.Count -gt 0)
	{
		for ($i = 0; $i -lt $HandlesView.Items.Count; $i++)
		{
			$item = $HandlesView.Items[$i];
			UnlockItem($item);
		}
		#UpdateHandlesList($target);
	}
}

function ListFormShown( $object ){
	UpdateHandlesList($target);
}

function BtnRefreshClick( $object ){
	UpdateHandlesList($target)
}

Main # This call must remain below all other event functions

#endregion

#region Script Settings
#<ScriptSettings xmlns="http://tempuri.org/ScriptSettings.xsd">
#  <ScriptPackager>
#    <process>powershell.exe</process>
#    <arguments />
#    <extractdir>%TEMP%</extractdir>
#    <files />
#    <usedefaulticon>true</usedefaulticon>
#    <showinsystray>false</showinsystray>
#    <altcreds>false</altcreds>
#    <efs>true</efs>
#    <ntfs>true</ntfs>
#    <local>false</local>
#    <abortonfail>true</abortonfail>
#    <product />
#    <version>1.0.0.1</version>
#    <versionstring />
#    <comments />
#    <company />
#    <includeinterpreter>false</includeinterpreter>
#    <forcecomregistration>false</forcecomregistration>
#    <consolemode>false</consolemode>
#    <EnableChangelog>false</EnableChangelog>
#    <AutoBackup>false</AutoBackup>
#    <snapinforce>false</snapinforce>
#    <snapinshowprogress>false</snapinshowprogress>
#    <snapinautoadd>2</snapinautoadd>
#    <snapinpermanentpath />
#    <cpumode>1</cpumode>
#    <hidepsconsole>false</hidepsconsole>
#  </ScriptPackager>
#</ScriptSettings>
#endregion