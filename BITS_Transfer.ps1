Add-Type -AssemblyName System.Windows.Forms

# Hide PowerShell Console
Add-Type -Name Window -Namespace Console -MemberDefinition '
[DllImport("Kernel32.dll")]
public static extern IntPtr GetConsoleWindow();
[DllImport("user32.dll")]
public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);
'
$consolePtr = [Console.Window]::GetConsoleWindow()
[Console.Window]::ShowWindow($consolePtr, 0)


$Form = New-Object system.Windows.Forms.Form 
$Form.Text = "BITS Tool"
$Form.TopMost = $true
$Form.FormBorderStyle = 'Fixed3D'
$Form.MaximizeBox = $false
$Form.Width = 535
$Form.Height = 280
$Form.add_FormClosing({[Console.Window]::ShowWindow($consolePtr, 1)})


$sourcelabel = New-Object system.windows.Forms.Label 
$sourcelabel.Text = "Source:"
$sourcelabel.AutoSize = $true
$sourcelabel.Width = 25
$sourcelabel.Height = 10
$sourcelabel.location = new-object system.drawing.point(10,10)
$sourcelabel.Font = "Microsoft Sans Serif,10"
$Form.controls.Add($sourcelabel) 

$destlabel = New-Object system.windows.Forms.Label 
$destlabel.Text = "Dest:"
$destlabel.AutoSize = $true
$destlabel.Width = 25
$destlabel.Height = 10
$destlabel.location = new-object system.drawing.point(10,40)
$destlabel.Font = "Microsoft Sans Serif,10"
$Form.controls.Add($destlabel) 


$sourcetext = New-Object system.windows.Forms.TextBox 
$sourcetext.Width = 350
$sourcetext.Height = 20
$sourcetext.location = new-object system.drawing.point(70,10)
$sourcetext.Font = "Microsoft Sans Serif,10"
$Form.controls.Add($sourcetext) 

$desttext = New-Object system.windows.Forms.TextBox 
$desttext.Width = 350
$desttext.Height = 20
$desttext.location = new-object system.drawing.point(70,40)
$desttext.Font = "Microsoft Sans Serif,10"
$Form.controls.Add($desttext) 


$sourcebutton = New-Object system.windows.Forms.Button 
$sourcebutton.Text = "Source"
$sourcebutton.Width = 70
$sourcebutton.Height = 20
$sourcebutton.location = new-object system.drawing.point(430,10)
$sourcebutton.Font = "Microsoft Sans Serif,10"
$Form.controls.Add($sourcebutton) 

$destbutton = New-Object system.windows.Forms.Button 
$destbutton.Text = "Dest"
$destbutton.Width = 70
$destbutton.Height = 20
$destbutton.location = new-object system.drawing.point(430,40)
$destbutton.Font = "Microsoft Sans Serif,10"
$Form.controls.Add($destbutton) 

$startbutton = New-Object system.windows.Forms.Button 
$startbutton.Text = "Start Transfer"
$startbutton.Width = 150
$startbutton.Height = 25
$startbutton.location = new-object system.drawing.point(10,70)
$startbutton.Font = "Microsoft Sans Serif,10"
$Form.controls.Add($startbutton) 

$getbutton = New-Object system.windows.Forms.Button 
$getbutton.Text = "Get Status"
$getbutton.enabled = $false
$getbutton.Width = 150
$getbutton.Height = 25
$getbutton.location = new-object system.drawing.point(180,70)
$getbutton.Font = "Microsoft Sans Serif,10"
$Form.controls.Add($getbutton) 

$finishbutton = New-Object system.windows.Forms.Button 
$finishbutton.Text = "Finish Transfer"
$finishbutton.enabled = $false
$finishbutton.Width = 150
$finishbutton.Height = 25
$finishbutton.location = new-object system.drawing.point(350,70)
$finishbutton.Font = "Microsoft Sans Serif,10"
$Form.controls.Add($finishbutton) 


$statusbox = New-Object system.windows.Forms.Label 
$statusbox.Text = "Awaiting transfer.."
$statusbox.AutoSize = $true
$statusbox.Width = 490
$statusbox.Height = 155
$statusbox.location = new-object system.drawing.point(10,104)
$statusbox.Font = "Consolas,10"
$Form.controls.Add($statusbox) 




$sourcebutton.Add_Click(
	{    
	$sourcetext.text=Get-OpenFileName
	}
)

$destbutton.Add_Click(
	{    
	$desttext.text=Get-SaveFileName
	}
)

$startbutton.Add_Click(
	{
		$source=$sourcetext.text
		$dest=$desttext.text
		$sourcebutton.enabled = $false
		$destbutton.enabled = $false
		$startbutton.enabled = $false
		$getbutton.enabled = $true
		$finishbutton.enabled = $false
		$desttext.enabled = $false
		$sourcetext.enabled = $false
		$Job = Start-BitsTransfer -Source "$source" -Destination "$dest" -Asynchronous -Priority low -DisplayName "scriptedbits"
		sleep 2;
		$BITSInfo = get-bitstransfer -Name "scriptedbits"
		$enlapsed=new-timespan -start $BITSInfo.creationtime -end $BITSInfo.ModificationTime
		$eta=(get-date).addseconds([math]::Round((($enlapsed.totalseconds * ($BITSInfo.bytestotal / $BITSInfo.bytestransferred)) - $enlapsed.totalseconds),0))
		$calcedbytes=[math]::Round(($BITSInfo.bytestotal - $BITSInfo.bytestransferred) / 1024 / $enlapsed.totalseconds,0)
		$lastbytes=$BITSInfo.bytestransferred
		$statusbox.Text=($BITSInfo `
			| Format-List `
				@{Name="Owner           ";Expression={$_.owneraccount}},
				@{Name="Creation Time   ";Expression={$_.creationtime}},
				@{Name="State           ";Expression={$_.jobstate}},
				@{Name="Transferred MB  ";Expression={[math]::Round($_.bytestransferred / 1MB,2)}},
				@{Name="Total MB        ";Expression={[math]::Round($_.bytestotal / 1MB,2)}},
				@{Name="Completed %     ";Expression={[math]::Round($_.bytestransferred / $_.bytestotal * 100,2)}},
				@{Name="Speed (KB/s)    ";Expression={$calcedbytes}},
				@{Name="Estimated ETA   ";Expression={$eta}}	`
			| Out-String).Trim()
	}
)

$getbutton.Add_Click(
	{
	# while (($Job.JobState -eq "Transferring") -or ($Job.JobState -eq "Connecting"))
	$BITSInfo = get-bitstransfer -Name "scriptedbits"
	if ($BITSInfo.jobstate -eq "Transferred") {
		$startbutton.enabled = $false
		$getbutton.enabled = $false
		$finishbutton.enabled = $true
	}
	$enlapsed=new-timespan -start $BITSInfo.creationtime -end $BITSInfo.ModificationTime
	$eta=(get-date).addseconds([math]::Round((($enlapsed.totalseconds * ($BITSInfo.bytestotal / $BITSInfo.bytestransferred)) - $enlapsed.totalseconds),0))
	$calcedbytes=[math]::Round($BITSInfo.bytestransferred / 1024 / $enlapsed.seconds,0)
	$statusbox.Text=($BITSInfo `
		| Format-List `
			@{Name="Owner           ";Expression={$_.owneraccount}},
			@{Name="Creation Time   ";Expression={$_.creationtime}},
			@{Name="State           ";Expression={$_.jobstate}},
			@{Name="Transferred MB  ";Expression={[math]::Round($_.bytestransferred / 1MB,2)}},
			@{Name="Total MB        ";Expression={[math]::Round($_.bytestotal / 1MB,2)}},
			@{Name="Completed %     ";Expression={[math]::Round($_.bytestransferred / $_.bytestotal * 100,2)}},
			@{Name="Speed (KB/s)    ";Expression={$calcedbytes}},
			@{Name="Estimated ETA   ";Expression={$eta}}	`
		| Out-String).Trim()
	}
)

$finishbutton.Add_Click(
	{
	get-bitstransfer -Name "scriptedbits" | Complete-BitsTransfer
	$sourcebutton.enabled = $true
	$destbutton.enabled = $true
	$startbutton.enabled = $true
	$finishbutton.enabled = $false
	$getbutton.enabled = $false
	$desttext.enabled = $true
	$sourcetext.enabled = $true
	$statusbox.Text="`n`n`n! ! ! ! ! ! ! ! !      Transfer Complete.      ! ! ! ! ! ! ! ! !"
	}
)



Function Get-OpenFileName($initialDirectory)
{   
 [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") |
 Out-Null

 $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
 $OpenFileDialog.initialDirectory = $initialDirectory
 $OpenFileDialog.filter = "All files (*.*)| *.*"
 $OpenFileDialog.ShowDialog() | Out-Null
 $global:partfilename = $OpenFileDialog.safefilename 
 $OpenFileDialog.filename
} 

Function Get-SaveFileName($initialDirectory)
{   
 [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") |
 Out-Null

 $OpenFileDialog = New-Object System.Windows.Forms.SaveFileDialog
 $OpenFileDialog.initialDirectory = $initialDirectory
 $OpenFileDialog.filename = $global:partfilename
 $OpenFileDialog.filter = "All files (*.*)| *.*"
 $OpenFileDialog.ShowDialog() | Out-Null
 $OpenFileDialog.filename
} 


[void]$Form.ShowDialog() 
$Form.Dispose() 