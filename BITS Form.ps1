Add-Type -AssemblyName System.Windows.Forms

$Form = New-Object system.Windows.Forms.Form 
$Form.Text = "BITS Tool"
$Form.TopMost = $true
$Form.FormBorderStyle = 'Fixed3D'
$Form.MaximizeBox = $false
$Form.Width = 535
$Form.Height = 225


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
$startbutton.Text = "Start BITS Transfer!"
$startbutton.Width = 490
$startbutton.Height = 25
$startbutton.location = new-object system.drawing.point(10,70)
$startbutton.Font = "Microsoft Sans Serif,10"
$Form.controls.Add($startbutton) 


$statusbox = New-Object system.windows.Forms.Label 
$statusbox.Text = "Awaiting transfer.."
#$statusbox.Text = "1`n2`n3`n4`n5"
$statusbox.AutoSize = $true
$statusbox.Width = 490
$statusbox.Height = 100
$statusbox.location = new-object system.drawing.point(10,100)
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
		$Job = Start-BitsTransfer -Source "$source" -Destination "$dest" -Asynchronous -Priority low
		while (($Job.JobState -eq "Transferring") -or ($Job.JobState -eq "Connecting")) `
		{ 
			$BITSInfo = get-bitstransfer
			$statusbox.Text=($BITSInfo `
				| Format-List `
					@{Name="Owner         ";Expression={$_.owneraccount}},
					@{Name="Creation Time ";Expression={$_.creationtime}},
					@{Name="Transferred MB";Expression={[math]::Round($_.bytestransferred / 1MB,2)}},
					@{Name="Total MB      ";Expression={[math]::Round($_.bytestotal / 1MB,2)}},
					@{Name="Completed %   ";Expression={[math]::Round($_.bytestransferred / $_.bytestotal * 100,2)}}	`
				| Out-String).Trim()
			sleep 2;
		Switch($Job.JobState)
		{
			"Transferred"	{
							Complete-BitsTransfer -BitsJob $Job
							$statusbox.Text="`n`n! ! ! ! ! ! ! ! !      Transfer Complete.      ! ! ! ! ! ! ! ! !"
							}
			"Error" {$statusbox.Text=$Job } # List the errors.
			default {"Other action"} #  Perform corrective action.
			
		}
		}
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