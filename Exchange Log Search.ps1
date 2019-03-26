Add-Type -AssemblyName System.Windows.Forms

$Form = New-Object system.Windows.Forms.Form 
$Form.Text = "Exchange Log Search Tool"
$Form.TopMost = $true
$Form.FormBorderStyle = 'Fixed3D'
$Form.MaximizeBox = $false
$Form.Width = 415
$Form.Height = 265



$datefrom = New-Object system.windows.Forms.Label 
$datefrom.Text = "From (date/time)"
$datefrom.AutoSize = $true
$datefrom.Width = 35
$datefrom.Height = 10
$datefrom.location = new-object system.drawing.point(10,10)
$datefrom.Font = "Microsoft Sans Serif,10"
$Form.controls.Add($datefrom) 

$datefromtext = New-Object system.windows.Forms.TextBox 
$datefromtext.text = [DateTime]::Now.AddMinutes(-1).ToString("MM/dd/yyy HH:mm:ss")
$datefromtext.Width = 130
$datefromtext.Height = 20
$datefromtext.location = new-object system.drawing.point(10,30)
$datefromtext.Font = "Microsoft Sans Serif,10"
$Form.controls.Add($datefromtext) 



$dateto = New-Object system.windows.Forms.Label 
$dateto.Text = "To (date/time)"
$dateto.AutoSize = $true
$dateto.Width = 35
$dateto.Height = 10
$dateto.location = new-object system.drawing.point(160,10)
$dateto.Font = "Microsoft Sans Serif,10"
$Form.controls.Add($dateto) 

$datetotext = New-Object system.windows.Forms.TextBox 
$datetotext.text = [DateTime]::Now.ToString("MM/dd/yyy HH:mm:ss")
$datetotext.Width = 130
$datetotext.Height = 20
$datetotext.location = new-object system.drawing.point(160,30)
$datetotext.Font = "Microsoft Sans Serif,10"
$Form.controls.Add($datetotext)



$searchbutton = New-Object system.windows.Forms.Button 
$searchbutton.Text = "Search!"
$searchbutton.Width = 65
$searchbutton.Height = 20
$searchbutton.location = new-object system.drawing.point(310,30)
$searchbutton.Font = "Microsoft Sans Serif,10"
$Form.controls.Add($searchbutton) 



$recipientslabel = New-Object system.windows.Forms.Label 
$recipientslabel.Text = "Recipient"
$recipientslabel.AutoSize = $true
$recipientslabel.Width = 35
$recipientslabel.Height = 10
$recipientslabel.location = new-object system.drawing.point(10,60)
$recipientslabel.Font = "Microsoft Sans Serif,10"
$Form.controls.Add($recipientslabel) 

$RecipientsCheckbox = New-Object System.Windows.Forms.Checkbox
$RecipientsCheckbox.Width = 25
$RecipientsCheckbox.Height = 25
$RecipientsCheckbox.Location = New-Object System.Drawing.point(15,80) 
$Form.Controls.Add($RecipientsCheckbox)

$recipientstext = New-Object system.windows.Forms.TextBox
$recipientstext.text = "joe.bloggs@accor.com"
$recipientstext.Width = 335
$recipientstext.Height = 20
$recipientstext.location = new-object system.drawing.point(40,80)
$recipientstext.Font = "Microsoft Sans Serif,10"
$Form.controls.Add($recipientstext)



$senderlabel = New-Object system.windows.Forms.Label 
$senderlabel.Text = "Sender"
$senderlabel.AutoSize = $true
$senderlabel.Width = 35
$senderlabel.Height = 10
$senderlabel.location = new-object system.drawing.point(10,110)
$senderlabel.Font = "Microsoft Sans Serif,10"
$Form.controls.Add($senderlabel) 

$SenderCheckbox = New-Object System.Windows.Forms.Checkbox
$SenderCheckbox.Width = 25
$SenderCheckbox.Height = 25
$SenderCheckbox.Location = New-Object System.Drawing.point(15,130) 
$Form.Controls.Add($SenderCheckbox)

$sendertext = New-Object system.windows.Forms.TextBox
$sendertext.text = "bob.smith@accor.com"
$sendertext.Width = 335
$sendertext.Height = 20
$sendertext.location = new-object system.drawing.point(40,130)
$sendertext.Font = "Microsoft Sans Serif,10"
$Form.controls.Add($sendertext)



$subjectlabel = New-Object system.windows.Forms.Label 
$subjectlabel.Text = "Subject"
$subjectlabel.AutoSize = $true
$subjectlabel.Width = 35
$subjectlabel.Height = 10
$subjectlabel.location = new-object system.drawing.point(10,160)
$subjectlabel.Font = "Microsoft Sans Serif,10"
$Form.controls.Add($subjectlabel) 

$subjectCheckbox = New-Object System.Windows.Forms.Checkbox
$subjectCheckbox.Width = 25
$subjectCheckbox.Height = 25
$subjectCheckbox.Location = New-Object System.Drawing.point(15,180) 
$Form.Controls.Add($subjectCheckbox)

$subjecttext = New-Object system.windows.Forms.TextBox
$subjecttext.text = "Hot Date!!"
$subjecttext.Width = 335
$subjecttext.Height = 20
$subjecttext.location = new-object system.drawing.point(40,180)
$subjecttext.Font = "Microsoft Sans Serif,10"
$Form.controls.Add($subjecttext)



$searchbutton.Add_Click(
	{
	# Disable fields while searching
	$searchbutton.enabled = $false
	$datefromtext.enabled = $false
	$datetotext.enabled = $false
	$SenderCheckbox.enabled = $false
	$sendertext.enabled = $false
	$subjectCheckbox.enabled = $false
	$subjecttext.enabled = $false
	$RecipientsCheckbox.enabled = $false
	$recipientstext.enabled = $false
	
	# Set hub server
	$hubserver1="colo-exhub1"
	$hubserver2="colo-exhub2"
	
	# Convert writtern date datetime
	$global:paramstart=[datetime]$datefromtext.text
	$global:paramend=[datetime]$datetotext.text

	# Set/Clear globals based on tickbox
	if ($RecipientsCheckbox.Checked -eq $true) {
		$global:recipientstext=$recipientstext.text	
	} else {
		$global:recipientstext=$null
	}
	if ($SenderCheckbox.Checked -eq $true) {
		$global:sendertext=$sendertext.text
	} else {
		$global:sendertext=$null
	}
	if ($subjectCheckbox.Checked -eq $true) {
		$global:subjecttext=$subjecttext.text
	} else {
		$global:subjecttext=$null
	}
	
	# Real big ugly if statement for tick box combos
	if ($global:recipientstext) {
		$logoutput=Get-MessageTrackingLog -server $hubserver1 -start $global:paramstart -end $global:paramend -recipients $global:recipientstext
		$logoutput+=Get-MessageTrackingLog -server $hubserver2 -start $global:paramstart -end $global:paramend -recipients $global:recipientstext
		$logoutput | select-object -property timestamp,serverhostname,eventid,messagesubject,sender,recipients,recipientstatus | sort-object -property timestamp | out-gridview
		
	} elseif ($global:sendertext) {
		$logoutput=Get-MessageTrackingLog -server $hubserver1 -start $global:paramstart -end $global:paramend -sender $global:sendertext
		$logoutput+=Get-MessageTrackingLog -server $hubserver2 -start $global:paramstart -end $global:paramend -sender $global:sendertext
		$logoutput | select-object -property timestamp,serverhostname,eventid,messagesubject,sender,recipients,recipientstatus | sort-object -property timestamp | out-gridview
		
	} elseif ($global:subjecttext) {
		$logoutput=Get-MessageTrackingLog -server $hubserver1 -start $global:paramstart -end $global:paramend -messagesubject $global:subjecttext
		$logoutput+=Get-MessageTrackingLog -server $hubserver2 -start $global:paramstart -end $global:paramend -messagesubject $global:subjecttext
		$logoutput | select-object -property timestamp,serverhostname,eventid,messagesubject,sender,recipients,recipientstatus | sort-object -property timestamp | out-gridview
		
	} elseif ($global:recipientstext -and $global:sendertext) {
		$logoutput=Get-MessageTrackingLog -server $hubserver1 -start $global:paramstart -end $global:paramend -recipients $global:recipientstext -sender $global:sendertext
		$logoutput+=Get-MessageTrackingLog -server $hubserver2 -start $global:paramstart -end $global:paramend -recipients $global:recipientstext -sender $global:sendertext
		$logoutput | select-object -property timestamp,serverhostname,eventid,messagesubject,sender,recipients,recipientstatus | sort-object -property timestamp | out-gridview
		
	} elseif ($global:recipientstext -and $global:subjecttext) {
		$logoutput=Get-MessageTrackingLog -server $hubserver1 -start $global:paramstart -end $global:paramend -recipients $global:recipientstext -messagesubject $global:subjecttext
		$logoutput+=Get-MessageTrackingLog -server $hubserver2 -start $global:paramstart -end $global:paramend -recipients $global:recipientstext -messagesubject $global:subjecttext
		$logoutput | select-object -property timestamp,serverhostname,eventid,messagesubject,sender,recipients,recipientstatus | sort-object -property timestamp | out-gridview
		
	} elseif ($global:sendertext -and $global:subjecttext) {
		$logoutput=Get-MessageTrackingLog -server $hubserver1 -start $global:paramstart -end $global:paramend -sender $global:sendertext -messagesubject $global:subjecttext
		$logoutput+=Get-MessageTrackingLog -server $hubserver2 -start $global:paramstart -end $global:paramend -sender $global:sendertext -messagesubject $global:subjecttext
		$logoutput | select-object -property timestamp,serverhostname,eventid,messagesubject,sender,recipients,recipientstatus | sort-object -property timestamp | out-gridview
		
	} elseif ($global:sendertext -and $global:subjecttext -and $global:recipientstext) {
		$logoutput=Get-MessageTrackingLog -server $hubserver1 -start $global:paramstart -end $global:paramend -recipients $global:recipientstext -sender $global:sendertext -messagesubject $global:subjecttext
		$logoutput+=Get-MessageTrackingLog -server $hubserver2 -start $global:paramstart -end $global:paramend -recipients $global:recipientstext -sender $global:sendertext -messagesubject $global:subjecttext
		$logoutput | select-object -property timestamp,serverhostname,eventid,messagesubject,sender,recipients,recipientstatus | sort-object -property timestamp | out-gridview
		
	} else {
		$logoutput=Get-MessageTrackingLog -server $hubserver1 -start $global:paramstart -end $global:paramend
		$logoutput+=Get-MessageTrackingLog -server $hubserver2 -start $global:paramstart -end $global:paramend
		$logoutput | select-object -property timestamp,serverhostname,eventid,messagesubject,sender,recipients,recipientstatus | sort-object -property timestamp | out-gridview
	}
	
	# Enable fields again
	$searchbutton.enabled = $true
	$datefromtext.enabled = $true
	$datetotext.enabled = $true
	$SenderCheckbox.enabled = $true
	$sendertext.enabled = $true
	$subjectCheckbox.enabled = $true
	$subjecttext.enabled = $true
	$RecipientsCheckbox.enabled = $true
	$recipientstext.enabled = $true
	}
)

if (!(Get-Command -Name get-exchangeserver -ErrorAction SilentlyContinue)) {
	Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010
	. $env:ExchangeInstallPath\bin\RemoteExchange.ps1 
	Connect-ExchangeServer -auto -AllowClobber 
}


[void]$Form.ShowDialog() 
$Form.Dispose() 


