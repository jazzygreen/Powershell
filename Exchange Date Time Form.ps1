Add-Type -AssemblyName System.Windows.Forms

$Form = New-Object system.Windows.Forms.Form 
$Form.Text = "OWA Timezone Tool"
$Form.TopMost = $true
$Form.FormBorderStyle = 'Fixed3D'
$Form.MaximizeBox = $false
$Form.Width = 360
$Form.Height = 240


$label2 = New-Object system.windows.Forms.Label 
$label2.Text = "Username"
$label2.AutoSize = $true
$label2.Width = 25
$label2.Height = 10
$label2.location = new-object system.drawing.point(10,10)
$label2.Font = "Microsoft Sans Serif,10"
$Form.controls.Add($label2) 


$unametext = New-Object system.windows.Forms.TextBox 
$unametext.Width = 250
$unametext.Height = 20
$unametext.location = new-object system.drawing.point(10,30)
$unametext.Font = "Microsoft Sans Serif,10"
$Form.controls.Add($unametext) 


$searchbutton = New-Object system.windows.Forms.Button 
$searchbutton.Text = "Search"
$searchbutton.Width = 60
$searchbutton.Height = 20
$searchbutton.location = new-object system.drawing.point(270,30)
$searchbutton.Font = "Microsoft Sans Serif,10"
$Form.controls.Add($searchbutton) 


$tzbox = New-Object system.windows.Forms.ComboBox 
$tzbox.Text = "Select Timezone"
$tzbox.items.add("Queensland")
$tzbox.items.add("Victoria / Tasmania / New South Wales")
$tzbox.items.add("South Australia")
$tzbox.items.add("Western Australia")
$tzbox.items.add("Northern Territory")
$tzbox.items.add("New Zealand")
$tzbox.items.add("Hawaii")
$tzbox.Width = 250
$tzbox.Height = 20
$tzbox.location = new-object system.drawing.point(10,70)
$tzbox.Font = "Microsoft Sans Serif,10"
$Form.controls.Add($tzbox) 


$setbutton = New-Object system.windows.Forms.Button 
$setbutton.Text = "Set"
$setbutton.Width = 60
$setbutton.Height = 20
$setbutton.location = new-object system.drawing.point(270,70)
$setbutton.Font = "Microsoft Sans Serif,10"
$Form.controls.Add($setbutton) 

$label2 = New-Object system.windows.Forms.Label 
$label2.Text = "Current Timezone Data:"
$label2.AutoSize = $true
$label2.Width = 25
$label2.Height = 10
$label2.location = new-object system.drawing.point(10,110)
$label2.Font = "Microsoft Sans Serif,10"
$Form.controls.Add($label2) 


$resultbox = New-Object system.windows.Forms.Label 
$resultbox.Text = "Please search for a user."
$resultbox.AutoSize = $true
$resultbox.Width = 220
$resultbox.Height = 120
$resultbox.location = new-object system.drawing.point(10,130)
$resultbox.Font = "Consolas,10"
$Form.controls.Add($resultbox) 



$searchbutton.Add_Click(
	{    
	$OSInfo = Get-MailboxRegionalConfiguration -Identity $unametext.text
	$resultbox.text=($OSInfo `
		| Format-List `
			@{Name="Date Format";Expression={$_.DateFormat}},
			@{Name="Language";Expression={$_.Language}},
			@{Name="Time Format";Expression={$_.TimeFormat}},
			@{Name="Time Zone";Expression={$_.TimeZone}}  `
		| Out-String).Trim()
	}
)


$setbutton.Add_Click(
	{    
	$displayresult=""
	start-sleep -milliseconds 300
	$displayresult="Setting Timezone to: `n" + $tzbox.SelectedItem
	$resultbox.text=$displayresult
	
	if ($tzbox.SelectedItem -eq "Queensland") {
		Set-MailboxRegionalConfiguration -Identity $unametext.text -DateFormat "d/MM/yyyy" -Language en-AU -TimeZone "E. Australia Standard Time"
	}
	
	if ($tzbox.SelectedItem -eq "Victoria / Tasmania / New South Wales") {
        Set-MailboxRegionalConfiguration -Identity $unametext.text -DateFormat "d/MM/yyyy" -Language en-AU -TimeZone "AUS Eastern Standard Time"
	}
	
    if ($tzbox.SelectedItem -eq "South Australia") {
		Set-MailboxRegionalConfiguration -Identity $unametext.text -DateFormat "d/MM/yyyy" -Language en-AU -TimeZone "Cen. Australia Standard Time"
	}
	
	if ($tzbox.SelectedItem -eq "Western Australia") {
		Set-MailboxRegionalConfiguration -Identity $unametext.text -DateFormat "d/MM/yyyy" -Language en-AU -TimeZone "W. Australia Standard Time"
    }
	
	if ($tzbox.SelectedItem -eq "Northern Territory") {
		Set-MailboxRegionalConfiguration -Identity $unametext.text -DateFormat "d/MM/yyyy" -Language en-AU -TimeZone "AUS Central Standard Time"		
    }
	
	if ($tzbox.SelectedItem -eq "New Zealand") {
		Set-MailboxRegionalConfiguration -Identity $unametext.text -DateFormat "d/MM/yyyy" -Language en-NZ -TimeZone "New Zealand Standard Time"
    }
	
	if ($tzbox.SelectedItem -eq "Hawaii") {
		Set-MailboxRegionalConfiguration -Identity $unametext.text -DateFormat "M/d/yyyy" -Language en-us -TimeZone "Hawaiian Standard Time"
    }
	start-sleep -seconds 2
	$OSInfo = Get-MailboxRegionalConfiguration -Identity $unametext.text
	$resultbox.text=($OSInfo `
		| Format-List `
			@{Name="Date Format";Expression={$_.DateFormat}},
			@{Name="Language";Expression={$_.Language}},
			@{Name="Time Format";Expression={$_.TimeFormat}},
			@{Name="Time Zone";Expression={$_.TimeZone}}  `
		| Out-String).Trim()
	
	}
)


Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010
. $env:ExchangeInstallPath\bin\RemoteExchange.ps1 
Connect-ExchangeServer -auto -AllowClobber 
cls

[void]$Form.ShowDialog() 
$Form.Dispose() 


