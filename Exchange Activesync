Add-Type -AssemblyName System.Windows.Forms

$Form = New-Object system.Windows.Forms.Form 
$Form.Text = "Activesync Tool"
$Form.TopMost = $true
$Form.FormBorderStyle = 'Fixed3D'
$Form.MaximizeBox = $false
$Form.Width = 400
$Form.Height = 415



$label1 = New-Object system.windows.Forms.Label 
$label1.Text = "Username"
$label1.AutoSize = $true
$label1.Width = 35
$label1.Height = 10
$label1.location = new-object system.drawing.point(10,10)
$label1.Font = "Microsoft Sans Serif,10"
$Form.controls.Add($label1) 

$unametext = New-Object system.windows.Forms.TextBox 
$unametext.Width = 150
$unametext.Height = 20
$unametext.location = new-object system.drawing.point(10,30)
$unametext.Font = "Microsoft Sans Serif,10"
$Form.controls.Add($unametext) 

$searchbutton = New-Object system.windows.Forms.Button 
$searchbutton.Text = "Search"
$searchbutton.Width = 60
$searchbutton.Height = 20
$searchbutton.location = new-object system.drawing.point(98,5)
$searchbutton.Font = "Microsoft Sans Serif,10"
$Form.controls.Add($searchbutton) 



$label3 = New-Object system.windows.Forms.Label 
$label3.Text = "DeviceID"
$label3.AutoSize = $true
$label3.Width = 35
$label3.Height = 10
$label3.location = new-object system.drawing.point(180,10)
$label3.Font = "Microsoft Sans Serif,10"
$Form.controls.Add($label3) 

$deviceidtext = New-Object system.windows.Forms.TextBox 
$deviceidtext.Width = 190
$deviceidtext.Height = 20
$deviceidtext.location = new-object system.drawing.point(180,30)
$deviceidtext.Font = "Microsoft Sans Serif,10"
$Form.controls.Add($deviceidtext)

$removebutton = New-Object system.windows.Forms.Button 
$removebutton.Text = "Remove"
$removebutton.Width = 70
$removebutton.Height = 20
$removebutton.location = new-object system.drawing.point(300,5)
$removebutton.Font = "Microsoft Sans Serif,10"
$Form.controls.Add($removebutton) 



$label2 = New-Object system.windows.Forms.Label 
$label2.Text = "Current Activesync Devices:"
$label2.AutoSize = $true
$label2.Width = 25
$label2.Height = 10
$label2.location = new-object system.drawing.point(10,60)
$label2.Font = "Microsoft Sans Serif,10"
$Form.controls.Add($label2) 


$resultbox = New-Object System.Windows.Forms.TextBox
$resultbox.Text = "Please search for a user."
$resultbox.AutoSize = $true
$resultbox.MultiLine = $True
$resultbox.ScrollBars = "Vertical"
$resultbox.location = new-object system.drawing.point(10,80)
$resultbox.Size = New-Object System.Drawing.Size(360,280)
$resultbox.Font = "Consolas,10"
$Form.controls.Add($resultbox) 



$searchbutton.Add_Click(
	{    
	$resultbox.text = ""
	$removebutton.enabled = $false
	$searchbutton.enabled = $false
	$unametext.enabled = $false
	$deviceidtext.enabled = $false
	$SyncInfo = Get-ActiveSyncDevice -mailbox $unametext.text | Get-ActiveSyncDeviceStatistics
			$resultbox.Text+=($SyncInfo `
			| Format-List `
				@{Name="Device   ";Expression={$_.DeviceType}},
				@{Name="Name     ";Expression={$_.DeviceFriendlyName}},
				@{Name="Added    ";Expression={$_.FirstSyncTime}},
				@{Name="Last seen";Expression={$_.LastSuccessSync}},
				@{Name="Device ID";Expression={$_.DeviceID}}  `
			| Out-String).Trim()
		$resultbox.Text+="`n"
	$global:guname=$unametext.text
	$deviceidtext.enabled = $true
	$removebutton.enabled = $true
	$searchbutton.enabled = $true
	$unametext.enabled = $true
	}
)


$removebutton.Add_Click(
	{    
	$resultbox.text=$displayresult
	$removebutton.enabled = $false
	$searchbutton.enabled = $false
	$unametext.enabled = $false
	$deviceidtext.enabled = $false
	$global:gdevid=$deviceidtext.text
	Get-ActiveSyncDevice -mailbox $global:guname -filter "DeviceId -eq ""$global:gdevid""" | Get-ActiveSyncDeviceStatistics | Remove-ActiveSyncDevice -Confirm:$false
	$resultbox.Text+=$global:gdevid + " removed from " + $global:guname
	sleep 1;
	sleep 1;
	$resultbox.Text=""
	$SyncInfo = Get-ActiveSyncDevice -mailbox $unametext.text | Get-ActiveSyncDeviceStatistics
			$resultbox.Text+=($SyncInfo `
			| Format-List `
				@{Name="Device   ";Expression={$_.DeviceType}},
				@{Name="Name     ";Expression={$_.DeviceFriendlyName}},
				@{Name="Added    ";Expression={$_.FirstSyncTime}},
				@{Name="Last seen";Expression={$_.LastSuccessSync}},
				@{Name="Device ID";Expression={$_.DeviceID}}  `
			| Out-String).Trim()
		$resultbox.Text+="`n"
	$deviceidtext.enabled = $true
	$removebutton.enabled = $true
	$searchbutton.enabled = $true
	$unametext.enabled = $true
	}
)


if (!(Get-Command -Name get-exchangeserver -ErrorAction SilentlyContinue)) {
	Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010
	. $env:ExchangeInstallPath\bin\RemoteExchange.ps1 
	Connect-ExchangeServer -auto -AllowClobber 
}


[void]$Form.ShowDialog() 
$Form.Dispose() 


