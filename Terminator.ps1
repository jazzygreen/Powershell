# Definitions
$imagepath = "C:\Temp\"
$logfile = "C:\Temp\terminator.log"
$destOU = "OU=_Old Accounts,DC=yourdomain,DC=ad"
$whodat = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name


Add-Type -AssemblyName System.Windows.Forms

# *************** Hide PowerShell Console ***************
Add-Type -Name Window -Namespace Console -MemberDefinition '
[DllImport("Kernel32.dll")]
public static extern IntPtr GetConsoleWindow();
[DllImport("user32.dll")]
public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);
'
$consolePtr = [Console.Window]::GetConsoleWindow()
[Console.Window]::ShowWindow($consolePtr, 0)


# *************** Build the GUI ***************
$Form = New-Object system.Windows.Forms.Form 
$Form.Text = "AD Terminator"
$Form.TopMost = $true
$Form.FormBorderStyle = 'Fixed3D'
$Form.MaximizeBox = $false
$Form.Width = 460
$Form.Height = 355
$form.add_FormClosing({[Console.Window]::ShowWindow($consolePtr, 1)})

$uname = New-Object system.windows.Forms.Label 
$uname.Text = "Username"
$uname.AutoSize = $true
$uname.Width = 25
$uname.Height = 10
$uname.location = new-object system.drawing.point(10,10)
$uname.Font = "Microsoft Sans Serif,10"
$Form.controls.Add($uname) 

$unametext = New-Object system.windows.Forms.TextBox 
$unametext.Width = 320
$unametext.Height = 20
$unametext.location = new-object system.drawing.point(10,30)
$unametext.Font = "Microsoft Sans Serif,10"
$Form.controls.Add($unametext) 

$searchbutton = New-Object system.windows.Forms.Button 
$searchbutton.Text = "Search"
$searchbutton.Width = 90
$searchbutton.Height = 30
$searchbutton.location = new-object system.drawing.point(340,20)
$searchbutton.Font = "Microsoft Sans Serif,10"
$Form.controls.Add($searchbutton) 

$userOUlabel = New-Object system.windows.Forms.Label 
$userOUlabel.Text = "OU"
$userOUlabel.Width = 40
$userOUlabel.Height = 20
$userOUlabel.location = new-object system.drawing.point(10,70)
$userOUlabel.Font = "Microsoft Sans Serif,10"
$userOUlabel.BorderStyle = 'Fixed3D'
$Form.controls.Add($userOUlabel)
$userOUoplabel = New-Object system.windows.Forms.Label 
$userOUoplabel.Text = " "
$userOUoplabel.Width = 280
$userOUoplabel.Height = 20
$userOUoplabel.location = new-object system.drawing.point(50,70)
$userOUoplabel.Font = "Microsoft Sans Serif,10"
$userOUoplabel.BorderStyle = 'Fixed3D'
$Form.controls.Add($userOUoplabel) 
$userOUbutton = New-Object system.windows.Forms.Button
$userOUbutton.Text = "Move OU"
$userOUbutton.Enabled = $false
$userOUbutton.Width = 90
$userOUbutton.Height = 20
$userOUbutton.location = new-object system.drawing.point(340,70)
$userOUbutton.Font = "Microsoft Sans Serif,10"
$Form.controls.Add($userOUbutton) 

$ustatuslabel = New-Object system.windows.Forms.Label 
$ustatuslabel.Text = "Account status"
$ustatuslabel.Width = 160
$ustatuslabel.Height = 20
$ustatuslabel.location = new-object system.drawing.point(10,90)
$ustatuslabel.Font = "Microsoft Sans Serif,10"
$ustatuslabel.BorderStyle = 'Fixed3D'
$Form.controls.Add($ustatuslabel)
$ustatusoplabel = New-Object system.windows.Forms.Label 
$ustatusoplabel.Text = " "
$ustatusoplabel.Width = 160
$ustatusoplabel.Height = 20
$ustatusoplabel.location = new-object system.drawing.point(170,90)
$ustatusoplabel.Font = "Microsoft Sans Serif,10"
$ustatusoplabel.BorderStyle = 'Fixed3D'
$Form.controls.Add($ustatusoplabel) 
$ustatusbutton = New-Object system.windows.Forms.Button
$ustatusbutton.Text = "Disable"
$ustatusbutton.Enabled = $false
$ustatusbutton.Width = 90
$ustatusbutton.Height = 20
$ustatusbutton.location = new-object system.drawing.point(340,90)
$ustatusbutton.Font = "Microsoft Sans Serif,10"
$Form.controls.Add($ustatusbutton) 

$pwsetlabel = New-Object system.windows.Forms.Label 
$pwsetlabel.Text = "Password last set"
$pwsetlabel.Width = 160
$pwsetlabel.Height = 20
$pwsetlabel.location = new-object system.drawing.point(10,110)
$pwsetlabel.Font = "Microsoft Sans Serif,10"
$pwsetlabel.BorderStyle = 'Fixed3D'
$Form.controls.Add($pwsetlabel) 
$pwsetoplabel = New-Object system.windows.Forms.Label 
$pwsetoplabel.Text = " "
$pwsetoplabel.Width = 160
$pwsetoplabel.Height = 20
$pwsetoplabel.location = new-object system.drawing.point(170,110)
$pwsetoplabel.Font = "Microsoft Sans Serif,10"
$pwsetoplabel.BorderStyle = 'Fixed3D'
$Form.controls.Add($pwsetoplabel) 
$pwsetbutton = New-Object system.windows.Forms.Button
$pwsetbutton.Text = "Randomise"
$pwsetbutton.Enabled = $false
$pwsetbutton.Width = 90
$pwsetbutton.Height = 20
$pwsetbutton.location = new-object system.drawing.point(340,110)
$pwsetbutton.Font = "Microsoft Sans Serif,10"
$Form.controls.Add($pwsetbutton)

$abooklabel = New-Object system.windows.Forms.Label 
$abooklabel.Text = "Address book"
$abooklabel.Width = 160
$abooklabel.Height = 20
$abooklabel.location = new-object system.drawing.point(10,130)
$abooklabel.Font = "Microsoft Sans Serif,10"
$abooklabel.BorderStyle = 'Fixed3D'
$Form.controls.Add($abooklabel) 
$abookoplabel = New-Object system.windows.Forms.Label 
$abookoplabel.Text = " "
$abookoplabel.Width = 160
$abookoplabel.Height = 20
$abookoplabel.location = new-object system.drawing.point(170,130)
$abookoplabel.Font = "Microsoft Sans Serif,10"
$abookoplabel.BorderStyle = 'Fixed3D'
$Form.controls.Add($abookoplabel) 
$abookbutton = New-Object system.windows.Forms.Button
$abookbutton.Text = "Hide"
$abookbutton.Enabled = $false
$abookbutton.Width = 90
$abookbutton.Height = 20
$abookbutton.location = new-object system.drawing.point(340,130)
$abookbutton.Font = "Microsoft Sans Serif,10"
$Form.controls.Add($abookbutton)

$mailfwdlabel = New-Object system.windows.Forms.Label 
$mailfwdlabel.Text = "Email redirect"
$mailfwdlabel.Width = 160
$mailfwdlabel.Height = 20
$mailfwdlabel.location = new-object system.drawing.point(10,150)
$mailfwdlabel.Font = "Microsoft Sans Serif,10"
$mailfwdlabel.BorderStyle = 'Fixed3D'
$Form.controls.Add($mailfwdlabel) 
$mailfwdoplabel = New-Object system.windows.Forms.Label 
$mailfwdoplabel.Text = " "
$mailfwdoplabel.Width = 160
$mailfwdoplabel.Height = 20
$mailfwdoplabel.location = new-object system.drawing.point(170,150)
$mailfwdoplabel.Font = "Microsoft Sans Serif,10"
$mailfwdoplabel.BorderStyle = 'Fixed3D'
$Form.controls.Add($mailfwdoplabel) 
$mailfwdbutton = New-Object system.windows.Forms.Button
$mailfwdbutton.Text = "Forward"
$mailfwdbutton.Enabled = $false
$mailfwdbutton.Width = 90
$mailfwdbutton.Height = 20
$mailfwdbutton.location = new-object system.drawing.point(340,150)
$mailfwdbutton.Font = "Microsoft Sans Serif,10"
$Form.controls.Add($mailfwdbutton)

$casowalabel = New-Object system.windows.Forms.Label 
$casowalabel.Text = "Mailbox OWA"
$casowalabel.Width = 160
$casowalabel.Height = 20
$casowalabel.location = new-object system.drawing.point(10,170)
$casowalabel.Font = "Microsoft Sans Serif,10"
$casowalabel.BorderStyle = 'Fixed3D'
$Form.controls.Add($casowalabel)
$casowaoplabel = New-Object system.windows.Forms.Label 
$casowaoplabel.Text = " "
$casowaoplabel.Width = 160
$casowaoplabel.Height = 20
$casowaoplabel.location = new-object system.drawing.point(170,170)
$casowaoplabel.Font = "Microsoft Sans Serif,10"
$casowaoplabel.BorderStyle = 'Fixed3D'
$Form.controls.Add($casowaoplabel) 
$casowabutton = New-Object system.windows.Forms.Button
$casowabutton.Text = "Disable"
$casowabutton.Enabled = $false
$casowabutton.Width = 90
$casowabutton.Height = 20
$casowabutton.location = new-object system.drawing.point(340,170)
$casowabutton.Font = "Microsoft Sans Serif,10"
$Form.controls.Add($casowabutton)

$caseaslabel = New-Object system.windows.Forms.Label 
$caseaslabel.Text = "Mailbox EAS"
$caseaslabel.Width = 160
$caseaslabel.Height = 20
$caseaslabel.location = new-object system.drawing.point(10,190)
$caseaslabel.Font = "Microsoft Sans Serif,10"
$caseaslabel.BorderStyle = 'Fixed3D'
$Form.controls.Add($caseaslabel) 
$caseasoplabel = New-Object system.windows.Forms.Label 
$caseasoplabel.Text = " "
$caseasoplabel.Width = 160
$caseasoplabel.Height = 20
$caseasoplabel.location = new-object system.drawing.point(170,190)
$caseasoplabel.Font = "Microsoft Sans Serif,10"
$caseasoplabel.BorderStyle = 'Fixed3D'
$Form.controls.Add($caseasoplabel) 
$caseasbutton = New-Object system.windows.Forms.Button
$caseasbutton.Text = "Disable"
$caseasbutton.Enabled = $false
$caseasbutton.Width = 90
$caseasbutton.Height = 20
$caseasbutton.location = new-object system.drawing.point(340,190)
$caseasbutton.Font = "Microsoft Sans Serif,10"
$Form.controls.Add($caseasbutton)

$casmapilabel = New-Object system.windows.Forms.Label 
$casmapilabel.Text = "Mailbox MAPI"
$casmapilabel.Width = 160
$casmapilabel.Height = 20
$casmapilabel.location = new-object system.drawing.point(10,210)
$casmapilabel.Font = "Microsoft Sans Serif,10"
$casmapilabel.BorderStyle = 'Fixed3D'
$Form.controls.Add($casmapilabel) 
$casmapioplabel = New-Object system.windows.Forms.Label 
$casmapioplabel.Text = " "
$casmapioplabel.Width = 160
$casmapioplabel.Height = 20
$casmapioplabel.location = new-object system.drawing.point(170,210)
$casmapioplabel.Font = "Microsoft Sans Serif,10"
$casmapioplabel.BorderStyle = 'Fixed3D'
$Form.controls.Add($casmapioplabel) 
$casmapibutton = New-Object system.windows.Forms.Button
$casmapibutton.Text = "Disable"
$casmapibutton.Enabled = $false
$casmapibutton.Width = 90
$casmapibutton.Height = 20
$casmapibutton.location = new-object system.drawing.point(340,210)
$casmapibutton.Font = "Microsoft Sans Serif,10"
$Form.controls.Add($casmapibutton)

$caspoplabel = New-Object system.windows.Forms.Label 
$caspoplabel.Text = "Mailbox POP"
$caspoplabel.Width = 160
$caspoplabel.Height = 20
$caspoplabel.location = new-object system.drawing.point(10,230)
$caspoplabel.Font = "Microsoft Sans Serif,10"
$caspoplabel.BorderStyle = 'Fixed3D'
$Form.controls.Add($caspoplabel) 
$caspopoplabel = New-Object system.windows.Forms.Label 
$caspopoplabel.Text = " "
$caspopoplabel.Width = 160
$caspopoplabel.Height = 20
$caspopoplabel.location = new-object system.drawing.point(170,230)
$caspopoplabel.Font = "Microsoft Sans Serif,10"
$caspopoplabel.BorderStyle = 'Fixed3D'
$Form.controls.Add($caspopoplabel) 
$caspopbutton = New-Object system.windows.Forms.Button
$caspopbutton.Text = "Disable"
$caspopbutton.Enabled = $false
$caspopbutton.Width = 90
$caspopbutton.Height = 20
$caspopbutton.location = new-object system.drawing.point(340,230)
$caspopbutton.Font = "Microsoft Sans Serif,10"
$Form.controls.Add($caspopbutton)

$casimaplabel = New-Object system.windows.Forms.Label 
$casimaplabel.Text = "Mailbox IMAP"
$casimaplabel.Width = 160
$casimaplabel.Height = 20
$casimaplabel.location = new-object system.drawing.point(10,250)
$casimaplabel.Font = "Microsoft Sans Serif,10"
$casimaplabel.BorderStyle = 'Fixed3D'
$Form.controls.Add($casimaplabel) 
$casimapoplabel = New-Object system.windows.Forms.Label 
$casimapoplabel.Text = " "
$casimapoplabel.Width = 160
$casimapoplabel.Height = 20
$casimapoplabel.location = new-object system.drawing.point(170,250)
$casimapoplabel.Font = "Microsoft Sans Serif,10"
$casimapoplabel.BorderStyle = 'Fixed3D'
$Form.controls.Add($casimapoplabel) 
$casimapbutton = New-Object system.windows.Forms.Button
$casimapbutton.Text = "Disable"
$casimapbutton.Enabled = $false
$casimapbutton.Width = 90
$casimapbutton.Height = 20
$casimapbutton.location = new-object system.drawing.point(340,250)
$casimapbutton.Font = "Microsoft Sans Serif,10"
$Form.controls.Add($casimapbutton)

$gentxtmembtn = New-Object system.windows.Forms.Button
$gentxtmembtn.Text = "Details (txt)"
$gentxtmembtn.Enabled = $false
$gentxtmembtn.Width = 100
$gentxtmembtn.Height = 30
$gentxtmembtn.location = new-object system.drawing.point(10,275)
$gentxtmembtn.Font = "Microsoft Sans Serif,10"
$Form.controls.Add($gentxtmembtn)

$genjpgmembtn = New-Object system.windows.Forms.Button
$genjpgmembtn.Text = "Details (jpg)"
$genjpgmembtn.Enabled = $false
$genjpgmembtn.Width = 100
$genjpgmembtn.Height = 30
$genjpgmembtn.location = new-object system.drawing.point(117,275)
$genjpgmembtn.Font = "Microsoft Sans Serif,10"
$Form.controls.Add($genjpgmembtn)

$sparebtn1 = New-Object system.windows.Forms.Button
$sparebtn1.Text = "Spare"
$sparebtn1.Enabled = $false
$sparebtn1.Width = 100
$sparebtn1.Height = 30
$sparebtn1.location = new-object system.drawing.point(223,275)
$sparebtn1.Font = "Microsoft Sans Serif,10"
$Form.controls.Add($sparebtn1)

$sparebtn2 = New-Object system.windows.Forms.Button
$sparebtn2.Text = "Spare"
$sparebtn2.Enabled = $false
$sparebtn2.Width = 100
$sparebtn2.Height = 30
$sparebtn2.location = new-object system.drawing.point(330,275)
$sparebtn2.Font = "Microsoft Sans Serif,10"
$Form.controls.Add($sparebtn2)


# *************** Import powershell moduled if required ***************
if (!(Get-Command -Name get-exchangeserver -ErrorAction SilentlyContinue)) {
	Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010
	. $env:ExchangeInstallPath\bin\RemoteExchange.ps1 
	Connect-ExchangeServer -auto -AllowClobber 
}

if (!(Get-Command -Name get-aduser -ErrorAction SilentlyContinue)) {
	Import-Module ActiveDirectory
}


# *************** Button definitions/code ***************
$unametext.Add_KeyDown({
    if ($_.KeyCode -eq "Enter") {
        $searchbutton.PerformClick()
    }
})


$searchbutton.Add_Click(
	{    
	$user = $unametext.Text
	
	# If the username too short, set a long name to search which will fail. a blank username will fail with error.
	if ($user.length -le 4) {
		$usearch="accountnamethatwillnevereverexist"
	} else {
		$usearch=$user
	}
	$usercheck = Get-ADUser -Filter {SamAccountName -eq $usearch}
	
	if ($user.length -le 4) {
	
	# If the specified account name is too short, ensure buttons diabled and prompt the user.
		$abookbutton.Enabled = $false
		$mailfwdbutton.Enabled = $false
		$casowabutton.Enabled = $false
		$caseasbutton.Enabled = $false
		$casmapibutton.Enabled = $false
		$caspopbutton.Enabled = $false
		$casimapbutton.Enabled = $false
		$ustatusbutton.Enabled = $false
		$pwsetbutton.Enabled = $false
		$genjpgmembtn.Enabled = $false
		$gentxtmembtn.Enabled = $false
		$userOUbutton.Enabled = $false
		[void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
		[Microsoft.VisualBasic.Interaction]::MsgBox("Username too short.", "OKOnly,SystemModal,Exclamation", "Warning")
	} elseif ($usercheck -eq $null) {
	
	# If the specified account name doesnt exist, ensure buttons diabled and prompt the user.
		$abookbutton.Enabled = $false
		$mailfwdbutton.Enabled = $false
		$casowabutton.Enabled = $false
		$caseasbutton.Enabled = $false
		$casmapibutton.Enabled = $false
		$caspopbutton.Enabled = $false
		$casimapbutton.Enabled = $false
		$ustatusbutton.Enabled = $false
		$pwsetbutton.Enabled = $false
		$genjpgmembtn.Enabled = $false
		$gentxtmembtn.Enabled = $false
		$userOUbutton.Enabled = $false
		[void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
		[Microsoft.VisualBasic.Interaction]::MsgBox("Username not found.", "OKOnly,SystemModal,Exclamation", "Warning")
	} else {
	
	# User has been found, get account details to display and enable or disable buttons accordingly.
	$userAD=get-aduser $user -properties *
	if ($userAD.msExchMailboxGuid) {
	
	# User has been found and is mail enabled, process mail related options.
		$userMB=get-mailbox $user
		$userCAS=get-casmailbox $user
		if ($userMB.HiddenFromAddressListsEnabled) {
			$abook="Hidden"
			$abookbutton.Enabled = $false
		} else {
			$abook="Showing"
			$abookbutton.Enabled = $true
		}
		if ([string]::IsNullOrEmpty($userMB.ForwardingAddress.Name)) {
			$mailfwd="No forward set"
			$mailfwdbutton.Enabled = $true
		} else {
			$mailfwdbutton.Enabled = $false
		}
		if ($userCAS.OWAEnabled) {
			$casowa="Enabled"
			$casowabutton.Enabled = $true
		} else {
			$casowa="Disabled"
			$casowabutton.Enabled = $false
		}
		if ($userCAS.ActiveSyncEnabled) {
			$caseas="Enabled"
			$caseasbutton.Enabled = $true
		} else {
			$caseas="Disabled"
			$caseasbutton.Enabled = $false
		}
		if ($userCAS.MAPIEnabled) {
			$casmapi="Enabled"
			$casmapibutton.Enabled = $true
		} else {
			$casmapi="Disabled"
			$casmapibutton.Enabled = $false
		}
		if ($userCAS.PopEnabled) {
			$caspop="Enabled"
			$caspopbutton.Enabled = $true
		} else {
			$caspop="Disabled"
			$caspopbutton.Enabled = $false
		}
		if ($userCAS.ImapEnabled) {
			$casimap="Enabled"
			$casimapbutton.Enabled = $true
		} else {
			$casimap="Disabled"
			$casimapbutton.Enabled = $false
		}
	} else {
	
	# User has been found but is not mail enabled, disable appropriate mail buttons.
		$abook="No email account"
		$abookbutton.Enabled = $false
		$mailfwd="No email account"
		$mailfwdbutton.Enabled = $false
		$casowa="No email account"
		$casowabutton.Enabled = $false
		$caseas="No email account"
		$caseasbutton.Enabled = $false
		$casmapi="No email account"
		$casmapibutton.Enabled = $false
		$caspop="No email account"
		$caspopbutton.Enabled = $false
		$casimap="No email account"
		$casimapbutton.Enabled = $false
	}
	
	# Generic account options are now displayed and buttons set accordingly.
	$userOU = ($userAD.DistinguishedName -split ",")[1]
	$userOUoplabel.Text = $userOU.substring(3)
	if ($userAD.enabled) {
		$ustatus="Enabled"
		$ustatusbutton.Enabled = $true
	} else {
		$ustatus="Disabled"
		$ustatusbutton.Enabled = $false
	}
	if ($userAD.pwdlastset -gt 0) {
		$pwtime=[datetime]::FromFileTime($userAD.pwdlastset)
		$pwsetbutton.Enabled = $true
	} else {
		$pwtime="Password set"
		$pwsetbutton.Enabled = $false
	}
	if (($userAD.DistinguishedName -split ",")[1] -eq ($destOU -split ",")[0]) {
		$userOUbutton.Enabled = $false
		$userOUoplabel.Text = ($userAD.DistinguishedName -split ",")[1].substring(3)
	} else {
		$userOUbutton.Enabled = $true
		$userOUoplabel.Text = ($userAD.DistinguishedName -split ",")[1].substring(3)
	}

	$genjpgmembtn.Enabled = $true
	$gentxtmembtn.Enabled = $true
	$ustatusoplabel.Text = $ustatus
	$pwsetoplabel.Text = $pwtime
	$abookoplabel.Text = $abook
	$mailfwdoplabel.Text = $mailfwd + $userMB.ForwardingAddress.Name
	$casowaoplabel.Text = $casowa
	$caseasoplabel.Text = $caseas
	$casmapioplabel.Text = $casmapi
	$caspopoplabel.Text = $caspop
	$casimapoplabel.Text = $casimap

	}
	}
)


$userOUbutton.Add_Click(
	{
	$user = $unametext.Text
	
	# Prompt to ensure a user should be moved to the old OU.
	[void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
	$response = [Microsoft.VisualBasic.Interaction]::MsgBox("Move $user ($destOU -split ",")[0].substring(3) ?",'YesNo,Question', "Respond please")
	
	if ($response -eq "Yes") {
	
		# Move the user.
		Get-ADUser $user | Move-ADObject -TargetPath $destOU
		
		# Log the action.
		(Get-Date).tostring("dd-MM-yyyy HH:mm:ss") + "`t" + $whodat + "`t" + "Move OU" + "`t" + $user | Out-File $logfile -Append
		
		# Sleep to allow for local replication.
		start-sleep 1
		start-sleep 1
		start-sleep 1
		start-sleep 1
	}

	$userAD=get-aduser $user
	
	# Update the display and buttons accordingly.
	if (($userAD.DistinguishedName -split ",")[1] -eq ($destOU -split ",")[0]) {
		$userOUbutton.Enabled = $false
		$userOUoplabel.Text = ($userAD.DistinguishedName -split ",")[1].substring(3)
	} else {
		$userOUbutton.Enabled = $true
		$userOUoplabel.Text = ($userAD.DistinguishedName -split ",")[1].substring(3)
	}

	}
)


$ustatusbutton.Add_Click(
	{
	$user = $unametext.Text
	
	# Disable the user account.
	Disable-ADAccount $user
	
	# Log the action.
	(Get-Date).tostring("dd-MM-yyyy HH:mm:ss") + "`t" + $whodat + "`t" + "Disable AD" + "`t" + $user | Out-File $logfile -Append
	$userAD=get-aduser $user -properties *
	
	# Update the display and buttons accordingly.
		if ($userAD.enabled) {
			$ustatus="Enabled"
			$ustatusbutton.Enabled = $true
			$ustatusoplabel.Text = $ustatus
		} else {
			$ustatus="Disabled"
			$ustatusbutton.Enabled = $false
			$ustatusoplabel.Text = $ustatus
		}
	}
)


$pwsetbutton.Add_Click(
	{
	$user = $unametext.Text
	
	# Generate a random complex password.
	$randompw=-join ((48..64) + (97..122) + (97..122) | Get-Random -Count 24 | % {[char]$_})
	
	# Set the password.
	Set-ADAccountPassword -Identity $user -Reset -NewPassword (ConvertTo-SecureString -AsPlainText "$randompw" -Force)
	
	# Force password change flag.
	Set-ADUser -Identity $user -ChangePasswordAtLogon $true
	
	# Log the action.
	(Get-Date).tostring("dd-MM-yyyy HH:mm:ss") + "`t" + $whodat + "`t" + "Change PW" + "`t" + $user | Out-File $logfile -Append
	
	$userAD=get-aduser $user -properties *
	
	# Update the display and buttons accordingly.
		if ($userAD.pwdlastset -gt 0) {
			$pwtime=[datetime]::FromFileTime($userAD.pwdlastset)
			$pwsetoplabel.Text = $pwtime
			$pwsetbutton.Enabled = $true
		} else {
			$pwtime="Password set"
			$pwsetoplabel.Text = $pwtime
			$pwsetbutton.Enabled = $false
		}
	}
)


$abookbutton.Add_Click(
	{
	$user = $unametext.Text
	
	# Hide the user from the GAL.
	set-mailbox -identity $user -HiddenFromAddressListsEnabled $true
	
	# Log the action.
	(Get-Date).tostring("dd-MM-yyyy HH:mm:ss") + "`t" + $whodat + "`t" + "Hide Mailbox" + "`t" + $user | Out-File $logfile -Append
	
	$userMB=get-mailbox $user
	
	# Update the display and buttons accordingly.
		if ($userMB.HiddenFromAddressListsEnabled) {
			$abook="Hidden"
			$abookbutton.Enabled = $false
			$abookoplabel.Text = $abook
		} else {
			$abook="Showing"
			$abookbutton.Enabled = $true
			$abookoplabel.Text = $abook
		}
	}
)


$mailfwdbutton.Add_Click(
	{
	$user = $unametext.Text
	
	# Prompt the for the forward destination account name.
	[void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
	$title = 'Email Forwarding address'
	$msg   = 'Enter the recipient username. (eg, bob.east)'
	$fwuser = [Microsoft.VisualBasic.Interaction]::InputBox($msg, $title)
	$fwaduser = get-mailbox -identity $fwuser
	$fwdemail = $fwaduser.primarysmtpaddress.address
	
	# Prompt to ensure the correct forward destination.
	$response = [Microsoft.VisualBasic.Interaction]::MsgBox("Set Email forward to $fwaduser ($fwdemail)?",'YesNo,Question', "Respond please")
		
	if ($response -eq "Yes") {
	
		# Set the email forward if user clicked Yes.
		set-mailbox -identity $user -ForwardingAddress $fwaduser
		
		# Log the action.
		(Get-Date).tostring("dd-MM-yyyy HH:mm:ss") + "`t" + $whodat + "`t" + "FW to " + $fwaduser + "`t" + $user | Out-File $logfile -Append
	}
	$userMB=get-mailbox $user
	
	# Update the display and buttons accordingly.
		if ([string]::IsNullOrEmpty($userMB.ForwardingAddress.Name)) {
			$mailfwd="No forward set"
			$mailfwdbutton.Enabled = $true
			$mailfwdoplabel.Text = $mailfwd
		} else {
			$mailfwdbutton.Enabled = $false
			$mailfwdoplabel.Text = $userMB.ForwardingAddress.Name
		}
	}
)


$casowabutton.Add_Click(
	{
	$user = $unametext.Text
	
	# Disable OWA.
	set-casmailbox -identity $user -OWAEnabled $false
	
	# Log the action.
	(Get-Date).tostring("dd-MM-yyyy HH:mm:ss") + "`t" + $whodat + "`t" + "Disable OWA" + "`t" + $user | Out-File $logfile -Append
	$userCAS=get-casmailbox $user
	
	# Update the display and buttons accordingly.
		if ($userCAS.OWAEnabled) {
			$casowa="Enabled"
			$casowabutton.Enabled = $true
			$casowaoplabel.Text = $casowa
		} else {
			$casowa="Disabled"
			$casowabutton.Enabled = $false
			$casowaoplabel.Text = $casowa
		}
	}
)


$caseasbutton.Add_Click(
	{
	$user = $unametext.Text
	
	# Disable EAS.
	set-casmailbox -identity $user -ActiveSyncEnabled $false
	
	# Log the action.
	(Get-Date).tostring("dd-MM-yyyy HH:mm:ss") + "`t" + $whodat + "`t" + "Disable EAS" + "`t" + $user | Out-File $logfile -Append
	$userCAS=get-casmailbox $user
	
	# Update the display and buttons accordingly.
		if ($userCAS.ActiveSyncEnabled) {
			$caseas="Enabled"
			$caseasbutton.Enabled = $true
			$caseasoplabel.Text = $caseas
		} else {
			$caseas="Disabled"
			$caseasbutton.Enabled = $false
			$caseasoplabel.Text = $caseas
		}
	}
)


$casmapibutton.Add_Click(
	{
	$user = $unametext.Text
	
	# Disable MAPI.
	set-casmailbox -identity $user -MAPIEnabled $false
	
	# Log the action.
	(Get-Date).tostring("dd-MM-yyyy HH:mm:ss") + "`t" + $whodat + "`t" + "Disable MAPI" + "`t" + $user | Out-File $logfile -Append
	$userCAS=get-casmailbox $user
	
	# Update the display and buttons accordingly.
		if ($userCAS.MAPIEnabled) {
			$casmapi="Enabled"
			$casmapibutton.Enabled = $true
			$casmapioplabel.Text = $casmapi
		} else {
			$casmapi="Disabled"
			$casmapibutton.Enabled = $false
			$casmapioplabel.Text = $casmapi
		}
	}
)


$caspopbutton.Add_Click(
	{
	$user = $unametext.Text
	
	# Disable POP.
	set-casmailbox -identity $user -PopEnabled $false
	
	# Log the action.
	(Get-Date).tostring("dd-MM-yyyy HH:mm:ss") + "`t" + $whodat + "`t" + "Disable POP" + "`t" + $user | Out-File $logfile -Append
	$userCAS=get-casmailbox $user
	
	# Update the display and buttons accordingly.
		if ($userCAS.PopEnabled) {
			$caspop="Enabled"
			$caspopbutton.Enabled = $true
			$caspopoplabel.Text = $caspop
		} else {
			$caspop="Disabled"
			$caspopbutton.Enabled = $false
			$caspopoplabel.Text = $caspop
		}
	}
)


$casimapbutton.Add_Click(
	{
	$user = $unametext.Text
	
	# Disable CAS.
	set-casmailbox -identity $user -ImapEnabled $false
	
	# Log the action.
	(Get-Date).tostring("dd-MM-yyyy HH:mm:ss") + "`t" + $whodat + "`t" + "Disable IMAP" + "`t" + $user | Out-File $logfile -Append
	$userCAS=get-casmailbox $user
	
	# Update the display and buttons accordingly.
		if ($userCAS.ImapEnabled) {
			$casimap="Enabled"
			$casimapbutton.Enabled = $true
			$casimapoplabel.Text = $casimap
		} else {
			$casimap="Disabled"
			$casimapbutton.Enabled = $false
			$casimapoplabel.Text = $casimap
		}
	}
)


$gentxtmembtn.Add_Click(
	{
	$user = $unametext.Text
	
	# Get users memberships.
	$usergroups = Get-ADPrincipalGroupMembership -identity $user
	
	# Sort the text.
	$groupgen = $usergroups.name | sort
	
	$userAD=get-aduser $user -properties *
	
	# If user is mail enabled, get addresses and format.
	if ($userAD.msExchMailboxGuid) {
		$groupgen += "`n"
		$mbox=Get-Mailbox $user
		$mailaddresses=$mbox.emailaddresses | sort -property prefixstring -CaseSensitive -descending
		$addresses=$mailaddresses.ProxyAddressString | select-string -pattern "smtp"
		foreach ($mailaddys in $addresses) {
			$usermail+=($mailaddys -split ":")[0] + " " + ($mailaddys -split ":")[1] + "`n"
		}
		$groupgen += $usermail.trim("`r`n") | fl
	}
	
	# Output groups to clipboard
	$groupgen | clip
	
	# Log the action.
	(Get-Date).tostring("dd-MM-yyyy HH:mm:ss") + "`t" + $whodat + "`t" + "Gen Member/Email (txt)" + "`t" + $user | Out-File $logfile -Append
	[void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
	
	# Prompt the user when complete.
	$response = [Microsoft.VisualBasic.Interaction]::MsgBox("Details copied to clipboard.",'OKOnly,Information', "FYI")
	}
)


$genjpgmembtn.Add_Click(
	{
	$user = $unametext.Text
	
	# Get users memberships.
	$usergroups = Get-ADPrincipalGroupMembership -identity $user
	
	# Sort the text.
	$srtgrp = $usergroups.name | sort

	# Add new line to groups only if in more than 1 (required for the image generation)
	if ($usergroups.name.count -gt 1) {
		foreach($ugrp in $srtgrp) {
			$groupgen+=$ugrp
			$groupgen+="`n"
		} 
	} else {
		$groupgen=$usergroups.name
	}

	$userAD=get-aduser $user -properties *
	
	# If user is mail enabled, get addresses and format.
	if ($userAD.msExchMailboxGuid) {
		$groupgen += "`n"
		$mbox=Get-Mailbox $user
		
		# Get mail/x400/x500 addresses and sort them.
		$mailaddresses=$mbox.emailaddresses | sort -property prefixstring -CaseSensitive -descending
		$addresses=$mailaddresses.ProxyAddressString | select-string -pattern "smtp"
		$addresses+=$mailaddresses.ProxyAddressString | select-string -pattern "x400"
		$addresses+=$mailaddresses.ProxyAddressString | select-string -pattern "x500"
		$i=0
		foreach ($mailaddys in $addresses) {
			$usermail+=($mailaddys -split ":")[0] + "`t" + ($mailaddys -split ":")[1] + "`n"
			$i++
		}
		$groupgen += $usermail.trim("`r`n") | fl | out-string
	}
	
	# Count the number of lines that will be generated in the image
	$linecount= $i + $usergroups.name.count + 1
	
	# Build and save the image based off of the data collected
	$picheight = ($linecount * 15) + 20
	Add-Type -AssemblyName System.Drawing
	$filename = $imagepath + (Get-Date).tostring("yyyy-MM-dd HH.mm.ss") + " " + $user + ".png"
	$bmp = new-object System.Drawing.Bitmap 450,$picheight 
	$font = new-object System.Drawing.Font SegoeUI,10 
	$brushBg = [System.Drawing.Brushes]::White
	$brushFg = [System.Drawing.Brushes]::Black 
	$graphics = [System.Drawing.Graphics]::FromImage($bmp) 
	$graphics.FillRectangle($brushBg,0,0,$bmp.Width,$bmp.Height) 
	$graphics.DrawString($groupgen,$font,$brushFg,10,10) 
	$graphics.Dispose() 
	$bmp.Save($filename) 
	[Reflection.Assembly]::LoadWithPartialName('System.Drawing');
	[Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms');
	$file = get-item($filename);
	$img = [System.Drawing.Image]::Fromfile($file);
	[System.Windows.Forms.Clipboard]::SetImage($img);

	# Log the action.
	(Get-Date).tostring("dd-MM-yyyy HH:mm:ss") + "`t" + $whodat + "`t" + "Gen Member/Email (jpg)" + "`t" + $user | Out-File $logfile -Append
	[void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
	
	# Prompt the user when complete.
	$response = [Microsoft.VisualBasic.Interaction]::MsgBox("Details generated to $filename",'OKOnly,Information', "FYI")
	}
)



[void]$Form.ShowDialog() 
$Form.Dispose() 


