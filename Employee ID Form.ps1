Add-Type -AssemblyName System.Windows.Forms

$pshost = get-host
$pswindow = $pshost.ui.rawui
$originalwindow = $pswindow.windowsize
$newsize = $pswindow.windowsize
$newsize.height = 5
$newsize.width = 10
$pswindow.windowsize = $newsize

$Form = New-Object system.Windows.Forms.Form 
$Form.Text = "Employee ID Tool"
$Form.TopMost = $true
$Form.FormBorderStyle = 'Fixed3D'
$Form.MaximizeBox = $false
$Form.Width = 290
$Form.Height = 175
$form.add_FormClosing({$pswindow.windowsize = $originalwindow})


$listbutton = New-Object system.windows.Forms.Button 
$listbutton.Text = "List"
$listbutton.Width = 40
$listbutton.Height = 25
$listbutton.location = new-object system.drawing.point(10,60)
$listbutton.Font = "Microsoft Sans Serif,10"
$Form.controls.Add($listbutton) 



$label2 = New-Object system.windows.Forms.Label 
$label2.Text = "Username"
$label2.AutoSize = $true
$label2.Width = 25
$label2.Height = 10
$label2.location = new-object system.drawing.point(10,10)
$label2.Font = "Microsoft Sans Serif,10"
$Form.controls.Add($label2) 

$unametext = New-Object system.windows.Forms.TextBox 
$unametext.Width = 160
$unametext.Height = 20
$unametext.location = new-object system.drawing.point(10,30)
$unametext.Font = "Microsoft Sans Serif,10"
$Form.controls.Add($unametext) 

$searchbutton = New-Object system.windows.Forms.Button 
$searchbutton.Text = "Search"
$searchbutton.Width = 95
$searchbutton.Height = 25
$searchbutton.location = new-object system.drawing.point(60,60)
$searchbutton.Font = "Microsoft Sans Serif,10"
$Form.controls.Add($searchbutton) 



$label2 = New-Object system.windows.Forms.Label 
$label2.Text = "ID"
$label2.AutoSize = $true
$label2.Width = 10
$label2.Height = 10
$label2.location = new-object system.drawing.point(190,10)
$label2.Font = "Microsoft Sans Serif,10"
$Form.controls.Add($label2) 

$empidtext = New-Object system.windows.Forms.TextBox 
$empidtext.Width = 70
$empidtext.Height = 20
$empidtext.location = new-object system.drawing.point(190,30)
$empidtext.Font = "Microsoft Sans Serif,10"
$Form.controls.Add($empidtext) 

$setbutton = New-Object system.windows.Forms.Button 
$setbutton.Text = "Set"
$setbutton.Width = 95
$setbutton.Height = 25
$setbutton.location = new-object system.drawing.point(165,60)
$setbutton.Font = "Microsoft Sans Serif,10"
$Form.controls.Add($setbutton) 



$resultbox = New-Object system.windows.Forms.Label 
$resultbox.Text = "Please search for a user."
$resultbox.AutoSize = $true
$resultbox.Width = 220
$resultbox.Height = 120
$resultbox.location = new-object system.drawing.point(10,100)
$resultbox.Font = "Consolas,10"
$Form.controls.Add($resultbox) 



$tzbox.Add_KeyDown({
    if ($_.KeyCode -eq "Enter") {
        $searchbutton.PerformClick()
    }
})

$unametext.Add_KeyDown({
    if ($_.KeyCode -eq "Enter") {
        $searchbutton.PerformClick()
    }
})


$listbutton.Add_Click(
	{
	
	$setbutton.enabled = $false
	$empidtext.enabled = $false
	$searchbutton.enabled = $false
	$listbutton.enabled = $false
	$unametext.enabled = $false
	$resultbox.text="Searching, please wait.."
	$date = (Get-Date).AddDays(-45)
	Get-ADUser -server gc-dc3.stella.ad -SearchBase "OU=Stella Hospitality,DC=Stella,DC=ad" -Filter 'employeeID -notlike "*"' -Properties * | where { $_.whenCreated -ge $date } | select displayName, sAMAccountName, physicalDeliveryOfficeName, employeeID, whenCreated | Out-gridview
	$resultbox.text="Please search for a user."
	$setbutton.enabled = $true
	$empidtext.enabled = $true
	$searchbutton.enabled = $true
	$listbutton.enabled = $true
	$unametext.enabled = $true
	}
)


$searchbutton.Add_Click(
	{    
	$setbutton.enabled = $false
	$empidtext.enabled = $false
	$searchbutton.enabled = $false
	$listbutton.enabled = $false
	$unametext.enabled = $false
	$username=$unametext.text
	
	if (Get-ADUser -filter "SamAccountName -eq '$username'"){
	$UserInfo = Get-ADUser -server gc-dc3.stella.ad -Identity $unametext.text -property employeeID
	$resultbox.text="Current Employee ID: "
	$resultbox.text+=$UserInfo.employeeid
	$setbutton.enabled = $true
	$searchbutton.enabled = $true
	$listbutton.enabled = $true
	$unametext.enabled = $true
	$empidtext.enabled = $true
	} else {
	$resultbox.text="Not found: "
	$resultbox.text+=$unametext.text
	$setbutton.enabled = $true
	$searchbutton.enabled = $true
	$listbutton.enabled = $true
	$unametext.enabled = $true
	$empidtext.enabled = $true
	}
	cls
	}
)


$setbutton.Add_Click(
	{    
	$setbutton.enabled = $false
	$empidtext.enabled = $false
	$searchbutton.enabled = $false
	$listbutton.enabled = $false
	$unametext.enabled = $false
	$username=$unametext.text
	$employee=$empidtext.text
	
	if (Get-ADUser -filter "SamAccountName -eq '$username'"){
	Set-ADUser -server gc-dc3.stella.ad -identity $username -EmployeeID $employee
	$resultbox.text="$username set to $employee"
	$setbutton.enabled = $true
	$searchbutton.enabled = $true
	$tzbox.enabled = $true
	$unametext.enabled = $true
	$empidtext.enabled = $true
	} else {
	$resultbox.text="Not found: "
	$resultbox.text+=$unametext.text
	$setbutton.enabled = $true
	$searchbutton.enabled = $true
	$listbutton.enabled = $true
	$unametext.enabled = $true
	$empidtext.enabled = $true
	}
	cls
	}
)


cls

$pswindow.windowsize = $newsize

[void]$Form.ShowDialog() 
$Form.Dispose() 
