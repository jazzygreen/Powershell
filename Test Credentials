Add-Type -AssemblyName System.Windows.Forms

$Form = New-Object system.Windows.Forms.Form 
$Form.Text = "AD Credential Checker"
$Form.TopMost = $true
$Form.FormBorderStyle = 'Fixed3D'
$Form.MaximizeBox = $false
$Form.Width = 350
$Form.Height = 160


$uname = New-Object system.windows.Forms.Label 
$uname.Text = "Username"
$uname.AutoSize = $true
$uname.Width = 25
$uname.Height = 10
$uname.location = new-object system.drawing.point(10,10)
$uname.Font = "Microsoft Sans Serif,10"
$Form.controls.Add($uname) 

$unametext = New-Object system.windows.Forms.TextBox 
$unametext.Width = 210
$unametext.Height = 20
$unametext.location = new-object system.drawing.point(10,30)
$unametext.Font = "Microsoft Sans Serif,10"
$Form.controls.Add($unametext) 



$pword = New-Object system.windows.Forms.Label 
$pword.Text = "Password"
$pword.AutoSize = $true
$pword.Width = 25
$pword.Height = 10
$pword.location = new-object system.drawing.point(10,60)
$pword.Font = "Microsoft Sans Serif,10"
$Form.controls.Add($pword) 

$pwordtext = New-Object system.windows.Forms.TextBox 
$pwordtext.Width = 210
$pwordtext.Height = 20
$pwordtext.location = new-object system.drawing.point(10,80)
$pwordtext.Font = "Microsoft Sans Serif,10"
$Form.controls.Add($pwordtext) 



$searchbutton = New-Object system.windows.Forms.Button 
$searchbutton.Text = "Search"
$searchbutton.Width = 80
$searchbutton.Height = 30
$searchbutton.location = new-object system.drawing.point(240,20)
$searchbutton.Font = "Microsoft Sans Serif,10"
$Form.controls.Add($searchbutton) 


$passwordtext = New-Object system.windows.Forms.Label 
$passwordtext.Text = "Tested:"
$passwordtext.AutoSize = $true
$passwordtext.Width = 25
$passwordtext.Height = 10
$passwordtext.location = new-object system.drawing.point(240,70)
$passwordtext.Font = "Microsoft Sans Serif,10"
$Form.controls.Add($passwordtext) 

$resultbox = New-Object system.windows.Forms.Label 
$resultbox.Text = ""
$resultbox.AutoSize = $true
$resultbox.Width = 80
$resultbox.Height = 30
$resultbox.location = new-object system.drawing.point(250,90)
$resultbox.Font = "Consolas,10"
$Form.controls.Add($resultbox) 



$pwordtext.Add_KeyDown({
    if ($_.KeyCode -eq "Enter") {
        $searchbutton.PerformClick()
    }
})



$searchbutton.Add_Click(
	{    
	$unametext.enabled = $false
	$pwordtext.enabled = $false
	$searchbutton.enabled = $false

	$Domain = "stella"
	Add-Type -AssemblyName System.DirectoryServices.AccountManagement
	$ContextType = [System.DirectoryServices.AccountManagement.ContextType]::Domain
	$PrincipalContext = New-Object System.DirectoryServices.AccountManagement.PrincipalContext $ContextType,$Domain
	$resultbox.text = $PrincipalContext.ValidateCredentials($unametext.text,$pwordtext.text)

	$unametext.enabled = $true
	$pwordtext.enabled = $true
	$searchbutton.enabled = $true
	}
)


[void]$Form.ShowDialog() 
$Form.Dispose() 


