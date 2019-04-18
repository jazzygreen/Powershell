Add-Type -AssemblyName System.Windows.Forms

$pshost = get-host
$pswindow = $pshost.ui.rawui
$originalwindow = $pswindow.windowsize
$newsize = $pswindow.windowsize
$newsize.height = 5
$newsize.width = 10
$pswindow.windowsize = $newsize


$Form = New-Object system.Windows.Forms.Form 
$Form.Text = "Email access searcher"
$Form.TopMost = $true
$Form.FormBorderStyle = 'Fixed3D'
$Form.MaximizeBox = $false
$Form.Width = 360
$Form.Height = 240
$form.add_FormClosing({$pswindow.windowsize = $originalwindow})


$label2 = New-Object system.windows.Forms.Label 
$label2.Text = "Username"
$label2.AutoSize = $true
$label2.Width = 25
$label2.Height = 10
$label2.location = new-object system.drawing.point(10,10)
$label2.Font = "Microsoft Sans Serif,10"
$Form.controls.Add($label2) 


$unametext = New-Object system.windows.Forms.TextBox 
$unametext.Width = 180
$unametext.Height = 20
$unametext.location = new-object system.drawing.point(10,30)
$unametext.Font = "Microsoft Sans Serif,10"
$Form.controls.Add($unametext) 


$searchbutton = New-Object system.windows.Forms.Button 
$searchbutton.Text = "Search"
$searchbutton.enabled = $false
$searchbutton.Width = 60
$searchbutton.Height = 20
$searchbutton.location = new-object system.drawing.point(270,30)
$searchbutton.Font = "Microsoft Sans Serif,10"
$Form.controls.Add($searchbutton) 


$setbutton = New-Object system.windows.Forms.Button 
$setbutton.Text = "Set OU"
$setbutton.Width = 60
$setbutton.Height = 20
$setbutton.location = new-object system.drawing.point(200,30)
$setbutton.Font = "Microsoft Sans Serif,10"
$Form.controls.Add($setbutton) 


$label2 = New-Object system.windows.Forms.Label 
$label2.Text = "Email search (OU not set):"
$label2.AutoSize = $true
$label2.Width = 25
$label2.Height = 10
$label2.location = new-object system.drawing.point(10,60)
$label2.Font = "Microsoft Sans Serif,10"
$Form.controls.Add($label2) 


$resultbox = New-Object system.windows.Forms.Label 
$resultbox.AutoSize = $true
$resultbox.Width = 220
$resultbox.Height = 170
$resultbox.location = new-object system.drawing.point(10,80)
$resultbox.Font = "Consolas,10"
$Form.controls.Add($resultbox) 


function Browse-AD()
{
    $dc_hash = @{}
    $selected_ou = $null

    
    $forest = Get-ADForest
    [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") | Out-Null
    [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null

    function Get-NodeInfo($sender, $dn_textbox)
    {
        $selected_node = $sender.Node
        $dn_textbox.Text = $selected_node.Name
    }

    function Add-ChildNodes($sender)
    {
        $expanded_node = $sender.Node

        if ($expanded_node.Name -eq "root") {
            return
        }

        $expanded_node.Nodes.Clear() | Out-Null

        $dc_hostname = $dc_hash[$($expanded_node.Name -replace '(OU=[^,]+,)*((DC=\w+,?)+)','$2')]
        if($dc_hostname -ne $null) {$child_OUs = Get-ADObject -Server $dc_hostname -Filter 'ObjectClass -eq "organizationalUnit" -or ObjectClass -eq "container"' -SearchScope OneLevel -SearchBase $expanded_node.Name}
        if($child_OUs -eq $null) {
            $sender.Cancel = $true
        } else {
            foreach($ou in $child_OUs) {
                $ou_node = New-Object Windows.Forms.TreeNode
                $ou_node.Text = $ou.Name
                $ou_node.Name = $ou.DistinguishedName
                $ou_node.Nodes.Add('') | Out-Null
                $expanded_node.Nodes.Add($ou_node) | Out-Null
            }
        }
    }

    function Add-ForestNodes($forest, [ref]$dc_hash)
    {
        $ad_root_node = New-Object Windows.Forms.TreeNode
        $ad_root_node.Text = $forest.RootDomain
        $ad_root_node.Name = "root"
        $ad_root_node.Expand()

        $i = 1
        foreach ($ad_domain in $forest.Domains) {
            # Write-Progress -Activity "Querying AD forest for domains and hostnames..." -Status $ad_domain -PercentComplete ($i++ / $forest.Domains.Count * 100)
            $dc = Get-ADDomainController -Server $ad_domain
            $dn = $dc.DefaultPartition
            $dc_hash.Value.Add($dn, $dc.Hostname)
            $dc_node = New-Object Windows.Forms.TreeNode
            $dc_node.Name = $dn
            $dc_node.Text = $dc.Domain
            $dc_node.Nodes.Add("") | Out-Null
            $ad_root_node.Nodes.Add($dc_node) | Out-Null
        }

        return $ad_root_node
    }
    
    $main_dlg_box = New-Object System.Windows.Forms.Form
    $main_dlg_box.ClientSize = New-Object System.Drawing.Size(400,600)
	$main_dlg_box.TopMost = $true
    $main_dlg_box.MaximizeBox = $false
    $main_dlg_box.MinimizeBox = $false
    $main_dlg_box.FormBorderStyle = 'FixedSingle'

    # widget size and location variables
    $ctrl_width_col = $main_dlg_box.ClientSize.Width/20
    $ctrl_height_row = $main_dlg_box.ClientSize.Height/15
    $max_ctrl_width = $main_dlg_box.ClientSize.Width - $ctrl_width_col*2
    $max_ctrl_height = $main_dlg_box.ClientSize.Height - $ctrl_height_row
    $right_edge_x = $max_ctrl_width
    $left_edge_x = $ctrl_width_col
    $bottom_edge_y = $max_ctrl_height
    $top_edge_y = $ctrl_height_row

    # setup text box showing the distinguished name of the currently selected node
    $dn_text_box = New-Object System.Windows.Forms.TextBox
    # can not set the height for a single line text box, that's controlled by the font being used
    $dn_text_box.Width = (14 * $ctrl_width_col)
    $dn_text_box.Location = New-Object System.Drawing.Point($left_edge_x, ($bottom_edge_y - $dn_text_box.Height))
    $main_dlg_box.Controls.Add($dn_text_box)
    # /text box for dN

    # setup Ok button
    $ok_button = New-Object System.Windows.Forms.Button
    $ok_button.Size = New-Object System.Drawing.Size(($ctrl_width_col * 2), $dn_text_box.Height)
    $ok_button.Location = New-Object System.Drawing.Point(($right_edge_x - $ok_button.Width), ($bottom_edge_y - $ok_button.Height))
    $ok_button.Text = "Ok"
    $ok_button.DialogResult = 'OK'
    $main_dlg_box.Controls.Add($ok_button)
    # /Ok button

    # setup tree selector showing the domains
    $ad_tree_view = New-Object System.Windows.Forms.TreeView
    $ad_tree_view.Size = New-Object System.Drawing.Size($max_ctrl_width, ($max_ctrl_height - $dn_text_box.Height - $ctrl_height_row*1.5))
    $ad_tree_view.Location = New-Object System.Drawing.Point($left_edge_x, $top_edge_y)
    $ad_tree_view.Nodes.Add($(Add-ForestNodes $forest ([ref]$dc_hash))) | Out-Null
    $ad_tree_view.Add_BeforeExpand({Add-ChildNodes $_})
    $ad_tree_view.Add_AfterSelect({Get-NodeInfo $_ $dn_text_box})
    $main_dlg_box.Controls.Add($ad_tree_view)
    # /tree selector

    $main_dlg_box.ShowDialog() | Out-Null

    return  $dn_text_box.Text
	
	cls
}


$unametext.Add_KeyDown({
    if ($_.KeyCode -eq "Enter") {
        $searchbutton.PerformClick()
    }
})


$searchbutton.Add_Click({
	$resultbox.text=""
	$setbutton.enabled = $false
	$searchbutton.enabled = $false
	$unamesearch=$unametext.text
	If ($unamesearch.length -gt 4) {
		$User = Get-Mailbox -Identity $unamesearch | Select-Object Alias
		$mball = Get-Mailbox -OrganizationalUnit "$global:pickedOU" -ResultSize Unlimited
		$mbfull = Get-Mailbox -OrganizationalUnit "$global:pickedOU" -ResultSize Unlimited | Get-MailboxPermission -User $User.Alias
		$mbsend = Get-Mailbox -OrganizationalUnit "$global:pickedOU" -ResultSize Unlimited | Get-RecipientPermission -trustee $User.alias
		foreach($line in $mball) {

			$hasfull=$null
			$hassend=$null
			if ($line | Get-MailboxPermission -user $user.alias) {$hasfull=" Full"}
			if ($line | Get-RecipientPermission -trustee $user.alias) {$hassend=" Send"}
			
			if (($hasfull -eq " Full") -or ($hassend -eq " Send")) {$resultbox.text+= $line.name + ":" + $hasfull + $hassend + "`n"}

		}
		# $resultbox.text+=$mbfull.identity.name
	} else {
		$resultbox.text="Please supply a valid username"
	}
	$setbutton.enabled = $true
	$searchbutton.enabled = $true
})


$setbutton.Add_Click({
	$global:pickedOU=Browse-AD
	$searchbutton.enabled = $true
	$label2.Text = "Email search:"
})


if (!(Get-Command -Name get-exchangeserver -ErrorAction SilentlyContinue)) {
	Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010
	. $env:ExchangeInstallPath\bin\RemoteExchange.ps1 
	Connect-ExchangeServer -auto -AllowClobber 
}

# Import-Module ActiveDirectory

# cls

$pswindow.windowsize = $newsize


[void]$Form.ShowDialog() 
$Form.Dispose() 


