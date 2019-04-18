#if ($PSVersionTable.PSVersion -gt [Version]"2.0") {
#  powershell -Version 2 -File $MyInvocation.MyCommand.Definition
#  exit
#}

if (!(Get-Command -Name get-exchangeserver -ErrorAction SilentlyContinue)) {
	Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010
	. $env:ExchangeInstallPath\bin\RemoteExchange.ps1 
	Connect-ExchangeServer -auto -AllowClobber 
}


$Username = Read-Host 'Enter username of the user'
# $file =  Read-Host 'location for the output file to be saves to'      Export-csv "$file\$username-folderstats.csv"
$Mailbox = Get-Mailbox $Username 
$Mailbox | Select-Object alias | foreach-object {Get-MailboxFolderStatistics -Identity $_.alias | select-object Identity, ItemsInFolder, FolderSize } | sort-object FolderSize -desc | out-gridview
