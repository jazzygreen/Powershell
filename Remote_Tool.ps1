Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# Devine application locations
$psexec="C:\MG-IT\Applications\psexec.exe"
$sccmviewer="C:\Program Files (x86)\Microsoft Configuration Manager\AdminConsole\bin\i386\CmRcViewer.exe"
$vncviewer="C:\MG-IT\Applications\vncviewer.exe"

$pshost = get-host
$pswindow = $pshost.ui.rawui
$originalwindow = $pswindow.windowsize
$newsize = $pswindow.windowsize
$newsize.height = 10
$newsize.width = 50
$pswindow.windowsize = $newsize


$Form = New-Object system.Windows.Forms.Form 
$Form.Text = "Remote Tool"
$Form.TopMost = $true
$Form.FormBorderStyle = 'Fixed3D'
$Form.MaximizeBox = $false
$Form.Width = 415
$Form.Height = 437
$form.add_FormClosing({
	$pswindow.windowsize = $originalwindow
	cls
})


$label2 = New-Object system.windows.Forms.Label 
$label2.Text = "Hostname or IP"
$label2.AutoSize = $true
$label2.Width = 25
$label2.Height = 10
$label2.location = new-object system.drawing.point(10,10)
$label2.Font = "Microsoft Sans Serif,10"
$Form.controls.Add($label2) 


$compname = New-Object system.windows.Forms.TextBox 
$compname.Width = 163
$compname.Height = 20
$compname.location = new-object system.drawing.point(10,30)
$compname.Font = "Microsoft Sans Serif,10"
$Form.controls.Add($compname) 


$selectr = New-Object system.windows.Forms.ComboBox 
$selectr.Text = "Select Operation"
$selectr.items.add("View system information")
$selectr.items.add("Monitor stepping for 30 seconds")
$selectr.items.add("Check server logons")
$selectr.items.add("Check server open files")
$selectr.items.add("Check live-na1 open files")
$selectr.items.add("Check remote connections / Session ID (WMI)")
$selectr.items.add("Check remote connections / Session ID")
$selectr.items.add("Check Citrix connections")
$selectr.items.add("Logoff remote user")
$selectr.items.add("List local profiles")
$selectr.items.add("Cleanup local profile files")
$selectr.items.add("Delete lastlogonuser")
$selectr.items.add("Set lastlogonuser registry key")
$selectr.items.add("Remove lastlogonuser registry key")
$selectr.items.add("Set CSC FormatDatabase registry key")
$selectr.items.add("Check Security.evtx")
$selectr.items.add("Clear Security.evtx")
$selectr.items.add("Rebuild PC GPO")
$selectr.items.add("Close Screen Connect")
$selectr.items.add("Close SCCM remote tool")
$selectr.items.add("Fix SCCM Remote Control")
$selectr.items.add("Enable remote registry")
$selectr.items.add("Enable WinRM")
$selectr.items.add("Initiate interactive CMD")
$selectr.items.add("Initiate remote session (rdp view)")
$selectr.items.add("Initiate remote session (rdp control)")
$selectr.items.add("Initiate remote session (sccm)")
$selectr.items.add("Initiate remote session (vnc)")
$selectr.Width = 375
$selectr.Height = 20
$selectr.location = new-object system.drawing.point(10,60)
$selectr.Font = "Microsoft Sans Serif,10"
$Form.controls.Add($selectr) 


$label4 = New-Object system.windows.Forms.Label 
$label4.Text = "Sesh ID"
$label4.AutoSize = $true
$label4.location = new-object system.drawing.point(200,10)
$label4.Font = "Microsoft Sans Serif,10"
$Form.controls.Add($label4) 

$seshtext = New-Object system.windows.Forms.TextBox 
$seshtext.Width = 60
$seshtext.Height = 20
$seshtext.location = new-object system.drawing.point(200,30)
$seshtext.Font = "Microsoft Sans Serif,10"
$Form.controls.Add($seshtext) 


$executebutton = New-Object system.windows.Forms.Button 
$executebutton.Text = "Execute"
$executebutton.Width = 100
$executebutton.Height = 35
$executebutton.location = new-object system.drawing.point(286,19)
$executebutton.Font = "Microsoft Sans Serif,10"
$executebutton.FlatStyle ="Standard"
$Form.controls.Add($executebutton) 


$resultbox = New-Object System.Windows.Forms.TextBox
$resultbox.Text = "Please search for a user."
$resultbox.AutoSize = $true
$resultbox.MultiLine = $True
# $resultbox.ScrollBars = "Vertical"
$resultbox.location = new-object system.drawing.point(10,100)
$resultbox.Size = New-Object System.Drawing.Size(375,284)
$resultbox.Font = "Consolas,10"
$Form.controls.Add($resultbox) 




$seshtext.Add_KeyDown({
    if ($_.KeyCode -eq "Enter") {
        $executebutton.PerformClick()
    }
})

$selectr.Add_KeyDown({
    if ($_.KeyCode -eq "Enter") {
        $executebutton.PerformClick()
    }
})

$compname.Add_KeyDown({
    if ($_.KeyCode -eq "Enter") {
        $executebutton.PerformClick()
    }
})


$executebutton.Add_Click(
	{
	$UN=$compname.text
	$SID=$seshtext.text
	$compname.enabled = $false
	$selectr.enabled = $false
	$seshtext.enabled = $false
	$executebutton.enabled = $false
	
	$isonline=test-connection -buffersize 32 -count 1 -computername $un -quiet
	
	if ($selectr.SelectedItem -eq "View system information" -and $isonline) {
		$resultbox.Text="Please wait.."
		$SYSInfo = get-wmiobject -computername $UN -class win32_computersystemproduct
		$BIOSInfo = get-wmiobject -computername $UN -class win32_bios
		$CPUInfo = get-wmiobject -computername $UN -class win32_processor
		# $CPUInfo2 = Get-WmiObject -computername $UN Win32_PerfFormattedData_Counters_ProcessorInformation | where-object {$_.name -eq "_total"}
		$OSInfo = get-wmiobject -computername $UN -class win32_operatingsystem
		$Logoncount = @(Get-WmiObject -computer $UN -Query "Select * FROM Win32_Process WHERE Name='explorer.exe'" -ErrorAction SilentlyContinue)
		
		$resultbox.Text=$OSInfo.Caption+"`r`n"
		
		$cpuz=0
		foreach($line in $CPUInfo) {$cpuz++}
		$resultbox.Text+=$cpuz
		$resultbox.Text+=" x "
		if($cpuz -eq 1) {$resultbox.Text+=$CPUInfo.Name} else {$resultbox.Text+=$CPUInfo[0].Name}
		$resultbox.Text+="`r`n"
		
		$resultbox.Text+=($SYSInfo `
			| Format-List `
				@{Name="Hostname       ";Expression={$_.PSComputerName}},
				@{Name="PC Model       ";Expression={$_.Name}}  `
			| Out-String).Trim()
		$resultbox.Text+="`r`n"
		
		$resultbox.Text+=($BIOSInfo `
			| Format-List `
				@{Name="Service Tag   ";Expression={$_.SerialNumber}}, 
				@{Name="BIOS Version   ";Expression={$_.SMBIOSBIOSVersion}}  `
			| Out-String).Trim()
		$resultbox.Text+="`r`n"
		
		$resultbox.Text+="Memory          : "
		$resultbox.Text+=$OSInfo | % {[Math]::Round(($_.TotalVisibleMemorySize / 1MB))}
		$resultbox.Text+="GB, "
		$resultbox.Text+=$OSInfo | % {[Math]::Round(($_.FreePhysicalMemory / 1KB))}
		$resultbox.Text+="MB Free"
		$resultbox.Text+="`r`n"

		$resultbox.Text+="Cur Speed (MHz) : "+$CPUInfo.CurrentClockSpeed
		$resultbox.Text+="`r`n"
		
		#$resultbox.Text+=($CPUInfo2 `
		#	| Format-List `
		#		@{Name="CPU usage %    ";Expression={$_.PercentProcessorTime}}  `
		#	| Out-String).Trim()						
		#$resultbox.Text+="`r`n"
		$resultbox.Text+="Users logged on : "
		if ($Logoncount.Count -eq 0) {
		$resultbox.Text+="None`r`n"
		} else {
		$resultbox.Text+=$Logoncount.Count
		$resultbox.Text+="`r`n"
		}
		$LastBootedTime = [DateTimeOffset]::Parse([Management.ManagementDateTimeConverter]::ToDateTime($OSInfo.LastBootUpTime).ToString())
		$SystemUpTime = New-TimeSpan -Start $LastBootedTime.UtcDateTime -End ([DateTime]::UtcNow)
		$resultbox.Text+="Uptime          : "
		$resultbox.Text+=$SystemUpTime.Days
		$resultbox.Text+=" Days, "
		$resultbox.Text+=$SystemUpTime.Hours
		$resultbox.Text+=" Hours, "
		$resultbox.Text+=$SystemUpTime.Minutes
		$resultbox.Text+=" Mins`r`n"	
		
		$resultbox.Text+=($OSInfo `
			| Format-List `
				@{Name="OS Build Date  ";Expression={$_.ConvertToDateTime($_.InstallDate)}}  `
			| Out-String).Trim()
	}
	
	if ($selectr.SelectedItem -eq "Monitor stepping for 30 seconds" -and $isonline) {
	$looping=0
        while ($looping -ne 10) {
			$CPUInfo = get-wmiobject -computername $UN -class win32_processor
			$resultbox.Text=$CPUInfo.Name+"`r`n"
			$resultbox.Text+=($CPUInfo `
				| Format-List `
					@{Name="Max Speed (MHz)";Expression={$_.MaxClockSpeed}}, 
					@{Name="Current Speed (MHz)";Expression={$_.CurrentClockSpeed}}  `
				| Out-String).Trim()
			$CPUInfo2 = Get-WmiObject -computername $UN Win32_PerfFormattedData_Counters_ProcessorInformation | where-object {$_.name -eq "_total"}
			$resultbox.Text+="`r`n"
			$resultbox.Text+=($CPUInfo2 `
				| Format-List `
					@{Name="Processor usage %  ";Expression={$_.PercentProcessorTime}}  `
				| Out-String).Trim()
			start-sleep -seconds 3
			$looping++
		}
	# $resultbox.Text = "Please select an action.."
	}
	
    if ($selectr.SelectedItem -eq "Check remote connections / Session ID (WMI)" -and $isonline) {
		$resultbox.Text=""
		$sesh=""
		$looping=0
		$explorerprocesses = @(Get-WmiObject -computer $UN -Query "Select * FROM Win32_Process WHERE Name='explorer.exe'" -ErrorAction SilentlyContinue)
		if ($explorerprocesses.Count -eq 0)
		{
			$resultbox.Text= "`r`n`r`nNobody interactively logged on"
		}
		else
		{
			ForEach ($i in $explorerprocesses)
			{
				$sesh+="Username / Session : "+$i.GetOwner().User+" / "+$i.SessionId+"`r`n"
			}
		}


		$resultbox.Text=$sesh
	}

	if ($selectr.SelectedItem -eq "Check remote connections / Session ID" -and $isonline) {
		$resultbox.Text=""
		$sesh=""
		$looping=0
		Start-Process -Wait -Filepath "$psexec" -ArgumentList "-accepteula -s \\$UN query user" -WindowStyle Hidden -RedirectStandardOutput c:\temp\psout.txt
		if (Test-Path c:\temp\psout.txt) {
			foreach ($line in Get-Content -Path c:\temp\psout.txt) {
				$resultbox.Text+=$line.substring(1,20) + $line.substring(23,7) + $line.substring(40,12) + "`r`n"
			}
			# start-sleep -seconds 8
			} else {
			$resultbox.Text = "No processes found."
			# start-sleep -seconds 3
		}
		if (Test-Path c:\temp\psout.txt) {remove-item c:\temp\psout.txt -force}
		# $resultbox.Text = "Please select an action.."
	}
	
	if ($selectr.SelectedItem -eq "Check Citrix connections" -and $isonline) {
		$resultbox.Text="Server                 Connection      PID`r`n"
		$sesh=""
		$looping=0
		Start-Process -Wait -Filepath "$psexec" -ArgumentList "-accepteula -s \\$UN cmd.exe /s /c ""for /f ""tokens=2"" %g in ('tasklist ^| find /i ""wfica32.exe""') do @netstat -o | find /i ""%g""""" -WindowStyle Hidden -RedirectStandardOutput c:\temp\psout.txt
		if (Test-Path c:\temp\psout.txt) {
			foreach ($line in Get-Content -Path c:\temp\psout.txt) {
				$resultbox.Text+=$line.substring(32,43)+"`r`n"
			}
			# start-sleep -seconds 8
			} else {
			$resultbox.Text = "No processes found."
			# start-sleep -seconds 3
		}
		if (Test-Path c:\temp\psout.txt) {remove-item c:\temp\psout.txt -force}
		# $resultbox.Text = "Please select an action.."
	}
	
	if ($selectr.SelectedItem -eq "Check server logons" -and $isonline) {
		$resultbox.text="Checking sessions on: " + $UN
		Get-WmiObject -computername $UN -class win32_serverconnection | sort-object username -unique | select-object computername, username | out-gridview
		# $resultbox.Text = "Please select an action.."
	}

	if ($selectr.SelectedItem -eq "Check server open files" -and $isonline) {
		 function Get-OpenFile {
			Param (
				[string]$ComputerName
			)

			$openfiles = openfiles.exe /query /s $ComputerName /fo csv /V

			$openfiles | ForEach-Object {
				$line = $_
				if ($line -match '","') {$line}
			} | ConvertFrom-Csv
		}

		$Output = Get-OpenFile -ComputerName $UN

		$Output | Out-GridView
	}
	
	if ($selectr.SelectedItem -eq "Check live-na1 open files" -and $isonline) {
		 function Get-OpenFile {
			Param (
				[string]$ComputerName
			)

			$openfiles = openfiles.exe /query /s $ComputerName /fo csv /V

			$openfiles | ForEach-Object {
				$line = $_
				if ($line -match '","') {$line}
			} | ConvertFrom-Csv
		}

		#Add Hostfile Records for the Servers
		If ((Get-Content "$($env:windir)\system32\Drivers\etc\hosts") -notcontains "172.25.32.161 LIVE-NA1-1") {ac -Encoding UTF8 "$($env:windir)\system32\Drivers\etc\hosts" "172.25.32.161 LIVE-NA1-1"}
		If ((Get-Content "$($env:windir)\system32\Drivers\etc\hosts") -notcontains "172.25.32.162 LIVE-NA1-2") {ac -Encoding UTF8 "$($env:windir)\system32\Drivers\etc\hosts" "172.25.32.162 LIVE-NA1-2"}
		If ((Get-Content "$($env:windir)\system32\Drivers\etc\hosts") -notcontains "172.25.32.163 LIVE-NA1-3") {ac -Encoding UTF8 "$($env:windir)\system32\Drivers\etc\hosts" "172.25.32.163 LIVE-NA1-3"}
		If ((Get-Content "$($env:windir)\system32\Drivers\etc\hosts") -notcontains "172.25.32.164 LIVE-NA1-4") {ac -Encoding UTF8 "$($env:windir)\system32\Drivers\etc\hosts" "172.25.32.164 LIVE-NA1-4"}

		$Server1 = Get-OpenFile -ComputerName "Live-NA1"
		$Server2 = Get-OpenFile -ComputerName "Live-NA1-1"
		$Server3 = Get-OpenFile -ComputerName "Live-NA1-2"
		$Server4 = Get-OpenFile -ComputerName "Live-NA1-3"
		$Server5 = Get-OpenFile -ComputerName "Live-NA1-4"

		$Output = $Server1 + $Server2 + $Server3 + $Server4 +$Server5 
		$Output | Out-GridView
	}

	if ($selectr.SelectedItem -eq "Logoff remote user" -and $isonline) {
		if ($SID -ne "0") {
			$resultbox.text="Sending loogoff command to: " + $UN + " for Session ID: "+$SID
			Start-Process -Wait -Filepath "$psexec" -ArgumentList "-accepteula -d -s \\$UN logoff  $SID"  -WindowStyle Hidden
			$resultbox.text+="`r`nProcess complete."
			# start-sleep -seconds 3
			# $resultbox.text="Please select an action.."
		}
		else { 
			$resultbox.text= "Please set a Session ID first."
			# start-sleep -seconds 3
			# $resultbox.text="Please select an action.."
		}
	}
	
	if ($selectr.SelectedItem -eq "List local profiles" -and $isonline) {
			$resultbox.text="Checking local profiles on: " + $UN
			$ProfileList = Get-ChildItem \\$UN\c$\Users
			$resultbox.text=""
			ForEach ($i in $ProfileList)
			{
				$UserList+="User: "+$i.Name+"`r`n"
			}
			$resultbox.text=$UserList
	}

	if ($selectr.SelectedItem -eq "Cleanup local profile files" -and $isonline) {
		[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null
		$TargetProfile = [Microsoft.VisualBasic.Interaction]::InputBox("Enter a username to remove (ie; joe.smith)", "User")
		$resultbox.text="Removing folders for: "+$TargetProfile+" on: "+$UN+"`r`n`r`n"
		$resultbox.text+="Removing Office Recent`r`n"
		Start-Process -Wait -Filepath "$psexec" -ArgumentList "-accepteula -d -s \\$UN cmd /c rmdir /q /s c:\users\$TargetProfile\Appdata\Roaming\Microsoft\Office"  -WindowStyle Hidden
		$resultbox.text+="Removing Windows Recent`r`n"
		Start-Process -Wait -Filepath "$psexec" -ArgumentList "-accepteula -d -s \\$UN cmd /c rmdir /q /s c:\users\$TargetProfile\Appdata\Roaming\Microsoft\Windows\Recent"  -WindowStyle Hidden
		$resultbox.text+="Removing IE Cookies`r`n"
		Start-Process -Wait -Filepath "$psexec" -ArgumentList "-accepteula -d -s \\$UN cmd /c rmdir /q /s c:\users\$TargetProfile\Appdata\Roaming\Microsoft\Windows\Cookies"  -WindowStyle Hidden
		$resultbox.text+="Removing Crypto`r`n"
		Start-Process -Wait -Filepath "$psexec" -ArgumentList "-accepteula -d -s \\$UN cmd /c rmdir /q /s c:\users\$TargetProfile\Appdata\Roaming\Microsoft\Windows\Crypto"  -WindowStyle Hidden
		$resultbox.text+="`r`nProcess complete."
	}
	
	if ($selectr.SelectedItem -eq "Delete lastlogonuser" -and $isonline) {
		$resultbox.text= "Deleting lastlogonuser.."
		REG DELETE \\$UN\HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Authentication\LogonUI /v LastLoggedOnDisplayName /f
		REG DELETE \\$UN\HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Authentication\LogonUI /v LastLoggedOnProvider /f
		REG DELETE \\$UN\HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Authentication\LogonUI /v LastLoggedOnSAMUser /f
		REG DELETE \\$UN\HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Authentication\LogonUI /v LastLoggedOnUser /f
		REG DELETE \\$UN\HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Authentication\LogonUI /v LastLoggedOnUserSID /f
		# start-sleep -seconds 1
		$resultbox.text+= "`r`n`r`nKeys deleted."
		# start-sleep -seconds 3
		# $resultbox.text="Please select an action.."
    }
	
	if ($selectr.SelectedItem -eq "Set lastlogonuser registry key" -and $isonline) {
			$resultbox.text= "Ensuring registry access.."
			Set-Service -name remoteregistry -computername $UN -startuptype manual
			Get-Service -name remoteregistry -computername $UN | start-service
			# start-sleep -seconds 1
			$regtest=REG query \\$UN\HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System /v dontdisplaylastusername
			$regtest | foreach {
				$items = $_.split(" ")
				if ($items[12] -eq "0") {$global:regvar = $items[12]}
				if ($items[12] -eq "1") {$global:regvar = $items[12]}
			}
			if ($global:regvar -eq "0") {
				$resultbox.text+= "`r`n`r`nSetting registry key.."
				REG ADD \\$UN\HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System /v dontdisplaylastusername /t REG_SZ /d 1 /f
				# start-sleep -seconds 1
				$resultbox.text+= "`r`n`r`nKey set."
				# start-sleep -seconds 3
			}
			elseif ($global:regvar -eq "1") {
				$resultbox.text+= "`r`n`r`nKey is already set not to display."
				# start-sleep -seconds 3
			}
			else {
				$resultbox.text+= "`r`n`r`nKey not found.`r`n`r`nSetting registry key.."
				REG ADD \\$UN\HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System /v dontdisplaylastusername /t REG_SZ /d 1 /f
				# start-sleep -seconds 1
				$resultbox.text+= "`r`n`r`nKey set."
				# start-sleep -seconds 3
			}
			# $resultbox.text="Please select an action.."
    }
	
	if ($selectr.SelectedItem -eq "Remove lastlogonuser registry key" -and $isonline) {
			$resultbox.text= "Ensuring registry access.."
			Set-Service -name remoteregistry -computername $UN -startuptype manual
			Get-Service -name remoteregistry -computername $UN | start-service
			# start-sleep -seconds 1
			$regtest=REG query \\$UN\HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System /v dontdisplaylastusername
			$regtest | foreach {
				$items = $_.split(" ")
				if ($items[12] -eq "0") {$global:regvar = $items[12]}
				if ($items[12] -eq "1") {$global:regvar = $items[12]}
			}
			if ($global:regvar -eq "1") {
				$resultbox.text+= "`r`n`r`nSetting registry key.."
				REG ADD \\$UN\HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System /v dontdisplaylastusername /t REG_SZ /d 0 /f
				# start-sleep -seconds 1
				$resultbox.text+= "`r`n`r`nKey set."
				# start-sleep -seconds 3
			}
			elseif ($global:regvar -eq "0") {
				$resultbox.text+= "`r`n`r`nKey is already set to display."
				# start-sleep -seconds 3
			}
			else {
				$resultbox.text+= "`r`n`r`nKey not found.`r`n`r`nSetting registry key.."
				REG ADD \\$UN\HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System /v dontdisplaylastusername /t REG_SZ /d 0 /f
				# start-sleep -seconds 1
				$resultbox.text+= "`r`n`r`nKey set."
				# start-sleep -seconds 3
			}
			# $resultbox.text="Please select an action.."
    }
	
	if ($selectr.SelectedItem -eq "Set CSC FormatDatabase registry key" -and $isonline) {
		$resultbox.text= "Ensuring registry access.."
		Set-Service -name remoteregistry -computername $UN -startuptype manual
		Get-Service -name remoteregistry -computername $UN | start-service
		# start-sleep -seconds 1
		$resultbox.text+= "`r`n`r`nSetting CSC FormatDatabase registry key.."
		REG ADD \\$UN\HKLM\SYSTEM\CurrentControlSet\Services\Csc\Parameters /v FormatDatabase /t REG_DWORD /d 1 /f
		# start-sleep -seconds 1
		$resultbox.text+= "`r`n`r`nKey set."
		# start-sleep -seconds 3
		# $resultbox.text="Please select an action.."
    }
	
	if ($selectr.SelectedItem -eq "Check Security.evtx" -and $isonline) {
		$resultbox.text= "Current log size: "
		$resultbox.text+=[math]::round(((get-item \\$UN\c$\Windows\System32\winevt\Logs\Security.evtx).length)/1MB,2)
		$resultbox.text+="MB"
		# start-sleep -seconds 1
		# start-sleep -seconds 5
		# $resultbox.text="Please select an action.."
    }	
	
	if ($selectr.SelectedItem -eq "Clear Security.evtx" -and $isonline) {
		$resultbox.text= "Current log size: "
		$resultbox.text+=[math]::round(((get-item \\$UN\c$\Windows\System32\winevt\Logs\Security.evtx).length)/1MB,2)
		$resultbox.text+="MB"
		$dateme=(Get-Date).ToString('HHmm-dd-MM-yy')
		$evtbk="c:\windows\system32\winevt\logs\Security_$dateme.evtx"
		$resultbox.text+= "`r`n`r`nBacking up and clearing Security.evtx.."
		# start-sleep -seconds 1
		Start-Process -Wait -Filepath "$psexec" -ArgumentList "-accepteula -d -s \\$UN wevtutil cl Security" -WindowStyle Hidden
		$resultbox.text+= "`r`n`r`nSecurity events cleared."
		# start-sleep -seconds 1
		# start-sleep -seconds 5
		# $resultbox.text="Please select an action.."
    }	
	
	if ($selectr.SelectedItem -eq "Rebuild PC GPO" -and $isonline) {
		$resultbox.text= "Removing GPO database secedit.sdb.."
		remove-item \\$UN\C$\WINDOWS\security\Database\secedit.sdb -force
		# start-sleep -seconds 2
		$resultbox.text+= "`r`n`r`nRemoving GPO policies from HKLM.."
		REG DELETE \\$UN\HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies /f
		# start-sleep -seconds 2
		$resultbox.text+= "`r`n`r`nComplete."
		# start-sleep -seconds 3
		# $resultbox.text="Please select an action.."
    }
	
	if ($selectr.SelectedItem -eq "Close Screen Connect" -and $isonline) {
		Start-Process -Wait -Filepath "$psexec" -ArgumentList "-accepteula -s \\$UN cmd.exe /s /c ""for /f ""tokens=2"" %g in ('tasklist ^| find /i ""Connect""') do @taskkill /f /pid %g""" -WindowStyle Hidden -RedirectStandardOutput c:\temp\psout.txt
		if (Test-Path c:\temp\psout.txt) {
			$resultbox.Text="Machine: "+$UN
			foreach ($line in Get-Content -Path c:\temp\psout.txt) {
				$resultbox.Text+="`r`n"+$line
			}
			# start-sleep -seconds 1
		}
		if (Test-Path c:\temp\psout.txt) {remove-item c:\temp\psout.txt -force}
		# start-sleep -seconds 5
		# $resultbox.text="Please select an action.."
    }	
	
	if ($selectr.SelectedItem -eq "Close SCCM remote tool" -and $isonline) {
		(Get-WmiObject -computer $UN win32_process -Filter "name = 'scnotification.exe'").Terminate()
		Get-Service -name CmRcService -computername $UN | restart-service 
		# $resultbox.text="Please select an action.."
    }
	
	if ($selectr.SelectedItem -eq "Fix SCCM Remote Control" -and $isonline) {
		$sccmsid=Get-WMIObject -ComputerName $UN -Query "Select sid from Win32_Group where name = 'ConfigMgr Remote Control Users'"
		$sccmgroup="ConfigMgr Remote Control Users"
		if (!$sccmsid.sid) {
			$resultbox.text="No Remote Control group found, Creating.."
			$cn = [ADSI]"WinNT://$UN"
			$group = $cn.Create("Group",$sccmgroup)
			$group.setinfo()
			$group.description = "SCCM Remote Control access"
			$group.SetInfo()
			$sccmsid=Get-WMIObject -ComputerName $UN -Query "Select sid from Win32_Group where name = 'ConfigMgr Remote Control Users'"
			$resultbox.Text+="`r`n`r`nGroup added with sid: " + $sccmsid.sid
			
			$resultbox.Text+="`r`n`r`nAdding users added to group.."
			$grouptoadd="SG-IT Operations"
			([adsi]"WinNT://$UN/$sccmgroup,group").add("WinNT://STELLA/$grouptoadd")
			$grouptoadd="SG-IT TAC Users"
			([adsi]"WinNT://$UN/$sccmgroup,group").add("WinNT://STELLA/$grouptoadd")
			$grouptoadd="SG-SCCM-RemoteControl-Gubse"
			([adsi]"WinNT://$UN/$sccmgroup,group").add("WinNT://STELLA/$grouptoadd")
			$sccmsid=Get-WMIObject -ComputerName $UN -Query "Select sid from Win32_Group where name = 'ConfigMgr Remote Control Users'"
		} else {
			$resultbox.Text="Sid found: " + $sccmsid.sid
		}
			
		$resultbox.Text+="`r`n`r`nKerjiggering reg keys.."
		Set-Service -name remoteregistry -computername $UN -startuptype manual
		Get-Service -name remoteregistry -computername $UN | start-service
		$sccmkey = 'SOFTWARE\Microsoft\SMS\Client\Client Components\Remote Control'
		$sccmreg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $UN)
		$sccmregkey = $sccmreg.opensubkey($sccmkey,$true)
		$sccmregkey.SetValue("Access Level",2,[Microsoft.Win32.RegistryValueKind]::DWORD)
		[string[]]$sccmremoteusers = ("STELLA\SG-IT Operations","STELLA\SG-IT TAC Users","STELLA\SG-SCCM-RemoteControl-Gubse")
		$sccmregkey.SetValue("PermittedViewers",$sccmremoteusers,[Microsoft.Win32.RegistryValueKind]::MultiString)
		$sccmregkey.SetValue("Remote Control Security Group SID",$sccmsid.sid,[Microsoft.Win32.RegistryValueKind]::String)
		
		Set-Service -name CmRcService -computername $UN -StartupType Automatic
		Get-Service -name CmRcService -computername $UN | restart-service 
		# $resultbox.text="Please select an action.."
    }
	
	if ($selectr.SelectedItem -eq "Enable remote registry" -and $isonline) {
		$resultbox.text="Enabling/starting Remote Registry.."
		Set-Service -name remoteregistry -computername $UN -startuptype manual
		Get-Service -name remoteregistry -computername $UN | start-service
		# start-sleep -seconds 3
		# $resultbox.text="Please select an action.."
    }

	if ($selectr.SelectedItem -eq "Enable WinRM" -and $isonline) {
		$resultbox.text="Enabling/starting Remote Registry.."
		REG ADD \\$UN\HKLM\SOFTWARE\Policies\Microsoft\Windows\WinRM\Service /v AllowAutoConfig /t REG_DWORD /d 1 /f
		REG ADD \\$UN\HKLM\SOFTWARE\Policies\Microsoft\Windows\WinRM\Service /v IPv4Filter /t REG_SZ /d "*" /f
		REG ADD \\$UN\HKLM\SOFTWARE\Policies\Microsoft\Windows\WinRM\Service /v IPv6Filter /t REG_SZ /d "*" /f
		Set-Service -name remoteregistry -computername $UN -startuptype automatic
		Get-Service -name remoteregistry -computername $UN | start-service
		# start-sleep -seconds 3
		# $resultbox.text="Please select an action.."
    }
	
	if ($selectr.SelectedItem -eq "Initiate interactive CMD" -and $isonline) {
		$resultbox.text="Opening command prompt to: " + $UN
		Start-Process -Filepath "$psexec" -ArgumentList "-accepteula -s \\$UN cmd.exe"
		$resultbox.text+="`r`n`r`nCommand prompt opened."
		# start-sleep -seconds 3
		# $resultbox.text="Please select an action.."

	}	
	
	if ($selectr.SelectedItem -eq "Initiate remote session (rdp view)" -and $isonline) {
		if ($SID -ne "0") {
			mstsc /v:$UN /shadow:$SID
			# start-sleep -seconds 2
			# $resultbox.text="Please select an action.."
		}
		else { 
			$resultbox.text= "Please set a Session ID first."
			# start-sleep -seconds 2
			# $resultbox.text="Please select an action.."
		}
    }
	
	if ($selectr.SelectedItem -eq "Initiate remote session (rdp control)" -and $isonline) {
		if ($SID -ne "0") {
			mstsc /v:$UN /shadow:$SID /control
			# start-sleep -seconds 2
			# $resultbox.text="Please select an action.."
		}
		else { 
			$resultbox.text= "Please set a Session ID first."
			# start-sleep -seconds 2
			# $resultbox.text="Please select an action.."
		}
    }
	
	if ($selectr.SelectedItem -eq "Initiate remote session (sccm)" -and $isonline) {
		& "$sccmviewer" $UN
		# $resultbox.text="Please select an action.."
    }
	
	if ($selectr.SelectedItem -eq "Initiate remote session (vnc)" -and $isonline) {
		& "$vncviewer" $UN
		# $resultbox.text="Please select an action.."
    }
	
	if (-not $isonline) {
		$resultbox.text="$UN is not responding.."
		# start-sleep -seconds 3
		# $resultbox.text="Please select an action.."
	}
	
	$compname.enabled = $true
	$selectr.enabled = $true
	$seshtext.enabled = $true
	$executebutton.enabled = $true
	}
)

$seshtext.text="0"
$compname.text="localhost"

[void]$Form.ShowDialog() 
$Form.Dispose() 

