<#==============================================================================
         File Name : DHCP-Manager.ps1
   Original Author : Kenneth C. Mazie (kcmjr AT kcmjr.com)
                   : 
       Description : A multi-purpose tool for managing and monitoring DHCP.
                   : 
             Notes : Normal operation is with no command line options.  
                   : Commandline options intentionally left out to avoid accidents.
                   :
      Requirements : Requires the PowerShell DHCPServer extensions.
                   : 
                   : 
          Warnings : 
                   :   
             Legal : Public Domain. Modify and redistribute freely. No rights reserved.
                   : SCRIPT PROVIDED "AS IS" WITHOUT WARRANTIES OR GUARANTEES OF 
                   : ANY KIND. USE AT YOUR OWN RISK. NO TECHNICAL SUPPORT PROVIDED.
                   :
           Credits : Code snippets and/or ideas came from many sources including but 
                   : not limited to the following:
                   : https://github.com/n2501r/spiderzebra/blob/master/PowerShell/DHCP_Scope_Report.ps1
                   : 
    Last Update by : Kenneth C. Mazie                                           
   Version History : v1.00 - 08-16-22 - Original 
    Change History : v2.00 - 09-00-23 - Numerous operational & bug fixes
                   : v2.10 - 12-15-23 - Adjusted email options, report format, other minor bugs.
                   : v3.00 - 12-25-23 - Relocated private settings out to external config for publishing. 
                   : v3.10 - 01-25-24 - Altered email send so it always goes out if over 80 or 95 %
                   : v3.20 - 05-21-24 - Added color gradiations for % used column.  Added generalized 
                   :                    grey background.
                   : v4.00 - 07-12-24 - Fixed detection of failed purges.  Added new messaging on report
                   :                    for failed purges as well as attching log files to the email.  
                   :                    Altered detection of console verses IDE.  Shuffled order of 
                   :                    operations to better report stats.  Tweaked HTML report coloring.
                   : v4.10 - 07-15-24 - Added gradiated % colors back.  Accidentaly remove after v3.0
                   : v4.20 - 12-20-24 - Adjested order of header rows, minor verbiage. Added black bars.
                   : v4.21 - 12-26-24 - Adjusted current state table wording for clarity.
                   : v4.22 - 05-05-25 - Another adjustment to current state table wording for added clarity.  Added debug option.
                   #>$ScriptVer = "v4.22"<#
                   :                  
==============================================================================#>
Clear-Host

if (!(Get-Module -Name "dhcpserver")) {
    Try{
        Get-Module -ListAvailable "dhcpserver" | Import-Module | Out-Null
    }Catch{
        Write-host "DHCP Module was not Loaded.  Exiting..." -ForegroundColor Red
        Break
    }
}

#--[ Switch Run option to $True as appropriate ]-----------------
$Report = $true            #--[ Dump detected settings to a CSV file in script folder ]--
$FromFile = $false         #--[ Update scope names from a text file.  If enabled other options do not run ]--
$RunUpdate = $false        #--[ Apply new settings to selected scopes. ]--
$Reconcile = $false        #--[ If set to $true will reconcile any scope found to be over 90% utilized ]--
$Purge = $true             #--[ Will detect reservations with unique IDs longer than 26 characters and remove them ]-- 
$SendEmail = $False        #--[ Forces email to be sent ]--
$Debug = $False            #--[ When set to true this stops the log from generating when doing repeated testing ]-- 
#--[ If all are set to $false detected stats are displayed to screen and nothing else is done unless any scope is over 80% ]--

#$Credential = $host.ui.PromptForCredential("Encrypted credential file Not found:", "Please enter your Domain\UserName and Password.", "", "NetBiosUserName") 
$DateTime = Get-Date -Format MM-dd-yyyy_HHmmss 

Function GetConsoleHost ($ExtOption){  #--[ Detect if we are using a script editor or the console ]--
    $ExtOption | Add-Member -MemberType NoteProperty -Name "Console" -Value $False -force
    Switch ($Host.Name){
        'consolehost'{
            $ExtOption | Add-Member -MemberType NoteProperty -Name "Console" -Value $true -force
            Write-Host "PowerShell Console detected. Output suppressed." -ForegroundColor Cyan
        }
        'Windows PowerShell ISE Host'{
            $ExtOption | Add-Member -MemberType NoteProperty -Name "Console" -Value $True -force
            Write-Host "PowerShell ISE editor detected." -ForegroundColor Cyan
        }
        'PrimalScriptHostImplementation'{
            $ExtOption | Add-Member -MemberType NoteProperty -Name "Console" -Value $True -force
            Write-Host "PrimalScript or PowerShell Studio editor detected." -ForegroundColor Cyan
        }
        "Visual Studio Code Host" {
            $ExtOption | Add-Member -MemberType NoteProperty -Name "Console" -Value $True -force
            Write-Host "Visual Studio Code editor detected." -ForegroundColor Cyan
        }
    }
    Return $ExtOption
}

Function StatusMsg ($Msg, $Color, $Console){
    If (($Host.Name -ne "consolehost") -or ($Console -eq $True)){   #--[ Only write status to screen if in an editor ]--
        Write-Host $Msg -ForegroundColor $Color
    }
}

Function LoadConfig ($ConfigFile){  #--[ Read and load configuration file ]-------------------------------------
    if (Test-Path -Path $ConfigFile -PathType Leaf){                       #--[ Error out if configuration file doesn't exist ]--
        [xml]$Config = Get-Content $ConfigFile  #--[ Read & Load XML ]--  
        $ExtOption = New-Object -TypeName psobject   
        $ExtOption | Add-Member -Force -MemberType NoteProperty -Name "Recipient0" -Value $Config.Settings.Email.Recipient0
        $ExtOption | Add-Member -Force -MemberType NoteProperty -Name "Recipient1" -Value $Config.Settings.Email.Recipient1
        $ExtOption | Add-Member -Force -MemberType NoteProperty -Name "Recipient2" -Value $Config.Settings.Email.Recipient2
        $ExtOption | Add-Member -Force -MemberType NoteProperty -Name "Sender" -Value $Config.Settings.Email.Sender
        $ExtOption | Add-Member -Force -MemberType NoteProperty -Name "SmtpServer" -Value $Config.Settings.Email.SmtpServer
        $ExtOption | Add-Member -Force -MemberType NoteProperty -Name "SaveFile" -Value ($PSScriptRoot+$Config.Settings.General.SaveFile+$DateTime+".txt")
        $ExtOption | Add-Member -Force -MemberType NoteProperty -Name "PurgeFile" -Value ($PSScriptRoot+$Config.Settings.General.PurgeFile+$DateTime+".txt")
        $ExtOption | Add-Member -Force -MemberType NoteProperty -Name "SiteServerPrefix" -Value $Config.Settings.General.SiteServerPrefix
        $ExtOption | Add-Member -Force -MemberType NoteProperty -Name "NewPrefix" -Value $Config.Settings.General.NewPrefix
        $ExtOption | Add-Member -Force -MemberType NoteProperty -Name "DnsArray" -Value $Config.Settings.Update.DnsArray
        $ExtOption | Add-Member -Force -MemberType NoteProperty -Name "NtpArray" -Value $Config.Settings.Update.NtpArray
        $ExtOption | Add-Member -Force -MemberType NoteProperty -Name "DomainArray" -Value $Config.Settings.Update.DomainArray
        $ExtOption | Add-Member -Force -MemberType NoteProperty -Name "PurgeFail" -Value $false
        $ExtOption | Add-Member -Force -MemberType NoteProperty -Name "Debug" -Value $False
    }Else{
        StatusMsg "MISSING XML CONFIG FILE.  File is required.  Script aborted..." " Red" $ExtOption.Console
        break;break;break
    }
    Return $ExtOption
}

Function Detect1 ($ScopeName,$ScopeID,$Option,$OptionID,$OptionArray){
    Try{
        $Detect = Get-DhcpServerv4OptionValue -ScopeId $ScopeID -OptionId $OptionID -ErrorAction SilentlyContinue
        $Detected = $Detect.Value.Split(" ")
    }Catch{
        $Detected = $Null
    }
    $Msg = '   '+$Option+' Entries : '+$Detected.Count
    StatusMsg $Msg "Yellow" $ExtOption.Console
    If (($Detected.Count -lt $OptionArray.Count) -Or ($Detected.Count -gt $OptionArray.Count)){
        $Update = $true
    }
    $Count = 1
    ForEach ($Item in $Detected){
        If ($OptionArray.Contains($Item)){
            Write-host '        '$Option' '$Count' : '$Item
        }Else{
            Write-host '        '$Option' '$Count' : '$Item
            $Update = $true
        }
        $Count++
    }
    Return $Update
}

Function Detect2 ($Value,$ScopeID,$VendorClass){
    $Detect = Get-DhcpServerv4OptionValue -ScopeId $ScopeID -VendorClass $VendorClass -ErrorAction SilentlyContinue
    If ($Detect.Value -eq $Value){
        $Msg = " - Value Verified -"
        StatusMsg $Msg "Green" $ExtOption.Console
    }Else{
        $Msg = " - Value Not Verified -"
        StatusMsg $Msg "Red" $ExtOption.Console
    } 
}

Function OptionUpdate ($ScopeName,$ScopeID,$Option,$OptionID,$OptionArray){   #==[ Update Scope Option ]=====================
    $Update = Detect1 $ScopeName $ScopeID $Option $OptionID $OptionArray
    If (($ScopeName -notlike "*DMZ*") -And ($Update)){
        Remove-DhcpServerv4OptionValue -ScopeId $ScopeID -OptionId $OptionID -ErrorAction SilentlyContinue -WhatIf
        Start-Sleep -Milliseconds 500
        If ($Option -eq "DNS"){
            Set-DhcpServerv4OptionValue -ScopeId $ScopeID -DnsServer $OptionArray.Split(",") -WhatIf
        }Else{
            Set-DhcpServerv4OptionValue -ScopeId $ScopeID -OptionId $OptionID $OptionArray.Split(",") -WhatIf
        }
        $Update = Detect1 $ScopeName $ScopeID $Option $OptionID $OptionArray
    }
    $Update = $False    
}

Function Repair ($ScopeID){
    Repair-DhcpServerv4IPRecord -ScopeId $ScopeID -Force 
}

#--[ End of Functions ]-------------------------------------

#==[ Begin ]================================================

#--[ Load external XML options file ]--
$ConfigFile = $PSScriptRoot+"\"+($MyInvocation.MyCommand.Name.Split("_")[0]).Split(".")[0]+".xml"
#$ConfigFile = $PSScriptRoot+"\"+$MyInvocation.MyCommand.Name+".xml"
#--[ Note: The above line strips off anything after an underscore so that script copies will use a common config file.        ]--        
#--[       This is handy when creating copies of the working script for testing.  To disable, use the commented line instead. ]--
$ExtOption = LoadConfig $ConfigFile
$ExtOption = GetConsoleHost $ExtOption

$HexBlack = '#000000'
$HexWhite = '#ffffff'
$HexBlue = '#0000ff'
$HexLtBlue = '#5a5dfa'
$HexCyan = '#03f0fc'
$HexLtGrey = '#CCCCCC'
$HexDkGrey = '#646464'
$HexYellow = '#FFF284'
$HexRed = '#FF0000'
$HexMaroon = '#a31818'
$HexGreen = '#006600'

$Total = 0
$SiteServer = (Get-DhcpServerInDC).DnsName | Where-Object {$_.SubString(0,3) -eq $ExtOption.SiteServerPrefix}

If ($ExtOption.Console){
    Write-host "`n`n--[ Begin ]------------------------------------" -foregroundcolor Yellow
    Write-host "`n`n   DHCP Server : " -ForegroundColor Yellow -NoNewline
    Write-Host $SiteServer -ForegroundColor White
}

If ($Debug){
    $ExtOption | Add-Member -Force -MemberType NoteProperty -Name "Debug" -Value $true
}

Try {
    $SiteScopes = $SiteServer | ForEach-Object {Get-DhcpServerv4Scope }
}Catch{
    Write-host "No Scopes were detected.  Is this system a DHCP server?" -ForegroundColor Red
    Write-host "Exiting..." -foregroundcolor red
    Break
}
$ScopeCount = $SiteScopes.count

If ($ExtOption.Console){
    write-host "   Scope Count : " -ForegroundColor Yellow -NoNewline
    Write-Host $ScopeCount -ForegroundColor White
}

If ($FromFile){  #--[ Load Text File With Updated Scope Names ]--
    $NewNames = Get-Content $PSScriptRoot'\new-scope-names.txt'
    ForEach ($Item in $NewNames){
        $ScopeID = $Item.Split(",")[0]
        $Name = $Item.Split(",")[1]        
        Try{
            Set-DhcpServerv4Scope -ScopeId $ScopeID -Name $Name #-whatif
            write-host $ScopeID"    "$Name
        }Catch{

        }
    }
    break
}

If ($Report){
    Add-Content -Path $ExtOption.SaveFile -Value "ScopeID,ScopeName,ScopeDescription,ScopeState,PercentUsed"
}

#--[ HTML Report Header ]--
$Data = '<table border="3" width="100%" ><tbody>
    <tr bgcolor='+$HexDkGrey+'>
    <td align="center"> <strong> <font color="white" size="2" face="tahoma" >Scope ID</font></strong></td>
    <td align="center"> <strong> <font color="white" size="2" face="tahoma" >Scope Name</font></strong></td>
    <td align="center"> <strong> <font color="white" size="2" face="tahoma" >Scope Description</font></strong></td>
    <td align="center"> <strong> <font color="white" size="2" face="tahoma" >Scope State</font></strong></td>
    <td align="center"> <strong> <font color="white" size="2" face="tahoma" >Total Addr</font></strong></td>
    <td align="center"> <strong> <font color="white" size="2" face="tahoma" >Addr In Use</font></strong></td>
    <td align="center"> <strong> <font color="white" size="2" face="tahoma" >Addr Free</font></strong></td>
    <td align="center"> <strong> <font color="white" size="2" face="tahoma" >% In Use</font></strong></td>
    <td align="center"> <strong> <font color="white" size="2" face="tahoma" >Reserved</font></strong></td>
    <td align="center"> <strong> <font color="white" size="2" face="tahoma" >Subnet Mask</font></strong></td>
    <td align="center"> <strong> <font color="white" size="2" face="tahoma" >Start of Range</font></strong></td>
    <td align="center"> <strong> <font color="white" size="2" face="tahoma" >End of Range</font></strong></td>
    <td align="center"> <strong> <font color="white" size="2" face="tahoma" >Lease Duration</font></strong></td>
    </tr>'

#--[ Cycle through detected scopes ]--
$Over80 = 0
$Over95 = 0
$Disabled = 0
$PurgeFailCount = 0
foreach ($Scope in $SiteScopes ){
    $PurgeFail = $False
    $ScopeID = $Scope.ScopeID.ipaddresstostring
    $ScopeName = $Scope.name
    $ScopeDescription = $Scope.Description
    $ScopeStatus = $Scope.State
    $ScopeStats = Get-DhcpServerv4ScopeStatistics -ScopeId $ScopeID
    $AddrInUse = [int]$ScopeStats.InUse 
    $AddrFree = [int]$ScopeStats.Free
    $AddrPercent = [int]$ScopeStats.PercentageInUse
    $AddrTotal = $AddrFree+$AddrInUse
    $Total = $Total+[int]$ScopeStats.InUse 

    If ($ExtOption.Console){  #--[ Display running results if console is enabled ]--
        write-host `n"        Scope ID : " -ForegroundColor Yellow -NoNewline
        write-host $ScopeID"   " -ForegroundColor White        
        write-host "     Scope Name  : " -ForegroundColor Yellow -NoNewline
        Write-host $ScopeName -ForegroundColor White
        write-host "    Description  : " -ForegroundColor Yellow -NoNewline       
        Write-host $ScopeDescription -ForegroundColor White
        write-host "   Scope Status  : " -ForegroundColor Yellow -NoNewline
        If ($ScopeStatus -like "*inactive*"){
            Write-host $ScopeStatus"   " -ForegroundColor Red
        }Else{
            Write-host $ScopeStatus"   " -ForegroundColor Green
        }
        write-host "     Statistics  : " -ForegroundColor Yellow -NoNewline
        write-host "Total Addr  ="$AddrTotal -ForegroundColor white
        write-host "                   Addr In Use ="$AddrInUse -ForegroundColor white
        write-host "                   Addr Free   ="$AddrFree -ForegroundColor White 
        If ($AddrPercent -gt 90){
            write-host "                   % In Use    = " -NoNewline
            Write-host $AddrPercent -ForegroundColor White -NoNewline
            write-host "                   % Free      = " -NoNewline
            Write-host (100-$addrPercent) -ForegroundColor Red 
            If ($Reconcile){
                Write-host "  -- Reconciling Scope --" -forground -ForegroundColor Magenta
                Repair $ScopeID
            }Else{
                Write-host " "
            }
        }Else{
            write-host "                   % In Use    = " -NoNewline
            Write-host $AddrPercent -ForegroundColor White 
            write-host "                   % Free      = " -NoNewline
            Write-host (100-$addrPercent) -ForegroundColor Green 
        }
        write-host "  Lease Duration : " -ForegroundColor Yellow -NoNewline
        [string]$lease = $Scope.LeaseDuration

        If($Lease -like "*.*"){
            $Days = [String]$Lease.Split(".")[0]
            $Hours = [String]($Lease.Split(".")[1]).Split(":")[0]
        }Else{
            $Days = 0
            $Hours = [String]$Lease.Split(":")[0]
        }
        $Min = [String]$Lease.Split(":")[1]
        $Sec = [String]$Lease.Split(":")[2]
        Write-host "$Days (Days)   $Hours (Hours)   $Min  (Min)   $Sec  (Sec)" -ForegroundColor White
    }

    If ($Purge){
        If ($ExtOption.Console){
            Write-Host "                 -- Bad Lease Purge Enabled ---" -ForegroundColor Cyan 
        }
        $LeaseInfo = Get-dhcpserverv4lease -ScopeId $ScopeID  | Select-object Ipaddress,addressstate,clientid,hostname,leaseexpirytime
        ForEach ($Item in $LeaseInfo){
            If ($Item.ClientID.Length -gt 26){
                $Timeout = 100
                $Ping = New-Object System.Net.NetworkInformation.Ping
                $Response = $Ping.Send($Item.Ipaddress,$Timeout)
                If ($Response.Status -eq "Success"){
                    If ($ExtOption.Console){
                        write-host "  --"$Item.Ipaddress"Is Alive." -ForegroundColor Green -NoNewline
                        write-host "    ("$Item.clientid")`n"
                    }
                }Else{   
                    $SendEmail = $true               
                    Try{
                        Remove-DhcpServerv4Lease -ScopeId $ScopeID -ClientId $Item.clientid -confirm:$false -ErrorAction:Stop  #-whatif
                        $Msg = 'Deleted lease: "'+$Item.Ipaddress+'" to: "'+$Item.clientID+'" from scope: "'+$ScopeDescription
                        Add-Content -Path $ExtOption.PurgeFile -Value $Msg
                        If ($ExtOption.Console){
                            write-host $Item.Ipaddress -ForegroundColor red -NoNewline
                            write-host "    "$Item.clientid
                        }
                    }Catch{
                        $ErrorMessage = $_.Exception.Message
                        $FailedItem = $_.Exception.ItemName
                        $PurgeFail = $True
                        $PurgeFailCount++
                        $Msg = 'FAILED to delete lease: "'+$Item.Ipaddress+'" to: "'+$Item.clientID+'" from scope: "'+$ScopeDescription
                        Add-Content -Path $ExtOption.PurgeFile -Value $Msg
                        $ErrorMAC = $ErrorMessage.Split(" ")[4]
                        If ($ExtOption.Console){
                            Write-Host "  -- Failed to delete IP lease for"$ErrorMAC -ForegroundColor Yellow
                            #write-host $ErrorMessage -ForegroundColor yellow
                            write-host $FailedItem -ForegroundColor cyan
                        }
                    }
                }
            }
        }
    }

    If ($Report){
        $SaveValue = $ScopeID+","+$ScopeName+","+$ScopeDescription+","+$ScopeStatus+","+$AddrPercent
        Add-Content -Path $ExtOption.SaveFile -Value $SaveValue
    }
    
    If ($RunUpdate){    #--[ Assorted scope adjustments.  !!! Comment these in or out as needed !!! ]--
        #==[ DNS Update ]================================================================
        OptionUpdate $ScopeName $ScopeID "DNS" "006" $ExtOption.DnsArray
        $Msg = "               : Updating DNS Option 006..."
        StatusMsg $Msg "Cyan" $ExtOption.Console

        #==[ NTP Update ]================================================================
        OptionUpdate $ScopeName $ScopeID "NTP" "042" $ExtOption.NtpArray
        $Msg = "               : Updating NTP Option 042..."
        StatusMsg $Msg "Cyan" $ExtOption.Console

        #==[ Domain Update ]=============================================================
        OptionUpdate $ScopeName $ScopeID "DOM" "015" $ExtOption.DomainArray
        $Msg = "               : Updating Domain Option 015..."
        StatusMsg $Msg "Cyan" $ExtOption.Console
  
        #==[ Scope Name Adjustment ]=====================================================
        If ($ScopeName -notlike "$ExtOption.NewPrefix*"){
            $NewName = "$ExtOption.NewPrefix "+$ScopeName
            Set-DhcpServerv4Scope -ScopeId $ScopeID -Name $NewName -whatif
            $Msg = "               : Updating Scope Name..."
            StatusMsg $Msg "Cyan" $ExtOption.Console
        } #>

        #==[ Description Update (Forces description to match scope name) ]===========================================
        #--[  ]--
        Set-DhcpServerv4Scope -ScopeId $ScopeID -Description $ScopeName -whatif
        $Msg = "               : Updating Description..."
        StatusMsg $Msg "Cyan" $ExtOption.Console
        #>

        <#==[ Wireless Extras ]==============================================================
        If (($ScopeName -like "*WIFI*") -or (($ScopeID -like "10.10*") -and ($ScopeName -like "*DATA*"))){
            $Value = "10.10.10.8"
            $OptionID = "241"

            $VendorClass = "Cisco AP 3700"
            write-host "     Applying : " -ForegroundColor Yellow -NoNewline
            write-host 'OptionId:'$OptionID' - VendorClass: '$VendorClass' - Value: '$Value -ForegroundColor Magenta -nonewline
           # Set-DhcpServerv4OptionValue -ScopeId $ScopeID -OptionId $OptionID -VendorClass $VendorClass -Value $Value #-WhatIf
            Detect2 $Value $ScopeID "Cisco AP 3700" 
           
            Start-Sleep -milliseconds 500    
            $VendorClass = "Cisco AP 2800"
            write-host "     Applying : " -ForegroundColor Yellow -NoNewline
            write-host 'OptionId:'$OptionID' - VendorClass: '$VendorClass' - Value: '$Value -ForegroundColor Magenta -nonewline
            #Set-DhcpServerv4OptionValue -ScopeId $ScopeID -OptionId $OptionID -VendorClass $VendorClass -Value $Value #-WhatIf
            Detect2 $Value $ScopeID $VendorClass        
        } #>
    }
    If ($AddrPercent -ge 95){
        $Data = $Data + '<tr bgcolor='+$HexRed+'><font color='+$HexYellow+'><strong>'
        $SendEmail = $True
        $Over95++
    }ElseIf ($AddrPercent -ge 80){
        $Data = $Data + '<tr bgcolor='+$HexYellow+'><strong>' 
        $SendEmail = $True
        $Over80++
    }Else{
        $Data = $Data + '<tr bgcolor='+$HexLtGrey+'>'
    }
    $Data = $Data + "<td  align='center'>$($ScopeId)</td>"
    If ($PurgeFail){
        $Data = $Data + "<td  align='center'><span style=background-color:$HexRed>
        <font color=$HexYellow>$($ScopeName)</Font></td>"
    }Else{
        $Data = $Data + "<td  align='center'><span style=background-color:$HexLtGrey>
        <font color=$HexBlack>$($ScopeName)</Font></td>"
    }    
    $Data = $Data + "<td  align='center'>$($ScopeDescription)</td>"
    If ($ScopeStatus -eq "Inactive"){
        $Data = $Data + "<td align='center'><font color='$HexMaroon'><strong>$($ScopeStatus)</td>"
        $Disabled++
    }Else{
        $Data = $Data + "<td  align='center'><font color='$HexGreen'>$($ScopeStatus)</td>"
    }
    
    $Data = $Data + "<td  align='center'>$($AddrTotal)</td>"
    $Data = $Data + "<td  align='center'>$($AddrInUse)</td>"
    $Data = $Data + "<td  align='center'>$($AddrFree)</td>"

    [int]$Percentage = [Int]$AddrPercent
            
    #--[ See https://www.w3schools.com/colors/colors_mixer.asp for color mix info ]--
    If ($Percentage -gt 95){$BGColor = "#FF0000"}
    If ($Percentage -le 95){$BGColor = "#ff4000"}
    If ($Percentage -le 90){$BGColor = "#ff6600"}    
    If ($Percentage -le 85){$BGColor = "#ef8300"}
    If ($Percentage -le 80){$BGColor = "#e29a00"}
    If ($Percentage -le 75){$BGColor = "#d9ab00"}
    If ($Percentage -le 70){$BGColor = "#cfbc00"}
    If ($Percentage -le 65){$BGColor = "#c9c800"}
    If ($Percentage -le 60){$BGColor = "#c2d300"}    
    If ($Percentage -le 55){$BGColor = "#bfd900"}
    If ($Percentage -le 50){$BGColor = "#b2d100"}
    If ($Percentage -le 45){$BGColor = "#a6c900"}
    If ($Percentage -le 40){$BGColor = "#99c200"}
    If ($Percentage -le 35){$BGColor = "#8cba00"}
    If ($Percentage -le 30){$BGColor = "#80b200"}
    If ($Percentage -le 25){$BGColor = "#73ab00"}
    If ($Percentage -le 20){$BGColor = "#66a300"}
    If ($Percentage -le 15){$BGColor = "#4c9400"}
    If ($Percentage -le 10){$BGColor = "#338500"}
    If ($Percentage -le 5) {$BGColor = "#1a7500"}
    If ($Percentage -le 1) {$BGColor = "#006600"}

    If ($Percentage -ge 85){                                               
        $Data = $Data += '<td bgcolor=' + $BGColorYel + '><font color=#ff0000><strong>' + $Percentage + ' %</strong></td>'       #--[ Add yellow on red if <20% volume percent free ]--
    #}ElseIf ($Percentage -lt 85){
    # $RowData += '<td bgcolor=' + $BGColor + '><font color=#ffffff><strong>' + $Percentage + ' %</strong></td>' #--[ Add yellow on red volume percent free ]--
    }ElseIf ($Percentage -lt 10){                             
        $Data = $Data + '<td bgcolor=' + $BGColor + '><font color=#ffffff><strong>' + $Percentage + ' %</strong></td>'           #--[ Add yellow on red volume percent free ]--
    }Else{                                                               
        $Data = $Data += '<td bgcolor=' + $BGColor + '><font color=#000000><strong>' + $Percentage + ' %</strong></td>'           #--[ Add volume percent free ]--
    }    

    $Data = $Data + "<td  align='center'>$($ScopeStats.Reserved)</td>"
    $Data = $Data + "<td  align='center'>$($Scope.SubnetMask)</td>"
    $Data = $Data + "<td  align='center'>$($Scope.StartRange)</td>"
    $Data = $Data + "<td align='center'>$($Scope.EndRange)</td>"
    $Data = $Data + "<td align='center'>$($Scope.LeaseDuration)</td>"
    $Data = $Data + "</tr>"     
}
$Data = $Data + "</table></table>
Script Version: "+$ScriptVer+" 
</body></html>"

#--[ HTML Header ]-----------------------------------------------------
$HTMLHeader = "
<html>
    <head>
        <meta http-equiv='Content-Type' content='text/html; charset=iso-8859-1'>
        <title>DHCP Status Report</title>
        <STYLE TYPE='text/css'>
            table {
                width: 100%
                padding: 1px;
                border: 1px solid black;
                border-collapse: collapse;
            }
            th {
                width 100%
            }
            td {
                padding: 1px;
                border: 1px solid black;
                border-collapse: collapse;
            }
            tr#one {
              background:black;
              border: 6px solid white;
              border-collapse: collapse;
            }
        </style>
    </head>
<body>
<table-layout: fixed>
"

#--[ Report Headers ]------------------------------------------------------
#--[ HTML title bar ]--
$ReportHeader = "
<table border='0' width='100%'>
    <table border='0' width='100%'>
    <tr bgcolor=$HexCyan>       
            <td colspan='7' height='25' align='center'><strong>
            <font color=$HexBlack size='4' face='tahoma'>DHCP Scope Statistics Report for $SiteServer&nbsp;&nbsp;&nbsp;&nbsp;</font>
            <font color=$HexBlack size='4' face='tahoma'> ($(Get-Date))</font>
            <font color=$HexBlack size='2' face='tahoma'> <BR>Total Scope Count = $ScopeCount</font>
        </tr>
    </table>"

#--[ Black bar ]--    
$ReportHeader = $ReportHeader +"<table bgcolor='black' border='0' width='100%'></table>"

#--[ HTML Current state bar ]--    
$ReportHeader = $ReportHeader +"<table border='0' width='100%'>"
If (($Over80 -ge 1) -or ($Over95 -ge 1)){
    $ReportHeader = $ReportHeader +"  
    <tr bgcolor=$HexLtGrey>
         <td colspan='5' height='5' align='center'><strong><font color=$HexBlack size='2' face='tahoma'>Current System State:&nbsp;&nbsp;<font color=$HexRed size='2' face='tahoma'>At least one scope is out of tolerance.</td><strong>
    </tr>"
}Else{
    $ReportHeader = $ReportHeader +"  
    <tr bgcolor=$HexLtGrey>
         <td colspan='5' height='5' align='center'><strong><font color=$HexBlack size='2' face='tahoma'>Current System State:&nbsp;&nbsp;<font color=$HexGreen size='2' face='tahoma'>All scopes are within tolerance.</td><strong>
    </tr>" 
}

$ReportHeader = $ReportHeader +"<tr bgcolor=$HexLtGrey><td></td>"
If ($Over80 -ge 1){
    $SendEmail = $True
    $ReportHeader = $ReportHeader +"<td width='20%' height='5' align='center'><span style=background-color:$HexYellow>$Over80 scopes are over 80%</span></td>"
}Else{
    $ReportHeader = $ReportHeader +"<td width='20%' height='5' align='center'>$Over80 scopes are over 80%</td>"
}

If ($Over95 -ge 1){
    $SendEmail = $true
    $ReportHeader = $ReportHeader +"<td width='20%' height='5' align='center'><span style=background-color:$HexRed>
    <font color=$HexWhite>$Over95 scopes are over 95%</span></font></td>"
}Else{
    $ReportHeader = $ReportHeader +"<td width='20%' height='5' align='center'>$Over95 scopes are over 95%</td>"
}

$ReportHeader = $ReportHeader +"
        <td width='20%' height='5' align='center'>$Disabled scopes are disabled. </td>
        <td></td></strong>
    </tr>
        <tr bgcolor=$HexLtGrey>
        <td colspan='5' height='5' align='center'><font color=$HexBlack size='2' face='tahoma'>Scopes over 80% in use are flagged as 
        <span style=background-color:$HexYellow>WARNINGS</span>. &nbsp;&nbsp;&nbsp;&nbsp;Scopes over 95% in use are flagged as
        <span style=background-color:$HexRed><font color=$HexWhite>CRITICALS</font></span>.</font>
    </tr>"
If ($PurgeFailCount -gt 0){
    $ReportHeader = $ReportHeader +"<tr bgcolor=$HexLtGrey>
         <td colspan='5' height='5' align='center'><strong><font color=$HexYellow size='2' face='tahoma'>
         <span style=background-color:$HexRed>
         WARNING: Invalid MAC addresses were detected in the DHCP databse and automatic cleanup has failed.&nbsp;&nbsp;Scope name(s) highlighted below.
         <strong></td>
    </tr>    
    </table>"
}Else{
    $ReportHeader = $ReportHeader +"</table>"
}

#--[ Black bar ]--
$ReportHeader = $ReportHeader + "<table bgcolor='black' border='0' width='100%'></table>"

#--[ Assemble Final Report file ]----------------------------------
$Report = $HTMLHeader+$ReportHeader+$Data+ "</table></body></html>"
#------------------------------------------------------------------

#--[ Construct and send the email ]--
$Smtp = New-Object Net.Mail.SmtpClient($ExtOption.SmtpServer,25)    
$Email = New-Object System.Net.Mail.MailMessage  
If (Test-Path -path $ExtOption.SaveFile){
    $Attachment = New-Object System.Net.Mail.Attachment($ExtOption.SaveFile, 'text/plain')
    $Email.Attachments.Add($attachment)
}
If (Test-Path -path $ExtOption.PurgeFile){
    $Attachment = New-Object System.Net.Mail.Attachment($ExtOption.PurgeFile, 'text/plain')
    $Email.Attachments.Add($attachment)
}
$Email.IsBodyHTML = $true
$Email.From = $ExtOption.Sender

If (!($ExtOption.Debug)){
    If(((get-date).DayOfWeek -eq "Sunday") -or ($SendEmail)){  #--[ Sends to main recipient only on Sunday or if needed ]--
        $Email.To.Add($ExtOption.Recipient0)      
        #    $Email.To.Add($ExtOption.Recipient1) 
    }Else{
        $Email.To.Add($ExtOption.Recipient1)   #--[ Always send to additional recipients for daily status ]--
        #    $Email.To.Add($ExtOption.Recipient2)  
        #    $Email.To.Add($ExtOption.Recipient3)  
    }
}Else{
    $Email.To.Add($ExtOption.Recipient1)   #--[ Only send to additionl recipients when in debug mode ]--
}


If ($Over95 -gt 0){
    $Email.Subject = "DHCP Status ALERT"
}Else{
    $Email.Subject = "DHCP Status Report"
}
$ErrorActionPreference = "stop"
$Email.Body = $Report
$Msg = "`n--- Email Sent ---"
StatusMsg $Msg "Red" $ExtOption.Console
$Smtp.Send($Email)

#--[ Only keep 10 of the last of each log ]-- 
Get-ChildItem -Path $PSScriptRoot | Where-Object {(-not $_.PsIsContainer) -and ($_.Name -like "*PurgeLog*")} | Sort-Object -Descending -Property LastTimeWrite | Select-Object -Skip 10 | Remove-Item
Get-ChildItem -Path $PSScriptRoot | Where-Object {(-not $_.PsIsContainer) -and ($_.Name -like "*Details*")} | Sort-Object -Descending -Property LastTimeWrite | Select-Object -Skip 10 | Remove-Item

$Msg = "`n--- Completed ---`n"
StatusMsg $Msg "Red" $ExtOption.Console


<#--[ XML File Example ]-----------------------------------------
<?xml version="1.0" encoding="utf-8"?>
<Settings>
    <General>
        <SiteServerPrefix>XYZ</SiteServerPrefix>
		<SaveFile>\DHCP_Details_</SaveFile>
		<PurgeFile>\DHCP_PurgeLog_</PurgeFile>
		<NewPrefix>NewName</NewPrefix>
	</General>
	<Email>
        <SmtpServer>mailbag.org</SmtpServer>
        <SmtpPort>25</SmtpPort>
		<Sender>DHCP_Server@mailbag.org</Sender>
        <Recipien0t>helpdesk@mailbag.org</Recipient0>
		<Recipient1>bob@mailbag.org</Recipient1>
		<Recipient2>john@mailbag.org</Recipient2>
	</Email>
	<Update>
		<DnsArray>8.8.8.8,4.4.8.8,1.1.1.1</DnsArray>
		<NtpArray>10.10.10.10,20.20.20.20</NtpArray>
		<DomainArray>"mailbag.org"</DomainArray>
	</Update>
</Settings> 
#>