. $PSScriptRoot\PullList.ps1

#Gets remote admins from the computer. Uses default Administrators group SID value to retrive admins if no groupname is supplied.
#User calling this function must have admin priviledges on the remote computer.
Function Get-Admins {
    Param($ComputerName, $GroupName, [PSCredential]$Credential)
    if ($GroupName) {
        $GroupName = $GroupName -join ";"
    }
    $command = {
        param($GroupName)
        $localObject = @() 
        if (!($GroupName)) {
            $Computer = [ADSI]"WinNT://$env:COMPUTERNAME,computer"  
            $Groups = $Computer.psbase.Children | Where {$_.psbase.schemaClassName -eq "group"}
            foreach ($j in $Groups) {
                $b = $j.Path.split('/', [StringSplitOptions]::RemoveEmptyEntries)
                [string[]]$GroupName += $b[-1] 
            }
        }
        elseif ($GroupName -eq "S-1-5-32-544") {
            $objSID = New-Object System.Security.Principal.SecurityIdentifier($GroupName)
            $objgroup = $objSID.Translate( [System.Security.Principal.NTAccount]).Value.split("\")[1]
            $GroupName = $objgroup
        }
        else {$GroupName = $GroupName.split(";")}

        Foreach ($g in $GroupName) {
            Try {
       
                $Group = [ADSI]("WinNT://$env:COMPUTERNAME/$g,group")
                $Members = @($Group.psbase.Invoke("Members"))
                if ($Members) {
                    ForEach ($Member In $Members) {
                        $AdsPath = $Member.GetType().InvokeMember("Adspath", "GetProperty", $null, $Member, $null)
                        $a = $AdsPath.split('/', [StringSplitOptions]::RemoveEmptyEntries)
                        $Name = $a[-1]
                        $domain = $a[-2]
                        if ($domain -eq $env:COMPUTERNAME)
                        {$FullName = $Name}
                        elseif ($domain -eq "WinNT:") {
                            Try {
                                $objSID = New-Object System.Security.Principal.SecurityIdentifier($Name) -ErrorAction stop
                                $objUser = $objSID.Translate( [System.Security.Principal.NTAccount])
                                $FullName = $objUser.Value 
                            }
                            Catch {$FullName = $Name}
                        }  
                        Else
                        {$FullName = $domain + "\" + $Name}       
                        $Class = $Member.GetType().InvokeMember("Class", 'GetProperty', $Null, $Member, $Null)

                        $localObject += New-Object PSObject -Property @{ 
                            "Name"         = $FullName;
                            "Class"        = $class;
                            "GroupName"    = $g;
                            "ComputerName" = $env:COMPUTERNAME
                        }            
                    } #Forloop 2 close 
                } # if close
        
                else {
                    $localObject += New-Object PSObject -Property @{ 
                        "Name"         = "-";
                        "Class"        = "-";
                        "GroupName"    = $g;
                        "ComputerName" = $env:COMPUTERNAME
                    } #property close
                }# else close
            } #Try close
            Catch {
                $ErrorMessage = $_.Exception.Message
                Write-Warning "$g - $ErrorMessage"
                $localObject += New-Object PSObject -Property @{ 
                    "Name"         = "-";
                    "Class"        = "-";
                    "GroupName"    = "$g NOT found!";
                    "ComputerName" = $env:COMPUTERNAME
                } #property close
            }
        } #Forloop 1 close
        return $localObject
    } #command close

    $AdminMembers = @()
    if (!($Credential)) { 
        $AdminMembers += Invoke-Command -ComputerName $ComputerName -ScriptBlock $command -ArgumentList ($GroupName)
    }
    else {   
        $AdminMembers += Invoke-Command -ComputerName $ComputerName -ScriptBlock $command -ArgumentList ($GroupName) -Credential $Credential
    }
    $newObject = New-Object psobject
    $newObject = $AdminMembers | Select -Property Name, Class, GroupName, ComputerName -Unique
    return $newObject
}

#Adds users to the remote computer. Uses Administrators group SID value to add admins if no groupname is supplied.
#Will not attempt to add user if the user already exists on the computer.
#User calling this function must have admin priviledges on the remote computer.
Function Add-Admins {
    Param($goodAdmins, $ResolvedComputers, $GroupName, [PSCredential]$Credential)

    $command = {
        param ($goodAdmins, $GroupName)
        if ($GroupName -eq "S-1-5-32-544") {
            $objSID = New-Object System.Security.Principal.SecurityIdentifier($GroupName)
            $objgroup = $objSID.Translate( [System.Security.Principal.NTAccount]).Value.split("\")[1]
            $GroupName = $objgroup
        }
        function Get-LocalAdmins {
            $localObject = @()
            $Group = [ADSI]("WinNT://$env:COMPUTERNAME/$GroupName,group")
            $Members = @($Group.psbase.Invoke("Members"))
            if ($Members) {
                ForEach ($Member In $Members) {
                    $AdsPath = $Member.GetType().InvokeMember("Adspath", "GetProperty", $null, $Member, $null)
                    $a = $AdsPath.split('/', [StringSplitOptions]::RemoveEmptyEntries)
                    $Name = $a[-1]
                    $domain = $a[-2]
                    if ($domain -eq $env:COMPUTERNAME)
                    {$FullName = $Name}
                    elseif ($domain -eq "WinNT:") {
                        Try {
                            $objSID = New-Object System.Security.Principal.SecurityIdentifier($Name) -ErrorAction stop
                            $objUser = $objSID.Translate( [System.Security.Principal.NTAccount])
                            $FullName = $objUser.Value
                        }
                        Catch {$FullName = $Name}
                    }  
                    Else
                    {$FullName = $domain + "\" + $Name}
     
                    $localObject += New-Object PSObject -Property @{ 
                        "Name" = $FullName;
                    }            
                } 
            } # if close
            return $localObject
        } 
        $presentAdmins = Get-LocalAdmins 
        $goodAdmins = $goodAdmins.split(';')
        $toadd = @()
        $toAdd = $goodAdmins | Where {$_ -notin $presentAdmins.Name}
        [string[]]$toDisplay = $goodAdmins | Where {$_ -in $presentAdmins.Name}
        $toDisplay = $toDisplay -join ','
        if ($toDisplay) {Write-Host "$env:COMPUTERNAME : The User(s) Already Exists: $toDisplay" -ForegroundColor DarkRed}        
        $group = [ADSI]("WinNT://$env:COMPUTERNAME/$GroupName,group")
        $toAdd = $toAdd | Foreach {$_.Replace("\", "/")} 
        foreach ($i in $toAdd) {$group.Add("WinNT://" + $i)}
    }
    #net localgroup administrators $toAdd /ADD}
    if (!($Credential)) {
        Invoke-Command -ComputerName $ResolvedComputers -ScriptBlock $command -ArgumentList ($goodAdmins, $GroupName)
    }
    else {
        Invoke-Command -ComputerName $ResolvedComputers -ScriptBlock $command -ArgumentList ($goodAdmins, $GroupName) -Credential $Credential
    }
}

#Removes users from the remote computer. Uses Administrators group SID value to remove admins if no groupname is supplied.
#Will not attempt to remove a user from the computer if the user does not exist.
#User calling this function must have admin priviledges on the remote computer.
Function Remove-Admins {
    Param($BadAdmins, $ResolvedComputers, $GroupName, [PSCredential]$Credential)

    $command = {
        param ($BadAdmins, $GroupName)
        if ($GroupName -eq "S-1-5-32-544") {
            $objSID = New-Object System.Security.Principal.SecurityIdentifier($GroupName)
            $objgroup = $objSID.Translate( [System.Security.Principal.NTAccount]).Value.split("\")[1]
            $GroupName = $objgroup
        }
        function Get-LocalAdmins {
            $localObject = @()
            $Group = [ADSI]("WinNT://$env:COMPUTERNAME/$GroupName,group")
            $Members = @($Group.psbase.Invoke("Members"))
            if ($Members) {
                ForEach ($Member In $Members) {
                    $AdsPath = $Member.GetType().InvokeMember("Adspath", "GetProperty", $null, $Member, $null)
                    $a = $AdsPath.split('/', [StringSplitOptions]::RemoveEmptyEntries)
                    $Name = $a[-1]
                    $domain = $a[-2]
                    if ($domain -eq $env:COMPUTERNAME)
                    {$FullName = $Name}
                    elseif ($domain -eq "WinNT:") {
                        Try {
                            $objSID = New-Object System.Security.Principal.SecurityIdentifier($Name) -ErrorAction stop
                            $objUser = $objSID.Translate( [System.Security.Principal.NTAccount])
                            $FullName = $objUser.Value
                        }
                        Catch {$FullName = $Name}
                    }  
                    Else
                    {$FullName = $domain + "\" + $Name}
     
                    $localObject += New-Object PSObject -Property @{ 
                        "Name" = $FullName;
                    }            
                } 
            } # if close
            return $localObject
        }
        $presentAdmins = Get-LocalAdmins
        $BadAdmins = $BadAdmins.split(';')
        $toRemove = @()
        $toRemove = $BadAdmins | Where {$_ -in $presentAdmins.Name}
        $toDisplay = $BadAdmins | where {$_ -notin $presentAdmins.Name}
        $toDisplay = $toDisplay -join ','
        if ($toDisplay) {Write-Host "$env:COMPUTERNAME : The User(s) Does not Exist: $toDisplay" -ForegroundColor DarkRed}
        $group = [ADSI]("WinNT://$env:COMPUTERNAME/$GroupName,group")
        $toRemove = $toRemove | Foreach {$_.Replace("\", "/")} 
        Foreach ($i in $toRemove) {$group.Remove("WinNT://" + $i)}
    }

    if (!($Credential)) {
        Invoke-Command -ComputerName $ResolvedComputers -ScriptBlock $command -ArgumentList ($BadAdmins, $GroupName)
    }
    else {
        Invoke-Command -ComputerName $ResolvedComputers -ScriptBlock $command -ArgumentList ($BadAdmins, $GroupName) -Credential $Credential
    }
}

#Renames any Local group on the remote computer.
#User calling this function must have admin priviledges on the remote computer.
Function Rename-Group {
    Param($ComputerName, $Name, $NewName)
    $command = {
        param($Name, $NewName)
        try {
            $group = [ADSI]"WinNT://$ENV:COMPUTERNAME/$Name, group"
            if ($group) {
                $group.psbase.Rename($NewName)
                $group.psbase.CommitChanges()
            }
        }
        catch {
            $ErrorMessage = $_.Exception.Message
            Write-Warning "$ENV:COMPUTERNAME - $ErrorMessage"
        }
     
    }
    Invoke-Command -ComputerName $ComputerName -ScriptBlock $command -ArgumentList $Name, $NewName

}

#Accepts parameters and performs validation before calling the function Get-Admins to fetch users from remote computer.
#User calling this function must have admin priviledges on the remote computer.
Function Get-RemoteAdmins {

    <#
    .SYNOPSIS
        Get users on remote Computer.
 
    .DESCRIPTION 
        The Get-RemoteAdmins cmdlet fetches Admins on the remote Computer and lists the member type. This is done using the PowerShell ADSI adapter. 

    .PARAMETER ComputerName 
        Specifies the Remote Computer Name. This can be supplied as a string with comma separated values or as a path to an Excel sheet.
 
    .PARAMETER Export
        Specifies the path to export the results of the command. This is a path to a CSV file.
        
    .PARAMETER GroupName
        Specifies the remote Computer group to perform the Operation. The default group is "Administrators" if no value is supplied.

    .PARAMETER Unique
        Return a unique list of users.

    .PARAMETER AllGroups
        Return a list of users from all the local groups on the remote computer.

    .EXAMPLE
        Get-RemoteAdmins -ComputerName Computer1,Computer2 -GroupName "Remote Desktop Users"
        Fetches all the users in the group "Remote Desktop Users" on the Computers Computer1 and Computer2.
    
    .EXAMPLE
        Get-RemoteAdmins -ComputerName Computer1,Computer2 -GroupName "Remote Desktop Users" -Unique
        Fetches all the unique users in the group "Remote Desktop Users" on the Computers Computer1 and Computer2.
        
    .EXAMPLE
        Get-RemoteAdmins -ComputerName Computer1,Computer2 -Export "C:\users\user1\Desktop\Export.csv"
        Fetches all the users in the group Administrators on the Computers Computer1 and Computer2 and stores them in the file Export.csv

    .EXAMPLE
        Get-RemoteAdmins -ComputerName "C:\users\user1\Desktop\Computers.xlsx" -Export "C:\users\user1\Desktop\Export.csv" -GroupName Guests -Unique
        Fetches all the unique users in the group Guests on the list of Computers in "Computers.xlsx" and stores them in the file Export.csv  
        
    .EXAMPLE
        Get-RemoteAdmins -ComputerName Computer1,Computer2 -Export "C:\users\user1\Desktop\Export.csv" -AllGroups
        Fetches all the users from all the local groups on the Computers Computer1 and Computer2 and stores them in the file Export.csv

    .EXAMPLE    
        Get-RemoteAdmins -ComputerName Computer1,Computer2  -AllGroups -Unique
        Fetches all the users from the all the local groups on the Computers Computer1 and Computer2 and removes the duplicate users.

    .EXAMPLE
        Get-RemoteAdmins -ComputerName 'C:\users\user1\Desktop\Computers.xlsx' -AllGroups | Where {$_.Name -eq "DOMAIN\USER"}
        Fetches the user "DOMAIN\USER" from the all the Computers in the file "C:\users\user1\Desktop\Computers.xlsx". This can be used to see on what computers does "DOMAIN\USER" exists.
    #>
    [CmdletBinding(DefaultParameterSetName = 'singleGroup')]
    Param(

        [Parameter(Mandatory = $True,
            Position = 0, ParameterSetName = 'singleGroup')]
        [Parameter(Mandatory = $True,
            Position = 0, ParameterSetName = 'AllGroup')]
        [ValidateNotNullOrEmpty()]
        [string[]]$ComputerName,


        [Parameter(Mandatory = $false,
            Position = 1, ParameterSetName = 'singleGroup')]
        [Parameter(Mandatory = $false,
            Position = 1, ParameterSetName = 'AllGroup')]
        [ValidateNotNullOrEmpty()]
        [ValidatePattern("^*.csv$")]
        [string]$Export,

        [Parameter(Mandatory = $false,
            Position = 2, ParameterSetName = 'singleGroup')]
        [ValidateNotNullOrEmpty()]
        [string[]]$GroupName = 'S-1-5-32-544',

        [Parameter(Mandatory = $false,
            Position = 3, ParameterSetName = 'singleGroup')]
        [Parameter(Mandatory = $false,
            Position = 3, ParameterSetName = 'AllGroup')]
        [switch]$Unique,

        [Parameter(Mandatory = $false,
            Position = 2, ParameterSetName = 'AllGroup')]
        [switch]$AllGroups,

        [Parameter(Mandatory = $false,
            Position = 4, ParameterSetName = 'singleGroup')]
        [Parameter(Mandatory = $false,
            Position = 4, ParameterSetName = 'AllGroup')]
        [ValidateNotNullOrEmpty()]
        [PSCredential]$Credential


    )
    switch ($PsCmdlet.ParameterSetName) {
        "singleGroup" {

            Try {

                $newComputerName = ValidateComputerName($ComputerName)

                $ResolvedComputers = ResolveDNS ($newComputerName)
                Write-Verbose -Verbose "Getting Remote Admins..."
                $Admins = Get-Admins $ResolvedComputers $GroupName $Credential
                     
                if (($Export.Length -ne 0) -and ($Unique -eq $True)) {
                    $customFormat = New-Object psobject 
                    $customFormat = $Admins | select -Property Name, Class, GroupName -Unique                   
                    $customFormat | Export-CSV -Path $Export -NoTypeInformation
                    Write-Host "The admin list is stored in $Export" -ForegroundColor Yellow
                }
                     
                elseif (($Export.Length -eq 0) -and ($Unique -eq $True)) {
                    $customFormat = New-Object psobject 
                    $customFormat = $Admins | select -Property Name, Class, GroupName -Unique
                    $customFormat
                }

                elseif (($Export.Length -ne 0) -and ($Unique -eq $false)) {
                    $Admins | Export-CSV -Path $Export -NoTypeInformation
                    Write-Host "The admin list is stored in $Export" -ForegroundColor Yellow
                }
                     
                else {
                    $Admins                                         
                }

            }
            Catch {
              
                Write-Warning $_.Exception.Message
              
            }

        } #Single Group end


        "AllGroup" {

            Try {

                $newComputerName = ValidateComputerName($ComputerName)

                $ResolvedComputers = ResolveDNS ($newComputerName)
                Write-Verbose -Verbose "Getting Remote Admins..."
                $Admins = Get-Admins $ResolvedComputers $Credential
                     
                if (($Export.Length -ne 0) -and ($Unique -eq $True)) {
                    $customFormat = New-Object psobject 
                    $customFormat = $Admins | select -Property Name, Class, GroupName -Unique                   
                    $customFormat | Export-CSV -Path $Export -NoTypeInformation
                    Write-Host "The admin list is stored in $Export" -ForegroundColor Yellow
                }
                     
                elseif (($Export.Length -eq 0) -and ($Unique -eq $True)) {
                    $customFormat = New-Object psobject 
                    $customFormat = $Admins | select -Property Name, Class, GroupName -Unique
                    $customFormat
                }

                elseif (($Export.Length -ne 0) -and ($Unique -eq $false)) {
                    $Admins | Export-CSV -Path $Export -NoTypeInformation
                    Write-Host "The admin list is stored in $Export" -ForegroundColor Yellow
                }
                     
                else {
                    $Admins                                         
                }

            }
            Catch {
              
                Write-Warning $_.Exception.Message
              
            }

        } # Allgroup end
            
    } #Paramter set end

}

#Accepts parameters and performs validation before calling the function Add-Admins to add users on the remote computer.
#User calling this function must have admin priviledges on the remote computer.
Function Add-RemoteAdmins {
    <#
    .SYNOPSIS
        Add user on remote Computer.
 
    .DESCRIPTION 
        The Add-RemoteAdmins cmdlet adds Admins on the remote Computer. This is done using the PowerShell ADSI adapter. 

    .PARAMETER ComputerName 
        Specifies the Remote Computer Name. This can be supplied as a string with comma separated values or as a path to an Excel sheet.
 
    .PARAMETER AddAdmin
        Specifies the Admins to be added on the remote Computer. This can be supplied as a string with comma separated values or as a path to an Excel sheet.
        
    .PARAMETER GroupName
        Specifies the remote Computer group to perform the Operation. The default group is "Administrators" if no value is supplied.

    .EXAMPLE
        Add-RemoteAdmins -ComputerName Computer1,Computer2 -AddAdmin Domain\user1,Domain\user2 -GroupName "Remote Desktop Users"
        Adds user1 and user2 in the group "Remote Desktop Users" on the Computers Computer1 and Computer2.
        
    .EXAMPLE
        Add-RemoteAdmins -ComputerName Computer1,Computer2 -AddAdmin Domain\user1,Domain\user2
        Adds user1 and user2 in the group Administrators on the Computers Computer1 and Computer2

    .EXAMPLE
        Add-RemoteAdmins -ComputerName Computer1,Computer2 -AddAdmin "C:\users\user1\Desktop\AddAdmin.xlsx"
        Adds the users from "AddAdmin.xlsx" in the Administrators group on Computer1 and Computer2.

    .EXAMPLE
        Add-RemoteAdmins -ComputerName "C:\users\user1\Desktop\Computers.xlsx" -AddAdmin "C:\users\user1\Desktop\AddAdmin.xlsx"
        Adds the users from "AddAdmin.xlsx" in the Administrators group on the list of Computers in "Computers.xlsx"      

    #>
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $True,
            Position = 0)]
        [ValidateNotNullOrEmpty()]
        [string[]]$ComputerName,

        [Parameter(Mandatory = $true,
            Position = 1)]
        [ValidateNotNullOrEmpty()]
        [string[]]$AddAdmin,

        [Parameter(Mandatory = $false,
            Position = 2)]
        [ValidateNotNullOrEmpty()]
        [string]$GroupName = 'S-1-5-32-544',

        [Parameter(Mandatory = $false,
            Position = 3)]
        [ValidateNotNullOrEmpty()]
        [pscredential]$Credential

    )
    

    $newComputerName = ValidateComputerName($ComputerName)
  
    if ($AddAdmin -match "^*.xlsx$") {
        Resolve-Path $AddAdmin > $null  -ErrorAction Stop
        Write-Verbose -Verbose "Getting the list of Admins to add..." 
        $newAddAdmin = PullList ($AddAdmin)
        $newAddAdmin = $newAddAdmin -join ';'
    }
    else {
        $newAddAdmin = $AddAdmin
    }
                     

    Try {
        $ResolvedComputers = ResolveDNS ($newComputerName)

        Write-Verbose -Verbose "Adding Remote Admins..."
        Add-Admins $newAddAdmin $ResolvedComputers $GroupName $Credential
        Write-Host "Completed" -ForegroundColor Green
    
    }
    Catch {
        Write-Warning $_.Exception.Message
    }    

}

#Accepts parameters and performs validation before calling the function Remove-Admins to remove users from the remote computer.
#User calling this function must have admin priviledges on the remote computer.
Function Remove-RemoteAdmins {
    <#
    .SYNOPSIS
        Remove user on remote Computer.
 
    .DESCRIPTION 
        The Remove-RemoteAdmins cmdlet removes the Admins from the remote Computer. This is done using the PowerShell ADSI adapter. 

    .PARAMETER ComputerName 
        Specifies the Remote Computer Name. This can be supplied as a string with comma separated values or as a path to an Excel sheet.
 
    .PARAMETER RemoveAdmin
        Specifies the Admins to be removed from the remote Computer. This can be supplied as a string with comma separated values or as a path to an Excel sheet.
        
    .PARAMETER GroupName
        Specifies the remote Computer group to perform the Operation. The default group is "Administrators" if no value is supplied.

    .EXAMPLE
        Remove-RemoteAdmins -ComputerName Computer1,Computer2 -RemoveAdmin Domain\user1,Domain\user2 -GroupName "Remote Desktop Users"
        Removes user1 and user2 from the group "Remote Desktop Users" on the Computers Computer1 and Computer2.
        
    .EXAMPLE
        Remove-RemoteAdmins -ComputerName Computer1,Computer2 -RemoveAdmin Domain\user1,Domain\user2
        Removes user1 and user2 from the group Administrators on the Computers Computer1 and Computer2

    .EXAMPLE
        Remove-RemoteAdmins -ComputerName Computer1,Computer2 -RemoveAdmin "C:\users\user1\Desktop\RemoveAdmin.xlsx"
        Pulls the list of users from the given excel and removes them from the Administrator group on Computer1 and Computer2.

    .EXAMPLE
        Remove-RemoteAdmins -ComputerName "C:\users\user1\Desktop\Computers.xlsx" -RemoveAdmin "C:\users\user1\Desktop\RemoveAdmin.xlsx"
        Removes the users in "RemoveAdmin.xlsx" from the Administrators group on the list of Computers in "Computers.xlsx"      

    #>
    [CmdletBinding()]
    Param(

        [Parameter(Mandatory = $True,
            Position = 0)]
        [ValidateNotNullOrEmpty()]
        [string[]]$ComputerName,

        [Parameter(Mandatory = $true,
            Position = 1)]
        [ValidateNotNullOrEmpty()]
        [string[]]$RemoveAdmin,

        [Parameter(Mandatory = $false,
            Position = 2)]
        [ValidateNotNullOrEmpty()]
        [string]$GroupName = 'S-1-5-32-544',

        [Parameter(Mandatory = $false,
            Position = 3)]
        [ValidateNotNullOrEmpty()]
        [pscredential]$Credential
                
    )

    $newComputerName = ValidateComputerName($ComputerName)

    if ($RemoveAdmin -match "^*.xlsx$") {
        Resolve-Path $RemoveAdmin > $null  -ErrorAction Stop
        Write-Verbose -Verbose "Pulling Bad Admins list..."                        
        $newRemoveAdmin = PullList($RemoveAdmin)
        $newRemoveAdmin = $newRemoveAdmin -join ';'
    }
    else {
        $newRemoveAdmin = $RemoveAdmin
    }
                    
    Try {
        $ResolvedComputers = ResolveDNS ($newComputerName)

        Write-Verbose -Verbose "Removing Remote Admins..."
        Remove-Admins $newRemoveAdmin $ResolvedComputers $GroupName $Credential
        Write-Host "Completed" -ForegroundColor Green
    
    }
    Catch {
        Write-Warning $_.Exception.Message
    }

}

#Accepts parameters and performs validation before calling the function Rename-Admins on the remote computer.
#User calling this function must have admin priviledges on the remote computer.
Function Rename-RemoteGroup {
    <#
    .SYNOPSIS
        Rename a group on a remote Computer.
 
    .DESCRIPTION 
        The Rename-RemoteGroup cmdlet renames groups on the remote Computer. This is done using the PowerShell ADSI adapter. 

    .PARAMETER ComputerName 
        Specifies the Remote Computer Name. This can be supplied as a string with comma separated values or as a path to an Excel sheet.
 
    .PARAMETER Name
        Specifies the Name of the group to be renamed on the remote Computer.
        
    .PARAMETER NewName
        Specifies the new name of the remote Computer group.

    .EXAMPLE
        Rename-RemoteGroup -ComputerName Computer1,Computer2 -Name User -NewName "Remote Desktop Users"
        Renames the group User on Computer1 and Computer2 to "Remote Desktop Users"

    .EXAMPLE
        Rename-RemoteGroup -ComputerName "C:\users\user1\Desktop\Computers.xlsx" -Name User -NewName "Remote Desktop Users"
        Renames the group User to "Remote Desktop Users" from the list of Computers in "Computers.xlsx"      

    #>
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $True,
            Position = 0)]
        [ValidateNotNullOrEmpty()]
        [string[]]$ComputerName,

        [Parameter(Mandatory = $true,
            Position = 1)]
        [ValidateNotNullOrEmpty()]
        [string]$Name,

        [Parameter(Mandatory = $true,
            Position = 2)]
        [ValidateNotNullOrEmpty()]
        [string]$NewName

    )
    
    $newComputerName = ValidateComputerName($ComputerName)
    $resolvedComputers = ResolveDNS($newComputerName)
    
    Write-Verbose -Verbose "Renaming the group $Name to $NewName..."
    Rename-Group $resolvedComputers $Name $NewName
    Write-Host "Completed" -ForegroundColor Green

}
export-modulemember -function Get-RemoteAdmins
export-modulemember -function Add-RemoteAdmins
export-modulemember -function Remove-RemoteAdmins
Export-ModuleMember -Function Rename-RemoteGroup