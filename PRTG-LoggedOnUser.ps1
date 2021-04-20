<#
    .SYNOPSIS
    Monitors Logged on Users using WinRM and WMI

    .DESCRIPTION
    Using WinRM and WMI this script counts logged on users
    You can exclude users or include only one or two users.

    Includes can be made with the script parameter $IncludePattern

    Exceptions can be made within this script by changing the variable $ExcludeScript. This way, the change applies to all PRTG sensors 
    based on this script. If exceptions have to be made on a per sensor level, the script parameter $ExcludePattern can be used.

    Copy this script to the PRTG probe EXEXML scripts folder (${env:ProgramFiles(x86)}\PRTG Network Monitor\Custom Sensors\EXEXML)
    and create a "EXE/Script Advanced" sensor. Choose this script from the dropdown and set at least:

    + Parameters: -ComputerName %host -IncludePattern '(Tesuser)'
    + Security Context: Use Windows credentials of parent device or set Username and Password

    .PARAMETER ComputerName
    The hostname or IP address of the Windows machine to be checked. Should be set to %host in the PRTG parameter configuration.

    .PARAMETER IncludePattern
    Regular expression to describe the Username to Include
    Just the Username not the Domain or LocalPCName
     
      Example: ^(TestUser|UserContoso)$

      Example2: ^(Test123.*|Test555)$ excludes Test123, Test1234, Test12345 and Test555. 

    #https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_regular_expressions?view=powershell-7.1

    .PARAMETER ExcludePattern
    Regular expression to describe the Username to Exclude
    Same way as $IncludePattern

    .PARAMETER UserName
    Provide the Windows user name to connect to the target host via WinRM and WMI. Better way than explicit credentials is to set the PRTG sensor
    to launch the script in the security context that uses the "Windows credentials of parent device".

    .PARAMETER Password
    Provide the Windows password for the user specified to connect to the target machine using WinRM and WMI. Better way than explicit credentials is to set the PRTG sensor
    to launch the script in the security context that uses the "Windows credentials of parent device".

    .EXAMPLE
    Sample call from PRTG EXE/Script Advanced
    PRTG-LoggedOnUser.ps1 -ComputerName %host -IncludePattern '(TestUser|User123)' -ExcludePattern '(Contoso)'

    .NOTES
    This script is based on the sample by Paessler (https://kb.paessler.com/en/topic/67869-auto-starting-services) and debold (https://github.com/debold/PRTG-WindowsServices)

    Author:  Jannos-443
    https://github.com/Jannos-443/PRTG-LoggedOnUser
#>
param(
    [string]$ComputerName = "",
    [string]$IncludePattern = '',
    [string]$ExcludePattern = '',
    [string]$UserName = "",
    [string]$Password = ""
)

#Catch all unhandled Errors
$ErrorActionPreference = "Stop"
trap{
    if($session -ne $null)
        {
        Remove-CimSession -CimSession $session -ErrorAction SilentlyContinue
        }
    $Output = "line:$($_.InvocationInfo.ScriptLineNumber.ToString()) char:$($_.InvocationInfo.OffsetInLine.ToString()) --- message: $($_.Exception.Message.ToString()) --- line: $($_.InvocationInfo.Line.ToString()) "
    $Output = $Output.Replace("<","")
    $Output = $Output.Replace(">","")
    Write-Output "<prtg>"
    Write-Output "<error>1</error>"
    Write-Output "<text>$Output</text>"
    Write-Output "</prtg>"
    Exit
}

if ($ComputerName -eq "") 
    {
    Write-Output "<prtg>"
    Write-Output " <error>1</error>"
    Write-Output " <text>You must provide a computer name to connect to</text>"
    Write-Output "</prtg>"
    Exit
    }

# Generate Credentials Object, if provided via parameter
try{
    if($UserName -eq "" -or $Password -eq "") 
        {
        $Credentials = $null
        }
    else 
        {
        $SecPasswd  = ConvertTo-SecureString $Password -AsPlainText -Force
        $Credentials= New-Object System.Management.Automation.PSCredential ($UserName, $secpasswd)
        }
} catch {
    Write-Output "<prtg>"
    Write-Output " <error>1</error>"
    Write-Output " <text>Error Parsing Credentials ($($_.Exception.Message))</text>"
    Write-Output "</prtg>"
    Exit
}

$WmiClass = "Win32_LoggedOnUser"

# Connect and Execute
try 
    {
    if ($Credentials -eq $null) 
        {
        $users = Get-CimInstance -Namespace "root\CIMV2" -ClassName $WmiClass -ComputerName $ComputerName | Select Antecedent -Unique
        } 
    
    else 
        {
        $session = New-CimSession –ComputerName $ComputerName -Credential $Credentials
        $users = Get-CimInstance -Namespace "root\CIMV2" -ClassName $WmiClass -CimSession $session | Select Antecedent -Unique
        Start-Sleep -Seconds 1
        Remove-CimSession -CimSession $session
        }

    } 
catch 
    {
    Write-Output "<prtg>"
    Write-Output " <error>1</error>"
    Write-Output " <text>Error connecting to $ComputerName ($($_.Exception.Message))</text>"
    Write-Output "</prtg>"
    Exit
    }


# hardcoded list that applies to all hosts
$ExcludeScript = '^(dwm-.*|testuser123asd)$' 
#Remove Ignored
if ($ExcludePattern -ne "") 
    {
    $users = $users | where {$_.Antecedent.name -notmatch $ExcludePattern}  
    }

if ($ExcludeScript -ne "") 
    {
    $users = $users | where {$_.Antecedent.name -notmatch $ExcludeScript}  
    }

#If Include is set -> output only included
if ($IncludePattern -ne "")
    {
    $users = $users | where {$_.Antecedent.name -match $IncludePattern}
    }

$text = "Currently Logged in Users: "
$CountCurrentUsers = 0

foreach($user in $users)
    {
    $CountCurrentUsers += 1
    $user= $user.Antecedent
    $domain = $user.Domain
    $samname = $user.Name
    $Text += "$($domain)\$($samname) ; "
    }

$xmlOutput = '<prtg>'

if($CountCurrentUsers -gt 0)
    {
    $xmlOutput = $xmlOutput + "<text>$text</text>"
    }
else
    {
    $xmlOutput = $xmlOutput + "<text>No Logged in Users found</text>"
    }

$xmlOutput = $xmlOutput + "<result>
        <channel>Connected Users</channel>
        <value>$CountCurrentUsers</value>
        <unit>Count</unit>
        </result>"

$xmlOutput = $xmlOutput + "</prtg>"

$xmlOutput