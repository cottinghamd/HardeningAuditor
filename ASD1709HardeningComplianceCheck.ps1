#Authors: David Cottingham & Huda Minhaj
#Purpose: This script checks for compliance with the ASD Hardening Guide 1709 by checking registry keys on the local machine. Where checks are unable to be performed in this manner, either other methods of scanning are used or the user is prompted for manual checking.
#This script is designed to be used as a simple spot check of a endpoint to ensure the correct settings are applied, regardless of how complex an organisations group policy may be.
#The ASD hardening guide can be downloaded here: https://www.asd.gov.au/publications/protect/Hardening_Win10.pdf 


[CmdletBinding(SupportsShouldProcess=$true)]
param(
    [string[]]$ComputerName = $env:COMPUTERNAME,
    [switch]$ShowAllInstalledProducts,
    [System.Management.Automation.PSCredential]$Credentials
)

Function Get-OfficeVersion {
[CmdletBinding(SupportsShouldProcess=$true)]
param(
    [Parameter(ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true, Position=0)]
    [string[]]$ComputerName = $env:COMPUTERNAME,
    [switch]$ShowAllInstalledProducts,
    [System.Management.Automation.PSCredential]$Credentials
)

begin {
    $HKLM = [UInt32] "0x80000002"
    $HKCR = [UInt32] "0x80000000"

    $excelKeyPath = "Excel\DefaultIcon"
    $wordKeyPath = "Word\DefaultIcon"
   
    $installKeys = 'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall',
                   'SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall'

    $officeKeys = 'SOFTWARE\Microsoft\Office',
                  'SOFTWARE\Wow6432Node\Microsoft\Office'

    $defaultDisplaySet = 'DisplayName','Version', 'ComputerName'

    $defaultDisplayPropertySet = New-Object System.Management.Automation.PSPropertySet('DefaultDisplayPropertySet',[string[]]$defaultDisplaySet)
    $PSStandardMembers = [System.Management.Automation.PSMemberInfo[]]@($defaultDisplayPropertySet)
}

process {

 $results = new-object PSObject[] 0;
 $MSexceptionList = "mui","visio","project","proofing","visual"

 foreach ($computer in $ComputerName) {
    if ($Credentials) {
       $os=Get-WMIObject win32_operatingsystem -computername $computer -Credential $Credentials
    } else {
       $os=Get-WMIObject win32_operatingsystem -computername $computer
    }

    $osArchitecture = $os.OSArchitecture

    if ($Credentials) {
       $regProv = Get-Wmiobject -list "StdRegProv" -namespace root\default -computername $computer -Credential $Credentials
    } else {
       $regProv = Get-Wmiobject -list "StdRegProv" -namespace root\default -computername $computer
    }

    [System.Collections.ArrayList]$VersionList = New-Object -TypeName System.Collections.ArrayList
    [System.Collections.ArrayList]$PathList = New-Object -TypeName System.Collections.ArrayList
    [System.Collections.ArrayList]$PackageList = New-Object -TypeName System.Collections.ArrayList
    [System.Collections.ArrayList]$ClickToRunPathList = New-Object -TypeName System.Collections.ArrayList
    [System.Collections.ArrayList]$ConfigItemList = New-Object -TypeName  System.Collections.ArrayList
    $ClickToRunList = new-object PSObject[] 0;

    foreach ($regKey in $officeKeys) {
       $officeVersion = $regProv.EnumKey($HKLM, $regKey)
       foreach ($key in $officeVersion.sNames) {
          if ($key -match "\d{2}\.\d") {
            if (!$VersionList.Contains($key)) {
              $AddItem = $VersionList.Add($key)
            }

            $path = join-path $regKey $key

            $configPath = join-path $path "Common\Config"
            $configItems = $regProv.EnumKey($HKLM, $configPath)
            if ($configItems) {
               foreach ($configId in $configItems.sNames) {
                 if ($configId) {
                    $Add = $ConfigItemList.Add($configId.ToUpper())
                 }
               }
            }

            $cltr = New-Object -TypeName PSObject
            $cltr | Add-Member -MemberType NoteProperty -Name InstallPath -Value ""
            $cltr | Add-Member -MemberType NoteProperty -Name UpdatesEnabled -Value $false
            $cltr | Add-Member -MemberType NoteProperty -Name UpdateUrl -Value ""
            $cltr | Add-Member -MemberType NoteProperty -Name StreamingFinished -Value $false
            $cltr | Add-Member -MemberType NoteProperty -Name Platform -Value ""
            $cltr | Add-Member -MemberType NoteProperty -Name ClientCulture -Value ""
            
            $packagePath = join-path $path "Common\InstalledPackages"
            $clickToRunPath = join-path $path "ClickToRun\Configuration"
            $virtualInstallPath = $regProv.GetStringValue($HKLM, $clickToRunPath, "InstallationPath").sValue

            [string]$officeLangResourcePath = join-path  $path "Common\LanguageResources"
            $mainLangId = $regProv.GetDWORDValue($HKLM, $officeLangResourcePath, "SKULanguage").uValue
            if ($mainLangId) {
                $mainlangCulture = [globalization.cultureinfo]::GetCultures("allCultures") | where {$_.LCID -eq $mainLangId}
                if ($mainlangCulture) {
                    $cltr.ClientCulture = $mainlangCulture.Name
                }
            }

            [string]$officeLangPath = join-path  $path "Common\LanguageResources\InstalledUIs"
            $langValues = $regProv.EnumValues($HKLM, $officeLangPath);
            if ($langValues) {
               foreach ($langValue in $langValues) {
                  $langCulture = [globalization.cultureinfo]::GetCultures("allCultures") | where {$_.LCID -eq $langValue}
               } 
            }

            if ($virtualInstallPath) {

            } else {
              $clickToRunPath = join-path $regKey "ClickToRun\Configuration"
              $virtualInstallPath = $regProv.GetStringValue($HKLM, $clickToRunPath, "InstallationPath").sValue
            }

            if ($virtualInstallPath) {
               if (!$ClickToRunPathList.Contains($virtualInstallPath.ToUpper())) {
                  $AddItem = $ClickToRunPathList.Add($virtualInstallPath.ToUpper())
               }

               $cltr.InstallPath = $virtualInstallPath
               $cltr.StreamingFinished = $regProv.GetStringValue($HKLM, $clickToRunPath, "StreamingFinished").sValue
               $cltr.UpdatesEnabled = $regProv.GetStringValue($HKLM, $clickToRunPath, "UpdatesEnabled").sValue
               $cltr.UpdateUrl = $regProv.GetStringValue($HKLM, $clickToRunPath, "UpdateUrl").sValue
               $cltr.Platform = $regProv.GetStringValue($HKLM, $clickToRunPath, "Platform").sValue
               $cltr.ClientCulture = $regProv.GetStringValue($HKLM, $clickToRunPath, "ClientCulture").sValue
               $ClickToRunList += $cltr
            }

            $packageItems = $regProv.EnumKey($HKLM, $packagePath)
            $officeItems = $regProv.EnumKey($HKLM, $path)

            foreach ($itemKey in $officeItems.sNames) {
              $itemPath = join-path $path $itemKey
              $installRootPath = join-path $itemPath "InstallRoot"

              $filePath = $regProv.GetStringValue($HKLM, $installRootPath, "Path").sValue
              if (!$PathList.Contains($filePath)) {
                  $AddItem = $PathList.Add($filePath)
              }
            }

            foreach ($packageGuid in $packageItems.sNames) {
              $packageItemPath = join-path $packagePath $packageGuid
              $packageName = $regProv.GetStringValue($HKLM, $packageItemPath, "").sValue
            
              if (!$PackageList.Contains($packageName)) {
                if ($packageName) {
                   $AddItem = $PackageList.Add($packageName.Replace(' ', '').ToLower())
                }
              }
            }

          }
       }
    }

    foreach ($regKey in $installKeys) {
        $keyList = new-object System.Collections.ArrayList
        $keys = $regProv.EnumKey($HKLM, $regKey)

        foreach ($key in $keys.sNames) {
           $path = join-path $regKey $key
           $installPath = $regProv.GetStringValue($HKLM, $path, "InstallLocation").sValue
           if (!($installPath)) { continue }
           if ($installPath.Length -eq 0) { continue }

           $buildType = "64-Bit"
           if ($osArchitecture -eq "32-bit") {
              $buildType = "32-Bit"
           }

           if ($regKey.ToUpper().Contains("Wow6432Node".ToUpper())) {
              $buildType = "32-Bit"
           }

           if ($key -match "{.{8}-.{4}-.{4}-1000-0000000FF1CE}") {
              $buildType = "64-Bit" 
           }

           if ($key -match "{.{8}-.{4}-.{4}-0000-0000000FF1CE}") {
              $buildType = "32-Bit" 
           }

           if ($modifyPath) {
               if ($modifyPath.ToLower().Contains("platform=x86")) {
                  $buildType = "32-Bit"
               }

               if ($modifyPath.ToLower().Contains("platform=x64")) {
                  $buildType = "64-Bit"
               }
           }

           $primaryOfficeProduct = $false
           $officeProduct = $false
           foreach ($officeInstallPath in $PathList) {
             if ($officeInstallPath) {
                try{
                $installReg = "^" + $installPath.Replace('\', '\\')
                $installReg = $installReg.Replace('(', '\(')
                $installReg = $installReg.Replace(')', '\)')
                if ($officeInstallPath -match $installReg) { $officeProduct = $true }
                } catch {}
             }
           }

           if (!$officeProduct) { continue };
           
           $name = $regProv.GetStringValue($HKLM, $path, "DisplayName").sValue          

           $primaryOfficeProduct = $true
           if ($ConfigItemList.Contains($key.ToUpper()) -and $name.ToUpper().Contains("MICROSOFT OFFICE")) {
              foreach($exception in $MSexceptionList){
                 if($name.ToLower() -match $exception.ToLower()){
                    $primaryOfficeProduct = $false
                 }
              }
           } else {
              $primaryOfficeProduct = $false
           }

           $clickToRunComponent = $regProv.GetDWORDValue($HKLM, $path, "ClickToRunComponent").uValue
           $uninstallString = $regProv.GetStringValue($HKLM, $path, "UninstallString").sValue
           if (!($clickToRunComponent)) {
              if ($uninstallString) {
                 if ($uninstallString.Contains("OfficeClickToRun")) {
                     $clickToRunComponent = $true
                 }
              }
           }

           $modifyPath = $regProv.GetStringValue($HKLM, $path, "ModifyPath").sValue 
           $version = $regProv.GetStringValue($HKLM, $path, "DisplayVersion").sValue

           $cltrUpdatedEnabled = $NULL
           $cltrUpdateUrl = $NULL
           $clientCulture = $NULL;

           [string]$clickToRun = $false

           if ($clickToRunComponent) {
               $clickToRun = $true
               if ($name.ToUpper().Contains("MICROSOFT OFFICE")) {
                  $primaryOfficeProduct = $true
               }

               foreach ($cltr in $ClickToRunList) {
                 if ($cltr.InstallPath) {
                   if ($cltr.InstallPath.ToUpper() -eq $installPath.ToUpper()) {
                       $cltrUpdatedEnabled = $cltr.UpdatesEnabled
                       $cltrUpdateUrl = $cltr.UpdateUrl
                       if ($cltr.Platform -eq 'x64') {
                           $buildType = "64-Bit" 
                       }
                       if ($cltr.Platform -eq 'x86') {
                           $buildType = "32-Bit" 
                       }
                       $clientCulture = $cltr.ClientCulture
                   }
                 }
               }
           }
           
           if (!$primaryOfficeProduct) {
              if (!$ShowAllInstalledProducts) {
                  continue
              }
           }

           $object = New-Object PSObject -Property @{DisplayName = $name; Version = $version; InstallPath = $installPath; ClickToRun = $clickToRun; 
                     Bitness=$buildType; ComputerName=$computer; ClickToRunUpdatesEnabled=$cltrUpdatedEnabled; ClickToRunUpdateUrl=$cltrUpdateUrl;
                     ClientCulture=$clientCulture }
           $object | Add-Member MemberSet PSStandardMembers $PSStandardMembers
           $results += $object

        }
    }
  }

  $results = Get-Unique -InputObject $results 

  return $results;
}

}

$officetemp = Get-OfficeVersion | select -ExpandProperty version
$officeversion = $officetemp.Substring(0,4)

#This registry paths assume that policies have been applied in group policy in user preferences
Get-ChildItem -Path "Registry::HKEY_CURRENT_USER\Software\Policies\Microsoft\Office\$officeversion\" | Select-Object -ExpandProperty Name | ForEach-Object{
$officename = ($_).Split('\')[6]
if ($officename.Contains("outlook") -or $officename.Contains("common") -or $officename.Contains("firstrun") -or $officename.Contains("onenote") -or $officename.Contains("Registration"))
{
    #donothing
}
else
{
    $appsetting = Get-ItemProperty -Path Registry::$_\Security -ErrorAction SilentlyContinue| Select-Object -ExpandProperty VBAWarnings -ErrorAction SilentlyContinue

If ($appsetting -eq $null)
{
    write-host "Macro settings have not been configured in $officename"
}
    elseif ($appsetting -eq "4")
    {
        write-host "Macros are disabled in $officename" -ForegroundColor Green
    }
    elseif ($appsetting -eq "1")
      {
            Write-Host "Macros are not disabled in $officename, set to Enable all Macros ($appsetting)" -ForegroundColor Red
      }
      elseif ($appsetting -eq "2")
      {
            Write-Host "Macros are not disabled in $officename, Disable all Macros with notification ($appsetting)" -ForegroundColor Red
      }
      elseif ($appsetting -eq "3")
      {
            Write-Host "Macros are not disabled in $officename, Disable all Macros except those digitally signed ($appsetting)" -ForegroundColor Red
      }
      else 
      {
            Write-Host "Macros are not disabled in $officename, value is unknown and set to $appsetting" -ForegroundColor Red
      }

$apptoscan = $_

$tldisable = Get-ItemProperty -Path "Registry::$apptoscan\Security\Trusted Locations" -Name alllocationsdisabled -ErrorAction SilentlyContinue|Select-Object -ExpandProperty alllocationsdisabled

if ($tldisable -eq '1')
{
write-host "Trusted Locations for $officename are disabled" -ForegroundColor Green
}
else
{

write-host "Trusted Locations For $officename are enabled" -ForegroundColor Yellow
foreach($_ in 1..50)
{
    $i++
    $trustedlocation = Get-ItemProperty -Path "Registry::$apptoscan\Security\Trusted Locations\location$_" -Name path -ErrorAction SilentlyContinue|Select-Object -ExpandProperty path
    If ($trustedlocation -ne $null)
    {
        write-host "$trustedlocation" -ForegroundColor Magenta
    }
}
}
}
}


#Outlook has unique macro settings so we check them separately here
$macrooutlook = Get-ItemProperty -Path Registry::HKEY_CURRENT_USER\Software\Policies\Microsoft\Office\$officeversion\outlook\Security -ErrorAction SilentlyContinue| Select-Object -ExpandProperty level -ErrorAction SilentlyContinue

If ($macrooutlook -eq $null)
{
write-host "Macro settings have not been configured in Microsoft Outlook"
}
elseif ($macrooutlook -eq "4"){
    write-host "Macros are disabled in Microsoft Outlook" -ForegroundColor Green
    }
    elseif ($macrooutlook -eq"1")
      {Write-Host "Macros are not disabled in Microsoft Outlook, set to Enable all Macros" -ForegroundColor Red}
      elseif ($macrooutlook -eq"2")
      {Write-Host "Macros are not disabled in Microsoft Outlook, set to Disable all Macros with notification" -ForegroundColor Red}
      elseif ($macrooutlook -eq"3")
      {Write-Host "Macros are not disabled in Microsoft Outlook, set to Disable all Macros except those digitally signed" -ForegroundColor Red}
      else {Write-host "Macros are not disabled in Microsoft Outlook, value is unknown and set to $macrooutlook" -ForegroundColor Red}

#MS Outlook

$tldisable = Get-ItemProperty -Path "Registry::HKEY_CURRENT_USER\Software\Policies\Microsoft\Office\$officeversion\Security\Trusted Locations" -Name alllocationsdisabled -ErrorAction SilentlyContinue|Select-Object -ExpandProperty alllocationsdisabled

if ($tldisable -eq '1')
{
write-host "Trusted Locations for Outlook are disabled" -ForegroundColor Green
}
else
{

write-host "Trusted Locations For Outlook are enabled" -ForegroundColor Yellow
foreach($_ in 1..50)
{
    $i++
    $trustedlocation = Get-ItemProperty -Path "Registry::HKEY_CURRENT_USER\Software\Policies\Microsoft\Office\$officeversion\Outlook\Security\Trusted Locations\location$_" -Name path -ErrorAction SilentlyContinue|Select-Object -ExpandProperty path
    If ($trustedlocation -ne $null)
    {
        write-host "$trustedlocation" -ForegroundColor Magenta
    }
}
}


write-host "`r`n####################### CREDENTIAL CACHING #######################`r`n"
write-host "This script is unable to check Number of Previous Logons to cache, this is because the setting is in the security registry hive, please check the GPO located at Computer Configuration\Policies\Windows Settings\Security Settings\Local Policies\Security Options\Interactive Logon" -ForegroundColor Cyan

#Check Network Access: Do not allow storage of passwords and credentials for network authentication
$networkaccess = Get-ItemProperty -Path "Registry::HKEY_LOCAL_MACHINE\System\CurrentControlSet\Control\Lsa\" -Name disabledomaincreds -ErrorAction SilentlyContinue|Select-Object -ExpandProperty disabledomaincreds

if ($networkaccess -eq $null)
{
write-host "Do not allow storage of passwords and credentials for network authentication is not configured" -ForegroundColor Yellow
}
    elseif ($networkaccess -eq '1')
    {
        write-host "Do not allow storage of passwords and credentials for network authentication is enabled" -ForegroundColor Green
    }
    else
    {
        write-host "Do not allow storage of passwords and credentials for network authentication is disabled" -ForegroundColor Red
    }

#Check WDigestAuthentication is disabled
$wdigest = Get-ItemProperty -Path "Registry::HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\SecurityProviders\Wdigest\" -Name uselogoncredential -ErrorAction SilentlyContinue|Select-Object -ExpandProperty uselogoncredential

if ($wdigest -eq $null)
{
write-host "WDigest is not configured" -ForegroundColor Yellow
}
    elseif ($wdigest -eq '0')
    {
        write-host "WDigest is disabled" -ForegroundColor Green
    }
    else
    {
        write-host "WDigest is enabled" -ForegroundColor Red
    }

#Check Turn on Virtualisation Based Security
$vbsecurity = Get-ItemProperty -Path "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\Windows\DeviceGuard\" -Name EnableVirtualizationBasedSecurity -ErrorAction SilentlyContinue|Select-Object -ExpandProperty EnableVirtualizationBasedSecurity

if ($vbsecurity -eq $null)
{
write-host "Virtualisation Based Security is not configured" -ForegroundColor Yellow
}
    elseif ($vbsecurity -eq '1')
    {
        write-host "Virtualisation Based security is enabled" -ForegroundColor Green
    }
    else
    {
        write-host "Virtualisation Based security is disabled" -ForegroundColor Red
    }

#Check Secure Boot and DMA Protection
$sbdmaprot = Get-ItemProperty -Path "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\Windows\DeviceGuard" -Name RequirePlatformSecurityFeatures -ErrorAction SilentlyContinue|Select-Object -ExpandProperty RequirePlatformSecurityFeatures

if ($sbdmaprot -eq $null)
{
write-host "Secure Boot and DMA Protection is not configured" -ForegroundColor Yellow
}
    elseif ($sbdmaprot -eq '3')
    {
        write-host "Secure Boot and DMA Protection is enabled" -ForegroundColor Green
    }
    else
    {
        write-host "Secure Boot and DMA Protection is set to something non compliant" -ForegroundColor Red
    }

#Check UEFI Lock is enabled for device guard
$uefilock = Get-ItemProperty -Path "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\Windows\DeviceGuard" -Name LsaCfgFlags -ErrorAction SilentlyContinue|Select-Object -ExpandProperty LsaCfgFlags

if ($uefilock -eq $null)
{
write-host "Virtualisation Based Protection of Code Integrity with UEFI lock is not configured" -ForegroundColor Yellow
}
    elseif ($uefilock -eq '1')
    {
        write-host "Virtualisation Based Protection of Code Integrity with UEFI lock is enabled" -ForegroundColor Green
    }
    else
    {
        write-host "Virtualisation Based Protection of Code Integrity with UEFI lock is set to something non compliant" -ForegroundColor Red
    }

write-host "`r`n####################### CONTROLLED FOLDER ACCESS #######################`r`n"

#Check Controlled Folder Access for Exploit Guard is Enabled
$cfaccess = Get-ItemProperty -Path "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows Defender\Exploit Guard\Controlled Folder Access" -Name EnableControlledFolderAccess -ErrorAction SilentlyContinue|Select-Object -ExpandProperty EnableControlledFolderAccess

if ($cfaccess -eq $null)
{
write-host "Controlled Folder Access for Exploit Guard is not configured" -ForegroundColor Yellow
}
    elseif ($cfaccess -eq '1')
    {
        write-host "Controlled Folder Access for Exploit Guard is enabled" -ForegroundColor Green
    }
    else
    {
        write-host "Controlled Folder Access for Exploit Guard is disabled" -ForegroundColor Red
    }

write-host "`r`n####################### CREDENTIAL ENTRY #######################`r`n"

#Check Do not display network selection UI

$netselectui = Get-ItemProperty -Path "Registry::HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Windows\System" -Name DontDisplayNetworkSelectionUI -ErrorAction SilentlyContinue|Select-Object -ExpandProperty DontDisplayNetworkSelectionUI

if ($netselectui -eq $null)
{
write-host "Do not display network selection UI is not configured" -ForegroundColor Yellow
}
    elseif ($netselectui -eq '1')
    {
        write-host "Do not display network selection UI is enabled" -ForegroundColor Green
    }
    else
    {
        write-host "Do not display network selection UI is disabled" -ForegroundColor Red
    }

#Check Enumerate local users on domain joined computers

$enumlocalusers = Get-ItemProperty -Path "Registry::HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Windows\System" -Name EnumerateLocalUsers -ErrorAction SilentlyContinue|Select-Object -ExpandProperty EnumerateLocalUsers

if ($enumlocalusers -eq $null)
{
write-host "Enumerate local users on domain joined computers is not configured" -ForegroundColor Yellow
}
    elseif ($enumlocalusers -eq '0')
    {
        write-host "Enumerate local users on domain joined computers is enabled" -ForegroundColor Green
    }
    else
    {
        write-host "Enumerate local users on domain joined computers is disabled" -ForegroundColor Red
    }


#Check Do not display the password reveal button

$disablepassreveal = Get-ItemProperty -Path "Registry::HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Windows\CredUI" -Name DisablePasswordReveal -ErrorAction SilentlyContinue|Select-Object -ExpandProperty DisablePasswordReveal

if ($disablepassreveal -eq $null)
{
write-host "Do not display the password reveal button is not configured" -ForegroundColor Yellow
}
    elseif ($disablepassreveal -eq '1')
    {
        write-host "Do not display the password reveal button is enabled" -ForegroundColor Green
    }
    else
    {
        write-host "Do not display the password reveal button is disabled" -ForegroundColor Red
    }

#Check Enumerate administrator accounts on elevation

$enumerateadmins = Get-ItemProperty -Path "Registry::HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Windows\CredUI" -Name EnumerateAdministrators -ErrorAction SilentlyContinue|Select-Object -ExpandProperty EnumerateAdministrators

if ($enumerateadmins -eq $null)
{
write-host "Enumerate administrator accounts on elevation is not configured" -ForegroundColor Yellow
}
    elseif ($enumerateadmins -eq '0')
    {
        write-host "Enumerate administrator accounts on elevation is disabled" -ForegroundColor Green
    }
    else
    {
        write-host "Enumerate administrator accounts on elevation is enabled" -ForegroundColor Red
    }

#Check Require trusted path for credential entry 

$credentry = Get-ItemProperty -Path "Registry::HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Windows\CredUI" -Name EnableSecureCredentialPrompting -ErrorAction SilentlyContinue|Select-Object -ExpandProperty EnableSecureCredentialPrompting

if ($credentry -eq $null)
{
write-host "Require trusted path for credential entry is not configured" -ForegroundColor Yellow
}
    elseif ($credentry -eq '1')
    {
        write-host "Require trusted path for credential entry is enabled" -ForegroundColor Green
    }
    else
    {
        write-host "Require trusted path for credential entry is disabled" -ForegroundColor Red
    }

#Check Disable or enable software Secure Attention Sequence  

$sasgeneration = Get-ItemProperty -Path "Registry::HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Policies\System" -Name SoftwareSASGeneration -ErrorAction SilentlyContinue|Select-Object -ExpandProperty SoftwareSASGeneration

if ($sasgeneration -eq $null)
{
write-host "Disable or enable software Secure Attention Sequence is not configured" -ForegroundColor Yellow
}
    elseif ($sasgeneration -eq '0')
    {
        write-host "Disable or enable software Secure Attention Sequence is disabled" -ForegroundColor Green
    }
    else
    {
        write-host "Disable or enable software Secure Attention Sequence is enabled" -ForegroundColor Red
    }

#Check Sign-in last interactive user automatically after a system-initiated restart 

$systeminitiated = Get-ItemProperty -Path "Registry::HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Policies\System" -Name DisableAutomaticRestartSignOn -ErrorAction SilentlyContinue|Select-Object -ExpandProperty DisableAutomaticRestartSignOn

if ($systeminitiated -eq $null)
{
write-host "Sign-in last interactive user automatically after a system-initiated restart is not configured" -ForegroundColor Yellow
}
    elseif ($systeminitiated -eq '0')
    {
        write-host "Sign-in last interactive user automatically after a system-initiated restart is disabled" -ForegroundColor Green
    }
    else
    {
        write-host "Sign-in last interactive user automatically after a system-initiated restart is enabled" -ForegroundColor Red
    }

#Check Interactive logon: Do not require CTRL+ALT+DEL 

$ctrlaltdel = Get-ItemProperty -Path "Registry::HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Policies\System" -Name DisableCAD -ErrorAction SilentlyContinue|Select-Object -ExpandProperty DisableCAD

if ($ctrlaltdel -eq $null)
{
write-host "Interactive logon: Do not require CTRL+ALT+DEL  is not configured" -ForegroundColor Yellow
}
    elseif ($ctrlaltdel -eq '0')
    {
        write-host "Interactive logon: Do not require CTRL+ALT+DEL is disabled" -ForegroundColor Green
    }
    else
    {
        write-host "Interactive logon: Do not require CTRL+ALT+DEL is enabled" -ForegroundColor Red
    }

#Check Interactive logon: Don't display username at sign-in 

$dontdisplaylastuser = Get-ItemProperty -Path "Registry::HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Policies\System" -Name DontDisplayLastUserName -ErrorAction SilentlyContinue|Select-Object -ExpandProperty DontDisplayLastUserName

if ($dontdisplaylastuser -eq $null)
{
write-host "Interactive logon: Don't display username at sign-in is not configured" -ForegroundColor Yellow
}
    elseif ($dontdisplaylastuser -eq '1')
    {
        write-host "Interactive logon: Don't display username at sign-in is enabled" -ForegroundColor Green
    }
    else
    {
        write-host "Interactive logon: Don't display username at sign-in is disabled" -ForegroundColor Red
    }


write-host "`r`n####################### EARLY LAUNCH ANTI MALWARE #######################`r`n"

#Check ELAM Configuration

$elam = Get-ItemProperty -Path "Registry::HKEY_LOCAL_MACHINE\System\CurrentControlSet\Policies\EarlyLaunch" -Name DriverLoadPolicy -ErrorAction SilentlyContinue|Select-Object -ExpandProperty DriverLoadPolicy

if ($elam -eq $null)
{
write-host "ELAM Boot-Start Driver Initialization Policy is not configured" -ForegroundColor Yellow
}
    elseif ($elam -eq '8')
    {
        write-host "ELAM Boot-Start Driver Initialization Policy is enabled and set to Good Only" -ForegroundColor Green
    }
    elseif ($elam -eq '2')
    {
        write-host "ELAM Boot-Start Driver Initialization Policy is enabled and set to Good and Unknown" -ForegroundColor Green
    }
    elseif ($elam -eq '3')
    {
        write-host "ELAM Boot-Start Driver Initialization Policy is enabled, but set to Good, Unknown, Bad but critical" -ForegroundColor Red
    }
    elseif ($elam -eq '7')
    {
        write-host "ELAM Boot-Start Driver Initialization Policy is enabled, but set allow All drivers" -ForegroundColor Red
    }
    else
    {
        write-host "ELAM Boot-Start Driver Initialization Policy is disabled" -ForegroundColor Red
    }


write-host "`r`n####################### ELEVATING PRIVILEGES #######################`r`n"



#User Account Control: Admin Approval Mode for the Built-in Administrator account

$adminapprovalmode = Get-ItemProperty -Path "Registry::HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Policies\System" -Name FilterAdministratorToken -ErrorAction SilentlyContinue|Select-Object -ExpandProperty FilterAdministratorToken

if ($adminapprovalmode -eq $null)
{
write-host "Admin Approval Mode for the Built-in Administrator account is not configured" -ForegroundColor Yellow
}
    elseif ($adminapprovalmode -eq '1')
    {
        write-host "Admin Approval Mode for the Built-in Administrator account is enabled" -ForegroundColor Green
    }
    else
    {
        write-host "Admin Approval Mode for the Built-in Administrator account is disabled" -ForegroundColor Red
    }

#User Account Control: Allow UIAccess applications to prompt for elevation without using the secure desktop
$uiaccessapplications = Get-ItemProperty -Path "Registry::HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Policies\System" -Name EnableUIADesktopToggle -ErrorAction SilentlyContinue|Select-Object -ExpandProperty EnableUIADesktopToggle

if ($uiaccessapplications -eq $null)
{
write-host "Allow UIAccess applications to prompt for elevation without using the secure desktop is not configured" -ForegroundColor Yellow
}
    elseif ($uiaccessapplications -eq '0')
    {
        write-host "Allow UIAccess applications to prompt for elevation without using the secure desktop is disabled" -ForegroundColor Green
    }
    else
    {
        write-host "Allow UIAccess applications to prompt for elevation without using the secure desktop is enabled" -ForegroundColor Red
    }

#User Account Control: Behavior of the elevation prompt for administrators in Admin Approval Mode
$elevationprompt = Get-ItemProperty -Path "Registry::HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Policies\System" -Name ConsentPromptBehaviorAdmin -ErrorAction SilentlyContinue|Select-Object -ExpandProperty ConsentPromptBehaviorAdmin

if ($elevationprompt -eq $null)
{
write-host "Behavior of the elevation prompt for administrators in Admin Approval Mode is not configured" -ForegroundColor Yellow
}
    elseif ($elevationprompt -eq '0')
    {
        write-host "Behavior of the elevation prompt for administrators in Admin Approval Mode is configured, but set to Elevate without prompting" -ForegroundColor Red
    }
        elseif ($elevationprompt -eq '1')
    {
        write-host "Behavior of the elevation prompt for administrators in Admin Approval Mode is configured and set to Prompt for credentials on the secure desktop" -ForegroundColor Green
    }
        elseif ($elevationprompt -eq '2')
    {
        write-host "Behavior of the elevation prompt for administrators in Admin Approval Mode is configured, but set to Prompt for consent on the secure desktop" -ForegroundColor Red
    }
        elseif ($elevationprompt -eq '3')
    {
        write-host "Behavior of the elevation prompt for administrators in Admin Approval Mode is configured, but set to Prompt for credentials" -ForegroundColor Red
    }
        elseif ($elevationprompt -eq '4')
    {
        write-host "Behavior of the elevation prompt for administrators in Admin Approval Mode is configured, but set to Prompt for consent" -ForegroundColor Red
    }
        elseif ($elevationprompt -eq '5')
    {
        write-host "Behavior of the elevation prompt for administrators in Admin Approval Mode is configured, but set to Prompt for consent for non-Windows binaries" -ForegroundColor Red
    }
    else
    {
        write-host "Behavior of the elevation prompt for administrators in Admin Approval Mode is not configured" -ForegroundColor Red
    }


#User Account Control: Behavior of the elevation prompt for standard users
$standardelevationprompt = Get-ItemProperty -Path "Registry::HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Policies\System" -Name ConsentPromptBehaviorUser -ErrorAction SilentlyContinue|Select-Object -ExpandProperty ConsentPromptBehaviorUser

if ($standardelevationprompt -eq $null)
{
write-host "Behavior of the elevation prompt for standard users is not configured" -ForegroundColor Yellow
}
    elseif ($standardelevationprompt -eq '0')
    {
        write-host "Behavior of the elevation prompt for standard users is configured, but set to Automatically deny elevation requests" -ForegroundColor Yellow
    }
        elseif ($standardelevationprompt -eq '1')
    {
        write-host "Behavior of the elevation prompt for standard users is configured set to Prompt for credentials on the secure desktop" -ForegroundColor Green
    }
        elseif ($standardelevationprompt -eq '3')
    {
        write-host "Behavior of the elevation prompt for standard users is configured, but set to Prompt for credentials" -ForegroundColor Red
    }
    else
    {
        write-host "Behavior of the elevation prompt for administrators is not configured" -ForegroundColor Red
    }





#User Account Control: Detect application installations and prompt for elevation
$detectinstallelevate = Get-ItemProperty -Path "Registry::HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Policies\System" -Name EnableInstallerDetection -ErrorAction SilentlyContinue|Select-Object -ExpandProperty EnableInstallerDetection

if ($detectinstallelevate -eq $null)
{
write-host "Detect application installations and prompt for elevation is not configured" -ForegroundColor Yellow
}
    elseif ($detectinstallelevate -eq '1')
    {
        write-host "Detect application installations and prompt for elevation is enabled" -ForegroundColor Green
    }
    else
    {
        write-host "Detect application installations and prompt for elevation is disabled" -ForegroundColor Red
    }



#User Account Control: Only elevate UIAccess applications that are installed in secure locations
$onlyelevateapps = Get-ItemProperty -Path "Registry::HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Policies\System" -Name EnableSecureUIAPaths -ErrorAction SilentlyContinue|Select-Object -ExpandProperty EnableSecureUIAPaths

if ($onlyelevateapps -eq $null)
{
write-host "Only elevate UIAccess applications that are installed in secure locations is not configured" -ForegroundColor Yellow
}
    elseif ($onlyelevateapps -eq '1')
    {
        write-host "Only elevate UIAccess applications that are installed in secure locations is enabled" -ForegroundColor Green
    }
    else
    {
        write-host "Only elevate UIAccess applications that are installed in secure locations is disabled" -ForegroundColor Red
    }



#User Account Control: Run all administrators in Admin Approval Mode
$adminapprovalmode = Get-ItemProperty -Path "Registry::HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Policies\System" -Name EnableLUA -ErrorAction SilentlyContinue|Select-Object -ExpandProperty EnableLUA

if ($adminapprovalmode -eq $null)
{
write-host "Run all administrators in Admin Approval Mode is not configured" -ForegroundColor Yellow
}
    elseif ($adminapprovalmode -eq '1')
    {
        write-host "Run all administrators in Admin Approval Mode is enabled" -ForegroundColor Green
    }
    else
    {
        write-host "Run all administrators in Admin Approval Mode is disabled" -ForegroundColor Red
    }

#User Account Control: Switch to the secure desktop when prompting for elevation
$promptonsecuredesktop = Get-ItemProperty -Path "Registry::HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Policies\System" -Name PromptOnSecureDesktop -ErrorAction SilentlyContinue|Select-Object -ExpandProperty PromptOnSecureDesktop

if ($promptonsecuredesktop -eq $null)
{
write-host "Switch to the secure desktop when prompting for elevation is not configured" -ForegroundColor Yellow
}
    elseif ($promptonsecuredesktop -eq '1')
    {
        write-host "Switch to the secure desktop when prompting for elevation is enabled" -ForegroundColor Green
    }
    else
    {
        write-host "Switch to the secure desktop when prompting for elevation is disabled" -ForegroundColor Red
    }



# User Account Control: Virtualize file and registry write failures to per-user locations
$EnableVirtualization = Get-ItemProperty -Path "Registry::HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Policies\System" -Name EnableVirtualization -ErrorAction SilentlyContinue|Select-Object -ExpandProperty EnableVirtualization

if ($EnableVirtualization -eq $null)
{
write-host "Virtualize file and registry write failures to per-user locations is not configured" -ForegroundColor Yellow
}
    elseif ($EnableVirtualization -eq '1')
    {
        write-host "Virtualize file and registry write failures to per-user locations is enabled" -ForegroundColor Green
    }
    else
    {
        write-host "Virtualize file and registry write failures to per-user locations is disabled" -ForegroundColor Red
    }


write-host "`r`n####################### EXPLOIT PROTECTION #######################`r`n"



# Use a common set of exploit protection settings (this has more settings need to research)
#$ExploitProtectionSettings = Get-ItemProperty -Path "HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Windows Defender ExploitGuard\Exploit Protection" -Name ExploitProtectionSettings -ErrorAction SilentlyContinue|Select-Object -ExpandProperty ExploitProtectionSettings

#if ($ExploitProtectionSettings -eq $null)
#{
#write-host "Use a common set of exploit protection settings is not configured" -ForegroundColor Yellow
#}
#    elseif ($ExploitProtectionSettings -eq '1')
#    {
#        write-host "Use a common set of exploit protection settings is enabled" -ForegroundColor Green
#    }
#    else
#    {
#        write-host "Use a common set of exploit protection settings is disabled" -ForegroundColor Red
#    }

# Prevent users from modifying settings
$DisallowExploitProtectionOverride = Get-ItemProperty -Path "Registry::HKLM\SOFTWARE\Policies\Microsoft\Windows Defender Security Center\App and Browser protection" -Name DisallowExploitProtectionOverride -ErrorAction SilentlyContinue|Select-Object -ExpandProperty DisallowExploitProtectionOverride

if ($DisallowExploitProtectionOverride -eq $null)
{
write-host "Prevent users from modifying settings is not configured" -ForegroundColor Yellow
}
    elseif ($DisallowExploitProtectionOverride -eq '1')
    {
        write-host "Prevent users from modifying settings is enabled" -ForegroundColor Green
    }
    else
    {
        write-host "Prevent users from modifying settings is disabled" -ForegroundColor Red
    }

# Turn off Data Execution Prevention for Explorer
$NoDataExecutionPrevention = Get-ItemProperty -Path "Registry::HKLM\Software\Policies\Microsoft\Windows\Explorer" -Name NoDataExecutionPrevention -ErrorAction SilentlyContinue|Select-Object -ExpandProperty NoDataExecutionPrevention

if ($NoDataExecutionPrevention -eq $null)
{
write-host "Turn off Data Execution Prevention for Explorer is not configured" -ForegroundColor Yellow
}
    elseif ($NoDataExecutionPrevention -eq '0')
    {
        write-host "Turn off Data Execution Prevention for Explorer is disabled" -ForegroundColor Green
    }
    else
    {
        write-host "Turn off Data Execution Prevention for Explorer is enabled" -ForegroundColor Red
    }

# Enabled Structured Exception Handling Overwrite Protection (SEHOP)
$DisableExceptionChainValidation = Get-ItemProperty -Path "Registry::HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\kernel\" -Name DisableExceptionChainValidation -ErrorAction SilentlyContinue|Select-Object -ExpandProperty DisableExceptionChainValidation

if ($DisableExceptionChainValidation -eq $null)
{
write-host "Enabled Structured Exception Handling Overwrite Protection (SEHOP) is not configured" -ForegroundColor Yellow
}
    elseif ($DisableExceptionChainValidation -eq '0')
    {
        write-host "Enabled Structured Exception Handling Overwrite Protection (SEHOP) is enabled" -ForegroundColor Green
    }
    else
    {
        write-host "Enabled Structured Exception Handling Overwrite Protection (SEHOP) is disabled" -ForegroundColor Red
    }


write-host "`r`n####################### LOCAL ADMINISTRATOR ACCOUNTS #######################`r`n"

# Accounts: Administrator account status
# This is apparently not a registry key, need to implement a check using another method later


#Apply UAC restrictions to local accounts on network logons 

$LocalAccountTokenFilterPolicy = Get-ItemProperty -Path "Registry::HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System\" -Name LocalAccountTokenFilterPolicy -ErrorAction SilentlyContinue|Select-Object -ExpandProperty LocalAccountTokenFilterPolicy

if ($LocalAccountTokenFilterPolicy -eq $null)
{
write-host "Apply UAC restrictions to local accounts on network logons is not configured" -ForegroundColor Yellow
}
    elseif ($LocalAccountTokenFilterPolicy -eq '0')
    {
        write-host "Apply UAC restrictions to local accounts on network logons is enabled" -ForegroundColor Green
    }
    else
    {
        write-host "Apply UAC restrictions to local accounts on network logons is disabled" -ForegroundColor Red
    }


write-host "`r`n####################### MICROSOFT EDGE #######################`r`n"


#Allow Adobe Flash 

$FlashPlayerEnabledLM = Get-ItemProperty -Path "Registry::HKLM\Software\Policies\Microsoft\MicrosoftEdge\Addons\" -Name FlashPlayerEnabled -ErrorAction SilentlyContinue|Select-Object -ExpandProperty FlashPlayerEnabled
$FlashPlayerEnabledUP = Get-ItemProperty -Path "Registry::HKCU\Software\Policies\Microsoft\MicrosoftEdge\Addons\" -Name FlashPlayerEnabled -ErrorAction SilentlyContinue|Select-Object -ExpandProperty FlashPlayerEnabled

if ($FlashPlayerEnabledLM -eq $null -and $FlashPlayerEnabledUP -eq $null)
{
write-host "Flash Player is Not Configured" -ForegroundColor Yellow
}

if ($FlashPlayerEnabledLM -eq '0')
    {
        write-host "Flash Player is disabled in Local Machine GP" -ForegroundColor Green
    }
if ($FlashPlayerEnabledLM -eq '1')
    {
        write-host "Flash Player is enabled in Local Machine GP" -ForegroundColor Red
    }   
if ($FlashPlayerEnabledUP -eq '0')
    {
        write-host "Flash Player is disabled in User GP" -ForegroundColor Green
    }
if ($FlashPlayerEnabledUP -eq '1')
    {
        write-host "Flash Player is enabled in User GP" -ForegroundColor Red
    }

#Allow Developer Tools

$AllowDeveloperToolsLM = Get-ItemProperty -Path "Registry::HKLM\Software\Policies\Microsoft\MicrosoftEdge\F12\" -Name AllowDeveloperTools -ErrorAction SilentlyContinue|Select-Object -ExpandProperty AllowDeveloperTools
$AllowDeveloperToolsUP = Get-ItemProperty -Path "Registry::HKCU\Software\Policies\Microsoft\MicrosoftEdge\F12\" -Name AllowDeveloperTools -ErrorAction SilentlyContinue|Select-Object -ExpandProperty AllowDeveloperTools

if ($AllowDeveloperToolsLM -eq $null -and $AllowDeveloperToolsUP -eq $null)
{
write-host "Edge Developer Tools are Not Configured" -ForegroundColor Yellow
}

if ($AllowDeveloperToolsLM -eq '0')
    {
        write-host "Edge Developer Tools are disabled in Local Machine GP" -ForegroundColor Green
    }
if ($AllowDeveloperToolsLM -eq '1')
    {
        write-host "Edge Developer Tools are enabled in Local Machine GP" -ForegroundColor Red
    }   
if ($AllowDeveloperToolsUP -eq '0')
    {
        write-host "Edge Developer Tools are disabled in User GP" -ForegroundColor Green
    }
if ($AllowDeveloperToolsUP -eq '1')
    {
        write-host "Edge Developer Tools are enabled in User GP" -ForegroundColor Red
    }


#Configure Do Not Track

$DoNotTrackLM = Get-ItemProperty -Path "Software\Policies\Microsoft\MicrosoftEdge\Main\" -Name DoNotTrack -ErrorAction SilentlyContinue|Select-Object -ExpandProperty DoNotTrack
$DoNotTracksUP = Get-ItemProperty -Path "Software\Policies\Microsoft\MicrosoftEdge\Main\" -Name DoNotTrack -ErrorAction SilentlyContinue|Select-Object -ExpandProperty DoNotTrack

if ($DoNotTrackLM -eq $null -and $DoNotTrackUP -eq $null)
{
write-host "Edge Do Not Track is Not Configured" -ForegroundColor Yellow
}

if ($AllowDeveloperToolsLM -eq '0')
    {
        write-host "Edge Do Not Track is disabled in Local Machine GP" -ForegroundColor Red
    }
if ($AllowDeveloperToolsLM -eq '1')
    {
        write-host "Edge Do Not Track is enabled in Local Machine GP" -ForegroundColor Green
    }   
if ($AllowDeveloperToolsUP -eq '0')
    {
        write-host "Edge Do Not Track is disabled in User GP" -ForegroundColor Red
    }
if ($AllowDeveloperToolsUP -eq '1')
    {
        write-host "Edge Do Not Track is enabled in User GP" -ForegroundColor Green
    }

#Configure Password Manager

$FormSuggestPasswordsLM = Get-ItemProperty -Path "Software\Policies\Microsoft\MicrosoftEdge\Main\" -Name 'FormSuggest Passwords' -ErrorAction SilentlyContinue|Select-Object -ExpandProperty 'FormSuggest Passwords'
$FormSuggestPasswordsUP = Get-ItemProperty -Path "Software\Policies\Microsoft\MicrosoftEdge\Main\" -Name 'FormSuggest Passwords' -ErrorAction SilentlyContinue|Select-Object -ExpandProperty 'FormSuggest Passwords'

if ($FormSuggestPasswordsLM -eq $null -and $FormSuggestPasswordsUP -eq $null)
{
write-host "Edge Password Manager is Not Configured" -ForegroundColor Yellow
}

if ($FormSuggestPasswordsLM -eq 'no')
    {
        write-host "Edge Password Manager is disabled in Local Machine GP" -ForegroundColor Red
    }
if ($FormSuggestPasswordsLM -eq 'yes')
    {
        write-host "Edge Password Manager is enabled in Local Machine GP" -ForegroundColor Green
    }   
if ($FormSuggestPasswordsUP -eq 'no')
    {
        write-host "Edge Password Manager is disabled in User GP" -ForegroundColor Red
    }
if ($FormSuggestPasswordsUP -eq 'yes')
    {
        write-host "Edge Password Manager is enabled in User GP" -ForegroundColor Green
    }

#Configure Pop-up Blocker

$AllowPopupsLM = Get-ItemProperty -Path "Software\Policies\Microsoft\MicrosoftEdge\Main\" -Name AllowPopups -ErrorAction SilentlyContinue|Select-Object -ExpandProperty AllowPopups
$AllowPopupsUP = Get-ItemProperty -Path "Software\Policies\Microsoft\MicrosoftEdge\Main\" -Name AllowPopups -ErrorAction SilentlyContinue|Select-Object -ExpandProperty AllowPopups

if ($AllowPopupsLM -eq $null -and $AllowPopupsUP -eq $null)
{
write-host "Edge Pop-up Blocker is Not Configured" -ForegroundColor Yellow
}

if ($AllowPopupsLM -eq 'no')
    {
        write-host "Edge Pop-up Blocker is disabled in Local Machine GP" -ForegroundColor Red
    }
if ($AllowPopupsLM -eq 'yes')
    {
        write-host "Edge Pop-up Blocker is enabled in Local Machine GP" -ForegroundColor Green
    }   
if ($AllowPopupsUP -eq 'no')
    {
        write-host "Edge Pop-up Blocker is disabled in User GP" -ForegroundColor Red
    }
if ($AllowPopupsUP -eq 'yes')
    {
        write-host "Edge Pop-up Blocker is enabled in User GP" -ForegroundColor Green
    }

$EnableSmartScreen = Get-ItemProperty -Path Registry::HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Windows\System\ -Name EnableSmartScreen -ErrorAction SilentlyContinue|Select-Object -ExpandProperty EnableSmartScreen
if ( $EnableSmartScreen -eq $null)
{
write-host " Configure Windows Defender SmartScreen is not configured" -ForegroundColor Yellow
}
   elseif ( $EnableSmartScreen  -eq  '1' )
{
write-host " Configure Windows Defender SmartScreen is enabled" -ForegroundColor Green
}
  elseif ( $EnableSmartScreen  -eq  '0' )
{
write-host " Configure Windows Defender SmartScreen is disabled" -ForegroundColor Red
}
  else
{
write-host " Configure Windows Defender SmartScreen is set to an unknown setting" -ForegroundColor Red
}


$LMPreventAccessToAboutFlagsInMicrosoftEdge = Get-ItemProperty -Path Registry::HKLM\Software\Policies\Microsoft\MicrosoftEdge\Main\ -Name PreventAccessToAboutFlagsInMicrosoftEdge -ErrorAction SilentlyContinue|Select-Object -ExpandProperty PreventAccessToAboutFlagsInMicrosoftEdge
$UPPreventAccessToAboutFlagsInMicrosoftEdge = Get-ItemProperty -Path Registry::HKCU\Software\Policies\Microsoft\MicrosoftEdge\Main\ -Name PreventAccessToAboutFlagsInMicrosoftEdge -ErrorAction SilentlyContinue|Select-Object -ExpandProperty PreventAccessToAboutFlagsInMicrosoftEdge
if ( $LMPreventAccessToAboutFlagsInMicrosoftEdge -eq $null -and  $UPPreventAccessToAboutFlagsInMicrosoftEdge -eq $null)
{
write-host " Prevent access to the about:flags page in Microsoft Edge is not configured" -ForegroundColor Yellow
}
if ( $LMPreventAccessToAboutFlagsInMicrosoftEdge  -eq '1' )
{
write-host " Prevent access to the about:flags page in Microsoft Edge is enabled in Local Machine GP" -ForegroundColor Green
}
if ( $LMPreventAccessToAboutFlagsInMicrosoftEdge  -eq '0' )
{
write-host " Prevent access to the about:flags page in Microsoft Edge is disabled in Local Machine GP" -ForegroundColor Red
}
if ( $UPPreventAccessToAboutFlagsInMicrosoftEdge  -eq  '1' )
{
write-host " Prevent access to the about:flags page in Microsoft Edge is enabled in User GP" -ForegroundColor Green
}
if ( $UPPreventAccessToAboutFlagsInMicrosoftEdge  -eq  '0' )
{
write-host " Prevent access to the about:flags page in Microsoft Edge is disabled in User GP" -ForegroundColor Red
}

$LMPreventOverride = Get-ItemProperty -Path Registry::HKLM\Software\Policies\Microsoft\MicrosoftEdge\PhishingFilter\ -Name PreventOverride -ErrorAction SilentlyContinue|Select-Object -ExpandProperty PreventOverride
$UPPreventOverride = Get-ItemProperty -Path Registry::HKCU\Software\Policies\Microsoft\MicrosoftEdge\PhishingFilter\ -Name PreventOverride -ErrorAction SilentlyContinue|Select-Object -ExpandProperty PreventOverride
if ( $LMPreventOverride -eq $null -and  $UPPreventOverride -eq $null)
{
write-host " Prevent bypassing Windows Defender SmartScreen prompts for sites is not configured" -ForegroundColor Yellow
}
if ( $LMPreventOverride  -eq '1' )
{
write-host " Prevent bypassing Windows Defender SmartScreen prompts for sites is enabled in Local Machine GP" -ForegroundColor Green
}
if ( $LMPreventOverride  -eq '0' )
{
write-host " Prevent bypassing Windows Defender SmartScreen prompts for sites is disabled in Local Machine GP" -ForegroundColor Red
}
if ( $UPPreventOverride  -eq  '1' )
{
write-host " Prevent bypassing Windows Defender SmartScreen prompts for sites is enabled in User GP" -ForegroundColor Green
}
if ( $UPPreventOverride  -eq  '0' )
{
write-host " Prevent bypassing Windows Defender SmartScreen prompts for sites is disabled in User GP" -ForegroundColor Red
}

$EnableNetworkProtection = Get-ItemProperty -Path 'Registry::HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Windows Defender\Windows Defender Exploit Guard\Network Protection\' -Name EnableNetworkProtection -ErrorAction SilentlyContinue|Select-Object -ExpandProperty EnableNetworkProtection
if ( $EnableNetworkProtection -eq $null)
{
write-host " Prevent users and apps from accessing dangerous websites is not configured" -ForegroundColor Yellow
}
   elseif ( $EnableNetworkProtection  -eq  '1' )
{
write-host " Prevent users and apps from accessing dangerous websites is enabled" -ForegroundColor Green
}
  elseif ( $EnableNetworkProtection  -eq  '0' )
{
write-host " Prevent users and apps from accessing dangerous websites is disabled" -ForegroundColor Red
}
  else
{
write-host " Prevent users and apps from accessing dangerous websites is set to an unknown setting" -ForegroundColor Red
}

$AllowAppHVSI_ProviderSet = Get-ItemProperty -Path Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\AppHVSI\ -Name AllowAppHVSI_ProviderSet -ErrorAction SilentlyContinue|Select-Object -ExpandProperty AllowAppHVSI_ProviderSet
if ( $AllowAppHVSI_ProviderSet -eq $null)
{
write-host " Turn on Windows Defender Application Guard in Enterprise Mode is not configured" -ForegroundColor Yellow
}
   elseif ( $AllowAppHVSI_ProviderSet  -eq  '1' )
{
write-host " Turn on Windows Defender Application Guard in Enterprise Mode is enabled" -ForegroundColor Green
}
  elseif ( $AllowAppHVSI_ProviderSet  -eq  '0' )
{
write-host " Turn on Windows Defender Application Guard in Enterprise Mode is disabled" -ForegroundColor Red
}
  else
{
write-host " Turn on Windows Defender Application Guard in Enterprise Mode is set to an unknown setting" -ForegroundColor Red
}

$LMEnabledV9 = Get-ItemProperty -Path Registry::HKLM\Software\Policies\Microsoft\MicrosoftEdge\PhishingFilter\ -Name EnabledV9 -ErrorAction SilentlyContinue|Select-Object -ExpandProperty EnabledV9
$UPEnabledV9 = Get-ItemProperty -Path Registry::HKCU\Software\Policies\Microsoft\MicrosoftEdge\PhishingFilter\ -Name EnabledV9 -ErrorAction SilentlyContinue|Select-Object -ExpandProperty EnabledV9
if ( $LMEnabledV9 -eq $null -and  $UPEnabledV9 -eq $null)
{
write-host " Configure Windows Defender SmartScreen is not configured" -ForegroundColor Yellow
}
if ( $LMEnabledV9  -eq '1' )
{
write-host " Configure Windows Defender SmartScreen is enabled in Local Machine GP" -ForegroundColor Green
}
if ( $LMEnabledV9  -eq '0' )
{
write-host " Configure Windows Defender SmartScreen is disabled in Local Machine GP" -ForegroundColor Red
}
if ( $UPEnabledV9  -eq  '1' )
{
write-host " Configure Windows Defender SmartScreen is enabled in User GP" -ForegroundColor Green
}
if ( $UPEnabledV9  -eq  '0' )
{
write-host " Configure Windows Defender SmartScreen is disabled in User GP" -ForegroundColor Red
}

$LMPreventOverride = Get-ItemProperty -Path Registry::HKLM\Software\Policies\Microsoft\MicrosoftEdge\PhishingFilter\ -Name PreventOverride -ErrorAction SilentlyContinue|Select-Object -ExpandProperty PreventOverride
$UPPreventOverride = Get-ItemProperty -Path Registry::HKCU\Software\Policies\Microsoft\MicrosoftEdge\PhishingFilter\ -Name PreventOverride -ErrorAction SilentlyContinue|Select-Object -ExpandProperty PreventOverride
if ( $LMPreventOverride -eq $null -and  $UPPreventOverride -eq $null)
{
write-host " Prevent bypassing Windows Defender SmartScreen prompts for sites is not configured" -ForegroundColor Yellow
}
if ( $LMPreventOverride  -eq '1' )
{
write-host " Prevent bypassing Windows Defender SmartScreen prompts for sites is enabled in Local Machine GP" -ForegroundColor Green
}
if ( $LMPreventOverride  -eq '0' )
{
write-host " Prevent bypassing Windows Defender SmartScreen prompts for sites is disabled in Local Machine GP" -ForegroundColor Red
}
if ( $UPPreventOverride  -eq  '1' )
{
write-host " Prevent bypassing Windows Defender SmartScreen prompts for sites is enabled in User GP" -ForegroundColor Green
}
if ( $UPPreventOverride  -eq  '0' )
{
write-host " Prevent bypassing Windows Defender SmartScreen prompts for sites is disabled in User GP" -ForegroundColor Red
}

write-host "`r`n####################### MULTI-FACTOR AUTHENTICATION #######################`r`n"

write-host "There are no controls in this section that can be checked by a PowerShell script, this control requires manual auditing" -ForegroundColor Cyan

write-host "`r`n####################### OPERATING SYSTEM ARCHITECTURE #######################`r`n"

$architecture = $ENV:PROCESSOR_ARCHITECTURE
if ($architecture -Match '64')
{
write-host "Operating System Architecture is 64-Bit" -ForegroundColor Green
}
elseif ($architecture -Match '32')
{
write-host "Operating System Architecture is 32-Bit" -ForegroundColor Red
}
else
{
write-host "Operating System Architecture was unable to be determined" -ForegroundColor Red
}


write-host "`r`n####################### OPERATING SYSTEM PATCHING #######################`r`n"


$AutoInstallMinorUpdates = Get-ItemProperty -Path Registry::HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Windows\WindowsUpdate\AU\ -Name AutoInstallMinorUpdates -ErrorAction SilentlyContinue|Select-Object -ExpandProperty AutoInstallMinorUpdates
if ( $AutoInstallMinorUpdates -eq $null)
{
write-host " Allow Automatic Updates immediate installation is not configured" -ForegroundColor Yellow
}
   elseif ( $AutoInstallMinorUpdates  -eq  '1' )
{
write-host " Allow Automatic Updates immediate installation is enabled" -ForegroundColor Green
}
  elseif ( $AutoInstallMinorUpdates  -eq  '0' )
{
write-host " Allow Automatic Updates immediate installation is disabled" -ForegroundColor Red
}
  else
{
write-host " Allow Automatic Updates immediate installation is set to an unknown setting" -ForegroundColor Red
}

$NoAutoUpdate = Get-ItemProperty -Path Registry::HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Windows\WindowsUpdate\AU\ -Name NoAutoUpdate -ErrorAction SilentlyContinue|Select-Object -ExpandProperty NoAutoUpdate
if ( $NoAutoUpdate -eq $null)
{
write-host " Configure Automatic Updates is not configured" -ForegroundColor Yellow
}
   elseif ( $NoAutoUpdate  -eq  '0' )
{
write-host " Configure Automatic Updates is enabled" -ForegroundColor Green
}
  elseif ( $NoAutoUpdate  -eq  '1' )
{
write-host " Configure Automatic Updates is disabled" -ForegroundColor Red
}
  else
{
write-host " Configure Automatic Updates is set to an unknown setting" -ForegroundColor Red
}

$ExcludeWUDriversInQualityUpdate = Get-ItemProperty -Path Registry::HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Windows\WindowsUpdate\ -Name ExcludeWUDriversInQualityUpdate -ErrorAction SilentlyContinue|Select-Object -ExpandProperty ExcludeWUDriversInQualityUpdate
if ( $ExcludeWUDriversInQualityUpdate -eq $null)
{
write-host " Do not include drivers with Windows Updates is not configured" -ForegroundColor Yellow
}
   elseif ( $ExcludeWUDriversInQualityUpdate  -eq  '0' )
{
write-host " Do not include drivers with Windows Updates is disabled" -ForegroundColor Green
}
  elseif ( $ExcludeWUDriversInQualityUpdate  -eq  '1' )
{
write-host " Do not include drivers with Windows Updates is enabled" -ForegroundColor Red
}
  else
{
write-host " Do not include drivers with Windows Updates is set to an unknown setting" -ForegroundColor Red
}

$NoAutoRebootWithLoggedOnUsers = Get-ItemProperty -Path Registry::HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Windows\WindowsUpdate\AU\ -Name NoAutoRebootWithLoggedOnUsers -ErrorAction SilentlyContinue|Select-Object -ExpandProperty NoAutoRebootWithLoggedOnUsers
if ( $NoAutoRebootWithLoggedOnUsers -eq $null)
{
write-host " No auto-restart with logged on users for scheduled automatic updates installations is not configured" -ForegroundColor Yellow
}
   elseif ( $NoAutoRebootWithLoggedOnUsers  -eq  '1' )
{
write-host " No auto-restart with logged on users for scheduled automatic updates installations is enabled" -ForegroundColor Green
}
  elseif ( $NoAutoRebootWithLoggedOnUsers  -eq  '0' )
{
write-host " No auto-restart with logged on users for scheduled automatic updates installations is disabled" -ForegroundColor Red
}
  else
{
write-host " No auto-restart with logged on users for scheduled automatic updates installations is set to an unknown setting" -ForegroundColor Red
}


$SetDisableUXWUAccess = Get-ItemProperty -Path Registry::HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Windows\WindowsUpdate\ -Name SetDisableUXWUAccess -ErrorAction SilentlyContinue|Select-Object -ExpandProperty SetDisableUXWUAccess
if ( $SetDisableUXWUAccess -eq $null)
{
write-host " Remove access to use all Windows Update features is not configured" -ForegroundColor Yellow
}
   elseif ( $SetDisableUXWUAccess  -eq  '0' )
{
write-host " Remove access to use all Windows Update features is disabled" -ForegroundColor Green
}
  elseif ( $SetDisableUXWUAccess  -eq  '1' )
{
write-host " Remove access to use all Windows Update features is enabled" -ForegroundColor Red
}
  else
{
write-host " Remove access to use all Windows Update features is set to an unknown setting" -ForegroundColor Red
}

$IncludeRecommendedUpdates = Get-ItemProperty -Path Registry::HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Windows\WindowsUpdate\AU\ -Name IncludeRecommendedUpdates -ErrorAction SilentlyContinue|Select-Object -ExpandProperty IncludeRecommendedUpdates
if ( $IncludeRecommendedUpdates -eq $null)
{
write-host " Turn on recommended updates via Automatic Updates is not configured" -ForegroundColor Yellow
}
   elseif ( $IncludeRecommendedUpdates  -eq  '1' )
{
write-host " Turn on recommended updates via Automatic Updates is enabled" -ForegroundColor Green
}
  elseif ( $IncludeRecommendedUpdates  -eq  '0' )
{
write-host " Turn on recommended updates via Automatic Updates is disabled" -ForegroundColor Red
}
  else
{
write-host " Turn on recommended updates via Automatic Updates is set to an unknown setting" -ForegroundColor Red
}

$UseWUServer = Get-ItemProperty -Path Registry::HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Windows\WindowsUpdate\AU\ -Name UseWUServer -ErrorAction SilentlyContinue|Select-Object -ExpandProperty UseWUServer
if ( $UseWUServer -eq $null)
{
write-host " Specify intranet Microsoft update service location is not configured" -ForegroundColor Yellow
}
   elseif ( $UseWUServer  -eq  '1' )
{
write-host " Specify intranet Microsoft update service location is enabled" -ForegroundColor Green
}
  elseif ( $UseWUServer  -eq  '0' )
{
write-host " Specify intranet Microsoft update service location is disabled" -ForegroundColor Red
}
  else
{
write-host " Specify intranet Microsoft update service location is set to an unknown setting" -ForegroundColor Red
}

write-host "`r`n####################### PASSWORD POLICY #######################`r`n"

$BlockDomainPicturePassword = Get-ItemProperty -Path Registry::HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Windows\System\ -Name BlockDomainPicturePassword -ErrorAction SilentlyContinue|Select-Object -ExpandProperty BlockDomainPicturePassword
if ( $BlockDomainPicturePassword -eq $null)
{
write-host " Turn off picture password sign-in is not configured" -ForegroundColor Yellow
}
   elseif ( $BlockDomainPicturePassword  -eq  '1' )
{
write-host " Turn off picture password sign-in is enabled" -ForegroundColor Green
}
  elseif ( $BlockDomainPicturePassword  -eq  '0' )
{
write-host " Turn off picture password sign-in is disabled" -ForegroundColor Red
}
  else
{
write-host " Turn off picture password sign-in is set to an unknown setting" -ForegroundColor Red
}

$AllowDomainPINLogon = Get-ItemProperty -Path Registry::HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Windows\System\ -Name AllowDomainPINLogon -ErrorAction SilentlyContinue|Select-Object -ExpandProperty AllowDomainPINLogon
if ( $AllowDomainPINLogon -eq $null)
{
write-host " Turn on convenience PIN sign-in is not configured" -ForegroundColor Yellow
}
   elseif ( $AllowDomainPINLogon  -eq  '0' )
{
write-host " Turn on convenience PIN sign-in is disabled" -ForegroundColor Green
}
  elseif ( $AllowDomainPINLogon  -eq  '1' )
{
write-host " Turn on convenience PIN sign-in is enabled" -ForegroundColor Red
}
  else
{
write-host " Turn on convenience PIN sign-in is set to an unknown setting" -ForegroundColor Red
}

write-host "Enforce Password History is unable to be checked using PowerShell, as the setting is not a registry key. Please check Computer Configuration\Policies\Administrative Templates\System\Logon" -ForegroundColor Cyan
write-host "Maximum password age is unable to be checked using PowerShell, as the setting is not a registry key. Please check Computer Configuration\Policies\Administrative Templates\System\Logon" -ForegroundColor Cyan
write-host "Minimum password age is unable to be checked using PowerShell, as the setting is not a registry key. Please check Computer Configuration\Policies\Administrative Templates\System\Logon" -ForegroundColor Cyan
write-host "Store passwords using reversible encryption is unable to be checked using PowerShell, as the setting is not a registry key. Please check Computer Configuration\Policies\Administrative Templates\System\Logon" -ForegroundColor Cyan

$LimitBlankPasswordUse = Get-ItemProperty -Path Registry::HKEY_LOCAL_MACHINE\System\CurrentControlSet\Control\Lsa\ -Name LimitBlankPasswordUse -ErrorAction SilentlyContinue|Select-Object -ExpandProperty LimitBlankPasswordUse
if ( $LimitBlankPasswordUse -eq $null)
{
write-host "Limit local account use of blank passwords to console logon only is not configured" -ForegroundColor Yellow
}
   elseif ( $LimitBlankPasswordUse  -eq  '0' )
{
write-host "Limit local account use of blank passwords to console logon only is disabled" -ForegroundColor Red
}
  elseif ( $LimitBlankPasswordUse  -eq  '1' )
{
write-host "Limit local account use of blank passwords to console logon only is enabled" -ForegroundColor Green
}
  else
{
write-host "Limit local account use of blank passwords to console logon only is set to an unknown setting" -ForegroundColor Red
}

write-host "`r`n####################### RESTRICTING PRIVILEGED ACCOUNTS #######################`r`n"

write-host "`r`n####################### SECURE BOOT #######################`r`n"

write-host "`r`n####################### ACCOUNT LOCKOUT POLICIES #######################`r`n"

write-host "Account Lockout Duration is unable to be checked using PowerShell, as the setting is not a registry key. Please check Computer Configuration\Policies\Windows Settings\Security Settings\Account Policies\Account Lockout Policy" -ForegroundColor Cyan
write-host "Account Lockout Threshold is unable to be checked using PowerShell, as the setting is not a registry key. Please check Computer Configuration\Policies\Windows Settings\Security Settings\Account Policies\Account Lockout Policy" -ForegroundColor Cyan
write-host "Reset Account Lockout Counter After is unable to be checked using PowerShell, as the setting is not a registry key. Please check Computer Configuration\Policies\Windows Settings\Security Settings\Account Policies\Account Lockout Policy" -ForegroundColor Cyan


write-host "`r`n####################### ANONYMOUS CONNECTIONS #######################`r`n"

$AllowInsecureGuestAuth = Get-ItemProperty -Path Registry::HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Windows\LanmanWorkstation\ -Name AllowInsecureGuestAuth -ErrorAction SilentlyContinue|Select-Object -ExpandProperty AllowInsecureGuestAuth
if ( $AllowInsecureGuestAuth -eq $null)
{
write-host " Enable insecure guest logons is not configured" -ForegroundColor Yellow
}
   elseif ( $AllowInsecureGuestAuth  -eq  '0' )
{
write-host " Enable insecure guest logons is disabled" -ForegroundColor Green
}
  elseif ( $AllowInsecureGuestAuth  -eq  '1' )
{
write-host " Enable insecure guest logons is enabled" -ForegroundColor Red
}
  else
{
write-host " Enable insecure guest logons is set to an unknown setting" -ForegroundColor Red
}

write-host "Network access: Allow anonymous SID/Name translation can't check is unable to be checked using PowerShell, as the setting is not a registry key. Please check Computer Configuration\Policies\Windows Settings\Local Policies\Security Options\" -ForegroundColor Cyan

$RestrictAnonymousSAM = Get-ItemProperty -Path Registry::HKEY_LOCAL_MACHINE\System\CurrentControlSet\Control\Lsa\ -Name RestrictAnonymousSAM -ErrorAction SilentlyContinue|Select-Object -ExpandProperty RestrictAnonymousSAM
if ( $RestrictAnonymousSAM -eq $null)
{
write-host " Network access: Do not allow anonymous enumeration of SAM accounts is not configured" -ForegroundColor Yellow
}
   elseif ( $RestrictAnonymousSAM  -eq  '1' )
{
write-host " Network access: Do not allow anonymous enumeration of SAM accounts is enabled" -ForegroundColor Green
}
  elseif ( $RestrictAnonymousSAM  -eq  '0' )
{
write-host " Network access: Do not allow anonymous enumeration of SAM accounts is disabled" -ForegroundColor Red
}
  else
{
write-host " Network access: Do not allow anonymous enumeration of SAM accounts is set to an unknown setting" -ForegroundColor Red
}

$RestrictAnonymous = Get-ItemProperty -Path Registry::HKEY_LOCAL_MACHINE\System\CurrentControlSet\Control\Lsa\ -Name RestrictAnonymous -ErrorAction SilentlyContinue|Select-Object -ExpandProperty RestrictAnonymous
if ( $RestrictAnonymous -eq $null)
{
write-host " Network access: Do not allow anonymous enumeration of SAM accounts and shares is not configured" -ForegroundColor Yellow
}
   elseif ( $RestrictAnonymous  -eq  '1' )
{
write-host " Network access: Do not allow anonymous enumeration of SAM accounts and shares is enabled" -ForegroundColor Green
}
  elseif ( $RestrictAnonymous  -eq  '0' )
{
write-host " Network access: Do not allow anonymous enumeration of SAM accounts and shares is disabled" -ForegroundColor Red
}
  else
{
write-host " Network access: Do not allow anonymous enumeration of SAM accounts and shares is set to an unknown setting" -ForegroundColor Red
}


$EveryoneIncludesAnonymous = Get-ItemProperty -Path Registry::HKEY_LOCAL_MACHINE\System\CurrentControlSet\Control\Lsa\ -Name EveryoneIncludesAnonymous -ErrorAction SilentlyContinue|Select-Object -ExpandProperty EveryoneIncludesAnonymous
if ( $EveryoneIncludesAnonymous -eq $null)
{
write-host " Network access: Let Everyone permissions apply to anonymous users is not configured" -ForegroundColor Yellow
}
   elseif ( $EveryoneIncludesAnonymous  -eq  '0' )
{
write-host " Network access: Let Everyone permissions apply to anonymous users is disabled" -ForegroundColor Green
}
  elseif ( $EveryoneIncludesAnonymous  -eq  '1' )
{
write-host " Network access: Let Everyone permissions apply to anonymous users is enabled" -ForegroundColor Red
}
  else
{
write-host " Network access: Let Everyone permissions apply to anonymous users is set to an unknown setting" -ForegroundColor Red
}

$RestrictNullSessAccess = Get-ItemProperty -Path Registry::HKEY_LOCAL_MACHINE\System\CurrentControlSet\Services\LanManServer\Parameters\ -Name RestrictNullSessAccess -ErrorAction SilentlyContinue|Select-Object -ExpandProperty RestrictNullSessAccess
if ( $RestrictNullSessAccess -eq $null)
{
write-host " Network access: Do not allow anonymous enumeration of SAM accounts and shares is not configured" -ForegroundColor Yellow
}
   elseif ( $RestrictNullSessAccess  -eq  '1' )
{
write-host " Network access: Restrict clients allowed to make remote calls to SAM is configured correctly" -ForegroundColor Green
}
   else
{
write-host " Network access: Restrict clients allowed to make remote calls to SAM is configured incorrectly" -ForegroundColor Red
}


$RestrictRemoteSAM = Get-ItemProperty -Path Registry::HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Lsa\ -Name RestrictRemoteSAM -ErrorAction SilentlyContinue|Select-Object -ExpandProperty RestrictRemoteSAM
if ( $RestrictRemoteSAM -eq $null)
{
write-host " Network access: Do not allow anonymous enumeration of SAM accounts and shares is not configured" -ForegroundColor Yellow
}
   elseif ( $RestrictRemoteSAM  -eq  'O:BAG:BAD:(A;;RC;;;BA)' )
{
write-host " Network access: Do not allow anonymous enumeration of SAM accounts and shares is enabled" -ForegroundColor Green
}
  elseif ( $RRestrictRemoteSAM  -eq  '0' )
{
write-host " Network access: Do not allow anonymous enumeration of SAM accounts and shares is disabled" -ForegroundColor Red
}
  else
{
write-host " Network access: Do not allow anonymous enumeration of SAM accounts and shares is set to an unknown setting" -ForegroundColor Red
}


write-host "Network security: Allow Local System to use computer identity for NTLM"



$allownullsessionfallback = Get-ItemProperty -Path Registry::HKEY_LOCAL_MACHINE\System\CurrentControlSet\Control\Lsa\MSV1_0\ -Name allownullsessionfallback -ErrorAction SilentlyContinue|Select-Object -ExpandProperty allownullsessionfallback
if ( $allownullsessionfallback -eq $null)
{
write-host "Network security: Allow LocalSystem NULL session fallback is not configured" -ForegroundColor Yellow
}
   elseif ( $allownullsessionfallback  -eq  '0' )
{
write-host "Network security: Allow LocalSystem NULL session fallback is disabled" -ForegroundColor Green
}
  elseif ( $allownullsessionfallback  -eq  '1' )
{
write-host "Network security: Allow LocalSystem NULL session fallback is enabled" -ForegroundColor Red
}
  else
{
write-host "Network security: Allow LocalSystem NULL session fallback is set to an unknown setting" -ForegroundColor Red
}

write-host "Access this computer from the network is unable to be checked using PowerShell, as the setting is not a registry key. Please check Computer Configuration\Policies\Windows Settings\Security Settings\Local Policies\User Rights Assignment\. ASD Recommendation is to only have 'Administrators & Remote Desktop Users' present" -ForegroundColor Cyan
write-host "Deny Access to this computer from the network is unable to be checked using PowerShell, as the setting is not a registry key. Please check Computer Configuration\Policies\Windows Settings\Security Settings\Local Policies\User Rights Assignment\. ASD Recommendation is to only have 'Guests & NT AUTHORITY\Local Account' present" -ForegroundColor Cyan


write-host "`r`n####################### ANTI-VIRUS SOFTWARE #######################`r`n"

$DisableAntiSpyware = Get-ItemProperty -Path "Registry::HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Windows Defender\" -Name DisableAntiSpyware -ErrorAction SilentlyContinue|Select-Object -ExpandProperty DisableAntiSpyware
if ( $DisableAntiSpyware -eq $null)
{
write-host " Turn off Windows Defender Antivirus is not configured" -ForegroundColor Yellow
}
   elseif ( $DisableAntiSpyware  -eq  '0' )
{
write-host " Turn off Windows Defender Antivirus is disabled" -ForegroundColor Green
}
  elseif ( $DisableAntiSpyware  -eq  '1' )
{
write-host " Turn off Windows Defender Antivirus is enabled" -ForegroundColor Red
}
  else
{
write-host " Turn off Windows Defender Antivirus is set to an unknown setting" -ForegroundColor Red
}

$LocalSettingOverrideSpyNetReporting = Get-ItemProperty -Path 'Registry::HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Microsoft Antimalware\SpyNet\' -Name LocalSettingOverrideSpyNetReporting -ErrorAction SilentlyContinue|Select-Object -ExpandProperty LocalSettingOverrideSpyNetReporting
if ( $LocalSettingOverrideSpyNetReporting -eq $null)
{
write-host " Configure local setting override for reporting to Microsoft Active Protection Service (MAPS). is not configured" -ForegroundColor Yellow
}
   elseif ( $LocalSettingOverrideSpyNetReporting  -eq  '0' )
{
write-host " Configure local setting override for reporting to Microsoft Active Protection Service (MAPS). is disabled" -ForegroundColor Green
}
  elseif ( $LocalSettingOverrideSpyNetReporting  -eq  '1' )
{
write-host " Configure local setting override for reporting to Microsoft Active Protection Service (MAPS). is enabled" -ForegroundColor Red
}
  else
{
write-host " Configure local setting override for reporting to Microsoft Active Protection Service (MAPS). is set to an unknown setting" -ForegroundColor Red
}

$DisableBlockAtFirstSeen = Get-ItemProperty -Path 'Registry::HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Windows Defender\Spynet\' -Name DisableBlockAtFirstSeen -ErrorAction SilentlyContinue|Select-Object -ExpandProperty DisableBlockAtFirstSeen
if ( $DisableBlockAtFirstSeen -eq $null)
{
write-host " Configure the 'Block at First Sight' feature is not configured" -ForegroundColor Yellow
}
   elseif ( $DisableBlockAtFirstSeen  -eq  '0' )
{
write-host " Configure the 'Block at First Sight' feature is enabled" -ForegroundColor Green
}
  elseif ( $DisableBlockAtFirstSeen  -eq  '1' )
{
write-host " Configure the 'Block at First Sight' feature is disabled" -ForegroundColor Red
}
  else
{
write-host " Configure the 'Block at First Sight' feature is set to an unknown setting" -ForegroundColor Red
}

$SpyNetReporting = Get-ItemProperty -Path 'Registry::HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Microsoft Antimalware\SpyNet\' -Name SpyNetReporting -ErrorAction SilentlyContinue|Select-Object -ExpandProperty SpyNetReporting
if ( $SpyNetReporting -eq $null)
{
write-host " Join Microsoft Active Protection Service (MAPS). is not configured" -ForegroundColor Yellow
}
   elseif ( $SpyNetReporting  -eq  '1' )
{
write-host " Join Microsoft Active Protection Service (MAPS). is enabled" -ForegroundColor Green
}
  elseif ( $SpyNetReporting  -eq  '0' )
{
write-host " Join Microsoft Active Protection Service (MAPS). is disabled" -ForegroundColor Red
}
  else
{
write-host " Join Microsoft Active Protection Service (MAPS). is set to an unknown setting" -ForegroundColor Red
}

$SubmitSamplesConsent = Get-ItemProperty -Path 'Registry::HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Windows Defender\Spynet\' -Name SubmitSamplesConsent -ErrorAction SilentlyContinue|Select-Object -ExpandProperty SubmitSamplesConsent
if ( $SubmitSamplesConsent -eq $null)
{
write-host " Send file samples when further analysis is required is not configured" -ForegroundColor Yellow
}
   elseif ( $SubmitSamplesConsent  -eq  '1' )
{
write-host " Send file samples when further analysis is required is enabled" -ForegroundColor Green
}
  elseif ( $SubmitSamplesConsent  -eq  '0' )
{
write-host " Send file samples when further analysis is required is disabled" -ForegroundColor Red
}
  else
{
write-host " Send file samples when further analysis is required is set to an unknown setting" -ForegroundColor Red
}

$MpBafsExtendedTimeout = Get-ItemProperty -Path 'Registry::HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Windows Defender\MpEngine\' -Name MpBafsExtendedTimeout -ErrorAction SilentlyContinue|Select-Object -ExpandProperty MpBafsExtendedTimeout
if ( $MpBafsExtendedTimeout -eq $null)
{
write-host " Configure extended cloud check is not configured" -ForegroundColor Yellow
}
   elseif ( $MpBafsExtendedTimeout  -eq  '1' )
{
write-host " Configure extended cloud check is enabled" -ForegroundColor Green
}
  elseif ( $MpBafsExtendedTimeout  -eq  '0' )
{
write-host " Configure extended cloud check is disabled" -ForegroundColor Red
}
  else
{
write-host " Configure extended cloud check is set to an unknown setting" -ForegroundColor Red
}

$MpCloudBlockLevel = Get-ItemProperty -Path 'Registry::HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Windows Defender\MpEngine\' -Name MpCloudBlockLevel -ErrorAction SilentlyContinue|Select-Object -ExpandProperty MpCloudBlockLevel
if ( $MpCloudBlockLevel -eq $null)
{
write-host " Select cloud protection level is not configured" -ForegroundColor Yellow
}
   elseif ( $MpCloudBlockLevel  -eq  '1' )
{
write-host " Select cloud protection level is enabled" -ForegroundColor Green
}
  elseif ( $MpCloudBlockLevel  -eq  '0' )
{
write-host " Select cloud protection level is disabled" -ForegroundColor Red
}
  else
{
write-host " Select cloud protection level is set to an unknown setting" -ForegroundColor Red
}

$LocalSettingOverrideDisableIOAVProtection = Get-ItemProperty -Path 'Registry::HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Windows Defender\Real-Time Protection\' -Name LocalSettingOverrideDisableIOAVProtection -ErrorAction SilentlyContinue|Select-Object -ExpandProperty LocalSettingOverrideDisableIOAVProtection
if ( $LocalSettingOverrideDisableIOAVProtection -eq $null)
{
write-host " Configure local setting override for scanning all downloaded files and attachments is not configured" -ForegroundColor Yellow
}
   elseif ( $LocalSettingOverrideDisableIOAVProtection  -eq  '1' )
{
write-host " Configure local setting override for scanning all downloaded files and attachments is enabled" -ForegroundColor Green
}
  elseif ( $LocalSettingOverrideDisableIOAVProtection  -eq  '0' )
{
write-host " Configure local setting override for scanning all downloaded files and attachments is disabled" -ForegroundColor Red
}
  else
{
write-host " Configure local setting override for scanning all downloaded files and attachments is set to an unknown setting" -ForegroundColor Red
}

$DisableRealtimeMonitoring = Get-ItemProperty -Path 'Registry::HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Windows Defender\Real-Time Protection\' -Name DisableRealtimeMonitoring -ErrorAction SilentlyContinue|Select-Object -ExpandProperty DisableRealtimeMonitoring
if ( $DisableRealtimeMonitoring -eq $null)
{
write-host " Turn off real-time protection is not configured" -ForegroundColor Yellow
}
   elseif ( $DisableRealtimeMonitoring  -eq  '0' )
{
write-host " Turn off real-time protection is disabled" -ForegroundColor Green
}
  elseif ( $DisableRealtimeMonitoring  -eq  '1' )
{
write-host " Turn off real-time protection is enabled" -ForegroundColor Red
}
  else
{
write-host " Turn off real-time protection is set to an unknown setting" -ForegroundColor Red
}

$DisableBehaviorMonitoring = Get-ItemProperty -Path 'Registry::HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Windows Defender\Real-Time Protection\' -Name DisableBehaviorMonitoring -ErrorAction SilentlyContinue|Select-Object -ExpandProperty DisableBehaviorMonitoring
if ( $DisableBehaviorMonitoring -eq $null)
{
write-host " Turn on behavior monitoring is not configured" -ForegroundColor Yellow
}
   elseif ( $DisableBehaviorMonitoring  -eq  '0' )
{
write-host " Turn on behavior monitoring is enabled" -ForegroundColor Green
}
  elseif ( $DisableBehaviorMonitoring  -eq  '1' )
{
write-host " Turn on behavior monitoring is disabled" -ForegroundColor Red
}
  else
{
write-host " Turn on behavior monitoring is set to an unknown setting" -ForegroundColor Red
}

$DisableScanOnRealtimeEnable = Get-ItemProperty -Path 'Registry::HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Windows Defender\Real-Time Protection\' -Name DisableScanOnRealtimeEnable -ErrorAction SilentlyContinue|Select-Object -ExpandProperty DisableScanOnRealtimeEnable
if ( $DisableScanOnRealtimeEnable -eq $null)
{
write-host " Turn on process scanning whenever real-time protection is enabled is not configured" -ForegroundColor Yellow
}
   elseif ( $DisableScanOnRealtimeEnable  -eq  '0' )
{
write-host " Turn on process scanning whenever real-time protection is enabled is enabled" -ForegroundColor Green
}
  elseif ( $DisableScanOnRealtimeEnable  -eq  '1' )
{
write-host " Turn on process scanning whenever real-time protection is enabled is disabled" -ForegroundColor Red
}
  else
{
write-host " Turn on process scanning whenever real-time protection is enabled is set to an unknown setting" -ForegroundColor Red
}

$PurgeItemsAfterDelay = Get-ItemProperty -Path 'Registry::HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Windows Defender\Quarantine\' -Name PurgeItemsAfterDelay -ErrorAction SilentlyContinue|Select-Object -ExpandProperty PurgeItemsAfterDelay
if ( $PurgeItemsAfterDelay -eq $null)
{
write-host " Configure removal of items from Quarantine folder is not configured" -ForegroundColor Yellow
}
   elseif ( $PurgeItemsAfterDelay  -eq  '0' )
{
write-host " Configure removal of items from Quarantine folder is disabled" -ForegroundColor Green
}
  elseif ( $PurgeItemsAfterDelay  -eq  '1' )
{
write-host " Configure removal of items from Quarantine folder is enabled" -ForegroundColor Red
}
  else
{
write-host " Configure removal of items from Quarantine folder is set to an unknown setting" -ForegroundColor Red
}

$AllowPause = Get-ItemProperty -Path 'Registry::HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Microsoft Antimalware\Scan\' -Name AllowPause -ErrorAction SilentlyContinue|Select-Object -ExpandProperty AllowPause
if ( $AllowPause -eq $null)
{
write-host " Allow users to pause scan is not configured" -ForegroundColor Yellow
}
   elseif ( $AllowPause  -eq  '0' )
{
write-host " Allow users to pause scan is disabled" -ForegroundColor Green
}
  elseif ( $AllowPause  -eq  '1' )
{
write-host " Allow users to pause scan is enabled" -ForegroundColor Red
}
  else
{
write-host " Allow users to pause scan is set to an unknown setting" -ForegroundColor Red
}

$CheckForSignaturesBeforeRunningScan = Get-ItemProperty -Path 'Registry::HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Microsoft Antimalware\Scan\' -Name CheckForSignaturesBeforeRunningScan -ErrorAction SilentlyContinue|Select-Object -ExpandProperty CheckForSignaturesBeforeRunningScan
if ( $CheckForSignaturesBeforeRunningScan -eq $null)
{
write-host " Check for the latest virus and spyware definitions before running a scheduled scan is not configured" -ForegroundColor Yellow
}
   elseif ( $CheckForSignaturesBeforeRunningScan  -eq  '1' )
{
write-host " Check for the latest virus and spyware definitions before running a scheduled scan is enabled" -ForegroundColor Green
}
  elseif ( $CheckForSignaturesBeforeRunningScan  -eq  '0' )
{
write-host " Check for the latest virus and spyware definitions before running a scheduled scan is disabled" -ForegroundColor Red
}
  else
{
write-host " Check for the latest virus and spyware definitions before running a scheduled scan is set to an unknown setting" -ForegroundColor Red
}

$DisableArchiveScanning = Get-ItemProperty -Path 'Registry::HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Microsoft Antimalware\Scan\' -Name DisableArchiveScanning -ErrorAction SilentlyContinue|Select-Object -ExpandProperty DisableArchiveScanning
if ( $DisableArchiveScanning -eq $null)
{
write-host " Scan archive files is not configured" -ForegroundColor Yellow
}
   elseif ( $DisableArchiveScanning  -eq  '0' )
{
write-host " Scan archive files is enabled" -ForegroundColor Green
}
  elseif ( $DisableArchiveScanning  -eq  '1' )
{
write-host " Scan archive files is disabled" -ForegroundColor Red
}
  else
{
write-host " Scan archive files is set to an unknown setting" -ForegroundColor Red
}

$DisablePackedExeScanning = Get-ItemProperty -Path 'Registry::HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Microsoft Antimalware\Scan\' -Name DisablePackedExeScanning -ErrorAction SilentlyContinue|Select-Object -ExpandProperty DisablePackedExeScanning
if ( $DisablePackedExeScanning -eq $null)
{
write-host " Scan packed executables is not configured" -ForegroundColor Yellow
}
   elseif ( $DisablePackedExeScanning  -eq  '0' )
{
write-host " Scan packed executables is enabled" -ForegroundColor Green
}
  elseif ( $DisablePackedExeScanning  -eq  '1' )
{
write-host " Scan packed executables is disabled" -ForegroundColor Red
}
  else
{
write-host " Scan packed executables is set to an unknown setting" -ForegroundColor Red
}

$DisableRemovableDriveScanning = Get-ItemProperty -Path 'Registry::HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Microsoft Antimalware\Scan\' -Name DisableRemovableDriveScanning -ErrorAction SilentlyContinue|Select-Object -ExpandProperty DisableRemovableDriveScanning
if ( $DisableRemovableDriveScanning -eq $null)
{
write-host " Scan removable drives is not configured" -ForegroundColor Yellow
}
   elseif ( $DisableRemovableDriveScanning  -eq  '0' )
{
write-host " Scan removable drives is enabled" -ForegroundColor Green
}
  elseif ( $DisableRemovableDriveScanning  -eq  '1' )
{
write-host " Scan removable drives is disabled" -ForegroundColor Red
}
  else
{
write-host " Scan removable drives is set to an unknown setting" -ForegroundColor Red
}

$DisableEmailScanning = Get-ItemProperty -Path 'Registry::HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Microsoft Antimalware\Scan\' -Name DisableEmailScanning -ErrorAction SilentlyContinue|Select-Object -ExpandProperty DisableEmailScanning
if ( $DisableEmailScanning -eq $null)
{
write-host " Turn on e-mail scanning is not configured" -ForegroundColor Yellow
}
   elseif ( $DisableEmailScanning  -eq  '0' )
{
write-host " Turn on e-mail scanning is enabled" -ForegroundColor Green
}
  elseif ( $DisableEmailScanning  -eq  '1' )
{
write-host " Turn on e-mail scanning is disabled" -ForegroundColor Red
}
  else
{
write-host " Turn on e-mail scanning is set to an unknown setting" -ForegroundColor Red
}

$DisableHeuristics = Get-ItemProperty -Path 'Registry::HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Microsoft Antimalware\Scan\' -Name DisableHeuristics -ErrorAction SilentlyContinue|Select-Object -ExpandProperty DisableHeuristics
if ( $DisableHeuristics -eq $null)
{
write-host " Turn on heuristics is not configured" -ForegroundColor Yellow
}
   elseif ( $DisableHeuristics  -eq  '0' )
{
write-host " Turn on heuristics is enabled" -ForegroundColor Green
}
  elseif ( $DisableHeuristics  -eq  '1' )
{
write-host " Turn on heuristics is disabled" -ForegroundColor Red
}
  else
{
write-host " Turn on heuristics is set to an unknown setting" -ForegroundColor Red
}


write-host "`r`n####################### ATTACHMENT MANAGER #######################`r`n"

$SaveZoneInformation = Get-ItemProperty -Path Registry::HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Attachments\ -Name SaveZoneInformation -ErrorAction SilentlyContinue|Select-Object -ExpandProperty SaveZoneInformation
if ( $SaveZoneInformation -eq $null)
{
write-host " Do not preserve zone information in file attachments is not configured" -ForegroundColor Yellow
}
   elseif ( $SaveZoneInformation  -eq  '2' )
{
write-host " Do not preserve zone information in file attachments is disabled" -ForegroundColor Green
}
  elseif ( $SaveZoneInformation  -eq  '1' )
{
write-host " Do not preserve zone information in file attachments is enabled" -ForegroundColor Red
}
  else
{
write-host " Do not preserve zone information in file attachments is set to an unknown setting" -ForegroundColor Red
}

$HideZoneInfoOnProperties = Get-ItemProperty -Path Registry::HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Attachments\ -Name HideZoneInfoOnProperties -ErrorAction SilentlyContinue|Select-Object -ExpandProperty HideZoneInfoOnProperties
if ( $HideZoneInfoOnProperties -eq $null)
{
write-host " Hide mechanisms to remove zone information is not configured" -ForegroundColor Yellow
}
   elseif ( $HideZoneInfoOnProperties  -eq  '1' )
{
write-host " Hide mechanisms to remove zone information is enabled" -ForegroundColor Green
}
  elseif ( $HideZoneInfoOnProperties  -eq  '0' )
{
write-host " Hide mechanisms to remove zone information is disabled" -ForegroundColor Red
}
  else
{
write-host " Hide mechanisms to remove zone information is set to an unknown setting" -ForegroundColor Red
}

write-host "`r`n####################### AUDIT EVENT MANAGEMENT #######################`r`n"

$ProcessCreationIncludeCmdLine_Enabled = Get-ItemProperty -Path  'Registry::HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Policies\System\Audit\'  -Name ProcessCreationIncludeCmdLine_Enabled -ErrorAction SilentlyContinue|Select-Object -ExpandProperty ProcessCreationIncludeCmdLine_Enabled
if ( $ProcessCreationIncludeCmdLine_Enabled -eq $null)
{
write-host " Include command line in process creation events is not configured" -ForegroundColor Yellow
}
   elseif ( $ProcessCreationIncludeCmdLine_Enabled  -eq  '1' )
{
write-host " Include command line in process creation events is enabled" -ForegroundColor Green
}
  elseif ( $ProcessCreationIncludeCmdLine_Enabled  -eq  '0' )
{
write-host " Include command line in process creation events is disabled" -ForegroundColor Red
}
  else
{
write-host " Include command line in process creation events is set to an unknown setting" -ForegroundColor Red
}


$1AW2CfpSKiewv0 = Get-ItemProperty -Path  'Registry::HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Windows\EventLog\Application\'  -Name MaxSize -ErrorAction SilentlyContinue|Select-Object -ExpandProperty MaxSize
if ( $1AW2CfpSKiewv0 -eq $null)
{
write-host " Specify the maximum log file size (KB) for the Application Log is not configured" -ForegroundColor Yellow
}
   elseif ( $1AW2CfpSKiewv0  -eq  '65536' )
{
write-host " Specify the maximum log file size (KB) for the Application Log is set to a compliant setting" -ForegroundColor Green
}
  elseif ( $1AW2CfpSKiewv0  -lt  '65536' )
{
write-host " Specify the maximum log file size (KB) for the Application Log is set to $1AW2CfpSKiewv0 which is a lower value than 65536 required for compliance" -ForegroundColor Red
}
  elseif ( $1AW2CfpSKiewv0  -gt  '65536' )
{
write-host " Specify the maximum log file size (KB) for the Application Log is set to $1AW2CfpSKiewv0 which is a higher value than 65536 required for compliance" -ForegroundColor Green
}
  else
{
write-host " Specify the maximum log file size (KB) for the Application Log is set to an unknown setting" -ForegroundColor Red
}

$1AW2CfpSKiewv = Get-ItemProperty -Path  'Registry::HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Windows\EventLog\Security\'  -Name MaxSize -ErrorAction SilentlyContinue|Select-Object -ExpandProperty MaxSize
if ( $1AW2CfpSKiewv -eq $null)
{
write-host " Specify the maximum log file size (KB) for the Security Log is not configured" -ForegroundColor Yellow
}
   elseif ( $1AW2CfpSKiewv  -eq  '65536' )
{
write-host " Specify the maximum log file size (KB) for the Security Log is set to a compliant setting" -ForegroundColor Green
}
  elseif ( $1AW2CfpSKiewv  -lt  '65536' )
{
write-host " Specify the maximum log file size (KB) for the Security Log is set to $1AW2CfpSKiewv which is a lower value than 65536 required for compliance" -ForegroundColor Red
}
  elseif ( $1AW2CfpSKiewv  -gt  '65536' )
{
write-host " Specify the maximum log file size (KB) for the Security Log is set to $1AW2CfpSKiewv which is a higher value than 65536 required for compliance" -ForegroundColor Green
}
  else
{
write-host " Specify the maximum log file size (KB) for the Security Log is set to an unknown setting" -ForegroundColor Red
}

$1AW2CfpSKiew = Get-ItemProperty -Path  'Registry::HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Windows\EventLog\System\'  -Name MaxSize -ErrorAction SilentlyContinue|Select-Object -ExpandProperty MaxSize
if ( $1AW2CfpSKiew -eq $null)
{
write-host " Specify the maximum log file size (KB) for the System Log is not configured" -ForegroundColor Yellow
}
   elseif ( $1AW2CfpSKiew  -eq  '65536' )
{
write-host " Specify the maximum log file size (KB) for the System Log is set to a compliant setting" -ForegroundColor Green
}
  elseif ( $1AW2CfpSKiew  -lt  '65536' )
{
write-host " Specify the maximum log file size (KB) for the System Log is set to $1AW2CfpSKiew which is a lower value than 65536 required for compliance" -ForegroundColor Red
}
  elseif ( $1AW2CfpSKiew  -gt  '65536' )
{
write-host " Specify the maximum log file size (KB) for the System Log is set to $1AW2CfpSKiew which is a higher value than 65536 required for compliance" -ForegroundColor Green
}
  else
{
write-host " Specify the maximum log file size (KB) for the System Log is set to an unknown setting" -ForegroundColor Red
}

$1AW2CfpSKiewv0n = Get-ItemProperty -Path  'Registry::HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Windows\EventLog\Setup\'  -Name MaxSize -ErrorAction SilentlyContinue|Select-Object -ExpandProperty MaxSize
if ( $1AW2CfpSKiewv0n -eq $null)
{
write-host " Specify the maximum log file size (KB) for the Setup Log is not configured" -ForegroundColor Yellow
}
   elseif ( $1AW2CfpSKiewv0n  -eq  '65536' )
{
write-host " Specify the maximum log file size (KB) for the Setup Log is set to a compliant setting" -ForegroundColor Green
}
  elseif ( $1AW2CfpSKiewv0n  -lt  '65536' )
{
write-host " Specify the maximum log file size (KB) for the Setup Log is set to $1AW2CfpSKiewv0n which is a lower value than 65536 required for compliance" -ForegroundColor Red
}
  elseif ( $1AW2CfpSKiewv0n  -gt  '65536' )
{
write-host " Specify the maximum log file size (KB) for the Setup Log is set to $1AW2CfpSKiewv0n which is a higher value than 65536 required for compliance" -ForegroundColor Green
}
  else
{
write-host " Specify the maximum log file size (KB) for the Setup Log is set to an unknown setting" -ForegroundColor Red
}

write-host "`r`nManage Auditing and Security Log is unable to be checked using PowerShell, as the setting is not a registry key. Please check Computer Configuration\Policies\Windows Settings\Security Settings\Local Policies\User Rights Assignment. ASD Recommendation is to only have 'Administrators' present" -ForegroundColor Cyan

write-host "`r`nAudit Credential Validation is unable to be checked using PowerShell, as the setting is not a registry key. Please check Computer Configuration\Policies\Windows Settings\Security Settings\Advanced Audit Policy Configuration\Audit Policies\Account Logon. ASD Recommendation is to have 'Success and Failure' Present" -ForegroundColor Cyan

write-host "`r`nFor the below controls please check: Computer Configuration\Policies\Windows Settings\Security Settings\Advanced Audit Policy Configuration\Audit Policies\Account Management"  -ForegroundColor Cyan
write-host "    Audit Computer Account Management is unable to be checked using PowerShell, as the setting is not a registry key. ASD Recommendation is to have 'Success and Failure' Present" -ForegroundColor Cyan

write-host "    Audit Other Account Management Events is unable to be checked using PowerShell, as the setting is not a registry key. ASD Recommendation is to have 'Success and Failure' Present" -ForegroundColor Cyan

write-host "    Audit Security Group Management is unable to be checked using PowerShell, as the setting is not a registry key. ASD Recommendation is to have 'Success and Failure' Present" -ForegroundColor Cyan

write-host "    Audit User Account Management is unable to be checked using PowerShell, as the setting is not a registry key. ASD Recommendation is to have 'Success and Failure' Present" -ForegroundColor Cyan

write-host "`r`nFor the below controls please check: Computer Configuration\Policies\Windows Settings\Security Settings\Advanced Audit Policy Configuration\Audit Policies\Detailed Tracking"  -ForegroundColor Cyan

write-host "    Audit PNP Activity is unable to be checked using PowerShell, as the setting is not a registry key. ASD Recommendation is to have 'Success' Present" -ForegroundColor Cyan

write-host "    Audit Process Creation is unable to be checked using PowerShell, as the setting is not a registry key. ASD Recommendation is to have 'Success' Present" -ForegroundColor Cyan

write-host "    Audit Process Termination is unable to be checked using PowerShell, as the setting is not a registry key. ASD Recommendation is to have 'Success' Present" -ForegroundColor Cyan

write-host "`r`nFor the below controls please check: Computer Configuration\Policies\Windows Settings\Security Settings\Advanced Audit Policy Configuration\Audit Policies\Logon/Logoff"  -ForegroundColor Cyan

write-host "    Audit Account Lockout is unable to be checked using PowerShell, as the setting is not a registry key. Please check. ASD Recommendation is to have 'Success and Failure' Present" -ForegroundColor Cyan

write-host "    Audit Group Membership is unable to be checked using PowerShell, as the setting is not a registry key. ASD Recommendation is to have 'Success' Present" -ForegroundColor Cyan

write-host "    Audit Logoff is unable to be checked using PowerShell, as the setting is not a registry key. ASD Recommendation is to have 'Success' Present" -ForegroundColor Cyan

write-host "    Audit Logon is unable to be checked using PowerShell, as the setting is not a registry key. ASD Recommendation is to have 'Success and Failure' Present" -ForegroundColor Cyan

write-host "    Audit Other Logon/Logoff Events is unable to be checked using PowerShell, as the setting is not a registry key. ASD Recommendation is to have 'Success and Failure' Present" -ForegroundColor Cyan

write-host "    Audit Audit Special Logon is unable to be checked using PowerShell, as the setting is not a registry key. ASD Recommendation is to have 'Success and Failure' Present" -ForegroundColor Cyan

write-host "`r`nFor the below controls please check: Computer Configuration\Policies\Windows Settings\Security Settings\Advanced Audit Policy Configuration\Audit Policies\Object Access"  -ForegroundColor Cyan

write-host "    Audit File Share is unable to be checked using PowerShell, as the setting is not a registry key. ASD Recommendation is to have 'Success and Failure' Present" -ForegroundColor Cyan

write-host "    Audit Kernel Object is unable to be checked using PowerShell, as the setting is not a registry key. ASD Recommendation is to have 'Success and Failure' Present" -ForegroundColor Cyan

write-host "    Audit Other Object Access Events is unable to be checked using PowerShell, as the setting is not a registry key. ASD Recommendation is to have 'Success and Failure' Present" -ForegroundColor Cyan

write-host "    Audit Removable Storage is unable to be checked using PowerShell, as the setting is not a registry key ASD Recommendation is to have 'Success and Failure' Present" -ForegroundColor Cyan

write-host "`r`nFor the below controls please check: Computer Configuration\Policies\Windows Settings\Security Settings\Advanced Audit Policy Configuration\Audit Policies\Policy Change"  -ForegroundColor Cyan

write-host "    Audit Audit Policy Change is unable to be checked using PowerShell, as the setting is not a registry key. ASD Recommendation is to have 'Success and Failure' Present" -ForegroundColor Cyan

write-host "    Audit Authentication Policy Change is unable to be checked using PowerShell, as the setting is not a registry key. ASD Recommendation is to have 'Success' Present" -ForegroundColor Cyan

write-host "    Audit Authorization Policy Change is unable to be checked using PowerShell, as the setting is not a registry key. ASD Recommendation is to have 'Success' Present" -ForegroundColor Cyan

write-host "`r`nAudit Sensitive Privilege Use is unable to be checked using PowerShell, as the setting is not a registry key. Please check Computer Configuration\Policies\Windows Settings\Security Settings\Advanced Audit Policy Configuration\Audit Policies\Privilege Use. ASD Recommendation is to have 'Success and Failure' Present" -ForegroundColor Cyan

write-host "`r`nFor the below controls please check: Computer Configuration\Policies\Windows Settings\Security Settings\Advanced Audit Policy Configuration\Audit Policies\System"  -ForegroundColor Cyan

write-host "    Audit IPsec Driver is unable to be checked using PowerShell, as the setting is not a registry key. ASD Recommendation is to have 'Success and Failure' Present" -ForegroundColor Cyan

write-host "    Audit Other System Events is unable to be checked using PowerShell, as the setting is not a registry key. ASD Recommendation is to have 'Success and Failure' Present" -ForegroundColor Cyan

write-host "    Audit Security State Change is unable to be checked using PowerShell, as the setting is not a registry key. ASD Recommendation is to have 'Success' Present" -ForegroundColor Cyan

write-host "    Audit Security System Extension is unable to be checked using PowerShell, as the setting is not a registry key. ASD Recommendation is to have 'Success and Failure' Present" -ForegroundColor Cyan

write-host "    Audit System Integrity is unable to be checked using PowerShell, as the setting is not a registry key. ASD Recommendation is to have 'Success and Failure' Present" -ForegroundColor Cyan

$SCENoApplyLegacyAuditPolicy = Get-ItemProperty -Path  'Registry::HKEY_LOCAL_MACHINE\System\CurrentControlSet\Control\Lsa\'  -Name SCENoApplyLegacyAuditPolicy -ErrorAction SilentlyContinue|Select-Object -ExpandProperty SCENoApplyLegacyAuditPolicy
if ( $SCENoApplyLegacyAuditPolicy -eq $null)
{
write-host "Audit: Force audit policy subcategory settings (Windows Vista or later) to override audit policy category settings is not configured" -ForegroundColor Yellow
}
   elseif ( $SCENoApplyLegacyAuditPolicy -eq  '1' )
{
write-host "Audit: Force audit policy subcategory settings (Windows Vista or later) to override audit policy category settings is enabled" -ForegroundColor Green
}
  elseif ( $SCENoApplyLegacyAuditPolicy  -eq  '0' )
{
write-host "Audit: Force audit policy subcategory settings (Windows Vista or later) to override audit policy category settings is disabled" -ForegroundColor Red
}
  else
{
write-host "Audit: Force audit policy subcategory settings (Windows Vista or later) to override audit policy category settings is set to an unknown setting" -ForegroundColor Red
}


write-host "`r`n####################### AUTOPLAY AND AUTORUN #######################`r`n"

$LMNoAutoplayfornonVolume = Get-ItemProperty -Path  'Registry::HKLM\Software\Policies\Microsoft\Windows\Explorer\' -Name NoAutoplayfornonVolume -ErrorAction SilentlyContinue|Select-Object -ExpandProperty NoAutoplayfornonVolume
$UPNoAutoplayfornonVolume = Get-ItemProperty -Path  'Registry::HKCU\Software\Policies\Microsoft\Windows\Explorer\' -Name NoAutoplayfornonVolume -ErrorAction SilentlyContinue|Select-Object -ExpandProperty NoAutoplayfornonVolume
if ( $LMNoAutoplayfornonVolume -eq $null -and  $UPNoAutoplayfornonVolume -eq $null)
{
write-host " Disallow Autoplay for non-volume devices is not configured" -ForegroundColor Yellow
}
if ( $LMNoAutoplayfornonVolume  -eq '1' )
{
write-host " Disallow Autoplay for non-volume devices is enabled in Local Machine GP" -ForegroundColor Green
}
if ( $LMNoAutoplayfornonVolume  -eq '0' )
{
write-host " Disallow Autoplay for non-volume devices is disabled in Local Machine GP" -ForegroundColor Red
}
if ( $UPNoAutoplayfornonVolume  -eq  '1' )
{
write-host " Disallow Autoplay for non-volume devices is enabled in User GP" -ForegroundColor Green
}
if ( $UPNoAutoplayfornonVolume  -eq  '0' )
{
write-host " Disallow Autoplay for non-volume devices is disabled in User GP" -ForegroundColor Red
}

$LMNoAutorun = Get-ItemProperty -Path  'Registry::HKLM\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\' -Name NoAutorun -ErrorAction SilentlyContinue|Select-Object -ExpandProperty NoAutorun
$UPNoAutorun = Get-ItemProperty -Path  'Registry::HKCU\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\' -Name NoAutorun -ErrorAction SilentlyContinue|Select-Object -ExpandProperty NoAutorun
if ( $LMNoAutorun -eq $null -and  $UPNoAutorun -eq $null)
{
write-host " Set the default behavior for AutoRun is not configured" -ForegroundColor Yellow
}
if ( $LMNoAutorun  -eq '1' )
{
write-host " Set the default behavior for AutoRun is enabled in Local Machine GP" -ForegroundColor Green
}
if ( $LMNoAutorun  -eq '2' )
{
write-host " Set the default behavior for AutoRun is disabled in Local Machine GP" -ForegroundColor Red
}
if ( $UPNoAutorun  -eq  '1' )
{
write-host " Set the default behavior for AutoRun is enabled in User GP" -ForegroundColor Green
}
if ( $UPNoAutorun  -eq  '2' )
{
write-host " Set the default behavior for AutoRun is disabled in User GP" -ForegroundColor Red
}

$LMNoDriveTypeAutoRun = Get-ItemProperty -Path  'Registry::HKLM\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\' -Name NoDriveTypeAutoRun -ErrorAction SilentlyContinue|Select-Object -ExpandProperty NoDriveTypeAutoRun
$UPNoDriveTypeAutoRun = Get-ItemProperty -Path  'Registry::HKCU\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\' -Name NoDriveTypeAutoRun -ErrorAction SilentlyContinue|Select-Object -ExpandProperty NoDriveTypeAutoRun
if ( $LMNoDriveTypeAutoRun -eq $null -and  $UPNoDriveTypeAutoRun -eq $null)
{
write-host " Turn off Autoplay is not configured" -ForegroundColor Yellow
}
if ( $LMNoDriveTypeAutoRun  -eq '255' )
{
write-host " Turn off Autoplay is enabled in Local Machine GP" -ForegroundColor Green
}
if ( $LMNoDriveTypeAutoRun  -eq '181' )
{
write-host " Turn off Autoplay is disabled in Local Machine GP" -ForegroundColor Red
}
if ( $UPNoDriveTypeAutoRun  -eq  '255' )
{
write-host " Turn off Autoplay is enabled in User GP" -ForegroundColor Green
}
if ( $UPNoDriveTypeAutoRun  -eq  '181' )
{
write-host " Turn off Autoplay is disabled in User GP" -ForegroundColor Red
}

write-host "`r`n####################### BIOS AND UEFI PASSWORDS #######################`r`n"
write-host "`r`n####################### BOOT DEVICES #######################`r`n"
write-host "`r`n####################### BRIDGING NETWORKS #######################`r`n"

$NC_AllowNetBridge_NLA = Get-ItemProperty -Path  'Registry::HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Windows\Network Connections\'  -Name NC_AllowNetBridge_NLA -ErrorAction SilentlyContinue|Select-Object -ExpandProperty NC_AllowNetBridge_NLA
if ( $NC_AllowNetBridge_NLA -eq $null)
{
write-host " Prohibit installation and configuration of Network Bridge on your DNS domain network is not configured" -ForegroundColor Yellow
}
   elseif ( $NC_AllowNetBridge_NLA  -eq  '0' )
{
write-host " Prohibit installation and configuration of Network Bridge on your DNS domain network is enabled" -ForegroundColor Green
}
  elseif ( $NC_AllowNetBridge_NLA  -eq  '1' )
{
write-host " Prohibit installation and configuration of Network Bridge on your DNS domain network is disabled" -ForegroundColor Red
}
  else
{
write-host " Prohibit installation and configuration of Network Bridge on your DNS domain network is set to an unknown setting" -ForegroundColor Red
}

$Force_Tunneling = Get-ItemProperty -Path  'Registry::HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Windows\TCPIP\v6Transition\'  -Name Force_Tunneling -ErrorAction SilentlyContinue|Select-Object -ExpandProperty Force_Tunneling
if ( $Force_Tunneling -eq $null)
{
write-host " Route all traffic through the internal network is not configured" -ForegroundColor Yellow
}
   elseif ( $Force_Tunneling  -eq  'Enabled' )
{
write-host " Route all traffic through the internal network is enabled" -ForegroundColor Green
}
  elseif ( $Force_Tunneling  -eq  'Disabled' )
{
write-host " Route all traffic through the internal network is disabled" -ForegroundColor Red
}
  else
{
write-host " Route all traffic through the internal network is set to an unknown setting" -ForegroundColor Red
}

$fBlockNonDomain = Get-ItemProperty -Path  'Registry::HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Windows\WcmSvc\GroupPolicy\'  -Name fBlockNonDomain -ErrorAction SilentlyContinue|Select-Object -ExpandProperty fBlockNonDomain
if ( $fBlockNonDomain -eq $null)
{
write-host " Prohibit connection to non-domain networks when connected to domain authenticated network is not configured" -ForegroundColor Yellow
}
   elseif ( $fBlockNonDomain  -eq  '1' )
{
write-host " Prohibit connection to non-domain networks when connected to domain authenticated network is enabled" -ForegroundColor Green
}
  elseif ( $fBlockNonDomain  -eq  '0' )
{
write-host " Prohibit connection to non-domain networks when connected to domain authenticated network is disabled" -ForegroundColor Red
}
  else
{
write-host " Prohibit connection to non-domain networks when connected to domain authenticated network is set to an unknown setting" -ForegroundColor Red
}


write-host "`r`n####################### BUILT-IN GUEST ACCOUNTS #######################`r`n"

write-host "Unable to check Accounts: Guest account status"
write-host "Unable to Deny log on locally"

write-host "`r`n####################### CASE LOCKS #######################`r`n"
write-host "`r`n####################### CD BURNER ACCESS #######################`r`n"

$NoCDBurning = Get-ItemProperty -Path  'Registry::HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\'  -Name NoCDBurning -ErrorAction SilentlyContinue|Select-Object -ExpandProperty NoCDBurning
if ( $NoCDBurning -eq $null)
{
write-host " Remove CD Burning features is not configured" -ForegroundColor Yellow
}
   elseif ( $NoCDBurning  -eq  '1' )
{
write-host " Remove CD Burning features is enabled" -ForegroundColor Green
}
  elseif ( $NoCDBurning  -eq  '0' )
{
write-host " Remove CD Burning features is disabled" -ForegroundColor Red
}
  else
{
write-host " Remove CD Burning features is set to an unknown setting" -ForegroundColor Red
}

write-host "`r`n####################### CENTRALISED AUDIT EVENT LOGGING #######################`r`n"
write-host "`r`n####################### COMMAND PROMPT #######################`r`n"

$DisableCMD = Get-ItemProperty -Path  'Registry::HKEY_CURRENT_USER\Software\Policies\Microsoft\Windows\System\'  -Name DisableCMD -ErrorAction SilentlyContinue|Select-Object -ExpandProperty DisableCMD
if ( $DisableCMD -eq $null)
{
write-host " Prevent access to the command prompt is not configured" -ForegroundColor Yellow
}
   elseif ( $DisableCMD  -eq  '1' )
{
write-host " Prevent access to the command prompt is enabled" -ForegroundColor Green
}
  elseif ( $DisableCMD  -eq  '2' )
{
write-host " Prevent access to the command prompt is disabled" -ForegroundColor Red
}
  else
{
write-host " Prevent access to the command prompt is set to an unknown setting" -ForegroundColor Red
}

write-host "`r`n####################### DIRECT MEMORY ACCESS #######################`r`n"
write-host "Unable to check Prevent installation of devices that match any of these device IDs"
write-host "Unable to check Prevent installation of devices using drivers that match these device setup classes"

write-host "`r`n####################### ENDPOINT DEVICE CONTROL #######################`r`n"

$Deny_All = Get-ItemProperty -Path  'Registry::HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Windows\RemovableStorageDevices\'  -Name Deny_All -ErrorAction SilentlyContinue|Select-Object -ExpandProperty Deny_All
if ( $Deny_All -eq $null)
{
write-host " All Removable Storage classes: Deny all access is not configured" -ForegroundColor Yellow
}
   elseif ( $Deny_All  -eq  '1' )
{
write-host " All Removable Storage classes: Deny all access is enabled" -ForegroundColor Green
}
  elseif ( $Deny_All  -eq  '0' )
{
write-host " All Removable Storage classes: Deny all access is disabled" -ForegroundColor Red
}
  else
{
write-host " All Removable Storage classes: Deny all access is set to an unknown setting" -ForegroundColor Red
}

$Deny_Execute = Get-ItemProperty -Path  'Registry::HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Windows\RemovableStorageDevices\{53f56308-b6bf-11d0-94f2-00a0c91efb8b}\'  -Name Deny_Execute -ErrorAction SilentlyContinue|Select-Object -ExpandProperty Deny_Execute
if ( $Deny_Execute -eq $null)
{
write-host " CD and DVD: Deny execute access is not configured" -ForegroundColor Yellow
}
   elseif ( $Deny_Execute  -eq  '1' )
{
write-host " CD and DVD: Deny execute access is enabled" -ForegroundColor Green
}
  elseif ( $Deny_Execute  -eq  '0' )
{
write-host " CD and DVD: Deny execute access is disabled" -ForegroundColor Red
}
  else
{
write-host " CD and DVD: Deny execute access is set to an unknown setting" -ForegroundColor Red
}


$Deny_Read = Get-ItemProperty -Path  'Registry::HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Windows\RemovableStorageDevices\{53f56308-b6bf-11d0-94f2-00a0c91efb8b}\'  -Name Deny_Read -ErrorAction SilentlyContinue|Select-Object -ExpandProperty Deny_Read
if ( $Deny_Read -eq $null)
{
write-host " CD and DVD: Deny read access is not configured" -ForegroundColor Yellow
}
   elseif ( $Deny_Read  -eq  '0' )
{
write-host " CD and DVD: Deny read access is disabled" -ForegroundColor Green
}
  elseif ( $Deny_Read  -eq  '1' )
{
write-host " CD and DVD: Deny read access is enabled" -ForegroundColor Red
}
  else
{
write-host " CD and DVD: Deny read access is set to an unknown setting" -ForegroundColor Red
}

$Deny_Write = Get-ItemProperty -Path  'Registry::HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Windows\RemovableStorageDevices\{53f56308-b6bf-11d0-94f2-00a0c91efb8b}\'  -Name Deny_Write -ErrorAction SilentlyContinue|Select-Object -ExpandProperty Deny_Write
if ( $Deny_Write -eq $null)
{
write-host " CD and DVD: Deny write access is not configured" -ForegroundColor Yellow
}
   elseif ( $Deny_Write  -eq  '1' )
{
write-host " CD and DVD: Deny write access is enabled" -ForegroundColor Green
}
  elseif ( $Deny_Write  -eq  '0' )
{
write-host " CD and DVD: Deny write access is disabled" -ForegroundColor Red
}
  else
{
write-host " CD and DVD: Deny write access is set to an unknown setting" -ForegroundColor Red
}

$Deny_Read2 = Get-ItemProperty -Path  'Registry::HKEY_CURRENT_USER\Software\Policies\Microsoft\Windows\RemovableStorageDevices\Custom\Deny_Read\'  -Name Deny_Read -ErrorAction SilentlyContinue|Select-Object -ExpandProperty Deny_Read
if ( $Deny_Read2 -eq $null)
{
write-host " Custom Classes: Deny read access is not configured" -ForegroundColor Yellow
}
   elseif ( $Deny_Read2  -eq  '0' )
{
write-host " Custom Classes: Deny read access is disabled" -ForegroundColor Green
}
  elseif ( $Deny_Read2  -eq  '1' )
{
write-host " Custom Classes: Deny read access is enabled" -ForegroundColor Red
}
  else
{
write-host " Custom Classes: Deny read access is set to an unknown setting" -ForegroundColor Red
}

$Deny_Write2 = Get-ItemProperty -Path  'Registry::HKEY_CURRENT_USER\Software\Policies\Microsoft\Windows\RemovableStorageDevices\Custom\Deny_Write\'  -Name Deny_Write -ErrorAction SilentlyContinue|Select-Object -ExpandProperty Deny_Write
if ( $Deny_Write2 -eq $null)
{
write-host " Custom Classes: Deny write access is not configured" -ForegroundColor Yellow
}
   elseif ( $Deny_Write2  -eq  '1' )
{
write-host " Custom Classes: Deny write access is enabled" -ForegroundColor Green
}
  elseif ( $Deny_Write2  -eq  '0' )
{
write-host " Custom Classes: Deny write access is disabled" -ForegroundColor Red
}
  else
{
write-host " Custom Classes: Deny write access is set to an unknown setting" -ForegroundColor Red
}

$Deny_Execute3 = Get-ItemProperty -Path  'Registry::HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Windows\RemovableStorageDevices\{53f56311-b6bf-11d0-94f2-00a0c91efb8b}\'  -Name Deny_Execute -ErrorAction SilentlyContinue|Select-Object -ExpandProperty Deny_Execute
if ( $Deny_Execute3 -eq $null)
{
write-host " Floppy Drives: Deny execute access is not configured" -ForegroundColor Yellow
}
   elseif ( $Deny_Execute3  -eq  '1' )
{
write-host " Floppy Drives: Deny execute access is enabled" -ForegroundColor Green
}
  elseif ( $Deny_Execute3  -eq  '0' )
{
write-host " Floppy Drives: Deny execute access is disabled" -ForegroundColor Red
}
  else
{
write-host " Floppy Drives: Deny execute access is set to an unknown setting" -ForegroundColor Red
}

$Deny_Read3 = Get-ItemProperty -Path  'Registry::HKEY_CURRENT_USER\Software\Policies\Microsoft\Windows\RemovableStorageDevices\{53f56311-b6bf-11d0-94f2-00a0c91efb8b}\'  -Name Deny_Read -ErrorAction SilentlyContinue|Select-Object -ExpandProperty Deny_Read
if ( $Deny_Read3 -eq $null)
{
write-host " Floppy Drives: Deny read access is not configured" -ForegroundColor Yellow
}
   elseif ( $Deny_Read3  -eq  '0' )
{
write-host " Floppy Drives: Deny read access is disabled" -ForegroundColor Green
}
  elseif ( $Deny_Read3  -eq  '1' )
{
write-host " Floppy Drives: Deny read access is enabled" -ForegroundColor Red
}
  else
{
write-host " Floppy Drives: Deny read access is set to an unknown setting" -ForegroundColor Red
}

$Deny_Write3 = Get-ItemProperty -Path  'Registry::HKEY_CURRENT_USER\Software\Policies\Microsoft\Windows\RemovableStorageDevices\{53f56311-b6bf-11d0-94f2-00a0c91efb8b}\'  -Name Deny_Write -ErrorAction SilentlyContinue|Select-Object -ExpandProperty Deny_Write
if ( $Deny_Write3 -eq $null)
{
write-host " Floppy Drives: Deny write access is not configured" -ForegroundColor Yellow
}
   elseif ( $Deny_Write3  -eq  '1' )
{
write-host " Floppy Drives: Deny write access is enabled" -ForegroundColor Green
}
  elseif ( $Deny_Write3  -eq  '0' )
{
write-host " Floppy Drives: Deny write access is disabled" -ForegroundColor Red
}
  else
{
write-host " Floppy Drives: Deny write access is set to an unknown setting" -ForegroundColor Red
}

$Deny_Execute4 = Get-ItemProperty -Path  'Registry::HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Windows\RemovableStorageDevices\{53f5630d-b6bf-11d0-94f2-00a0c91efb8b}\'  -Name Deny_Execute -ErrorAction SilentlyContinue|Select-Object -ExpandProperty Deny_Execute
if ( $Deny_Execute4 -eq $null)
{
write-host " Removable Disks: Deny execute access is not configured" -ForegroundColor Yellow
}
   elseif ( $Deny_Execute4  -eq  '1' )
{
write-host " Removable Disks: Deny execute access is enabled" -ForegroundColor Green
}
  elseif ( $Deny_Execute4  -eq  '0' )
{
write-host " Removable Disks: Deny execute access is disabled" -ForegroundColor Red
}
  else
{
write-host " Removable Disks: Deny execute access is set to an unknown setting" -ForegroundColor Red
}

$Deny_Read4 = Get-ItemProperty -Path  'Registry::HKEY_CURRENT_USER\Software\Policies\Microsoft\Windows\RemovableStorageDevices\{53f5630d-b6bf-11d0-94f2-00a0c91efb8b}\'  -Name Deny_Read -ErrorAction SilentlyContinue|Select-Object -ExpandProperty Deny_Read
if ( $Deny_Read4 -eq $null)
{
write-host " Removable Disks: Deny read access is not configured" -ForegroundColor Yellow
}
   elseif ( $Deny_Read4  -eq  '0' )
{
write-host " Removable Disks: Deny read access is disabled" -ForegroundColor Green
}
  elseif ( $Deny_Read4  -eq  '1' )
{
write-host " Removable Disks: Deny read access is enabled" -ForegroundColor Red
}
  else
{
write-host " Removable Disks: Deny read access is set to an unknown setting" -ForegroundColor Red
}

$Deny_Write4 = Get-ItemProperty -Path  'Registry::HKEY_CURRENT_USER\Software\Policies\Microsoft\Windows\RemovableStorageDevices\{53f5630d-b6bf-11d0-94f2-00a0c91efb8b}\'  -Name Deny_Write -ErrorAction SilentlyContinue|Select-Object -ExpandProperty Deny_Write
if ( $Deny_Write4 -eq $null)
{
write-host " Removable Disks: Deny write access is not configured" -ForegroundColor Yellow
}
   elseif ( $Deny_Write4  -eq  '1' )
{
write-host " Removable Disks: Deny write access is enabled" -ForegroundColor Green
}
  elseif ( $Deny_Write4  -eq  '0' )
{
write-host " Removable Disks: Deny write access is disabled" -ForegroundColor Red
}
  else
{
write-host " Removable Disks: Deny write access is set to an unknown setting" -ForegroundColor Red
}

$Deny_Execute5 = Get-ItemProperty -Path  'Registry::HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Windows\RemovableStorageDevices\{53f5630b-b6bf-11d0-94f2-00a0c91efb8b}\'  -Name Deny_Execute -ErrorAction SilentlyContinue|Select-Object -ExpandProperty Deny_Execute
if ( $Deny_Execute5 -eq $null)
{
write-host " Tape Drives: Deny execute access is not configured" -ForegroundColor Yellow
}
   elseif ( $Deny_Execute5  -eq  '1' )
{
write-host " Tape Drives: Deny execute access is enabled" -ForegroundColor Green
}
  elseif ( $Deny_Execute5  -eq  '0' )
{
write-host " Tape Drives: Deny execute access is disabled" -ForegroundColor Red
}
  else
{
write-host " Tape Drives: Deny execute access is set to an unknown setting" -ForegroundColor Red
}

$Deny_Read5 = Get-ItemProperty -Path  'Registry::HKEY_CURRENT_USER\Software\Policies\Microsoft\Windows\RemovableStorageDevices\{53f5630b-b6bf-11d0-94f2-00a0c91efb8b}\'  -Name Deny_Read -ErrorAction SilentlyContinue|Select-Object -ExpandProperty Deny_Read
if ( $Deny_Read5 -eq $null)
{
write-host " Tape Drives: Deny read access is not configured" -ForegroundColor Yellow
}
   elseif ( $Deny_Read5  -eq  '0' )
{
write-host " Tape Drives: Deny read access is disabled" -ForegroundColor Green
}
  elseif ( $Deny_Read5  -eq  '1' )
{
write-host " Tape Drives: Deny read access is enabled" -ForegroundColor Red
}
  else
{
write-host " Tape Drives: Deny read access is set to an unknown setting" -ForegroundColor Red
}

$Deny_Write5 = Get-ItemProperty -Path  'Registry::HKEY_CURRENT_USER\Software\Policies\Microsoft\Windows\RemovableStorageDevices\{53f5630b-b6bf-11d0-94f2-00a0c91efb8b}\'  -Name Deny_Write -ErrorAction SilentlyContinue|Select-Object -ExpandProperty Deny_Write
if ( $Deny_Write5 -eq $null)
{
write-host " Tape Drives: Deny write access is not configured" -ForegroundColor Yellow
}
   elseif ( $Deny_Write5  -eq  '1' )
{
write-host " Tape Drives: Deny write access is enabled" -ForegroundColor Green
}
  elseif ( $Deny_Write5  -eq  '0' )
{
write-host " Tape Drives: Deny write access is disabled" -ForegroundColor Red
}
  else
{
write-host " Tape Drives: Deny write access is set to an unknown setting" -ForegroundColor Red
}
write-host "Unable to check WPD Devices: Deny read access"
write-host "Unable to check WPD Devices: Deny write access"

write-host "`r`n####################### FILE AND PRINT SHARING #######################`r`n"

$DisableHomeGroup = Get-ItemProperty -Path  'Registry::HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Windows\HomeGroup\'  -Name DisableHomeGroup -ErrorAction SilentlyContinue|Select-Object -ExpandProperty DisableHomeGroup
if ( $DisableHomeGroup -eq $null)
{
write-host " Prevent the computer from joining a homegroup is not configured" -ForegroundColor Yellow
}
   elseif ( $DisableHomeGroup  -eq  '1' )
{
write-host " Prevent the computer from joining a homegroup is enabled" -ForegroundColor Green
}
  elseif ( $DisableHomeGroup  -eq  '0' )
{
write-host " Prevent the computer from joining a homegroup is disabled" -ForegroundColor Red
}
  else
{
write-host " Prevent the computer from joining a homegroup is set to an unknown setting" -ForegroundColor Red
}

$NoInplaceSharing = Get-ItemProperty -Path  'Registry::HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\'  -Name NoInplaceSharing -ErrorAction SilentlyContinue|Select-Object -ExpandProperty NoInplaceSharing
if ( $NoInplaceSharing -eq $null)
{
write-host " Prevent users from sharing files within their profile. is not configured" -ForegroundColor Yellow
}
   elseif ( $NoInplaceSharing  -eq  '1' )
{
write-host " Prevent users from sharing files within their profile. is enabled" -ForegroundColor Green
}
  elseif ( $NoInplaceSharing  -eq  '0' )
{
write-host " Prevent users from sharing files within their profile. is disabled" -ForegroundColor Red
}
  else
{
write-host " Prevent users from sharing files within their profile. is set to an unknown setting" -ForegroundColor Red
}

write-host "`r`n####################### GROUP POLICY PROCESSING #######################`r`n"
write-host "Unable to check Hardened UNC Paths"

$NoBackgroundPolicy = Get-ItemProperty -Path  'Registry::HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Windows\Group Policy\{35378EAC-683F-11D2-A89A-00C04FBBCFA2}\'  -Name NoGPOListChanges -ErrorAction SilentlyContinue|Select-Object -ExpandProperty NoGPOListChanges
if ( $NoBackgroundPolicy -eq $null)
{
write-host " Configure registry policy processing is not configured" -ForegroundColor Yellow
}
   elseif ( $NoBackgroundPolicy  -eq  '0' )
{
write-host " Configure registry policy processing is enabled" -ForegroundColor Green
}
  elseif ( $NoBackgroundPolicy  -eq  '1' )
{
write-host " Configure registry policy processing is disabled" -ForegroundColor Red
}
  else
{
write-host " Configure registry policy processing is set to an unknown setting" -ForegroundColor Red
}

$NoBackgroundPolicy2 = Get-ItemProperty -Path  'Registry::HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Windows\Group Policy\{827D319E-6EAC-11D2-A4EA-00C04F79F83A}\'  -Name NoGPOListChanges -ErrorAction SilentlyContinue|Select-Object -ExpandProperty NoGPOListChanges
if ( $NoBackgroundPolicy2 -eq $null)
{
write-host " Configure security policy processing is not configured" -ForegroundColor Yellow
}
   elseif ( $NoBackgroundPolicy2  -eq  '0' )
{
write-host " Configure security policy processing is enabled" -ForegroundColor Green
}
  elseif ( $NoBackgroundPolicy2  -eq  '1' )
{
write-host " Configure security policy processing is disabled" -ForegroundColor Red
}
  else
{
write-host " Configure security policy processing is set to an unknown setting" -ForegroundColor Red
}

$DisableBkGndGroupPolicy = Get-ItemProperty -Path  'Registry::HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Policies\System\'  -Name DisableBkGndGroupPolicy -ErrorAction SilentlyContinue|Select-Object -ExpandProperty DisableBkGndGroupPolicy
if ( $DisableBkGndGroupPolicy -eq $null)
{
write-host " Turn off background refresh of Group Policy is not configured" -ForegroundColor Yellow
}
   elseif ( $DisableBkGndGroupPolicy  -eq  '0' )
{
write-host " Turn off background refresh of Group Policy is disabled" -ForegroundColor Green
}
  elseif ( $DisableBkGndGroupPolicy  -eq  '1' )
{
write-host " Turn off background refresh of Group Policy is enabled" -ForegroundColor Red
}
  else
{
write-host " Turn off background refresh of Group Policy is set to an unknown setting" -ForegroundColor Red
}

$DisableLGPOProcessing = Get-ItemProperty -Path  'Registry::HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Windows\System\'  -Name DisableLGPOProcessing -ErrorAction SilentlyContinue|Select-Object -ExpandProperty DisableLGPOProcessing
if ( $DisableLGPOProcessing -eq $null)
{
write-host " Turn off Local Group Policy Objects processing is not configured" -ForegroundColor Yellow
}
   elseif ( $DisableLGPOProcessing  -eq  '1' )
{
write-host " Turn off Local Group Policy Objects processing is enabled" -ForegroundColor Green
}
  elseif ( $DisableLGPOProcessing  -eq  '0' )
{
write-host " Turn off Local Group Policy Objects processing is disabled" -ForegroundColor Red
}
  else
{
write-host " Turn off Local Group Policy Objects processing is set to an unknown setting" -ForegroundColor Red
}

write-host "`r`n####################### HARD DRIVE ENCRYPTION #######################`r`n"

write-host "Unable to check Choose drive encryption method and cipher strength (Windows 10 [Version 1511] and later)"

$DisableExternalDMAUnderLock = Get-ItemProperty -Path  'Registry::HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\FVE\'  -Name DisableExternalDMAUnderLock -ErrorAction SilentlyContinue|Select-Object -ExpandProperty DisableExternalDMAUnderLock
if ( $DisableExternalDMAUnderLock -eq $null)
{
write-host " Disable new DMA devices when this computer is locked is not configured" -ForegroundColor Yellow
}
   elseif ( $DisableExternalDMAUnderLock  -eq  '1' )
{
write-host " Disable new DMA devices when this computer is locked is enabled" -ForegroundColor Green
}
  elseif ( $DisableExternalDMAUnderLock  -eq  '0' )
{
write-host " Disable new DMA devices when this computer is locked is disabled" -ForegroundColor Red
}
  else
{
write-host " Disable new DMA devices when this computer is locked is set to an unknown setting" -ForegroundColor Red
}

$MorBehavior = Get-ItemProperty -Path  'Registry::HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\FVE\'  -Name MorBehavior -ErrorAction SilentlyContinue|Select-Object -ExpandProperty MorBehavior
if ( $MorBehavior -eq $null)
{
write-host " Prevent memory overwrite on restart is not configured" -ForegroundColor Yellow
}
   elseif ( $MorBehavior  -eq  '0' )
{
write-host " Prevent memory overwrite on restart is disabled" -ForegroundColor Green
}
  elseif ( $MorBehavior  -eq  '1' )
{
write-host " Prevent memory overwrite on restart is enabled" -ForegroundColor Red
}
  else
{
write-host " Prevent memory overwrite on restart is set to an unknown setting" -ForegroundColor Red
}

write-host "Unable to check Choose how BitLocker-protected fixed drives can be recovered "
write-host "Unable to check Configure use of passwords for fixed data drives "

$FDVDenyWriteAccess = Get-ItemProperty -Path  'Registry::HKEY_LOCAL_MACHINE\System\CurrentControlSet\Policies\Microsoft\FVE\'  -Name FDVDenyWriteAccess -ErrorAction SilentlyContinue|Select-Object -ExpandProperty FDVDenyWriteAccess
if ( $FDVDenyWriteAccess -eq $null)
{
write-host " Deny write access to fixed drives not protected by BitLocker is not configured" -ForegroundColor Yellow
}
   elseif ( $FDVDenyWriteAccess  -eq  '1' )
{
write-host " Deny write access to fixed drives not protected by BitLocker is enabled" -ForegroundColor Green
}
  elseif ( $FDVDenyWriteAccess  -eq  '0' )
{
write-host " Deny write access to fixed drives not protected by BitLocker is disabled" -ForegroundColor Red
}
  else
{
write-host " Deny write access to fixed drives not protected by BitLocker is set to an unknown setting" -ForegroundColor Red
}

write-host "unable to check Enforce drive encryption type on fixed data drive "


$OSEnablePreBootPinExceptionOnDECapableDevice = Get-ItemProperty -Path  'Registry::HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\FVE\'  -Name OSEnablePreBootPinExceptionOnDECapableDevice -ErrorAction SilentlyContinue|Select-Object -ExpandProperty OSEnablePreBootPinExceptionOnDECapableDevice
if ( $OSEnablePreBootPinExceptionOnDECapableDevice -eq $null)
{
write-host " Allow devices compliant with InstantGo or HSTI to opt out of pre-boot PIN. is not configured" -ForegroundColor Yellow
}
   elseif ( $OSEnablePreBootPinExceptionOnDECapableDevice  -eq  '0' )
{
write-host " Allow devices compliant with InstantGo or HSTI to opt out of pre-boot PIN. is disabled" -ForegroundColor Green
}
  elseif ( $OSEnablePreBootPinExceptionOnDECapableDevice  -eq  '1' )
{
write-host " Allow devices compliant with InstantGo or HSTI to opt out of pre-boot PIN. is enabled" -ForegroundColor Red
}
  else
{
write-host " Allow devices compliant with InstantGo or HSTI to opt out of pre-boot PIN. is set to an unknown setting" -ForegroundColor Red
}

$UseEnhancedPin = Get-ItemProperty -Path  'Registry::HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\FVE\'  -Name UseEnhancedPin -ErrorAction SilentlyContinue|Select-Object -ExpandProperty UseEnhancedPin
if ( $UseEnhancedPin -eq $null)
{
write-host " Allow enhanced PINs for startup is not configured" -ForegroundColor Yellow
}
   elseif ( $UseEnhancedPin  -eq  '1' )
{
write-host " Allow enhanced PINs for startup is enabled" -ForegroundColor Green
}
  elseif ( $UseEnhancedPin  -eq  '0' )
{
write-host " Allow enhanced PINs for startup is disabled" -ForegroundColor Red
}
  else
{
write-host " Allow enhanced PINs for startup is set to an unknown setting" -ForegroundColor Red
}

$OSManageNKP = Get-ItemProperty -Path  'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\FVE\'  -Name OSManageNKP -ErrorAction SilentlyContinue|Select-Object -ExpandProperty OSManageNKP
if ( $OSManageNKP -eq $null)
{
write-host " Allow network unlock at startup is not configured" -ForegroundColor Yellow
}
   elseif ( $OSManageNKP  -eq  '1' )
{
write-host " Allow network unlock at startup is enabled" -ForegroundColor Green
}
  elseif ( $OSManageNKP  -eq  '0' )
{
write-host " Allow network unlock at startup is disabled" -ForegroundColor Red
}
  else
{
write-host " Allow network unlock at startup is set to an unknown setting" -ForegroundColor Red
}

$OSAllowSecureBootForIntegrity = Get-ItemProperty -Path  'Registry::HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\FVE\'  -Name OSAllowSecureBootForIntegrity -ErrorAction SilentlyContinue|Select-Object -ExpandProperty OSAllowSecureBootForIntegrity
if ( $OSAllowSecureBootForIntegrity -eq $null)
{
write-host " Allow Secure Boot for integrity validation is not configured" -ForegroundColor Yellow
}
   elseif ( $OSAllowSecureBootForIntegrity  -eq  '1' )
{
write-host " Allow Secure Boot for integrity validation is enabled" -ForegroundColor Green
}
  elseif ( $OSAllowSecureBootForIntegrity  -eq  '0' )
{
write-host " Allow Secure Boot for integrity validation is disabled" -ForegroundColor Red
}
  else
{
write-host " Allow Secure Boot for integrity validation is set to an unknown setting" -ForegroundColor Red
}

write-host "Unable to check Choose how BitLocker-protected operating system drives can be recovered"
write-host "Unable to check Configure minimum PIN length for startup "
write-host "Unable to check Configure use of passwords for operating system drives "

$DisallowStandardUserPINReset = Get-ItemProperty -Path  'Registry::HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\FVE\'  -Name DisallowStandardUserPINReset -ErrorAction SilentlyContinue|Select-Object -ExpandProperty DisallowStandardUserPINReset
if ( $DisallowStandardUserPINReset -eq $null)
{
write-host " Disallow standard users from changing the PIN or password is not configured" -ForegroundColor Yellow
}
   elseif ( $DisallowStandardUserPINReset  -eq  '0' )
{
write-host " Disallow standard users from changing the PIN or password is disabled" -ForegroundColor Green
}
  elseif ( $DisallowStandardUserPINReset  -eq  '1' )
{
write-host " Disallow standard users from changing the PIN or password is enabled" -ForegroundColor Red
}
  else
{
write-host " Disallow standard users from changing the PIN or password is set to an unknown setting" -ForegroundColor Red
}

write-host "Enforce drive encryption type on operating system drive unable to check "
write-host "Unable to check Require additional authentication at startup"

$TPMAutoReseal = Get-ItemProperty -Path  'Registry::HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\FVE\'  -Name TPMAutoReseal -ErrorAction SilentlyContinue|Select-Object -ExpandProperty TPMAutoReseal
if ( $TPMAutoReseal -eq $null)
{
write-host " Reset platform validation data after BitLocker recovery is not configured" -ForegroundColor Yellow
}
   elseif ( $TPMAutoReseal  -eq  '1' )
{
write-host " Reset platform validation data after BitLocker recovery is enabled" -ForegroundColor Green
}
  elseif ( $TPMAutoReseal  -eq  '0' )
{
write-host " Reset platform validation data after BitLocker recovery is disabled" -ForegroundColor Red
}
  else
{
write-host " Reset platform validation data after BitLocker recovery is set to an unknown setting" -ForegroundColor Red
}

write-host "Choose how BitLocker-protected removable drives can be recovered"
write-host "Configure use of passwords for removable data drives "
write-host "Control use of BitLocker on removable drives"

$RDVDenyWriteAccess = Get-ItemProperty -Path  'Registry::HKEY_LOCAL_MACHINE\System\CurrentControlSet\Policies\Microsoft\FVE\'  -Name RDVDenyWriteAccess -ErrorAction SilentlyContinue|Select-Object -ExpandProperty RDVDenyWriteAccess
if ( $RDVDenyWriteAccess -eq $null)
{
write-host " Deny write access to removable drives not protected by BitLocker is not configured" -ForegroundColor Yellow
}
   elseif ( $RDVDenyWriteAccess  -eq  '1' )
{
write-host " Deny write access to removable drives not protected by BitLocker is enabled" -ForegroundColor Green
}
  elseif ( $RDVDenyWriteAccess  -eq  '0' )
{
write-host " Deny write access to removable drives not protected by BitLocker is disabled" -ForegroundColor Red
}
  else
{
write-host " Deny write access to removable drives not protected by BitLocker is set to an unknown setting" -ForegroundColor Red
}

$RDVEncryptionType = Get-ItemProperty -Path  'Registry::HKLM\SOFTWARE\Policies\Microsoft\FVE\RDVEncryptionType HKLM\SOFTWARE\Policies\Microsoft\FVE\'  -Name  RDVEncryptionType -ErrorAction SilentlyContinue|Select-Object -ExpandProperty RDVEncryptionType
if ( $RDVEncryptionType -eq $null)
{
write-host "Enforce drive encryption type on removable data drive  is not configured" -ForegroundColor Yellow
}
   elseif ( $RDVEncryptionType  -eq  '1' )
{
write-host "Enforce drive encryption type on removable data drive  is enabled with full encryption" -ForegroundColor Green
}
  elseif ( $RDVEncryptionType  -eq  '2' )
{
write-host "Enforce drive encryption type on removable data drive  is enabled with Used Space Only encryption" -ForegroundColor Red
}
  else
{
write-host "Enforce drive encryption type on removable data drive  is set to Allow user to choose" -ForegroundColor Red
}

write-host "`r`n####################### INSTALLING APPLICATIONS #######################`r`n"

$EnableSmartScreen = Get-ItemProperty -Path  'Registry::HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Windows\System\'  -Name EnableSmartScreen -ErrorAction SilentlyContinue|Select-Object -ExpandProperty EnableSmartScreen
if ( $EnableSmartScreen -eq $null)
{
write-host " Configure Windows Defender SmartScreen is not configured" -ForegroundColor Yellow
}
   elseif ( $EnableSmartScreen  -eq  '1' )
{
write-host " Configure Windows Defender SmartScreen is enabled" -ForegroundColor Green

$ShellSmartScreenLevel = Get-ItemProperty -Path  'Registry::HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Windows\System\'  -Name ShellSmartScreenLevel -ErrorAction SilentlyContinue|Select-Object -ExpandProperty ShellSmartScreenLevel
if ( $ShellSmartScreenLevel -eq $null)
{
write-host "SmartScreen is not configured" -ForegroundColor Yellow
}
   elseif ( $ShellSmartScreenLevel  -eq  'Block' )
{
write-host "Windows Defender SmartScreen is set to Warn and Prevent Bypass" -ForegroundColor Green
}
  elseif ( $ShellSmartScreenLevel -eq  'Warn' )
{
write-host "SmartScreen is set to Warn" -ForegroundColor Red
}
  else
{
write-host "SmartScreen is set to an unknown setting" -ForegroundColor Red
}


}
  elseif ( $EnableSmartScreen  -eq  '0' )
{
write-host " Configure Windows Defender SmartScreen is disabled" -ForegroundColor Red
}
  else
{
write-host " Configure Windows Defender SmartScreen is set to an unknown setting" -ForegroundColor Red
}

$EnableUserControl = Get-ItemProperty -Path  'Registry::HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Windows\Installer\'  -Name EnableUserControl -ErrorAction SilentlyContinue|Select-Object -ExpandProperty EnableUserControl
if ( $EnableUserControl -eq $null)
{
write-host " Allow user control over installs is not configured" -ForegroundColor Yellow
}
   elseif ( $EnableUserControl  -eq  '0' )
{
write-host " Allow user control over installs is disabled" -ForegroundColor Green
}
  elseif ( $EnableUserControl  -eq  '1' )
{
write-host " Allow user control over installs is enabled" -ForegroundColor Red
}
  else
{
write-host " Allow user control over installs is set to an unknown setting" -ForegroundColor Red
}

$AlwaysInstallElevated = Get-ItemProperty -Path  'Registry::HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Windows\Installer\'  -Name AlwaysInstallElevated -ErrorAction SilentlyContinue|Select-Object -ExpandProperty AlwaysInstallElevated
if ( $AlwaysInstallElevated -eq $null)
{
write-host " Always install with elevated privileges is not configured in local machine policy" -ForegroundColor Yellow
}
   elseif ( $AlwaysInstallElevated  -eq  '0' )
{
write-host " Always install with elevated privileges is disabled in local machine policy" -ForegroundColor Green
}
  elseif ( $AlwaysInstallElevated  -eq  '1' )
{
write-host " Always install with elevated privileges is enabled in local machine policy" -ForegroundColor Red
}
  else
{
write-host " Always install with elevated privileges is set to an unknown setting in local machine policy" -ForegroundColor Red
}


$AlwaysInstallElevated1 = Get-ItemProperty -Path  'Registry::HKEY_CURRENT_USER\Software\Policies\Microsoft\Windows\Installer\'  -Name AlwaysInstallElevated -ErrorAction SilentlyContinue|Select-Object -ExpandProperty AlwaysInstallElevated
if ( $AlwaysInstallElevated1 -eq $null)
{
write-host " Always install with elevated privileges is not configured in user policy" -ForegroundColor Yellow
}
   elseif ( $AlwaysInstallElevated1  -eq  '0' )
{
write-host " Always install with elevated privileges is disabled in user policy" -ForegroundColor Green
}
  elseif ( $AlwaysInstallElevated1  -eq  '1' )
{
write-host " Always install with elevated privileges is enabled in user policy" -ForegroundColor Red
}
  else
{
write-host " Always install with elevated privileges is set to an unknown setting in user policy" -ForegroundColor Red
}

write-host "`r`n####################### INTERNET PRINTING #######################`r`n"

$DisableWebPnPDownload = Get-ItemProperty -Path  'Registry::HKEY_CURRENT_USER\Software\Policies\Microsoft\Windows NT\Printers\'  -Name DisableWebPnPDownload -ErrorAction SilentlyContinue|Select-Object -ExpandProperty DisableWebPnPDownload
if ( $DisableWebPnPDownload -eq $null)
{
write-host " Turn off downloading of print drivers over HTTP is not configured" -ForegroundColor Yellow
}
   elseif ( $DisableWebPnPDownload  -eq  '1' )
{
write-host " Turn off downloading of print drivers over HTTP is enabled" -ForegroundColor Green
}
  elseif ( $DisableWebPnPDownload  -eq  '0' )
{
write-host " Turn off downloading of print drivers over HTTP is disabled" -ForegroundColor Red
}
  else
{
write-host " Turn off downloading of print drivers over HTTP is set to an unknown setting" -ForegroundColor Red
}

$DisableHTTPPrinting = Get-ItemProperty -Path  'Registry::HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Windows NT\Printers\'  -Name DisableHTTPPrinting -ErrorAction SilentlyContinue|Select-Object -ExpandProperty DisableHTTPPrinting
if ( $DisableHTTPPrinting -eq $null)
{
write-host " Turn off printing over HTTP is not configured" -ForegroundColor Yellow
}
   elseif ( $DisableHTTPPrinting  -eq  '1' )
{
write-host " Turn off printing over HTTP is enabled" -ForegroundColor Green
}
  elseif ( $DisableHTTPPrinting  -eq  '0' )
{
write-host " Turn off printing over HTTP is disabled" -ForegroundColor Red
}
  else
{
write-host " Turn off printing over HTTP is set to an unknown setting" -ForegroundColor Red
}

write-host "`r`n####################### LEGACY AND RUN ONCE LISTS #######################`r`n"

$UN6ehVpmakAXClE = Get-ItemProperty -Path  'Registry::HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\'  -Name DisableCurrentUserRun -ErrorAction SilentlyContinue|Select-Object -ExpandProperty DisableCurrentUserRun
if ( $UN6ehVpmakAXClE -eq $null)
{
write-host " Do not process the legacy run list is not configured" -ForegroundColor Yellow
}
   elseif ( $UN6ehVpmakAXClE  -eq  '1' )
{
write-host " Do not process the legacy run list is enabled" -ForegroundColor Green
}
  elseif ( $UN6ehVpmakAXClE  -eq  '0' )
{
write-host " Do not process the legacy run list is disabled" -ForegroundColor Red
}
  else
{
write-host " Do not process the legacy run list is set to an unknown setting" -ForegroundColor Red
}

$keAWhyT9w1aMjVE = Get-ItemProperty -Path  'Registry::HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\'  -Name DisableLocalMachineRunOnce -ErrorAction SilentlyContinue|Select-Object -ExpandProperty DisableLocalMachineRunOnce
if ( $keAWhyT9w1aMjVE -eq $null)
{
write-host " Do not process the run once list is not configured" -ForegroundColor Yellow
}
   elseif ( $keAWhyT9w1aMjVE  -eq  '1' )
{
write-host " Do not process the run once list is enabled" -ForegroundColor Green
}
  elseif ( $keAWhyT9w1aMjVE  -eq  '0' )
{
write-host " Do not process the run once list is disabled" -ForegroundColor Red
}
  else
{
write-host " Do not process the run once list is set to an unknown setting" -ForegroundColor Red
}

foreach($_ in 1..50)
{
    $runkeys = Get-ItemProperty -Path "Registry::HKLM\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\Run\" -Name $_ -ErrorAction SilentlyContinue|Select-Object -ExpandProperty $_
    If ($runkeys -ne $null)
    {
        write-host "The following run key is set: $runkeys" -ForegroundColor Red

    }
}
foreach($_ in 1..50)
{
    $runkeys2 = Get-ItemProperty -Path "Registry::HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run\" -Name $_ -ErrorAction SilentlyContinue|Select-Object -ExpandProperty $_
    If ($runkeys2 -ne $null)
    {
        write-host "The following run key is set: $runkeys2" -ForegroundColor Red

    }
}
If ($runkeys -eq $null -and $runkeys2 -eq $runkeys2)
{

    write-host "Run These Programs At User Logon is disabled, no run keys are set" -ForegroundColor Green
}



write-host "`r`n####################### MICROSOFT ACCOUNTS #######################`r`n"

$7u6bAiHSjEa1L9F = Get-ItemProperty -Path  'Registry::HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\MicrosoftAccount\'  -Name DisableUserAuth -ErrorAction SilentlyContinue|Select-Object -ExpandProperty DisableUserAuth
if ( $7u6bAiHSjEa1L9F -eq $null)
{
write-host " Block all consumer Microsoft account user authentication is not configured" -ForegroundColor Yellow
}
   elseif ( $7u6bAiHSjEa1L9F  -eq  '1' )
{
write-host " Block all consumer Microsoft account user authentication is enabled" -ForegroundColor Green
}
  elseif ( $7u6bAiHSjEa1L9F  -eq  '0' )
{
write-host " Block all consumer Microsoft account user authentication is disabled" -ForegroundColor Red
}
  else
{
write-host " Block all consumer Microsoft account user authentication is set to an unknown setting" -ForegroundColor Red
}

$q69ocA0RwE3KT7D = Get-ItemProperty -Path  'Registry::HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Windows\OneDrive\'  -Name DisableFileSyncNGSC -ErrorAction SilentlyContinue|Select-Object -ExpandProperty DisableFileSyncNGSC
if ( $q69ocA0RwE3KT7D -eq $null)
{
write-host " Prevent the usage of OneDrive for file storage is not configured" -ForegroundColor Yellow
}
   elseif ( $q69ocA0RwE3KT7D  -eq  '1' )
{
write-host " Prevent the usage of OneDrive for file storage is enabled" -ForegroundColor Green
}
  elseif ( $q69ocA0RwE3KT7D  -eq  '0' )
{
write-host " Prevent the usage of OneDrive for file storage is disabled" -ForegroundColor Red
}
  else
{
write-host " Prevent the usage of OneDrive for file storage is set to an unknown setting" -ForegroundColor Red
}

write-host "Unable to check Accounts: Block Microsoft accounts"

write-host "`r`n####################### MSS SETTINGS #######################`r`n"

$fYg2RApMS8B3z4o = Get-ItemProperty -Path  'Registry::HKEY_LOCAL_MACHINE\System\CurrentControlSet\Services\Tcpip\Parameters\'  -Name DisableIPSourceRouting -ErrorAction SilentlyContinue|Select-Object -ExpandProperty DisableIPSourceRouting
if ( $fYg2RApMS8B3z4o -eq $null)
{
write-host " MSS: (DisableIPSourceRouting) IP source routing protection level (protects against packet spoofing) is not configured" -ForegroundColor Yellow
}
   elseif ( $fYg2RApMS8B3z4o  -eq  '2' )
{
write-host " MSS: (DisableIPSourceRouting) IP source routing protection level (protects against packet spoofing) is set to Highest protection, source routing is completely disabled " -ForegroundColor Green
}
  elseif ( $fYg2RApMS8B3z4o  -eq  '0' -or $fYg2RApMS8B3z4o  -eq  '1' )
{
write-host " MSS: (DisableIPSourceRouting) IP source routing protection level (protects against packet spoofing) is configured incorrectly" -ForegroundColor Red
}
  else
{
write-host " MSS: (DisableIPSourceRouting) IP source routing protection level (protects against packet spoofing) is set to an unknown setting" -ForegroundColor Red
}

$Yd9tFn6Q4UEIR8a = Get-ItemProperty -Path  'Registry::HKEY_LOCAL_MACHINE\System\CurrentControlSet\Services\Tcpip6\Parameters\'  -Name DisableIPSourceRouting -ErrorAction SilentlyContinue|Select-Object -ExpandProperty DisableIPSourceRouting
if ( $Yd9tFn6Q4UEIR8a -eq $null)
{
write-host " MSS: (DisableIPSourceRouting IPv6) IP source routing protection level (protects against packet spoofing) is not configured" -ForegroundColor Yellow
}
   elseif ( $Yd9tFn6Q4UEIR8a  -eq  '2' )
{
write-host " MSS: (DisableIPSourceRouting IPv6) IP source routing protection level (protects against packet spoofing) is set to Highest protection, source routing is completely disabled " -ForegroundColor Green
}
  elseif ( $Yd9tFn6Q4UEIR8a  -eq  '0' -or $Yd9tFn6Q4UEIR8a  -eq  '1' )
{
write-host " MSS: (DisableIPSourceRouting IPv6) IP source routing protection level (protects against packet spoofing) is configured incorrectly" -ForegroundColor Red
}
  else
{
write-host " MSS: (DisableIPSourceRouting IPv6) IP source routing protection level (protects against packet spoofing) is set to an unknown setting" -ForegroundColor Red
}

$ZqEKJnRyWQruTsH = Get-ItemProperty -Path  'Registry::HKEY_LOCAL_MACHINE\System\CurrentControlSet\Services\Tcpip\Parameters\'  -Name EnableICMPRedirect -ErrorAction SilentlyContinue|Select-Object -ExpandProperty EnableICMPRedirect
if ( $ZqEKJnRyWQruTsH -eq $null)
{
write-host " MSS: (EnableICMPRedirect) Allow ICMP redirects to override OSPF generated routes is not configured" -ForegroundColor Yellow
}
   elseif ( $ZqEKJnRyWQruTsH  -eq  '0' )
{
write-host " MSS: (EnableICMPRedirect) Allow ICMP redirects to override OSPF generated routes is disabled" -ForegroundColor Green
}
  elseif ( $ZqEKJnRyWQruTsH  -eq  '1' )
{
write-host " MSS: (EnableICMPRedirect) Allow ICMP redirects to override OSPF generated routes is enabled" -ForegroundColor Red
}
  else
{
write-host " MSS: (EnableICMPRedirect) Allow ICMP redirects to override OSPF generated routes is set to an unknown setting" -ForegroundColor Red
}

$JKYyPoEx63dhjZr = Get-ItemProperty -Path  'Registry::HKEY_LOCAL_MACHINE\System\CurrentControlSet\Services\Netbt\Parameters\'  -Name NoNameReleaseOnDemand -ErrorAction SilentlyContinue|Select-Object -ExpandProperty NoNameReleaseOnDemand
if ( $JKYyPoEx63dhjZr -eq $null)
{
write-host " MSS: (NoNameReleaseOnDemand) Allow the computer to ignore NetBIOS name release requests except from WINS servers is not configured" -ForegroundColor Yellow
}
   elseif ( $JKYyPoEx63dhjZr  -eq  '1' )
{
write-host " MSS: (NoNameReleaseOnDemand) Allow the computer to ignore NetBIOS name release requests except from WINS servers is enabled" -ForegroundColor Green
}
  elseif ( $JKYyPoEx63dhjZr  -eq  '0' )
{
write-host " MSS: (NoNameReleaseOnDemand) Allow the computer to ignore NetBIOS name release requests except from WINS servers is disabled" -ForegroundColor Red
}
  else
{
write-host " MSS: (NoNameReleaseOnDemand) Allow the computer to ignore NetBIOS name release requests except from WINS servers is set to an unknown setting" -ForegroundColor Red
}

write-host "`r`n####################### NETBIOS OVER TCP/IP #######################`r`n"

$servicenetbt = get-service netbt

if ($servicenetbt.Status -eq 'Running')
{
    write-host "NetBIOS Over TCP/IP service is running, NetBIOS over TCP/IP is likely enabled" -ForegroundColor Red
}
elseif ($servicenetbt.Status -eq 'Disabled')
{
    write-host "NetBIOS Over TCP/IP service is disabled, NetBIOS over TCP/IP is not running" -ForegroundColor Green
}
elseif ($servicenetbt.Status -eq 'Stopped')
{
    write-host "NetBIOS Over TCP/IP service is stopped but not disabled" -ForegroundColor Red
}
else
{
    write-host "NetBIOS Over TCP/IP service status was unable to be determined" -ForegroundColor Yellow
}



write-host "`r`n####################### NETWORK AUTHENTICATION #######################`r`n"

write-host "Unable to check Network security: Configure encryption types allowed for Kerberos"

write-host "Unable to check Network security: LAN Manager authentication level"

write-host "Unable to check Network security: Minimum session security for NTLM SSP based (including secure RPC) clients"

write-host "`r`n####################### NOLM HASH POLICY #######################`r`n"

write-host "Unable to check Network security: Do not store LAN Manager hash value on next password change"



write-host "`r`n####################### OPERATING SYSTEM FUNCTIONALITY #######################`r`n"


write-host "`r`n####################### POWER MANAGEMENT #######################`r`n"

$p86A1e2VhcGQKas = Get-ItemProperty -Path  'Registry::HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Power\PowerSettings\abfc2519-3608-4c2a-94ea-171b0ed546ab\'  -Name DCSettingIndex -ErrorAction SilentlyContinue|Select-Object -ExpandProperty DCSettingIndex
if ( $p86A1e2VhcGQKas -eq $null)
{
write-host " Allow standby states (S1-S3) when sleeping (on battery) is not configured" -ForegroundColor Yellow
}
   elseif ( $p86A1e2VhcGQKas  -eq  '0' )
{
write-host " Allow standby states (S1-S3) when sleeping (on battery) is disabled" -ForegroundColor Green
}
  elseif ( $p86A1e2VhcGQKas  -eq  '1' )
{
write-host " Allow standby states (S1-S3) when sleeping (on battery) is enabled" -ForegroundColor Red
}
  else
{
write-host " Allow standby states (S1-S3) when sleeping (on battery) is set to an unknown setting" -ForegroundColor Red
}

$w4PO3v6EaroqgUu = Get-ItemProperty -Path  'Registry::HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Power\PowerSettings\abfc2519-3608-4c2a-94ea-171b0ed546ab\'  -Name ACSettingIndex -ErrorAction SilentlyContinue|Select-Object -ExpandProperty ACSettingIndex
if ( $w4PO3v6EaroqgUu -eq $null)
{
write-host " Allow standby states (S1-S3) when sleeping (plugged in) is not configured" -ForegroundColor Yellow
}
   elseif ( $w4PO3v6EaroqgUu  -eq  '0' )
{
write-host " Allow standby states (S1-S3) when sleeping (plugged in) is disabled" -ForegroundColor Green
}
  elseif ( $w4PO3v6EaroqgUu  -eq  '1' )
{
write-host " Allow standby states (S1-S3) when sleeping (plugged in) is enabled" -ForegroundColor Red
}
  else
{
write-host " Allow standby states (S1-S3) when sleeping (plugged in) is set to an unknown setting" -ForegroundColor Red
}


$b9ePm1KdQUNf7tu = Get-ItemProperty -Path  'Registry::HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Power\PowerSettings\0e796bdb-100d-47d6-a2d5-f7d2daa51f51\'  -Name DCSettingIndex -ErrorAction SilentlyContinue|Select-Object -ExpandProperty DCSettingIndex
if ( $b9ePm1KdQUNf7tu -eq $null)
{
write-host " Require a password when a computer wakes (on battery) is not configured" -ForegroundColor Yellow
}
   elseif ( $b9ePm1KdQUNf7tu  -eq  '1' )
{
write-host " Require a password when a computer wakes (on battery) is enabled" -ForegroundColor Green
}
  elseif ( $b9ePm1KdQUNf7tu  -eq  '0' )
{
write-host " Require a password when a computer wakes (on battery) is disabled" -ForegroundColor Red
}
  else
{
write-host " Require a password when a computer wakes (on battery) is set to an unknown setting" -ForegroundColor Red
}

$GmlQKPgtw7i91Fx = Get-ItemProperty -Path  'Registry::HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Power\PowerSettings\0e796bdb-100d-47d6-a2d5-f7d2daa51f51\'  -Name ACSettingIndex -ErrorAction SilentlyContinue|Select-Object -ExpandProperty ACSettingIndex
if ( $GmlQKPgtw7i91Fx -eq $null)
{
write-host " Require a password when a computer wakes (plugged in) is not configured" -ForegroundColor Yellow
}
   elseif ( $GmlQKPgtw7i91Fx  -eq  '1' )
{
write-host " Require a password when a computer wakes (plugged in) is enabled" -ForegroundColor Green
}
  elseif ( $GmlQKPgtw7i91Fx  -eq  '0' )
{
write-host " Require a password when a computer wakes (plugged in) is disabled" -ForegroundColor Red
}
  else
{
write-host " Require a password when a computer wakes (plugged in) is set to an unknown setting" -ForegroundColor Red
}

$IDxPlKksMyvH3Xd = Get-ItemProperty -Path  'Registry::HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Power\PowerSettings\9D7815A6-7EE4-497E-8888-515A05F02364\'  -Name DCSettingIndex -ErrorAction SilentlyContinue|Select-Object -ExpandProperty DCSettingIndex
if ( $IDxPlKksMyvH3Xd -eq $null)
{
write-host " Specify the system hibernate timeout (on battery) is not configured" -ForegroundColor Yellow
}
   elseif ( $IDxPlKksMyvH3Xd  -eq  '0' )
{
write-host " Specify the system hibernate timeout (on battery) is enabled and set to 0 seconds" -ForegroundColor Green
}
   else
{
write-host " Specify the system hibernate timeout (on battery) is set to an unknown setting" -ForegroundColor Red
}

$wqSbpksEI7retQd = Get-ItemProperty -Path  'Registry::HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Power\PowerSettings\9D7815A6-7EE4-497E-8888-515A05F02364\'  -Name ACSettingIndex -ErrorAction SilentlyContinue|Select-Object -ExpandProperty ACSettingIndex
if ( $wqSbpksEI7retQd -eq $null)
{
write-host " Specify the system hibernate timeout (plugged in) is not configured" -ForegroundColor Yellow
}
   elseif ( $wqSbpksEI7retQd  -eq  '0' )
{
write-host " Specify the system hibernate timeout (plugged in) is enabled and set to 0 seconds" -ForegroundColor Green
}
 
  else
{
write-host " Specify the system hibernate timeout (plugged in) is set to an unknown setting" -ForegroundColor Red
}

$7QZf3kP5WXARGrt = Get-ItemProperty -Path  'Registry::HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Power\PowerSettings\29F6C1DB-86DA-48C5-9FDB-F2B67B1F44DA\'  -Name DCSettingIndex -ErrorAction SilentlyContinue|Select-Object -ExpandProperty DCSettingIndex
if ( $7QZf3kP5WXARGrt -eq $null)
{
write-host " Specify the system sleep timeout (on battery) is not configured" -ForegroundColor Yellow
}
   elseif ( $7QZf3kP5WXARGrt  -eq  '0' )
{
write-host " Specify the system sleep timeout (on battery) is enabled and set to 0 seconds" -ForegroundColor Green
}
 
  else
{
write-host " Specify the system sleep timeout (on battery) is set to an unknown setting" -ForegroundColor Red
}
$r5kh6s8qULHTAfD = Get-ItemProperty -Path  'Registry::HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Power\PowerSettings\29F6C1DB-86DA-48C5-9FDB-F2B67B1F44DA\'  -Name ACSettingIndex -ErrorAction SilentlyContinue|Select-Object -ExpandProperty ACSettingIndex
if ( $r5kh6s8qULHTAfD -eq $null)
{
write-host " Specify the system sleep timeout (plugged in) is not configured" -ForegroundColor Yellow
}
   elseif ( $r5kh6s8qULHTAfD  -eq  '0' )
{
write-host " Specify the system sleep timeout (plugged in) is enabled and set to 0 seconds" -ForegroundColor Green
}
  else
{
write-host " Specify the system sleep timeout (plugged in) is set to an unknown setting" -ForegroundColor Red
}

$BMbAhC2V4J0SpLD = Get-ItemProperty -Path  'Registry::HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Power\PowerSettings\7bc4a2f9-d8fc-4469-b07b-33eb785aaca0\'  -Name DCSettingIndex -ErrorAction SilentlyContinue|Select-Object -ExpandProperty DCSettingIndex
if ( $BMbAhC2V4J0SpLD -eq $null)
{
write-host " Specify the unattended sleep timeout (on battery) is not configured" -ForegroundColor Yellow
}
   elseif ( $BMbAhC2V4J0SpLD  -eq  '0' )
{
write-host " Specify the unattended sleep timeout (on battery) is enabled and set to 0 seconds" -ForegroundColor Green
}
  else
{
write-host " Specify the unattended sleep timeout (on battery) is set to an unknown setting" -ForegroundColor Red
}

$4lhpjTxyb92RsKJ = Get-ItemProperty -Path  'Registry::HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Power\PowerSettings\7bc4a2f9-d8fc-4469-b07b-33eb785aaca0\'  -Name ACSettingIndex -ErrorAction SilentlyContinue|Select-Object -ExpandProperty ACSettingIndex
if ( $4lhpjTxyb92RsKJ -eq $null)
{
write-host " Specify the unattended sleep timeout (plugged in) is not configured" -ForegroundColor Yellow
}
   elseif ( $4lhpjTxyb92RsKJ  -eq  '0' )
{
write-host " Specify the unattended sleep timeout (plugged in) is enabled" -ForegroundColor Green
}
    else
{
write-host " Specify the unattended sleep timeout (plugged in) is set to an unknown setting" -ForegroundColor Red
}

$bOEF2189wg3Dhzq = Get-ItemProperty -Path  'Registry::HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Power\PowerSettings\94ac6d29-73ce-41a6-809f-6363ba21b47e\'  -Name DCSettingIndex -ErrorAction SilentlyContinue|Select-Object -ExpandProperty DCSettingIndex
if ( $bOEF2189wg3Dhzq -eq $null)
{
write-host " Turn off hybrid sleep (on battery) is not configured" -ForegroundColor Yellow
}
   elseif ( $bOEF2189wg3Dhzq  -eq  '0' )
{
write-host " Turn off hybrid sleep (on battery) is enabled" -ForegroundColor Green
}
  elseif ( $bOEF2189wg3Dhzq  -eq  '1' )
{
write-host " Turn off hybrid sleep (on battery) is disabled" -ForegroundColor Red
}
  else
{
write-host " Turn off hybrid sleep (on battery) is set to an unknown setting" -ForegroundColor Red
}

$xcyp78VGK9RYUs0 = Get-ItemProperty -Path  'Registry::HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Power\PowerSettings\94ac6d29-73ce-41a6-809f-6363ba21b47e\'  -Name ACSettingIndex -ErrorAction SilentlyContinue|Select-Object -ExpandProperty ACSettingIndex
if ( $xcyp78VGK9RYUs0 -eq $null)
{
write-host " Turn off hybrid sleep (plugged in) is not configured" -ForegroundColor Yellow
}
   elseif ( $xcyp78VGK9RYUs0  -eq  '0' )
{
write-host " Turn off hybrid sleep (plugged in) is enabled" -ForegroundColor Green
}
  elseif ( $xcyp78VGK9RYUs0  -eq  '1' )
{
write-host " Turn off hybrid sleep (plugged in) is disabled" -ForegroundColor Red
}
  else
{
write-host " Turn off hybrid sleep (plugged in) is set to an unknown setting" -ForegroundColor Red
}

$LXGISnrDvyTAdjE = Get-ItemProperty -Path  'Registry::HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Windows\Explorer\'  -Name ShowHibernateOption -ErrorAction SilentlyContinue|Select-Object -ExpandProperty ShowHibernateOption
if ( $LXGISnrDvyTAdjE -eq $null)
{
write-host " Show hibernate in the power options menu is not configured" -ForegroundColor Yellow
}
   elseif ( $LXGISnrDvyTAdjE  -eq  '0' )
{
write-host " Show hibernate in the power options menu is disabled" -ForegroundColor Green
}
  elseif ( $LXGISnrDvyTAdjE  -eq  '1' )
{
write-host " Show hibernate in the power options menu is enabled" -ForegroundColor Red
}
  else
{
write-host " Show hibernate in the power options menu is set to an unknown setting" -ForegroundColor Red
}

$JwmcB8OLGS0loNP = Get-ItemProperty -Path  'Registry::HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Windows\Explorer\'  -Name ShowSleepOption -ErrorAction SilentlyContinue|Select-Object -ExpandProperty ShowSleepOption
if ( $JwmcB8OLGS0loNP -eq $null)
{
write-host " Show sleep in the power options menu is not configured" -ForegroundColor Yellow
}
   elseif ( $JwmcB8OLGS0loNP  -eq  '0' )
{
write-host " Show sleep in the power options menu is disabled" -ForegroundColor Green
}
  elseif ( $JwmcB8OLGS0loNP  -eq  '1' )
{
write-host " Show sleep in the power options menu is enabled" -ForegroundColor Red
}
  else
{
write-host " Show sleep in the power options menu is set to an unknown setting" -ForegroundColor Red
}

write-host "`r`n####################### POWERSHELL #######################`r`n"

$LMCJtZgR8FhxmbGke = Get-ItemProperty -Path  'Registry::HKLM\Software\Policies\Microsoft\Windows\PowerShell\ScriptBlockLogging\' -Name EnableScriptBlockLogging -ErrorAction SilentlyContinue|Select-Object -ExpandProperty EnableScriptBlockLogging
$UPCJtZgR8FhxmbGke = Get-ItemProperty -Path  'Registry::HKCU\Software\Policies\Microsoft\Windows\PowerShell\ScriptBlockLogging\' -Name EnableScriptBlockLogging -ErrorAction SilentlyContinue|Select-Object -ExpandProperty EnableScriptBlockLogging
if ( $LMCJtZgR8FhxmbGke -eq $null -and  $UPCJtZgR8FhxmbGke -eq $null)
{
write-host " Turn on PowerShell Script Block Logging is not configured" -ForegroundColor Yellow
}
if ( $LMCJtZgR8FhxmbGke  -eq '1' )
{
write-host " Turn on PowerShell Script Block Logging is enabled in Local Machine GP" -ForegroundColor Green
}
if ( $LMCJtZgR8FhxmbGke  -eq '0' )
{
write-host " Turn on PowerShell Script Block Logging is disabled in Local Machine GP" -ForegroundColor Red
}
if ( $UPCJtZgR8FhxmbGke  -eq  '1' )
{
write-host " Turn on PowerShell Script Block Logging is enabled in User GP" -ForegroundColor Green
}
if ( $UPCJtZgR8FhxmbGke  -eq  '0' )
{
write-host " Turn on PowerShell Script Block Logging is disabled in User GP" -ForegroundColor Red
}

$LMCJtZgR8FhxmbGked = Get-ItemProperty -Path  'Registry::HKLM\Software\Policies\Microsoft\Windows\PowerShell\ScriptBlockLogging\' -Name EnableScriptBlockInvocationLogging -ErrorAction SilentlyContinue|Select-Object -ExpandProperty EnableScriptBlockInvocationLogging
$UPCJtZgR8FhxmbGked = Get-ItemProperty -Path  'Registry::HKCU\Software\Policies\Microsoft\Windows\PowerShell\ScriptBlockLogging\' -Name EnableScriptBlockInvocationLogging -ErrorAction SilentlyContinue|Select-Object -ExpandProperty EnableScriptBlockInvocationLogging
if ( $LMCJtZgR8FhxmbGked -eq $null -and  $UPCJtZgR8FhxmbGked -eq $null)
{
write-host " Turn on PowerShell Script Block Invocation Logging is not configured" -ForegroundColor Yellow
}
if ( $LMCJtZgR8FhxmbGked  -eq '1' )
{
write-host " Turn on PowerShell Script Block Invocation Logging is enabled in Local Machine GP" -ForegroundColor Green
}
if ( $LMCJtZgR8FhxmbGked  -eq '0' )
{
write-host " Turn on PowerShell Script Block Invocation Logging is disabled in Local Machine GP" -ForegroundColor Red
}
if ( $UPCJtZgR8FhxmbGked  -eq  '1' )
{
write-host " Turn on PowerShell Script Block Invocation Logging is enabled in User GP" -ForegroundColor Green
}
if ( $UPCJtZgR8FhxmbGked  -eq  '0' )
{
write-host " Turn on PowerShell Script Block Invocation Logging is disabled in User GP" -ForegroundColor Red
}

$LMbMRxhAX7jTCJI2S = Get-ItemProperty -Path  'Registry::HKLM\Software\Policies\Microsoft\Windows\PowerShell\' -Name EnableScripts -ErrorAction SilentlyContinue|Select-Object -ExpandProperty EnableScripts
$UPbMRxhAX7jTCJI2S = Get-ItemProperty -Path  'Registry::HKCU\Software\Policies\Microsoft\Windows\PowerShell\' -Name EnableScripts -ErrorAction SilentlyContinue|Select-Object -ExpandProperty EnableScripts
if ( $LMbMRxhAX7jTCJI2S -eq $null -and  $UPbMRxhAX7jTCJI2S -eq $null)
{
write-host " Turn on Script Execution is not configured" -ForegroundColor Yellow
}
if ( $LMbMRxhAX7jTCJI2S  -eq '1' )
{
write-host " Turn on Script Execution is enabled in Local Machine GP" -ForegroundColor Green
}
if ( $LMbMRxhAX7jTCJI2S  -eq '0' )
{
write-host " Turn on Script Execution is disabled in Local Machine GP" -ForegroundColor Red
}
if ( $UPbMRxhAX7jTCJI2S  -eq  '1' )
{
write-host " Turn on Script Execution is enabled in User GP" -ForegroundColor Green
}
if ( $UPbMRxhAX7jTCJI2S  -eq  '0' )
{
write-host " Turn on Script Execution is disabled in User GP" -ForegroundColor Red
}


$LMbMRxhAX7jTCJI2 = Get-ItemProperty -Path  'Registry::HKLM\Software\Policies\Microsoft\Windows\PowerShell\' -Name ExecutionPolicy -ErrorAction SilentlyContinue|Select-Object -ExpandProperty ExecutionPolicy
$UPbMRxhAX7jTCJI2 = Get-ItemProperty -Path  'Registry::HKCU\Software\Policies\Microsoft\Windows\PowerShell\' -Name ExecutionPolicy -ErrorAction SilentlyContinue|Select-Object -ExpandProperty ExecutionPolicy
if ( $LMbMRxhAX7jTCJI2 -eq $null -and  $UPbMRxhAX7jTCJI2S -eq $null)
{
write-host " Script Execution is not configured" -ForegroundColor Yellow
}
if ( $LMbMRxhAX7jTCJI2  -eq '0' )
{
write-host " Allow only signed powershell scripts is enabled in Local Machine GP" -ForegroundColor Green
}
if ( $LMbMRxhAX7jTCJI2  -eq '1' -or $LMbMRxhAX7jTCJI2  -eq '2' )
{
write-host " Powershell scripts are set to allow all scripts or allow local scripts and remote signed scripts in Local Machine GP" -ForegroundColor Red
}
if ( $UPbMRxhAX7jTCJI2  -eq  '0' )
{
write-host " Allow only signed powershell scripts is enabled in User GP" -ForegroundColor Green
}
if ( $UPbMRxhAX7jTCJI2  -eq '1' -or $UPbMRxhAX7jTCJI2  -eq '2')
{
write-host "Powershell scripts are set to allow all scripts or allow local scripts and remote signed scripts in User GP" -ForegroundColor Red
}


write-host "`r`n####################### REGISTRY EDITING TOOLS #######################`r`n"

$ne3X0uL4lhqB1Ga = Get-ItemProperty -Path  'Registry::HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\System\'  -Name DisableRegistryTools -ErrorAction SilentlyContinue|Select-Object -ExpandProperty DisableRegistryTools
if ( $ne3X0uL4lhqB1Ga -eq $null)
{
write-host " Prevent access to registry editing tools is not configured" -ForegroundColor Yellow
}
   elseif ( $ne3X0uL4lhqB1Ga  -eq  '2' )
{
write-host " Prevent access to registry editing tools is enabled" -ForegroundColor Green
}
  elseif ( $ne3X0uL4lhqB1Ga  -eq  '1' )
{
write-host " Prevent access to registry editing tools is disabled" -ForegroundColor Red
}
  else
{
write-host " Prevent access to registry editing tools is set to an unknown setting" -ForegroundColor Red
}

write-host "`r`n####################### REMOTE ASSISTANCE #######################`r`n"

$4KQi6CmJpGgqVAs = Get-ItemProperty -Path  'Registry::HKEY_LOCAL_MACHINE\Software\policies\Microsoft\Windows NT\Terminal Services\'  -Name fAllowUnsolicited -ErrorAction SilentlyContinue|Select-Object -ExpandProperty fAllowUnsolicited
if ( $4KQi6CmJpGgqVAs -eq $null)
{
write-host " Configure Offer Remote Assistance is not configured" -ForegroundColor Yellow
}
   elseif ( $4KQi6CmJpGgqVAs  -eq  '0' )
{
write-host " Configure Offer Remote Assistance is disabled" -ForegroundColor Green
}
  elseif ( $4KQi6CmJpGgqVAs  -eq  '1' )
{
write-host " Configure Offer Remote Assistance is enabled" -ForegroundColor Red
}
  else
{
write-host " Configure Offer Remote Assistance is set to an unknown setting" -ForegroundColor Red
}

$ostWYT0pIug5Qcb = Get-ItemProperty -Path  'Registry::HKEY_LOCAL_MACHINE\Software\policies\Microsoft\Windows NT\Terminal Services\'  -Name fAllowToGetHelp -ErrorAction SilentlyContinue|Select-Object -ExpandProperty fAllowToGetHelp
if ( $ostWYT0pIug5Qcb -eq $null)
{
write-host " Configure Solicited Remote Assistance is not configured" -ForegroundColor Yellow
}
   elseif ( $ostWYT0pIug5Qcb  -eq  '0' )
{
write-host " Configure Solicited Remote Assistance is disabled" -ForegroundColor Green
}
  elseif ( $ostWYT0pIug5Qcb  -eq  '1' )
{
write-host " Configure Solicited Remote Assistance is enabled" -ForegroundColor Red
}
  else
{
write-host " Configure Solicited Remote Assistance is set to an unknown setting" -ForegroundColor Red
}

write-host "`r`n####################### REMOTE DESKTOP SERVICES #######################`r`n"

$kQwHe03XYWy17KG = Get-ItemProperty -Path  'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\Windows NT\Terminal Services\'  -Name fDenyTSConnections -ErrorAction SilentlyContinue|Select-Object -ExpandProperty fDenyTSConnections
if ( $kQwHe03XYWy17KG -eq $null)
{
write-host " Allow users to connect remotely by using Remote Desktop Services is not configured" -ForegroundColor Yellow
}
   elseif ( $kQwHe03XYWy17KG  -eq  '1' )
{
write-host " Allow users to connect remotely by using Remote Desktop Services is disabled" -ForegroundColor Green
}
  elseif ( $kQwHe03XYWy17KG  -eq  '0' )
{
write-host " Allow users to connect remotely by using Remote Desktop Services is enabled" -ForegroundColor Red
}
  else
{
write-host " Allow users to connect remotely by using Remote Desktop Services is set to an unknown setting" -ForegroundColor Red
}

write-host "Unable to check allow log on through remote desktop services"
write-host "Unable to check deny log on through remote desktop services"

$NQV54zJaxh6nOE0 = Get-ItemProperty -Path  'Registry::HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Windows\CredentialsDelegation\'  -Name AllowProtectedCreds -ErrorAction SilentlyContinue|Select-Object -ExpandProperty AllowProtectedCreds
if ( $NQV54zJaxh6nOE0 -eq $null)
{
write-host " Remote host allows delegation of non-exportable credentials is not configured" -ForegroundColor Yellow
}
   elseif ( $NQV54zJaxh6nOE0  -eq  '1' )
{
write-host " Remote host allows delegation of non-exportable credentials is enabled" -ForegroundColor Green
}
  elseif ( $NQV54zJaxh6nOE0  -eq  '0' )
{
write-host " Remote host allows delegation of non-exportable credentials is disabled" -ForegroundColor Red
}
  else
{
write-host " Remote host allows delegation of non-exportable credentials is set to an unknown setting" -ForegroundColor Red
}

$rhnwzd2NLqTAf8J = Get-ItemProperty -Path  'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\Windows NT\Terminal Services\'  -Name AuthenticationLevel -ErrorAction SilentlyContinue|Select-Object -ExpandProperty AuthenticationLevel
if ( $rhnwzd2NLqTAf8J -eq $null)
{
write-host " Configure server authentication for client is not configured" -ForegroundColor Yellow
}
   elseif ( $rhnwzd2NLqTAf8J  -eq  '1' )
{
write-host " Configure server authentication for client is enabled" -ForegroundColor Green
}
  elseif ( $rhnwzd2NLqTAf8J  -eq  '2' -or $rhnwzd2NLqTAf8J  -eq  '0' )
{
write-host " Configure server authentication for client is set to a non-compliant setting" -ForegroundColor Red
}
  else
{
write-host " Configure server authentication for client is set to an unknown setting" -ForegroundColor Red
}

$USPueEgdnK6yjIL = Get-ItemProperty -Path  'Registry::HKEY_CURRENT_USER\SOFTWARE\Policies\Microsoft\Windows NT\Terminal Services\'  -Name DisablePasswordSaving -ErrorAction SilentlyContinue|Select-Object -ExpandProperty DisablePasswordSaving
if ( $USPueEgdnK6yjIL -eq $null)
{
write-host " Do not allow passwords to be saved is not configured" -ForegroundColor Yellow
}
   elseif ( $USPueEgdnK6yjIL  -eq  '1' )
{
write-host " Do not allow passwords to be saved is enabled" -ForegroundColor Green
}
  elseif ( $USPueEgdnK6yjIL  -eq  '0' )
{
write-host " Do not allow passwords to be saved is disabled" -ForegroundColor Red
}
  else
{
write-host " Do not allow passwords to be saved is set to an unknown setting" -ForegroundColor Red
}

$7Zz8LPwJN6ky4gX = Get-ItemProperty -Path  'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\Windows NT\Terminal Services\'  -Name fDenyTSConnections -ErrorAction SilentlyContinue|Select-Object -ExpandProperty fDenyTSConnections
if ( $7Zz8LPwJN6ky4gX -eq $null)
{
write-host " Allow users to connect remotely by using Remote Desktop Services is not configured" -ForegroundColor Yellow
}
   elseif ( $7Zz8LPwJN6ky4gX  -eq  '0' )
{
write-host " Allow users to connect remotely by using Remote Desktop Services is enabled" -ForegroundColor Green
}
  elseif ( $7Zz8LPwJN6ky4gX  -eq  '1' )
{
write-host " Allow users to connect remotely by using Remote Desktop Services is disabled" -ForegroundColor Red
}
  else
{
write-host " Allow users to connect remotely by using Remote Desktop Services is set to an unknown setting" -ForegroundColor Red
}

$fYIVuDva8ER2A9M = Get-ItemProperty -Path  'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\Windows NT\Terminal Services\'  -Name fDisableForcibleLogoff -ErrorAction SilentlyContinue|Select-Object -ExpandProperty fDisableForcibleLogoff
if ( $fYIVuDva8ER2A9M -eq $null)
{
write-host " Deny logoff of an administrator logged in to the console session is not configured" -ForegroundColor Yellow
}
   elseif ( $fYIVuDva8ER2A9M  -eq  '1' )
{
write-host " Deny logoff of an administrator logged in to the console session is enabled" -ForegroundColor Green
}
  elseif ( $fYIVuDva8ER2A9M  -eq  '0' )
{
write-host " Deny logoff of an administrator logged in to the console session is disabled" -ForegroundColor Red
}
  else
{
write-host " Deny logoff of an administrator logged in to the console session is set to an unknown setting" -ForegroundColor Red
}

$RWGtm1iw4Pj0Mhs = Get-ItemProperty -Path  'Registry::HKEY_CURRENT_USER\SOFTWARE\Policies\Microsoft\Windows NT\Terminal Services\'  -Name fDisableClip -ErrorAction SilentlyContinue|Select-Object -ExpandProperty fDisableClip

if ( $RWGtm1iw4Pj0Mhs -eq $null)
{
write-host " Do not allow Clipboard redirection is not configured" -ForegroundColor Yellow
}
   elseif ( $RWGtm1iw4Pj0Mhs  -eq  '1' )
{
write-host " Do not allow Clipboard redirection is enabled" -ForegroundColor Green
}
  elseif ( $RWGtm1iw4Pj0Mhs  -eq  '0' )
{
write-host " Do not allow Clipboard redirection is disabled" -ForegroundColor Red
}
  else
{
write-host " Do not allow Clipboard redirection is set to an unknown setting" -ForegroundColor Red
}

$MJ2WdIt7mlhbckR = Get-ItemProperty -Path  'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\Windows NT\Terminal Services\'  -Name fDisableCdm -ErrorAction SilentlyContinue|Select-Object -ExpandProperty fDisableCdm
if ( $MJ2WdIt7mlhbckR -eq $null)
{
write-host " Do not allow drive redirection is not configured" -ForegroundColor Yellow
}
   elseif ( $MJ2WdIt7mlhbckR  -eq  '1' )
{
write-host " Do not allow drive redirection is enabled" -ForegroundColor Green
}
  elseif ( $MJ2WdIt7mlhbckR  -eq  '0' )
{
write-host " Do not allow drive redirection is disabled" -ForegroundColor Red
}
  else
{
write-host " Do not allow drive redirection is set to an unknown setting" -ForegroundColor Red
}

$lRPQ5MjsugZpCAI = Get-ItemProperty -Path  'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\Windows NT\Terminal Services\'  -Name fPromptForPassword -ErrorAction SilentlyContinue|Select-Object -ExpandProperty fPromptForPassword
if ( $lRPQ5MjsugZpCAI -eq $null)
{
write-host " Always prompt for password upon connection is not configured" -ForegroundColor Yellow
}
   elseif ( $lRPQ5MjsugZpCAI  -eq  '1' )
{
write-host " Always prompt for password upon connection is enabled" -ForegroundColor Green
}
  elseif ( $lRPQ5MjsugZpCAI  -eq  '0' )
{
write-host " Always prompt for password upon connection is disabled" -ForegroundColor Red
}
  else
{
write-host " Always prompt for password upon connection is set to an unknown setting" -ForegroundColor Red
}

$CPoKihTNYQpqsBz = Get-ItemProperty -Path  'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\Windows NT\Terminal Services\'  -Name fWritableTSCCPermTab -ErrorAction SilentlyContinue|Select-Object -ExpandProperty fWritableTSCCPermTab
if ( $CPoKihTNYQpqsBz -eq $null)
{
write-host " Do not allow local administrators to customize permissions is not configured" -ForegroundColor Yellow
}
   elseif ( $CPoKihTNYQpqsBz  -eq  '0' )
{
write-host " Do not allow local administrators to customize permissions is enabled" -ForegroundColor Green
}
  elseif ( $CPoKihTNYQpqsBz  -eq  '1' )
{
write-host " Do not allow local administrators to customize permissions is disabled" -ForegroundColor Red
}
  else
{
write-host " Do not allow local administrators to customize permissions is set to an unknown setting" -ForegroundColor Red
}

$k2FQDrJen34MOVg = Get-ItemProperty -Path  'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\Windows NT\Terminal Services\'  -Name fEncryptRPCTraffic -ErrorAction SilentlyContinue|Select-Object -ExpandProperty fEncryptRPCTraffic
if ( $k2FQDrJen34MOVg -eq $null)
{
write-host " Require secure RPC communication is not configured" -ForegroundColor Yellow
}
   elseif ( $k2FQDrJen34MOVg  -eq  '1' )
{
write-host " Require secure RPC communication is enabled" -ForegroundColor Green
}
  elseif ( $k2FQDrJen34MOVg  -eq  '0' )
{
write-host " Require secure RPC communication is disabled" -ForegroundColor Red
}
  else
{
write-host " Require secure RPC communication is set to an unknown setting" -ForegroundColor Red
}

$ycroPUFjHk1l4aq = Get-ItemProperty -Path  'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\Windows NT\Terminal Services\'  -Name SecurityLayer -ErrorAction SilentlyContinue|Select-Object -ExpandProperty SecurityLayer
if ( $ycroPUFjHk1l4aq -eq $null)
{
write-host " Require use of specific security layer for remote (RDP) connections is not configured" -ForegroundColor Yellow
}
   elseif ( $ycroPUFjHk1l4aq  -eq  '2' )
{
write-host " Require use of specific security layer for remote (RDP) connections is set to SSL" -ForegroundColor Green
}
  elseif ( $ycroPUFjHk1l4aq  -eq  '1' -or $ycroPUFjHk1l4aq  -eq  '0' )
{
write-host " Require use of specific security layer for remote (RDP) connections set to Negotiate or RDP" -ForegroundColor Red
}
  else
{
write-host " Require use of specific security layer for remote (RDP) connections is set to an unknown setting" -ForegroundColor Red
}

$vYkIVXt8CZfzRT3 = Get-ItemProperty -Path  'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\Windows NT\Terminal Services\'  -Name UserAuthentication -ErrorAction SilentlyContinue|Select-Object -ExpandProperty UserAuthentication
if ( $vYkIVXt8CZfzRT3 -eq $null)
{
write-host " Require user authentication for remote connections by using Network Level Authentication is not configured" -ForegroundColor Yellow
}
   elseif ( $vYkIVXt8CZfzRT3  -eq  '1' )
{
write-host " Require user authentication for remote connections by using Network Level Authentication is enabled" -ForegroundColor Green
}
  elseif ( $vYkIVXt8CZfzRT3  -eq  '0' )
{
write-host " Require user authentication for remote connections by using Network Level Authentication is disabled" -ForegroundColor Red
}
  else
{
write-host " Require user authentication for remote connections by using Network Level Authentication is set to an unknown setting" -ForegroundColor Red
}

$MXAzBSUFTGujfc1 = Get-ItemProperty -Path  'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\Windows NT\Terminal Services\'  -Name MinEncryptionLevel -ErrorAction SilentlyContinue|Select-Object -ExpandProperty MinEncryptionLevel
if ( $MXAzBSUFTGujfc1 -eq $null)
{
write-host " Set client connection encryption level is not configured" -ForegroundColor Yellow
}
   elseif ( $MXAzBSUFTGujfc1  -eq  '3' )
{
write-host " Set client connection encryption level is set to high" -ForegroundColor Green
}
  elseif ( $MXAzBSUFTGujfc1  -eq  '1' -or $MXAzBSUFTGujfc1  -eq  '2' )
{
write-host " Set client connection encryption level is set to client compatible or a low level" -ForegroundColor Red
}
  else
{
write-host " Set client connection encryption level is set to an unknown setting" -ForegroundColor Red
}

write-host "Unable to check allow log on through remote desktop services"
write-host "Unable to check deny log on through remote desktop services"


write-host "`r`n####################### REMOTE PROCEDURE CALL #######################`r`n"

$HWPLG72S8TrAqKk = Get-ItemProperty -Path  'Registry::HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Windows NT\Rpc\'  -Name RestrictRemoteClients -ErrorAction SilentlyContinue|Select-Object -ExpandProperty RestrictRemoteClients
if ( $HWPLG72S8TrAqKk -eq $null)
{
write-host " Restrict Unauthenticated RPC clients is not configured" -ForegroundColor Yellow
}
   elseif ( $HWPLG72S8TrAqKk  -eq  '1' )
{
write-host " Restrict Unauthenticated RPC clients is enabled" -ForegroundColor Green
}
  elseif ( $HWPLG72S8TrAqKk  -eq  '0' -or $HWPLG72S8TrAqKk  -eq  '2'  )
{
write-host " Restrict Unauthenticated RPC clients is disabled" -ForegroundColor Red
}
  else
{
write-host " Restrict Unauthenticated RPC clients is set to an unknown setting" -ForegroundColor Red
}


write-host "`r`n####################### REPORTING SYSTEM INFORMATION #######################`r`n"

$PNH7sOv6IUqTLd0 = Get-ItemProperty -Path  'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\Windows\ScriptedDiagnosticsProvider\Policy\'  -Name DisableQueryRemoteServer -ErrorAction SilentlyContinue|Select-Object -ExpandProperty DisableQueryRemoteServer
if ( $PNH7sOv6IUqTLd0 -eq $null)
{
write-host " Microsoft Support Diagnostic Tool: Turn on MSDT interactive communication with support provider is not configured" -ForegroundColor Yellow
}
   elseif ( $PNH7sOv6IUqTLd0  -eq  '0' )
{
write-host " Microsoft Support Diagnostic Tool: Turn on MSDT interactive communication with support provider is disabled" -ForegroundColor Green
}
  elseif ( $PNH7sOv6IUqTLd0  -eq  '1' )
{
write-host " Microsoft Support Diagnostic Tool: Turn on MSDT interactive communication with support provider is enabled" -ForegroundColor Red
}
  else
{
write-host " Microsoft Support Diagnostic Tool: Turn on MSDT interactive communication with support provider is set to an unknown setting" -ForegroundColor Red
}

$pB5HU3iuVdShzK9 = Get-ItemProperty -Path  'Registry::HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Windows\AppCompat\'  -Name DisableInventory -ErrorAction SilentlyContinue|Select-Object -ExpandProperty DisableInventory
if ( $pB5HU3iuVdShzK9 -eq $null)
{
write-host " Turn off Inventory Collector is not configured" -ForegroundColor Yellow
}
   elseif ( $pB5HU3iuVdShzK9  -eq  '1' )
{
write-host " Turn off Inventory Collector is enabled" -ForegroundColor Green
}
  elseif ( $pB5HU3iuVdShzK9  -eq  '0' )
{
write-host " Turn off Inventory Collector is disabled" -ForegroundColor Red
}
  else
{
write-host " Turn off Inventory Collector is set to an unknown setting" -ForegroundColor Red
}

$HhF0z6Ccr3LGPxd = Get-ItemProperty -Path  'Registry::HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Windows\AppCompat\'  -Name DisableUAR -ErrorAction SilentlyContinue|Select-Object -ExpandProperty DisableUAR
if ( $HhF0z6Ccr3LGPxd -eq $null)
{
write-host " Turn off Steps Recorder is not configured" -ForegroundColor Yellow
}
   elseif ( $HhF0z6Ccr3LGPxd  -eq  '1' )
{
write-host " Turn off Steps Recorder is enabled" -ForegroundColor Green
}
  elseif ( $HhF0z6Ccr3LGPxd  -eq  '0' )
{
write-host " Turn off Steps Recorder is disabled" -ForegroundColor Red
}
  else
{
write-host " Turn off Steps Recorder is set to an unknown setting" -ForegroundColor Red
}

$LM8KSMxRACOWXwybq = Get-ItemProperty -Path  'Registry::HKLM\Software\Policies\Microsoft\Windows\DataCollection\' -Name AllowTelemetry -ErrorAction SilentlyContinue|Select-Object -ExpandProperty AllowTelemetry
$UP8KSMxRACOWXwybq = Get-ItemProperty -Path  'Registry::HKCU\Software\Policies\Microsoft\Windows\DataCollection\' -Name AllowTelemetry -ErrorAction SilentlyContinue|Select-Object -ExpandProperty AllowTelemetry
if ( $LM8KSMxRACOWXwybq -eq $null -and  $UP8KSMxRACOWXwybq -eq $null)
{
write-host " Allow Telemetry is not configured" -ForegroundColor Yellow
}
if ( $LM8KSMxRACOWXwybq  -eq '0' )
{
write-host " Allow Telemetry is enabled in Local Machine GP" -ForegroundColor Green
}
if ( $LM8KSMxRACOWXwybq  -eq '1' -or $LM8KSMxRACOWXwybq  -eq '2' -or $LM8KSMxRACOWXwybq  -eq '3' )
{
write-host " Allow Telemetry is set to a non-compliant setting in Local Machine GP" -ForegroundColor Red
}
if ( $UP8KSMxRACOWXwybq  -eq  '0' )
{
write-host " Allow Telemetry is enabled in User GP" -ForegroundColor Green
}
if ( $LM8KSMxRACOWXwybq  -eq '1' -or $LM8KSMxRACOWXwybq  -eq '2' -or $LM8KSMxRACOWXwybq  -eq '3' )
{
write-host " Allow Telemetry is set to a non-compliant setting in User GP" -ForegroundColor Red
}


$KVHIZdcponOfwF7 = Get-ItemProperty -Path  'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\Windows\Windows Error Reporting\'  -Name CorporateWerServer -ErrorAction SilentlyContinue|Select-Object -ExpandProperty CorporateWerServer
if ( $KVHIZdcponOfwF7 -eq $null)
{
write-host " Configure Corporate Windows Error Reporting is not configured" -ForegroundColor Red
}
  else
{
write-host " The corporate WER server is configured as $KVHIZdcponOfwF7" -ForegroundColor Green
}

$KVHIZdcponOfwF = Get-ItemProperty -Path  'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\Windows\Windows Error Reporting\'  -Name CorporateWerUseSSL -ErrorAction SilentlyContinue|Select-Object -ExpandProperty CorporateWerUseSSL
if ( $KVHIZdcponOfwF -eq $null)
{
write-host "Connect using SSL is not configured" -ForegroundColor Yellow
}
   elseif ( $KVHIZdcponOfwF  -eq  '1' )
{
write-host "Connect using SSL is enabled" -ForegroundColor Green
}
  elseif ( $KVHIZdcponOfwF  -eq  '0' )
{
write-host "Connect using SSL is disabled" -ForegroundColor Red
}

write-host "`r`n####################### SAFE MODE #######################`r`n"


$HhF0z6Ccr3LGPx = Get-ItemProperty -Path  'Registry::HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Policies\System\'  -Name SafeModeBlockNonAdmins -ErrorAction SilentlyContinue|Select-Object -ExpandProperty SafeModeBlockNonAdmins
if ( $HhF0z6Ccr3LGPx -eq $null)
{
write-host "Block Non-Administrators in Safe Mode not configured" -ForegroundColor Yellow
}
   elseif ( $HhF0z6Ccr3LGPx  -eq  '1' )
{
write-host "Block Non-Administrators in Safe Mode is enabled" -ForegroundColor Green
}
  elseif ( $HhF0z6Ccr3LGPx  -eq  '0' )
{
write-host "Block Non-Administrators in Safe Mode is disabled" -ForegroundColor Red
}
  else
{
write-host "Block Non-Administrators in Safe Mode is set to an unknown setting" -ForegroundColor Red
}

write-host "`r`n####################### SECURE CHANNEL COMMUNICATIONS #######################`r`n"

write-host "Can't check this whole section yet"


write-host "`r`n####################### SECURITY POLICIES #######################`r`n"

$ZCprfnJQOVLF4wT = Get-ItemProperty -Path  'Registry::HKEY_LOCAL_MACHINE\Software\Microsoft\wcmsvc\wifinetworkmanager\config\'  -Name AutoConnectAllowedOEM -ErrorAction SilentlyContinue|Select-Object -ExpandProperty AutoConnectAllowedOEM
if ( $ZCprfnJQOVLF4wT -eq $null)
{
write-host " Allow Windows to automatically connect to suggested open hotspots, to networks shared by contacts, and to hotspots offering paid services is not configured" -ForegroundColor Yellow
}
   elseif ( $ZCprfnJQOVLF4wT  -eq  '0' )
{
write-host " Allow Windows to automatically connect to suggested open hotspots, to networks shared by contacts, and to hotspots offering paid services is disabled" -ForegroundColor Green
}
  elseif ( $ZCprfnJQOVLF4wT  -eq  '1' )
{
write-host " Allow Windows to automatically connect to suggested open hotspots, to networks shared by contacts, and to hotspots offering paid services is enabled" -ForegroundColor Red
}
  else
{
write-host " Allow Windows to automatically connect to suggested open hotspots, to networks shared by contacts, and to hotspots offering paid services is set to an unknown setting" -ForegroundColor Red
}


$x783w1bfW4nNCZV = Get-ItemProperty -Path  'Registry::HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Windows\CloudContent\'  -Name DisableWindowsConsumerFeatures -ErrorAction SilentlyContinue|Select-Object -ExpandProperty DisableWindowsConsumerFeatures
if ( $x783w1bfW4nNCZV -eq $null)
{
write-host " Turn off Microsoft consumer experiences is not configured" -ForegroundColor Yellow
}
   elseif ( $x783w1bfW4nNCZV  -eq  '1' )
{
write-host " Turn off Microsoft consumer experiences is enabled" -ForegroundColor Green
}
  elseif ( $x783w1bfW4nNCZV  -eq  '0' )
{
write-host " Turn off Microsoft consumer experiences is disabled" -ForegroundColor Red
}
  else
{
write-host " Turn off Microsoft consumer experiences is set to an unknown setting" -ForegroundColor Red
}

$PAch3CtoO9Ijfvr = Get-ItemProperty -Path  'Registry::HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Windows\Explorer\'  -Name NoHeapTerminationOnCorruption -ErrorAction SilentlyContinue|Select-Object -ExpandProperty NoHeapTerminationOnCorruption
if ( $PAch3CtoO9Ijfvr -eq $null)
{
write-host " Turn off heap termination on corruption is not configured" -ForegroundColor Yellow
}
   elseif ( $PAch3CtoO9Ijfvr  -eq  '0' )
{
write-host " Turn off heap termination on corruption is disabled" -ForegroundColor Green
}
  elseif ( $PAch3CtoO9Ijfvr  -eq  '1' )
{
write-host " Turn off heap termination on corruption is enabled" -ForegroundColor Red
}
  else
{
write-host " Turn off heap termination on corruption is set to an unknown setting" -ForegroundColor Red
}

$X7bBFV0iTPk6rYj = Get-ItemProperty -Path  'Registry::HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\'  -Name PreXPSP2ShellProtocolBehavior -ErrorAction SilentlyContinue|Select-Object -ExpandProperty PreXPSP2ShellProtocolBehavior
if ( $X7bBFV0iTPk6rYj -eq $null)
{
write-host " Turn off shell protocol protected mode is not configured" -ForegroundColor Yellow
}
   elseif ( $X7bBFV0iTPk6rYj  -eq  '0' )
{
write-host " Turn off shell protocol protected mode is disabled" -ForegroundColor Green
}
  elseif ( $X7bBFV0iTPk6rYj  -eq  '1' )
{
write-host " Turn off shell protocol protected mode is enabled" -ForegroundColor Red
}
  else
{
write-host " Turn off shell protocol protected mode is set to an unknown setting" -ForegroundColor Red
}

$LMwVsYrmNLSvR3156 = Get-ItemProperty -Path  'Registry::HKLM\Software\Policies\Microsoft\Internet Explorer\Feeds\' -Name DisableEnclosureDownload -ErrorAction SilentlyContinue|Select-Object -ExpandProperty DisableEnclosureDownload
$UPwVsYrmNLSvR3156 = Get-ItemProperty -Path  'Registry::HKCU\Software\Policies\Microsoft\Internet Explorer\Feeds\' -Name DisableEnclosureDownload -ErrorAction SilentlyContinue|Select-Object -ExpandProperty DisableEnclosureDownload
if ( $LMwVsYrmNLSvR3156 -eq $null -and  $UPwVsYrmNLSvR3156 -eq $null)
{
write-host " Prevent downloading of enclosures is not configured" -ForegroundColor Yellow
}
if ( $LMwVsYrmNLSvR3156  -eq '1' )
{
write-host " Prevent downloading of enclosures is enabled in Local Machine GP" -ForegroundColor Green
}
if ( $LMwVsYrmNLSvR3156  -eq '0' )
{
write-host " Prevent downloading of enclosures is disabled in Local Machine GP" -ForegroundColor Red
}
if ( $UPwVsYrmNLSvR3156  -eq  '1' )
{
write-host " Prevent downloading of enclosures is enabled in User GP" -ForegroundColor Green
}
if ( $UPwVsYrmNLSvR3156  -eq  '0' )
{
write-host " Prevent downloading of enclosures is disabled in User GP" -ForegroundColor Red
}

$g0OCVPTHarb4FiU = Get-ItemProperty -Path  'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\Windows\Windows Search\'  -Name AllowIndexingEncryptedStoresOrItems -ErrorAction SilentlyContinue|Select-Object -ExpandProperty AllowIndexingEncryptedStoresOrItems
if ( $g0OCVPTHarb4FiU -eq $null)
{
write-host " Allow indexing of encrypted files is not configured" -ForegroundColor Yellow
}
   elseif ( $g0OCVPTHarb4FiU  -eq  '0' )
{
write-host " Allow indexing of encrypted files is disabled" -ForegroundColor Green
}
  elseif ( $g0OCVPTHarb4FiU  -eq  '1' )
{
write-host " Allow indexing of encrypted files is enabled" -ForegroundColor Red
}
  else
{
write-host " Allow indexing of encrypted files is set to an unknown setting" -ForegroundColor Red
}

$OqU8k1BrR0gFnNz = Get-ItemProperty -Path  'Registry::HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Windows\GameDVR\'  -Name AllowGameDVR -ErrorAction SilentlyContinue|Select-Object -ExpandProperty AllowGameDVR
if ( $OqU8k1BrR0gFnNz -eq $null)
{
write-host " Enables or disables Windows Game Recording and Broadcasting is not configured" -ForegroundColor Yellow
}
   elseif ( $OqU8k1BrR0gFnNz  -eq  '0' )
{
write-host " Enables or disables Windows Game Recording and Broadcasting is disabled" -ForegroundColor Green
}
  elseif ( $OqU8k1BrR0gFnNz  -eq  '1' )
{
write-host " Enables or disables Windows Game Recording and Broadcasting is enabled" -ForegroundColor Red
}
  else
{
write-host " Enables or disables Windows Game Recording and Broadcasting is set to an unknown setting" -ForegroundColor Red
}

write-host "Unable to check some security options on page 41"

write-host "`r`n####################### SERVER MESSAGE BLOCK SESSIONS #######################`r`n"


$JZyMnHu1K3IXh40 = Get-ItemProperty -Path  'Registry::HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\MrxSmb10\'  -Name Start -ErrorAction SilentlyContinue|Select-Object -ExpandProperty Start
if ( $JZyMnHu1K3IXh40 -eq $null)
{
write-host " Configure SMB v1 client driver is not configured" -ForegroundColor Yellow
}
   elseif ( $JZyMnHu1K3IXh40  -eq  '4' )
{
write-host " Configure SMB v1 client driver is disabled" -ForegroundColor Green
}
  elseif ( $JZyMnHu1K3IXh40  -eq  '2' -or $JZyMnHu1K3IXh40  -eq  '3' )
{
write-host " Configure SMB v1 client driver is set to manual or automatic start" -ForegroundColor Red
}
  else
{
write-host " Configure SMB v1 client driver is set to an unknown setting" -ForegroundColor Red
}

$CJYvExedTmlj9OQ = Get-ItemProperty -Path  'Registry::HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\LanmanServer\Parameters\'  -Name SMB1 -ErrorAction SilentlyContinue|Select-Object -ExpandProperty SMB1
if ( $CJYvExedTmlj9OQ -eq $null)
{
write-host " Configure SMB v1 server is not configured" -ForegroundColor Yellow
}
   elseif ( $CJYvExedTmlj9OQ  -eq  '0' )
{
write-host " Configure SMB v1 server is disabled" -ForegroundColor Green
}
  elseif ( $CJYvExedTmlj9OQ  -eq  '1' )
{
write-host " Configure SMB v1 server is enabled" -ForegroundColor Red
}
  else
{
write-host " Configure SMB v1 server is set to an unknown setting" -ForegroundColor Red
}

write-host "Unable to check for remaining microsoft network server settings"

write-host "`r`n####################### SESSION LOCKING #######################`r`n"

$tMm2f35wdzqlIkg = Get-ItemProperty -Path  'Registry::HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Windows\Personalization\'  -Name NoLockScreenCamera -ErrorAction SilentlyContinue|Select-Object -ExpandProperty NoLockScreenCamera
if ( $tMm2f35wdzqlIkg -eq $null)
{
write-host " Prevent enabling lock screen camera is not configured" -ForegroundColor Yellow
}
   elseif ( $tMm2f35wdzqlIkg  -eq  '1' )
{
write-host " Prevent enabling lock screen camera is enabled" -ForegroundColor Green
}
  elseif ( $tMm2f35wdzqlIkg  -eq  '0' )
{
write-host " Prevent enabling lock screen camera is disabled" -ForegroundColor Red
}
  else
{
write-host " Prevent enabling lock screen camera is set to an unknown setting" -ForegroundColor Red
}

$9Ot0aqonKNiEm5b = Get-ItemProperty -Path  'Registry::HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Windows\Personalization\'  -Name NoLockScreenSlideshow -ErrorAction SilentlyContinue|Select-Object -ExpandProperty NoLockScreenSlideshow
if ( $9Ot0aqonKNiEm5b -eq $null)
{
write-host " Prevent enabling lock screen slide show is not configured" -ForegroundColor Yellow
}
   elseif ( $9Ot0aqonKNiEm5b  -eq  '1' )
{
write-host " Prevent enabling lock screen slide show is enabled" -ForegroundColor Green
}
  elseif ( $9Ot0aqonKNiEm5b  -eq  '0' )
{
write-host " Prevent enabling lock screen slide show is disabled" -ForegroundColor Red
}
  else
{
write-host " Prevent enabling lock screen slide show is set to an unknown setting" -ForegroundColor Red
}

$cbGLB9V2Rhk7fq5 = Get-ItemProperty -Path  'Registry::HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Windows\System\'  -Name AllowDomainDelayLock -ErrorAction SilentlyContinue|Select-Object -ExpandProperty AllowDomainDelayLock
if ( $cbGLB9V2Rhk7fq5 -eq $null)
{
write-host " Allow users to select when a password is required when resuming from connected standby is not configured" -ForegroundColor Yellow
}
   elseif ( $cbGLB9V2Rhk7fq5  -eq  '0' )
{
write-host " Allow users to select when a password is required when resuming from connected standby is disabled" -ForegroundColor Green
}
  elseif ( $cbGLB9V2Rhk7fq5  -eq  '1' )
{
write-host " Allow users to select when a password is required when resuming from connected standby is enabled" -ForegroundColor Red
}
  else
{
write-host " Allow users to select when a password is required when resuming from connected standby is set to an unknown setting" -ForegroundColor Red
}

$jrSiA6Xq2mBVpCZ = Get-ItemProperty -Path  'Registry::HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Windows\System\'  -Name DisableLockScreenAppNotifications -ErrorAction SilentlyContinue|Select-Object -ExpandProperty DisableLockScreenAppNotifications
if ( $jrSiA6Xq2mBVpCZ -eq $null)
{
write-host " Turn off app notifications on the lock screen is not configured" -ForegroundColor Yellow
}
   elseif ( $jrSiA6Xq2mBVpCZ  -eq  '1' )
{
write-host " Turn off app notifications on the lock screen is enabled" -ForegroundColor Green
}
  elseif ( $jrSiA6Xq2mBVpCZ  -eq  '0' )
{
write-host " Turn off app notifications on the lock screen is disabled" -ForegroundColor Red
}
  else
{
write-host " Turn off app notifications on the lock screen is set to an unknown setting" -ForegroundColor Red
}

$aBGYEMCPVjRLeFc = Get-ItemProperty -Path  'Registry::HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Windows\Explorer\'  -Name ShowLockOption -ErrorAction SilentlyContinue|Select-Object -ExpandProperty ShowLockOption
if ( $aBGYEMCPVjRLeFc -eq $null)
{
write-host " Show lock in the user tile menu is not configured" -ForegroundColor Yellow
}
   elseif ( $aBGYEMCPVjRLeFc  -eq  '1' )
{
write-host " Show lock in the user tile menu is enabled" -ForegroundColor Green
}
  elseif ( $aBGYEMCPVjRLeFc  -eq  '0' )
{
write-host " Show lock in the user tile menu is disabled" -ForegroundColor Red
}
  else
{
write-host " Show lock in the user tile menu is set to an unknown setting" -ForegroundColor Red
}

$oRJPdEy5i0DCqFX = Get-ItemProperty -Path  'Registry::HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\WindowsInkWorkspace\'  -Name AllowWindowsInkWorkspace -ErrorAction SilentlyContinue|Select-Object -ExpandProperty AllowWindowsInkWorkspace
if ( $oRJPdEy5i0DCqFX -eq $null)
{
write-host " Allow Windows Ink Workspace is not configured" -ForegroundColor Yellow
}
   elseif ( $oRJPdEy5i0DCqFX  -eq  '1' )
{
write-host " Allow Windows Ink Workspace is on but dissalow access above lock" -ForegroundColor Green
}
  elseif ( $oRJPdEy5i0DCqFX  -eq  '0' -or $oRJPdEy5i0DCqFX  -eq  '2' )
{
write-host " Allow Windows Ink Workspace is disabled or turned on, both not recommended settings" -ForegroundColor Red
}
  else
{
write-host " Allow Windows Ink Workspace is set to an unknown setting" -ForegroundColor Red
}

write-host "Unable to check machine inactivity limit"

$nKErRNAU3b4k6hI = Get-ItemProperty -Path  'Registry::HKEY_CURRENT_USER\Software\Policies\Microsoft\Windows\Control Panel\Desktop\'  -Name ScreenSaveActive -ErrorAction SilentlyContinue|Select-Object -ExpandProperty ScreenSaveActive
if ( $nKErRNAU3b4k6hI -eq $null)
{
write-host " Enable screen saver is not configured" -ForegroundColor Yellow
}
   elseif ( $nKErRNAU3b4k6hI  -eq  '1' )
{
write-host " Enable screen saver is enabled" -ForegroundColor Green
}
  elseif ( $nKErRNAU3b4k6hI  -eq  '0' )
{
write-host " Enable screen saver is disabled" -ForegroundColor Red
}
  else
{
write-host " Enable screen saver is set to an unknown setting" -ForegroundColor Red
}

$v692ozEayg53Lfs = Get-ItemProperty -Path  'Registry::HKEY_CURRENT_USER\Software\Policies\Microsoft\Windows\Control Panel\Desktop\'  -Name ScreenSaverIsSecure -ErrorAction SilentlyContinue|Select-Object -ExpandProperty ScreenSaverIsSecure
if ( $v692ozEayg53Lfs -eq $null)
{
write-host " Password protect the screen saver is not configured" -ForegroundColor Yellow
}
   elseif ( $v692ozEayg53Lfs  -eq  '1' )
{
write-host " Password protect the screen saver is enabled" -ForegroundColor Green
}
  elseif ( $v692ozEayg53Lfs  -eq  '0' )
{
write-host " Password protect the screen saver is disabled" -ForegroundColor Red
}
  else
{
write-host " Password protect the screen saver is set to an unknown setting" -ForegroundColor Red
}

$EWeBJdm8rjbwAo3 = Get-ItemProperty -Path  'Registry::HKEY_CURRENT_USER\Software\Policies\Microsoft\Windows\Control Panel\Desktop\'  -Name ScreenSaveTimeOut -ErrorAction SilentlyContinue|Select-Object -ExpandProperty ScreenSaveTimeOut
if ( $EWeBJdm8rjbwAo3 -eq $null)
{
write-host " Screen saver timeout is not configured" -ForegroundColor Yellow
}
   elseif ( $EWeBJdm8rjbwAo3  -eq  '900' )
{
write-host " Screen saver timeout is set compliant" -ForegroundColor Green
}
  elseif ( $EWeBJdm8rjbwAo3  -lt '900')
{
write-host " Screen saver timeout is lower than a compliant setting" -ForegroundColor Red
}
  elseif ( $EWeBJdm8rjbwAo3  -gt '900')
{
write-host " Screen saver timeout is higher than the compliant setting" -ForegroundColor Green
}
  else
{
write-host " Screen saver timeout is set to an unknown setting" -ForegroundColor Red
}

$7NdvQjghTrwKYW4 = Get-ItemProperty -Path  'Registry::HKEY_CURRENT_USER\SOFTWARE\Policies\Microsoft\Windows\CurrentVersion\PushNotifications\'  -Name NoToastApplicationNotificationOnLockScreen -ErrorAction SilentlyContinue|Select-Object -ExpandProperty NoToastApplicationNotificationOnLockScreen
if ( $7NdvQjghTrwKYW4 -eq $null)
{
write-host " Turn off toast notifications on the lock screen is not configured" -ForegroundColor Yellow
}
   elseif ( $7NdvQjghTrwKYW4  -eq  '1' )
{
write-host " Turn off toast notifications on the lock screen is enabled" -ForegroundColor Green
}
  elseif ( $7NdvQjghTrwKYW4  -eq  '0' )
{
write-host " Turn off toast notifications on the lock screen is disabled" -ForegroundColor Red
}
  else
{
write-host " Turn off toast notifications on the lock screen is set to an unknown setting" -ForegroundColor Red
}

$YcLMvmzxA0X3tu6 = Get-ItemProperty -Path  'Registry::HKEY_CURRENT_USER\Software\Policies\Microsoft\Windows\CloudContent\'  -Name DisableThirdPartySuggestions -ErrorAction SilentlyContinue|Select-Object -ExpandProperty DisableThirdPartySuggestions
if ( $YcLMvmzxA0X3tu6 -eq $null)
{
write-host " Do not suggest third-party content in Windows spotlight is not configured" -ForegroundColor Yellow
}
   elseif ( $YcLMvmzxA0X3tu6  -eq  '1' )
{
write-host " Do not suggest third-party content in Windows spotlight is enabled" -ForegroundColor Green
}
  elseif ( $YcLMvmzxA0X3tu6  -eq  '0' )
{
write-host " Do not suggest third-party content in Windows spotlight is disabled" -ForegroundColor Red
}
  else
{
write-host " Do not suggest third-party content in Windows spotlight is set to an unknown setting" -ForegroundColor Red
}


write-host "`r`n####################### SOFTWARE-BASED FIREWALLS #######################`r`n"

write-host "`r`n####################### SOUND RECORDER #######################`r`n"

$IAtVlOZ8HnEGCq5 = Get-ItemProperty -Path  'Registry::HKEY_CURRENT_USER\SOFTWARE\Policies\Microsoft\SoundRecorder\'  -Name Soundrec -ErrorAction SilentlyContinue|Select-Object -ExpandProperty Soundrec
if ( $IAtVlOZ8HnEGCq5 -eq $null)
{
write-host " Do not allow Sound Recorder to run is not configured" -ForegroundColor Yellow
}
   elseif ( $IAtVlOZ8HnEGCq5  -eq  '1' )
{
write-host " Do not allow Sound Recorder to run is enabled" -ForegroundColor Green
}
  elseif ( $IAtVlOZ8HnEGCq5  -eq  '0' )
{
write-host " Do not allow Sound Recorder to run is disabled" -ForegroundColor Red
}
  else
{
write-host " Do not allow Sound Recorder to run is set to an unknown setting" -ForegroundColor Red
}

write-host "`r`n####################### STANDARD OPERATING ENVIRONMENT #######################`r`n"

write-host "`r`n####################### SYSTEM BACKUP AND RESTORE #######################`r`n"

write-host "Can't check system backup and restore settings yet"

write-host "`r`n####################### SYSTEM CRYPTOGRAPHY #######################`r`n"

write-host "Unable to check force strong key protection for user keys stored ont eh computer"


$9UNpgi6osfkQlnF = Get-ItemProperty -Path  'Registry::HKEY_LOCAL_MACHINE\Software\Policies\Centrify\CentrifyDC\Settings\Fips\'  -Name fips.mode.enable -ErrorAction SilentlyContinue|Select-Object -ExpandProperty fips.mode.enable
if ( $9UNpgi6osfkQlnF -eq $null)
{
write-host " Use FIPS compliant algorithms for encryption, hashing and signing is not configured" -ForegroundColor Yellow
}
   elseif ( $9UNpgi6osfkQlnF  -eq  'true' )
{
write-host " Use FIPS compliant algorithms for encryption, hashing and signing is enabled" -ForegroundColor Green
}
  elseif ( $9UNpgi6osfkQlnF  -eq  'false' )
{
write-host " Use FIPS compliant algorithms for encryption, hashing and signing is disabled" -ForegroundColor Red
}
  else
{
write-host " Use FIPS compliant algorithms for encryption, hashing and signing is set to an unknown setting" -ForegroundColor Red
}

write-host "`r`n####################### USER RIGHTS POLICIES #######################`r`n"

write-host "Unable to check this chapter"

write-host "`r`n####################### VIRTUALISED WEB AND EMAIL ACCESS #######################`r`n"


write-host "`r`n####################### WINDOWS REMOTE MANAGEMENT #######################`r`n"

$q8Y9g4oz6TAULkJ = Get-ItemProperty -Path  'Registry::HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Windows\WinRM\Client\'  -Name AllowBasic -ErrorAction SilentlyContinue|Select-Object -ExpandProperty AllowBasic
if ( $q8Y9g4oz6TAULkJ -eq $null)
{
write-host " Allow Basic authentication is not configured" -ForegroundColor Yellow
}
   elseif ( $q8Y9g4oz6TAULkJ  -eq  '0' )
{
write-host " Allow Basic authentication is disabled" -ForegroundColor Green
}
  elseif ( $q8Y9g4oz6TAULkJ  -eq  '1' )
{
write-host " Allow Basic authentication is enabled" -ForegroundColor Red
}
  else
{
write-host " Allow Basic authentication is set to an unknown setting" -ForegroundColor Red
}


$svkG3Au1aOf5IwN = Get-ItemProperty -Path  'Registry::HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Windows\WinRM\Client\'  -Name AllowUnencryptedTraffic -ErrorAction SilentlyContinue|Select-Object -ExpandProperty AllowUnencryptedTraffic
if ( $svkG3Au1aOf5IwN -eq $null)
{
write-host " Allow unencrypted traffic is not configured" -ForegroundColor Yellow
}
   elseif ( $svkG3Au1aOf5IwN  -eq  '0' )
{
write-host " Allow unencrypted traffic is disabled" -ForegroundColor Green
}
  elseif ( $svkG3Au1aOf5IwN  -eq  '1' )
{
write-host " Allow unencrypted traffic is enabled" -ForegroundColor Red
}
  else
{
write-host " Allow unencrypted traffic is set to an unknown setting" -ForegroundColor Red
}

$Zvk72J5CFEsdqhg = Get-ItemProperty -Path  'Registry::HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Windows\WinRM\Client\'  -Name AllowDigest -ErrorAction SilentlyContinue|Select-Object -ExpandProperty AllowDigest
if ( $Zvk72J5CFEsdqhg -eq $null)
{
write-host " Disallow Digest authentication is not configured" -ForegroundColor Yellow
}
   elseif ( $Zvk72J5CFEsdqhg  -eq  '0' )
{
write-host " Disallow Digest authentication is enabled" -ForegroundColor Green
}
  elseif ( $Zvk72J5CFEsdqhg  -eq  '1' )
{
write-host " Disallow Digest authentication is disabled" -ForegroundColor Red
}
  else
{
write-host " Disallow Digest authentication is set to an unknown setting" -ForegroundColor Red
}

$R3rxMaJTWuI8Ggn = Get-ItemProperty -Path  'Registry::HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Windows\WinRM\Service\'  -Name AllowBasic -ErrorAction SilentlyContinue|Select-Object -ExpandProperty AllowBasic
if ( $R3rxMaJTWuI8Ggn -eq $null)
{
write-host " Allow Basic authentication is not configured" -ForegroundColor Yellow
}
   elseif ( $R3rxMaJTWuI8Ggn  -eq  '0' )
{
write-host " Allow Basic authentication is disabled" -ForegroundColor Green
}
  elseif ( $R3rxMaJTWuI8Ggn  -eq  '1' )
{
write-host " Allow Basic authentication is enabled" -ForegroundColor Red
}
  else
{
write-host " Allow Basic authentication is set to an unknown setting" -ForegroundColor Red
}

$WeNYH9rskqIXnld = Get-ItemProperty -Path  'Registry::HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Windows\WinRM\Service\'  -Name AllowUnencryptedTraffic -ErrorAction SilentlyContinue|Select-Object -ExpandProperty AllowUnencryptedTraffic
if ( $WeNYH9rskqIXnld -eq $null)
{
write-host " Allow unencrypted traffic is not configured" -ForegroundColor Yellow
}
   elseif ( $WeNYH9rskqIXnld  -eq  '0' )
{
write-host " Allow unencrypted traffic is disabled" -ForegroundColor Green
}
  elseif ( $WeNYH9rskqIXnld  -eq  '1' )
{
write-host " Allow unencrypted traffic is enabled" -ForegroundColor Red
}
  else
{
write-host " Allow unencrypted traffic is set to an unknown setting" -ForegroundColor Red
}

$Gl0HpCP1daqYn28 = Get-ItemProperty -Path  'Registry::HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Windows\WinRM\Service\'  -Name DisableRunAs -ErrorAction SilentlyContinue|Select-Object -ExpandProperty DisableRunAs
if ( $Gl0HpCP1daqYn28 -eq $null)
{
write-host " Disallow WinRM from storing RunAs credentials is not configured" -ForegroundColor Yellow
}
   elseif ( $Gl0HpCP1daqYn28  -eq  '1' )
{
write-host " Disallow WinRM from storing RunAs credentials is enabled" -ForegroundColor Green
}
  elseif ( $Gl0HpCP1daqYn28  -eq  '0' )
{
write-host " Disallow WinRM from storing RunAs credentials is disabled" -ForegroundColor Red
}
  else
{
write-host " Disallow WinRM from storing RunAs credentials is set to an unknown setting" -ForegroundColor Red
}


write-host "`r`n####################### WINDOWS REMOTE SHELL ACCESS #######################`r`n"

$traYJW4x86uMjUG = Get-ItemProperty -Path  'Registry::HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Windows\WinRM\Service\WinRS\'  -Name AllowRemoteShellAccess -ErrorAction SilentlyContinue|Select-Object -ExpandProperty AllowRemoteShellAccess
if ( $traYJW4x86uMjUG -eq $null)
{
write-host " Allow Remote Shell Access is not configured" -ForegroundColor Yellow
}
   elseif ( $traYJW4x86uMjUG  -eq  '0' )
{
write-host " Allow Remote Shell Access is disabled" -ForegroundColor Green
}
  elseif ( $traYJW4x86uMjUG  -eq  '1' )
{
write-host " Allow Remote Shell Access is enabled" -ForegroundColor Red
}
  else
{
write-host " Allow Remote Shell Access is set to an unknown setting" -ForegroundColor Red
}

write-host "`r`n####################### WINDOWS SEARCH AND CORTANA #######################`r`n"

$nCf3tP6YSFhcpD0 = Get-ItemProperty -Path  'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\Windows\Windows Search\'  -Name AllowCortana -ErrorAction SilentlyContinue|Select-Object -ExpandProperty AllowCortana
if ( $nCf3tP6YSFhcpD0 -eq $null)
{
write-host " Allow Cortana is not configured" -ForegroundColor Yellow
}
   elseif ( $nCf3tP6YSFhcpD0  -eq  '0' )
{
write-host " Allow Cortana is disabled" -ForegroundColor Green
}
  elseif ( $nCf3tP6YSFhcpD0  -eq  '1' )
{
write-host " Allow Cortana is enabled" -ForegroundColor Red
}
  else
{
write-host " Allow Cortana is set to an unknown setting" -ForegroundColor Red
}

$zKbSDWr3cMvUZu7 = Get-ItemProperty -Path  'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\Windows\Windows Search\'  -Name ConnectedSearchUseWeb -ErrorAction SilentlyContinue|Select-Object -ExpandProperty ConnectedSearchUseWeb
if ( $zKbSDWr3cMvUZu7 -eq $null)
{
write-host " Don't search the web or display web results in Search is not configured" -ForegroundColor Yellow
}
   elseif ( $zKbSDWr3cMvUZu7  -eq  '0' )
{
write-host " Don't search the web or display web results in Search is enabled" -ForegroundColor Green
}
  elseif ( $zKbSDWr3cMvUZu7  -eq  '1' )
{
write-host " Don't search the web or display web results in Search is disabled" -ForegroundColor Red
}
  else
{
write-host " Don't search the web or display web results in Search is set to an unknown setting" -ForegroundColor Red
}

write-host "`r`n####################### WINDOWS TO GO #######################`r`n"

$rbWyQvlG5TAVoS7 = Get-ItemProperty -Path  'Registry::HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\PortableOperatingSystem\'  -Name Launcher -ErrorAction SilentlyContinue|Select-Object -ExpandProperty Launcher
if ( $rbWyQvlG5TAVoS7 -eq $null)
{
write-host " Windows To Go Default Startup Options is not configured" -ForegroundColor Yellow
}
   elseif ( $rbWyQvlG5TAVoS7  -eq  '0' )
{
write-host " Windows To Go Default Startup Options is disabled" -ForegroundColor Green
}
  elseif ( $rbWyQvlG5TAVoS7  -eq  '1' )
{
write-host " Windows To Go Default Startup Options is enabled" -ForegroundColor Red
}
  else
{
write-host " Windows To Go Default Startup Options is set to an unknown setting" -ForegroundColor Red
}

write-host "`r`n####################### DISPLAYING FILE EXTENSIONS #######################`r`n"

$rbWyQvlG5TAVoS = Get-ItemProperty -Path  'Registry::HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced'  -Name HideFileExt -ErrorAction SilentlyContinue|Select-Object -ExpandProperty HideFileExt
if ( $rbWyQvlG5TAVoS -eq $null)
{
write-host " Display file extensions is not configured" -ForegroundColor Yellow
}
   elseif ( $rbWyQvlG5TAVoS  -eq  '1' )
{
write-host "Display file extensions is enabled" -ForegroundColor Green
}
  elseif ( $rbWyQvlG5TAVoS  -eq  '1' )
{
write-host " Display file extensions is disabled" -ForegroundColor Red
}
  else
{
write-host " Display file extensions is set to an unknown setting" -ForegroundColor Red
}

write-host "`r`n####################### FILE AND FOLDER SECURITY PROPERTIES #######################`r`n"

$7DTmwyr9KIcjvMi = Get-ItemProperty -Path  'Registry::HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\'  -Name NoSecurityTab -ErrorAction SilentlyContinue|Select-Object -ExpandProperty NoSecurityTab
if ( $7DTmwyr9KIcjvMi -eq $null)
{
write-host " Remove Security tab is not configured" -ForegroundColor Yellow
}
   elseif ( $7DTmwyr9KIcjvMi  -eq  '1' )
{
write-host " Remove Security tab is enabled" -ForegroundColor Green
}
  elseif ( $7DTmwyr9KIcjvMi  -eq  '0' )
{
write-host " Remove Security tab is disabled" -ForegroundColor Red
}
  else
{
write-host " Remove Security tab is set to an unknown setting" -ForegroundColor Red
}

write-host "`r`n####################### LOCATION AWARENESS #######################`r`n"

$L0t3zDQOWT82Yjk = Get-ItemProperty -Path  'Registry::HKEY_CURRENT_USER\Software\Policies\Microsoft\Windows\LocationAndSensors\'  -Name DisableLocation -ErrorAction SilentlyContinue|Select-Object -ExpandProperty DisableLocation
if ( $L0t3zDQOWT82Yjk -eq $null)
{
write-host " Turn off location is not configured" -ForegroundColor Yellow
}
   elseif ( $L0t3zDQOWT82Yjk  -eq  '1' )
{
write-host " Turn off location is enabled" -ForegroundColor Green
}
  elseif ( $L0t3zDQOWT82Yjk  -eq  '0' )
{
write-host " Turn off location is disabled" -ForegroundColor Red
}
  else
{
write-host " Turn off location is set to an unknown setting" -ForegroundColor Red
}

$wOWZP5iF8Ah2HLn = Get-ItemProperty -Path  'Registry::HKEY_CURRENT_USER\Software\Policies\Microsoft\Windows\LocationAndSensors\'  -Name DisableLocationScripting -ErrorAction SilentlyContinue|Select-Object -ExpandProperty DisableLocationScripting
if ( $wOWZP5iF8Ah2HLn -eq $null)
{
write-host " Turn off location scripting is not configured" -ForegroundColor Yellow
}
   elseif ( $wOWZP5iF8Ah2HLn  -eq  '1' )
{
write-host " Turn off location scripting is enabled" -ForegroundColor Green
}
  elseif ( $wOWZP5iF8Ah2HLn  -eq  '0' )
{
write-host " Turn off location scripting is disabled" -ForegroundColor Red
}
  else
{
write-host " Turn off location scripting is set to an unknown setting" -ForegroundColor Red
}

$SbtA61CokgvnOKE = Get-ItemProperty -Path  'Registry::HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Windows\LocationAndSensors\'  -Name DisableWindowsLocationProvider -ErrorAction SilentlyContinue|Select-Object -ExpandProperty DisableWindowsLocationProvider
if ( $SbtA61CokgvnOKE -eq $null)
{
write-host " Turn off Windows Location Provider is not configured" -ForegroundColor Yellow
}
   elseif ( $SbtA61CokgvnOKE  -eq  '1' )
{
write-host " Turn off Windows Location Provider is enabled" -ForegroundColor Green
}
  elseif ( $SbtA61CokgvnOKE  -eq  '0' )
{
write-host " Turn off Windows Location Provider is disabled" -ForegroundColor Red
}
  else
{
write-host " Turn off Windows Location Provider is set to an unknown setting" -ForegroundColor Red
}


write-host "`r`n####################### MICROSOFT STORE #######################`r`n"

$64GduoTfcmp2iqY = Get-ItemProperty -Path  'Registry::HKEY_CURRENT_USER\Software\Policies\Microsoft\Windows\Explorer\'  -Name NoUseStoreOpenWith -ErrorAction SilentlyContinue|Select-Object -ExpandProperty NoUseStoreOpenWith
if ( $64GduoTfcmp2iqY -eq $null)
{
write-host " Turn off access to the Store is not configured" -ForegroundColor Yellow
}
   elseif ( $64GduoTfcmp2iqY  -eq  '1' )
{
write-host " Turn off access to the Store is enabled" -ForegroundColor Green
}
  elseif ( $64GduoTfcmp2iqY  -eq  '0' )
{
write-host " Turn off access to the Store is disabled" -ForegroundColor Red
}
  else
{
write-host " Turn off access to the Store is set to an unknown setting" -ForegroundColor Red
}

$2D3fnVsKR9pBEYm = Get-ItemProperty -Path  'Registry::HKEY_CURRENT_USER\Software\Policies\Microsoft\WindowsStore\'  -Name RemoveWindowsStore -ErrorAction SilentlyContinue|Select-Object -ExpandProperty RemoveWindowsStore
if ( $2D3fnVsKR9pBEYm -eq $null)
{
write-host " Turn off the Store application is not configured" -ForegroundColor Yellow
}
   elseif ( $2D3fnVsKR9pBEYm  -eq  '1' )
{
write-host " Turn off the Store application is enabled" -ForegroundColor Green
}
  elseif ( $2D3fnVsKR9pBEYm  -eq  '0' )
{
write-host " Turn off the Store application is disabled" -ForegroundColor Red
}
  else
{
write-host " Turn off the Store application is set to an unknown setting" -ForegroundColor Red
}


write-host "`r`n####################### PUBLISHING INFORMATION TO THE WEB #######################`r`n"


$8Ak7NpxH5Vs3bWE = Get-ItemProperty -Path  'Registry::HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\'  -Name NoWebServices -ErrorAction SilentlyContinue|Select-Object -ExpandProperty NoWebServices
if ( $8Ak7NpxH5Vs3bWE -eq $null)
{
write-host " Turn off Internet download for Web publishing and online ordering wizards is not configured" -ForegroundColor Yellow
}
   elseif ( $8Ak7NpxH5Vs3bWE  -eq  '1' )
{
write-host " Turn off Internet download for Web publishing and online ordering wizards is enabled" -ForegroundColor Green
}
  elseif ( $8Ak7NpxH5Vs3bWE  -eq  '0' )
{
write-host " Turn off Internet download for Web publishing and online ordering wizards is disabled" -ForegroundColor Red
}
  else
{
write-host " Turn off Internet download for Web publishing and online ordering wizards is set to an unknown setting" -ForegroundColor Red
}

write-host "`r`n####################### RESULTANT SET OF POLICY REPORTING #######################`r`n"

$dc04uCRS6vJGiNf = Get-ItemProperty -Path  'Registry::HKEY_CURRENT_USER\Software\Policies\Microsoft\Windows\System\'  -Name DenyRsopToInteractiveUser -ErrorAction SilentlyContinue|Select-Object -ExpandProperty DenyRsopToInteractiveUser
if ( $dc04uCRS6vJGiNf -eq $null)
{
write-host " Determine if interactive users can generate Resultant Set of Policy data is not configured" -ForegroundColor Yellow
}
   elseif ( $dc04uCRS6vJGiNf  -eq  '1' )
{
write-host " Determine if interactive users can generate Resultant Set of Policy data is enabled" -ForegroundColor Green
}
  elseif ( $dc04uCRS6vJGiNf  -eq  '0' )
{
write-host " Determine if interactive users can generate Resultant Set of Policy data is disabled" -ForegroundColor Red
}
  else
{
write-host " Determine if interactive users can generate Resultant Set of Policy data is set to an unknown setting" -ForegroundColor Red
}

Get-Variable -Exclude PWD,*Preference | Remove-Variable -EA 0
