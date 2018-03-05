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
write-host "Unable to check Number of Previous Logons to cache, due to setting being in the Security hive, look at setting in Computer Configuration\Policies\Windows Settings\Security Settings\Local Policies\Security Options\Interactive Logon"

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