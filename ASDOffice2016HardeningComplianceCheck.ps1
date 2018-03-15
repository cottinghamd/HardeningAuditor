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


$officeuserhive = Get-ChildItem -Path "Registry::HKCU\Software\Policies\Microsoft\Office\$officeversion\" -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Name 
$officelocalhive = Get-ChildItem -Path "Registry::HKLM\Software\Policies\Microsoft\Office\$officeversion\" -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Name 

if ($officeuserhive -eq $null -and $officelocalhive -eq $null)
{
write-host "No Microsoft Office group policies were detected, this script will now exit" -ForegroundColor Yellow
pause
break
}

write-host "`r`n####################### ATTACK SURFACE REDUCTION #######################`r`n"

write-host "`r`n####################### MACROS #######################`r`n"

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
    write-host "Macro settings have not been configured in $officename" -ForegroundColor Yellow
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
write-host "Macro settings have not been configured in Microsoft Outlook" -ForegroundColor Yellow
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

write-host "`r`n####################### PATCHING #######################`r`n"

write-host "`r`n####################### ACTIVE-X #######################`r`n"

$disableallactivex = Get-ItemProperty -Path "Registry::HKCU\software\policies\microsoft\office\common\security" -Name disableallactivex -ErrorAction SilentlyContinue|Select-Object -ExpandProperty disableallactivex

if ($disableallactivex -eq $null)
{
write-host "Disable All ActiveX is not configured" -ForegroundColor Yellow
}

elseif ($disableallactivex -eq '1')
{
write-host "Disable All ActiveX is enabled" -ForegroundColor Green
}
elseif ($disableallactivex -eq '0')
{
write-host "Disable All ActiveX is disabled" -ForegroundColor Red
}
else
{
write-host "Disable All ActiveX is configured to an unknown setting" -ForegroundColor Red
}


write-host "`r`n####################### ADD-INS #######################`r`n"

write-host "`r`n####################### EXTENSION HARDENING #######################`r`n"

$extensionhardening = Get-ItemProperty -Path "Registry::HKCU\software\policies\microsoft\office\$officeversion\excel\security" -Name extensionhardening -ErrorAction SilentlyContinue|Select-Object -ExpandProperty extensionhardening

if ($extensionhardening -eq $null)
{
write-host "Make hidden markup visible for Powerpoint is not configured" -ForegroundColor Yellow
}
elseif ($extensionhardening -eq '0')
{
write-host "Extension hardening for Excel is enabled, however it is set to Allow Different which is a non-compliant setting. The compliant setting is always match file type" -ForegroundColor Red
}
elseif ($extensionhardening -eq '1')
{
write-host "Extension hardening for Excel is enabled, however it is set to Allow Different, but warn which is a non-compliant setting. The compliant setting is always match file type" -ForegroundColor Red
}
elseif ($extensionhardening -eq '2')
{
write-host "Extension hardening for Excel is enabled and set to Always match file type" -ForegroundColor Green
}
else
{
write-host "Extension hardening for Excel is set to an unknown setting" -ForegroundColor Red
}


write-host "`r`n####################### FILE TYPE BLOCKING #######################`r`n"

$dbasefiles = Get-ItemProperty -Path "Registry::HKCU\software\policies\microsoft\office\$officeversion\excel\security\fileblock" -Name dbasefiles -ErrorAction SilentlyContinue|Select-Object -ExpandProperty dbasefiles

if ($dbasefiles -eq $null)
{
write-host "File Type Blocking for dBase III / IV files in Excel is not configured" -ForegroundColor Yellow
}
elseif ($dbasefiles -eq '0')
{
write-host "Do not block for dBase III / IV files in Excel is set to 'do not block'" -ForegroundColor Red
}
elseif ($dbasefiles -eq '2')
{
write-host "Do not block for dBase III / IV files in Excel is set to 'Open/Save blocked, use open policy'" -ForegroundColor Red
}
else
{
write-host "Do not block for dBase III / IV files in Excel is set to an unknown setting" -ForegroundColor Red
}

$difandsylkfiles = Get-ItemProperty -Path "Registry::HKCU\software\policies\microsoft\office\$officeversion\excel\security\fileblock" -Name difandsylkfiles -ErrorAction SilentlyContinue|Select-Object -ExpandProperty difandsylkfiles

if ($difandsylkfiles -eq $null)
{
write-host "File Type Blocking for Dif and Sylk files in Excel is not configured" -ForegroundColor Yellow
}
elseif ($difandsylkfiles -eq '0')
{
write-host "File Type Blocking for Dif and Sylk files in Excel is set to 'do not block'" -ForegroundColor Red
}
elseif ($difandsylkfiles -eq '1')
{
write-host "File Type Blocking for Dif and Sylk files in Excel is set to 'Save Blocked''" -ForegroundColor Red
}
elseif ($difandsylkfiles -eq '2')
{
write-host "File Type Blocking for Dif and Sylk files in Excel is set to 'Open/Save blocked, use open policy'" -ForegroundColor Red
}
else
{
write-host "File Type Blocking for Dif and Sylk files in Excel is set to an unknown setting" -ForegroundColor Red
}

$xl2macros = Get-ItemProperty -Path "Registry::HKCU\software\policies\microsoft\office\$officeversion\excel\security\fileblock" -Name xl2macros -ErrorAction SilentlyContinue|Select-Object -ExpandProperty xl2macros

if ($xl2macros -eq $null)
{
write-host "File Type Blocking for Excel 2 macrosheets and add-in files is not configured" -ForegroundColor Yellow
}
elseif ($xl2macros -eq '0')
{
write-host "File Type Blocking for Excel 2 macrosheets and add-in files is set to 'do not block'" -ForegroundColor Red
}
elseif ($xl2macros -eq '1')
{
write-host "File Type Blocking for Excel 2 macrosheets and add-in files is set to 'Save Blocked''" -ForegroundColor Red
}
elseif ($xl2macros -eq '2')
{
write-host "File Type Blocking for Excel 2 macrosheets and add-in files is set to 'Open/Save blocked, use open policy'" -ForegroundColor Red
}
elseif ($xl2macros -eq '3')
{
write-host "File Type Blocking for Excel 2 macrosheets and add-in files is set to 'Block'" -ForegroundColor Green
}
elseif ($xl2macros -eq '4')
{
write-host "File Type Blocking for Excel 2 macrosheets and add-in files is set to 'Open in Protected View'" -ForegroundColor Red
}
elseif ($xl2macros -eq '5')
{
write-host "File Type Blocking for Excel 2 macrosheets and add-in files is set to 'Allow editing and open in Protected View'" -ForegroundColor Red
}
else
{
write-host "File Type Blocking for Excel 2 macrosheets and add-in files is set to an unknown setting" -ForegroundColor Red
}

$xl2worksheets = Get-ItemProperty -Path "Registry::HKCU\software\policies\microsoft\office\$officeversion\excel\security\fileblock" -Name xl2worksheets -ErrorAction SilentlyContinue|Select-Object -ExpandProperty xl2worksheets

if ($xl2worksheets -eq $null)
{
write-host "File Type Blocking for Excel 2 worksheets in Excel is not configured" -ForegroundColor Yellow
}
elseif ($xl2worksheets -eq '0')
{
write-host "File Type Blocking for Excel 2 worksheets is set to 'do not block'" -ForegroundColor Red
}
elseif ($xl2worksheets -eq '1')
{
write-host "File Type Blocking for Excel 2 worksheets is set to 'Save Blocked''" -ForegroundColor Red
}
elseif ($xl2worksheets -eq '2')
{
write-host "File Type Blocking for Excel 2 worksheets is set to 'Open/Save blocked, use open policy'" -ForegroundColor Red
}
elseif ($xl2worksheets -eq '3')
{
write-host "File Type Blocking for Excel 2 worksheets is set to 'Block'" -ForegroundColor Green
}
elseif ($xl2worksheets -eq '4')
{
write-host "File Type Blocking for Excel 2 worksheets is set to 'Open in Protected View'" -ForegroundColor Red
}
elseif ($xl2worksheets -eq '5')
{
write-host "File Type Blocking for Excel 2 worksheets is set to 'Allow editing and open in Protected View'" -ForegroundColor Red
}
else
{
write-host "File Type Blocking for Excel 2 worksheets is set to an unknown setting" -ForegroundColor Red
}

$xlamfiles = Get-ItemProperty -Path "Registry::HKCU\software\policies\microsoft\office\$officeversion\excel\security\fileblock" -Name xlamfiles -ErrorAction SilentlyContinue|Select-Object -ExpandProperty xlamfiles

if ($xlamfiles -eq $null)
{
write-host "File Type Blocking for Excel 2007 and later add-in files is not configured" -ForegroundColor Yellow
}
elseif ($xlamfiles -eq '0')
{
write-host "File Type Blocking for Excel 2007 and later add-in files is set to 'do not block'" -ForegroundColor Red
}
elseif ($xlamfiles -eq '1')
{
write-host "File Type Blocking for Excel 2007 and later add-in files is set to 'Save Blocked''" -ForegroundColor Red
}
elseif ($xlamfiles -eq '2')
{
write-host "File Type Blocking for Excel 2007 and later add-in files is set to 'Open/Save blocked, use open policy'" -ForegroundColor Red
}
else
{
write-host "File Type Blocking for Excel 2007 and later add-in files is set to an unknown setting" -ForegroundColor Red
}

$xlsbfiles = Get-ItemProperty -Path "Registry::HKCU\software\policies\microsoft\office\$officeversion\excel\security\fileblock" -Name xlsbfiles -ErrorAction SilentlyContinue|Select-Object -ExpandProperty xlsbfiles

if ($xlsbfiles -eq $null)
{
write-host "File Type Blocking for Excel 2007 and later binary workbooks is not configured" -ForegroundColor Yellow
}
elseif ($xlsbfiles -eq '0')
{
write-host "File Type Blocking for Excel 2007 and later binary workbooks is set to 'do not block'" -ForegroundColor Red
}
elseif ($xlsbfiles -eq '1')
{
write-host "File Type Blocking for Excel 2007 and later binary workbooks is set to 'Save Blocked''" -ForegroundColor Red
}
elseif ($xlsbfiles -eq '2')
{
write-host "File Type Blocking for Excel 2007 and later binary workbooks is set to 'Open/Save blocked, use open policy'" -ForegroundColor Red
}
elseif ($xlsbfiles -eq '3')
{
write-host "File Type Blocking for Excel 2007 and later binary workbooks is set to 'Block'" -ForegroundColor Green
}
elseif ($xlsbfiles -eq '4')
{
write-host "File Type Blocking for Excel 2007 and later binary workbooks is set to 'Open in Protected View'" -ForegroundColor Red
}
elseif ($xlsbfiles -eq '5')
{
write-host "File Type Blocking for Excel 2007 and later binary workbooks is set to 'Allow editing and open in Protected View'" -ForegroundColor Red
}
else
{
write-host "File Type Blocking for Excel 2007 and later binary workbooks is set to an unknown setting" -ForegroundColor Red
}

$xl3macros = Get-ItemProperty -Path "Registry::HKCU\software\policies\microsoft\office\$officeversion\excel\security\fileblock" -Name xl3macros -ErrorAction SilentlyContinue|Select-Object -ExpandProperty xl3macros

if ($xl3macros -eq $null)
{
write-host "File Type Blocking for Excel 3 macrosheets and add-in files workbooks in Excel is not configured" -ForegroundColor Yellow
}
elseif ($xl3macros -eq '0')
{
write-host "File Type Blocking for Excel 3 macrosheets and add-in files is set to 'do not block'" -ForegroundColor Red
}
elseif ($xl3macros -eq '1')
{
write-host "File Type Blocking for Excel 3 macrosheets and add-in files is set to 'Save Blocked''" -ForegroundColor Red
}
elseif ($xl3macros -eq '2')
{
write-host "File Type Blocking for Excel 3 macrosheets and add-in files is set to 'Open/Save blocked, use open policy'" -ForegroundColor Red
}
elseif ($xl3macros -eq '3')
{
write-host "File Type Blocking for Excel 3 macrosheets and add-in files is set to 'Block'" -ForegroundColor Green
}
elseif ($xl3macros -eq '4')
{
write-host "File Type Blocking for Excel 3 macrosheets and add-in files is set to 'Open in Protected View'" -ForegroundColor Red
}
elseif ($xl3macros -eq '5')
{
write-host "File Type Blocking for Excel 3 macrosheets and add-in files is set to 'Allow editing and open in Protected View'" -ForegroundColor Red
}
else
{
write-host "File Type Blocking for Excel 3 macrosheets and add-in files is set to an unknown setting" -ForegroundColor Red
}

$xl3worksheets = Get-ItemProperty -Path "Registry::HKCU\software\policies\microsoft\office\$officeversion\excel\security\fileblock" -Name xl3worksheets -ErrorAction SilentlyContinue|Select-Object -ExpandProperty xl3worksheets

if ($xl3worksheets -eq $null)
{
write-host "File Type Blocking for Excel 3 worksheets is not configured" -ForegroundColor Yellow
}
elseif ($xl3worksheets -eq '0')
{
write-host "File Type Blocking for Excel 3 worksheets is set to 'do not block'" -ForegroundColor Red
}
elseif ($xl3worksheets -eq '1')
{
write-host "File Type Blocking for Excel 3 worksheets is set to 'Save Blocked''" -ForegroundColor Red
}
elseif ($xl3worksheets -eq '2')
{
write-host "File Type Blocking for Excel 3 worksheets is set to 'Open/Save blocked, use open policy'" -ForegroundColor Red
}
elseif ($xl3worksheets -eq '3')
{
write-host "File Type Blocking for Excel 3 worksheets is set to 'Block'" -ForegroundColor Green
}
elseif ($xl3worksheets -eq '4')
{
write-host "File Type Blocking for Excel 3 worksheets is set to 'Open in Protected View'" -ForegroundColor Red
}
elseif ($xl3worksheets -eq '5')
{
write-host "File Type Blocking for Excel 3 worksheets is set to 'Allow editing and open in Protected View'" -ForegroundColor Red
}
else
{
write-host "File Type Blocking for Excel 3 worksheets is set to an unknown setting" -ForegroundColor Red
}

$xl4workbooks = Get-ItemProperty -Path "Registry::HKCU\software\policies\microsoft\office\$officeversion\excel\security\fileblock" -Name xl4workbooks -ErrorAction SilentlyContinue|Select-Object -ExpandProperty xl4workbooks

if ($xl4workbooks -eq $null)
{
write-host "File Type Blocking for Excel 4 workbooks is not configured" -ForegroundColor Yellow
}
elseif ($xl4workbooks -eq '0')
{
write-host "File Type Blocking for Excel 4 workbooks is set to 'do not block'" -ForegroundColor Red
}
elseif ($xl4workbooks -eq '1')
{
write-host "File Type Blocking for Excel 4 workbooks is set to 'Save Blocked''" -ForegroundColor Red
}
elseif ($xl4workbooks -eq '2')
{
write-host "File Type Blocking for Excel 4 workbooks is set to 'Open/Save blocked, use open policy'" -ForegroundColor Red
}
elseif ($xl4workbooks -eq '3')
{
write-host "File Type Blocking for Excel 4 workbooks is set to 'Block'" -ForegroundColor Green
}
elseif ($xl4workbooks -eq '4')
{
write-host "File Type Blocking for Excel 4 workbooks is set to 'Open in Protected View'" -ForegroundColor Red
}
elseif ($xl4workbooks -eq '5')
{
write-host "File Type Blocking for Excel 4 workbooks is set to 'Allow editing and open in Protected View'" -ForegroundColor Red
}
else
{
write-host "File Type Blocking for Excel 3 worksheets in Excel is set to an unknown setting" -ForegroundColor Red
}


Excel 4 workbooks	HKCU\software\policies\microsoft\office\16.0\excel\security\fileblock!	[0, Do not block] [2, Open/Save blocked, use open policy] [3, Block] [4, Open in Protected View] [5, Allow editing and open in Protected View]


Set default file block behavior	HKCU\software\policies\microsoft\office\16.0\excel\security\fileblock!openinprotectedview	[0, Blocked files are not opened] [1, Blocked files open in Protected View and can not be edited] [2, Blocked files open in Protected View and can be edited]
Excel 2007 and later workbooks and templates	HKCU\software\policies\microsoft\office\16.0\excel\security\fileblock!xlsxandxltxfiles	[0, Do not block] [1, Save blocked] [2, Open/Save blocked, use open policy] [3, Block] [4, Open in Protected View] [5, Allow editing and open in Protected View]
Excel 2007 and later macro-enabled workbooks and templates	HKCU\software\policies\microsoft\office\16.0\excel\security\fileblock!xlsmandxltmfiles	[0, Do not block] [1, Save blocked] [2, Open/Save blocked, use open policy] [3, Block] [4, Open in Protected View] [5, Allow editing and open in Protected View]


OpenDocument Spreadsheet files	HKCU\software\policies\microsoft\office\16.0\excel\security\fileblock!opendocumentspreadsheet	[0, Do not block] [1, Save blocked] [2, Open/Save blocked, use open policy] [3, Block] [4, Open in Protected View] [5, Allow editing and open in Protected View]
Excel 97-2003 add-in files	HKCU\software\policies\microsoft\office\16.0\excel\security\fileblock!xl97addins	[0, Do not block] [1, Save blocked] [2, Open/Save blocked, use open policy]
Excel 97-2003 workbooks and templates	HKCU\software\policies\microsoft\office\16.0\excel\security\fileblock!xl97workbooksandtemplates	[0, Do not block] [1, Save blocked] [2, Open/Save blocked, use open policy] [3, Block] [4, Open in Protected View] [5, Allow editing and open in Protected View]
Excel 95-97 workbooks and templates	HKCU\software\policies\microsoft\office\16.0\excel\security\fileblock!xl9597workbooksandtemplates	[0, Do not block] [2, Open/Save blocked, use open policy] [3, Block] [4, Open in Protected View] [5, Allow editing and open in Protected View]
Excel 95 workbooks	HKCU\software\policies\microsoft\office\16.0\excel\security\fileblock!xl95workbooks	[0, Do not block] [1, Save blocked] [2, Open/Save blocked, use open policy] [3, Block] [4, Open in Protected View] [5, Allow editing and open in Protected View]

Excel 4 worksheets	HKCU\software\policies\microsoft\office\16.0\excel\security\fileblock!xl4worksheets	[0, Do not block] [2, Open/Save blocked, use open policy] [3, Block] [4, Open in Protected View] [5, Allow editing and open in Protected View]

Excel 4 macrosheets and add-in files	HKCU\software\policies\microsoft\office\16.0\excel\security\fileblock!xl4macros	[0, Do not block] [2, Open/Save blocked, use open policy] [3, Block] [4, Open in Protected View] [5, Allow editing and open in Protected View]

Web pages and Excel 2003 XML spreadsheets	HKCU\software\policies\microsoft\office\16.0\excel\security\fileblock!htmlandxmlssfiles	[0, Do not block] [1, Save blocked] [2, Open/Save blocked, use open policy] [3, Block] [4, Open in Protected View] [5, Allow editing and open in Protected View]
XML files	HKCU\software\policies\microsoft\office\16.0\excel\security\fileblock!xmlfiles	[0, Do not block] [1, Save blocked] [2, Open/Save blocked, use open policy]
Text files	HKCU\software\policies\microsoft\office\16.0\excel\security\fileblock!textfiles	[0, Do not block] [1, Save blocked] [2, Open/Save blocked, use open policy]
Excel add-in files	HKCU\software\policies\microsoft\office\16.0\excel\security\fileblock!xllfiles	[0, Do not block] [2, Open/Save blocked, use open policy]

Microsoft Office query files	HKCU\software\policies\microsoft\office\16.0\excel\security\fileblock!officequeries	[0, Do not block] [1, Save blocked] [2, Open/Save blocked, use open policy] [3, Block] [4, Open in Protected View] [5, Allow editing and open in Protected View]
Microsoft Office data connection files	HKCU\software\policies\microsoft\office\16.0\excel\security\fileblock!officedataconnections	[0, Do not block] [2, Open/Save blocked, use open policy]
Other data source files	HKCU\software\policies\microsoft\office\16.0\excel\security\fileblock!officedatasources	[0, Do not block] [2, Open/Save blocked, use open policy]
Offline cube files	HKCU\software\policies\microsoft\office\16.0\excel\security\fileblock!offlinecubefiles	[0, Do not block] [2, Open/Save blocked, use open policy]

Legacy converters for Excel	HKCU\software\policies\microsoft\office\16.0\excel\security\fileblock!converters	[0, Do not block] [2, Open/Save blocked, use open policy] [3, Block] [4, Open in Protected View] [5, Allow editing and open in Protected View]
Microsoft Office Open XML converters for Excel	HKCU\software\policies\microsoft\office\16.0\excel\security\fileblock!ooxmlconverters	[0, Do not block] [1, Save blocked] [2, Open/Save blocked, use open policy] [3, Block] [4, Open in Protected View] [5, Allow editing and open in Protected View]








write-host "`r`n####################### HIDDEN MARKUP #######################`r`n"  -ForegroundColor Cyan

#Powerpoint - Make Hidden Markup Visible

$hiddenmarkupppt = Get-ItemProperty -Path "Registry::HKEY_CURRENT_USER\software\policies\microsoft\office\$officeversion\powerpoint\options" -Name markupopensave -ErrorAction SilentlyContinue|Select-Object -ExpandProperty markupopensave

if ($hiddenmarkupppt -eq $null)
{
write-host "Make hidden markup visible for Powerpoint is not configured" -ForegroundColor Yellow
}

elseif ($hiddenmarkupppt -eq '1')
{
write-host "Make hidden markup visible for Powerpoint is enabled" -ForegroundColor Green
}
else
{
write-host "Make hidden markup visible for Powerpoint is disabled" -ForegroundColor Red
}


#Word - Make Hidden Markup Visible

$hiddenmarkupword = Get-ItemProperty -Path "Registry::HKEY_CURRENT_USER\software\policies\microsoft\office\$officeversion\word\options" -Name showmarkupopensave -ErrorAction SilentlyContinue|Select-Object -ExpandProperty showmarkupopensave

if ($hiddenmarkupword -eq $null)
{
write-host "Make hidden markup visible for Word is not configured" -ForegroundColor Yellow
}

elseif ($hiddenmarkupword -eq '1')
{
write-host "Make hidden markup visible for Word is enabled" -ForegroundColor Green
}
else
{
write-host "Make hidden markup visible for Word is disabled" -ForegroundColor Red
}


write-host "`r`n####################### OFFICE FILE VALIDATION #######################`r`n"  -ForegroundColor Cyan

#Turn off error reporting for files that fail file validation

$disablereporting = Get-ItemProperty -Path "Registry::HKEY_CURRENT_USER\software\policies\microsoft\office\$officeversion\common\security\filevalidation" -Name disablereporting -ErrorAction SilentlyContinue|Select-Object -ExpandProperty disablereporting

if ($disablereporting -eq $null)
{
write-host "Turn off error reporting for files that fail file validation is not configured" -ForegroundColor Yellow
}

elseif ($disablereporting -eq '1')
{
write-host "Turn off error reporting for files that fail file validation is enabled" -ForegroundColor Green
}
else
{
write-host "Turn off error reporting for files that fail file validation is disabled" -ForegroundColor Red
}


#Turn off file validation - excel

$filevalidationexcel = Get-ItemProperty -Path "Registry::HKEY_CURRENT_USER\software\policies\microsoft\office\$officeversion\excel\security\filevalidation" -Name enableonload -ErrorAction SilentlyContinue|Select-Object -ExpandProperty enableonload

if ($filevalidationexcel -eq $null)
{
write-host "Turn off file validation is not configured in Excel" -ForegroundColor Yellow
}

elseif ($filevalidationexcel -eq '1')
{
write-host "Turn off file validation is disabled in Excel" -ForegroundColor Green
}
else
{
write-host "Turn off file validation is enabled in Excel" -ForegroundColor Red
}


#Turn off file validation - Powerpoint

$filevalidationppt = Get-ItemProperty -Path "Registry::HKEY_CURRENT_USER\software\policies\microsoft\office\$officeversion\powerpoint\security\filevalidation" -Name enableonload -ErrorAction SilentlyContinue|Select-Object -ExpandProperty enableonload

if ($filevalidationppt -eq $null)
{
write-host "Turn off file validation is not configured in Powepoint" -ForegroundColor Yellow
}

elseif ($filevalidationppt -eq '1')
{
write-host "Turn off file validation is disabled in Powepoint" -ForegroundColor Green
}
else
{
write-host "Turn off file validation is enabled in Powepoint" -ForegroundColor Red
}

#Turn off file validation - Word

$filevalidationword = Get-ItemProperty -Path "Registry::HKEY_CURRENT_USER\software\policies\microsoft\office\$officeversion\word\security\filevalidation" -Name enableonload -ErrorAction SilentlyContinue|Select-Object -ExpandProperty enableonload

if ($filevalidationword -eq $null)
{
write-host "Turn off file validation is not configured in Word" -ForegroundColor Yellow
}

elseif ($filevalidationppt -eq '1')
{
write-host "Turn off file validation is disabled in Word" -ForegroundColor Green
}
else
{
write-host "Turn off file validation is enabled in Word" -ForegroundColor Red
}


write-host "`r`n####################### PROTECTED VIEW #######################`r`n"  -ForegroundColor Cyan

#Do not open files from the Internet zone in Protected View - Excel

$disableifexcel = Get-ItemProperty -Path "Registry::HKEY_CURRENT_USER\software\policies\microsoft\office\$officeversion\excel\security\protectedview" -Name disableinternetfilesinpv -ErrorAction SilentlyContinue|Select-Object -ExpandProperty disableinternetfilesinpv

if ($disableifexcel -eq $null)
{
write-host "Do not open files from the Internet zone in Protected View is not configured in Excel" -ForegroundColor Yellow
}

elseif ($disableifexcel -eq '0')
{
write-host "Do not open files from the Internet zone in Protected View is disabled in Excel" -ForegroundColor Green
}
elseif ($disableifexcel -eq '1')
{
write-host "Do not open files from the Internet zone in Protected View is enabled in Excel" -ForegroundColor Red
}
else
{
write-host "Do not open files from the Internet zone in Protected View is set to an unknown configuration in Excel" -ForegroundColor Red
}



#Do not open files from the Internet zone in Protected View - Powerpoint

$disableifpowerpoint = Get-ItemProperty -Path "Registry::HKEY_CURRENT_USER\software\policies\microsoft\office\$officeversion\powerpoint\security\protectedview" -Name disableinternetfilesinpv -ErrorAction SilentlyContinue|Select-Object -ExpandProperty disableinternetfilesinpv

if ($disableifpowerpoint -eq $null)
{
write-host "Do not open files from the Internet zone in Protected View is not configured in powerpoint" -ForegroundColor Yellow
}

elseif ($disableifpowerpoint -eq '0')
{
write-host "Do not open files from the Internet zone in Protected View is disabled in powerpoint" -ForegroundColor Green
}
elseif ($disableifpowerpoint -eq '1')
{
write-host "Do not open files from the Internet zone in Protected View is enabled in powerpoint" -ForegroundColor Red
}
else
{
write-host "Do not open files from the Internet zone in Protected View is set to an unknown configuration in powerpoint" -ForegroundColor Red
}


#Do not open files from the Internet zone in Protected View - word

$disableifword = Get-ItemProperty -Path "Registry::HKEY_CURRENT_USER\software\policies\microsoft\office\$officeversion\word\security\protectedview" -Name disableinternetfilesinpv -ErrorAction SilentlyContinue|Select-Object -ExpandProperty disableinternetfilesinpv

if ($disableifword -eq $null)
{
write-host "Do not open files from the Internet zone in Protected View is not configured in word" -ForegroundColor Yellow
}

elseif ($disableifword -eq '0')
{
write-host "Do not open files from the Internet zone in Protected View is disabled in word" -ForegroundColor Green
}
elseif ($disableifword -eq '1')
{
write-host "Do not open files from the Internet zone in Protected View is enabled in word" -ForegroundColor Red
}
else
{
write-host "Do not open files from the Internet zone in Protected View is set to an unknown configuration in word" -ForegroundColor Red
}


#Do not open files in unsafe locations in Protected View - Excel

$disableifulexcel = Get-ItemProperty -Path "Registry::HKEY_CURRENT_USER\software\policies\microsoft\office\$officeversion\excel\security\protectedview" -Name disableunsafelocationsinpv -ErrorAction SilentlyContinue|Select-Object -ExpandProperty disableunsafelocationsinpv

if ($disableifulexcel -eq $null)
{
write-host "Do not open files in unsafe locations in Protected View is not configured in Excel" -ForegroundColor Yellow
}

elseif ($disableifulexcel -eq '0')
{
write-host "Do not open files in unsafe locations in Protected View is disabled in Excel" -ForegroundColor Green
}
elseif ($disableifulexcel -eq '1')
{
write-host "Do not open files in unsafe locations in Protected View is enabled in Excel" -ForegroundColor Red
}
else
{
write-host "Do not open files in unsafe locations in Protected View is set to an unknown configuration in Excel" -ForegroundColor Red
}



#Do not open files in unsafe locations in Protected View - Powerpoint

$disableifulpowerpoint = Get-ItemProperty -Path "Registry::HKEY_CURRENT_USER\software\policies\microsoft\office\$officeversion\powerpoint\security\protectedview" -Name disableunsafelocationsinpv -ErrorAction SilentlyContinue|Select-Object -ExpandProperty disableunsafelocationsinpv

if ($disableifulpowerpoint -eq $null)
{
write-host "Do not open files in unsafe locations in Protected View is not configured in powerpoint" -ForegroundColor Yellow
}

elseif ($disableifulpowerpoint -eq '0')
{
write-host "Do not open files in unsafe locations in Protected View is disabled in powerpoint" -ForegroundColor Green
}
elseif ($disableifulpowerpoint -eq '1')
{
write-host "Do not open files in unsafe locations in Protected View is enabled in powerpoint" -ForegroundColor Red
}
else
{
write-host "Do not open files in unsafe locations in Protected View is set to an unknown configuration in powerpoint" -ForegroundColor Red
}


#Do not open files in unsafe locations in Protected View - word

$disableifulword = Get-ItemProperty -Path "Registry::HKEY_CURRENT_USER\software\policies\microsoft\office\$officeversion\word\security\protectedview" -Name disableunsafelocationsinpv -ErrorAction SilentlyContinue|Select-Object -ExpandProperty disableunsafelocationsinpv

if ($disableifulword -eq $null)
{
write-host "Do not open files in unsafe locations in Protected View is not configured in word" -ForegroundColor Yellow
}

elseif ($disableifulword -eq '0')
{
write-host "Do not open files in unsafe locations in Protected View is disabled in word" -ForegroundColor Green
}
elseif ($disableifulword -eq '1')
{
write-host "Do not open files in unsafe locations in Protected View is enabled in word" -ForegroundColor Red
}
else
{
write-host "Do not open files in unsafe locations in Protected View is set to an unknown configuration in word" -ForegroundColor Red
}


#Turn off Protected View for attachments opened from Outlook - Excel

$disableattachmentsexcel = Get-ItemProperty -Path "Registry::HKEY_CURRENT_USER\software\policies\microsoft\office\$officeversion\excel\security\protectedview" -Name disableattachmentsinpv -ErrorAction SilentlyContinue|Select-Object -ExpandProperty disableattachmentsinpv

if ($disableattachmentsexcel -eq $null)
{
write-host "Turn off Protected View for attachments opened from Outlook is not configured in Excel" -ForegroundColor Yellow
}

elseif ($disableattachmentsexcel -eq '0')
{
write-host "Turn off Protected View for attachments opened from Outlook is disabled in Excel" -ForegroundColor Green
}
elseif ($disableattachmentsexcel -eq '1')
{
write-host "Turn off Protected View for attachments opened from Outlook is enabled in Excel" -ForegroundColor Red
}
else
{
write-host "Turn off Protected View for attachments opened from Outlook is set to an unknown configuration in Excel" -ForegroundColor Red
}



#Turn off Protected View for attachments opened from Outlook - Powerpoint

$disableattachmentspowerpoint = Get-ItemProperty -Path "Registry::HKEY_CURRENT_USER\software\policies\microsoft\office\$officeversion\powerpoint\security\protectedview" -Name disableattachmentsinpv -ErrorAction SilentlyContinue|Select-Object -ExpandProperty disableattachmentsinpv

if ($disableattachmentspowerpoint -eq $null)
{
write-host "Turn off Protected View for attachments opened from Outlook is not configured in powerpoint" -ForegroundColor Yellow
}

elseif ($disableattachmentspowerpoint -eq '0')
{
write-host "Turn off Protected View for attachments opened from Outlook is disabled in powerpoint" -ForegroundColor Green
}
elseif ($disableattachmentspowerpoint -eq '1')
{
write-host "Turn off Protected View for attachments opened from Outlook is enabled in powerpoint" -ForegroundColor Red
}
else
{
write-host "Turn off Protected View for attachments opened from Outlook is set to an unknown configuration in powerpoint" -ForegroundColor Red
}


#Turn off Protected View for attachments opened from Outlook - word

$disableattachmentsword = Get-ItemProperty -Path "Registry::HKEY_CURRENT_USER\software\policies\microsoft\office\$officeversion\word\security\protectedview" -Name disableattachmentsinpv -ErrorAction SilentlyContinue|Select-Object -ExpandProperty disableattachmentsinpv

if ($disableattachmentsword -eq $null)
{
write-host "Turn off Protected View for attachments opened from Outlook is not configured in word" -ForegroundColor Yellow
}

elseif ($disableattachmentsword -eq '0')
{
write-host "Turn off Protected View for attachments opened from Outlook is disabled in word" -ForegroundColor Green
}
elseif ($disableattachmentsword -eq '1')
{
write-host "Turn off Protected View for attachments opened from Outlook is enabled in word" -ForegroundColor Red
}
else
{
write-host "Turn off Protected View for attachments opened from Outlook is set to an unknown configuration in word" -ForegroundColor Red
}

write-host "`r`n####################### TRUSTED DOCUMENTS #######################`r`n"  -ForegroundColor Cyan

#Turn off trusted documents - Excel

$trusteddocsexcel = Get-ItemProperty -Path "Registry::HKEY_CURRENT_USER\software\policies\microsoft\office\$officeversion\excel\security\trusted documents" -Name disabletrusteddocuments -ErrorAction SilentlyContinue|Select-Object -ExpandProperty disabletrusteddocuments

if ($trusteddocsexcel -eq $null)
{
write-host "Turn off trusted documents is not configured in Excel" -ForegroundColor Yellow
}

elseif ($trusteddocsexcel -eq '1')
{
write-host "Turn off trusted documents is enabled in Excel" -ForegroundColor Green
}
elseif ($trusteddocsexcel -eq '0')
{
write-host "Turn off trusted documents is disabled in Excel" -ForegroundColor Red
}
else
{
write-host "Turn off trusted documents is set to an unknown configuration in Excel" -ForegroundColor Red
}



#Turn off trusted documents - Powerpoint

$trusteddocspowerpoint = Get-ItemProperty -Path "Registry::HKEY_CURRENT_USER\software\policies\microsoft\office\$officeversion\powerpoint\security\trusted documents" -Name disabletrusteddocuments -ErrorAction SilentlyContinue|Select-Object -ExpandProperty disabletrusteddocuments

if ($trusteddocspowerpoint -eq $null)
{
write-host "Turn off trusted documents is not configured in powerpoint" -ForegroundColor Yellow
}

elseif ($trusteddocspowerpoint -eq '1')
{
write-host "Turn off trusted documents is enabled in powerpoint" -ForegroundColor Green
}
elseif ($trusteddocspowerpoint -eq '0')
{
write-host "Turn off trusted documents is disabled in powerpoint" -ForegroundColor Red
}
else
{
write-host "Turn off trusted documents is set to an unknown configuration in powerpoint" -ForegroundColor Red
}


#Turn off trusted documents - word

$trusteddocsword = Get-ItemProperty -Path "Registry::HKEY_CURRENT_USER\software\policies\microsoft\office\$officeversion\word\security\trusted documents" -Name disabletrusteddocuments -ErrorAction SilentlyContinue|Select-Object -ExpandProperty disabletrusteddocuments

if ($trusteddocsword -eq $null)
{
write-host "Turn off trusted documents is not configured in word" -ForegroundColor Yellow
}

elseif ($trusteddocsword -eq '1')
{
write-host "Turn off trusted documents is enabled in word" -ForegroundColor Green
}
elseif ($trusteddocsword -eq '0')
{
write-host "Turn off trusted documents is disabled in word" -ForegroundColor Red
}
else
{
write-host "Turn off trusted documents is set to an unknown configuration in word" -ForegroundColor Red
}


#Turn off Trusted Documents on the network - Excel

$trusteddocsnetworkexcel = Get-ItemProperty -Path "Registry::HKEY_CURRENT_USER\software\policies\microsoft\office\$officeversion\excel\security\trusted documents" -Name disablenetworktrusteddocuments -ErrorAction SilentlyContinue|Select-Object -ExpandProperty disablenetworktrusteddocuments

if ($trusteddocsnetworkexcel -eq $null)
{
write-host "Turn off Trusted Documents on the network is not configured in Excel" -ForegroundColor Yellow
}

elseif ($trusteddocsnetworkexcel -eq '1')
{
write-host "Turn off Trusted Documents on the network is enabled in Excel" -ForegroundColor Green
}
elseif ($trusteddocsnetworkexcel -eq '0')
{
write-host "Turn off Trusted Documents on the network is disabled in Excel" -ForegroundColor Red
}
else
{
write-host "Turn off Trusted Documents on the network is set to an unknown configuration in Excel" -ForegroundColor Red
}



#Turn off Trusted Documents on the network - Powerpoint

$trusteddocsnetworkpowerpoint = Get-ItemProperty -Path "Registry::HKEY_CURRENT_USER\software\policies\microsoft\office\$officeversion\powerpoint\security\trusted documents" -Name disablenetworktrusteddocuments -ErrorAction SilentlyContinue|Select-Object -ExpandProperty disablenetworktrusteddocuments

if ($trusteddocsnetworkpowerpoint -eq $null)
{
write-host "Turn off Trusted Documents on the network is not configured in powerpoint" -ForegroundColor Yellow
}

elseif ($trusteddocsnetworkpowerpoint -eq '1')
{
write-host "Turn off Trusted Documents on the network is enabled in powerpoint" -ForegroundColor Green
}
elseif ($trusteddocsnetworkpowerpoint -eq '0')
{
write-host "Turn off Trusted Documents on the network is disabled in powerpoint" -ForegroundColor Red
}
else
{
write-host "Turn off Trusted Documents on the network is set to an unknown configuration in powerpoint" -ForegroundColor Red
}


#Turn off Trusted Documents on the network - word

$trusteddocsnetworkword = Get-ItemProperty -Path "Registry::HKEY_CURRENT_USER\software\policies\microsoft\office\$officeversion\word\security\trusted documents" -Name disablenetworktrusteddocuments -ErrorAction SilentlyContinue|Select-Object -ExpandProperty disablenetworktrusteddocuments

if ($trusteddocsnetworkword -eq $null)
{
write-host "Turn off Trusted Documents on the network is not configured in word" -ForegroundColor Yellow
}

elseif ($trusteddocsnetworkword -eq '1')
{
write-host "Turn off Trusted Documents on the network is enabled in word" -ForegroundColor Green
}
elseif ($trusteddocsnetworkword -eq '0')
{
write-host "Turn off Trusted Documents on the network is disabled in word" -ForegroundColor Red
}
else
{
write-host "Turn off Trusted Documents on the network is set to an unknown configuration in word" -ForegroundColor Red
}

write-host "`r`n####################### REPORTING INFORMATION #######################`r`n" -ForegroundColor Cyan


#Allow including screenshot with Office Feedback

$includescreenshot = Get-ItemProperty -Path "Registry::HKEY_CURRENT_USER\software\policies\microsoft\office\$officeversion\common\feedback" -Name includescreenshot -ErrorAction SilentlyContinue|Select-Object -ExpandProperty includescreenshot

if ($includescreenshot -eq $null)
{
write-host "Allow including screenshot with Office Feedback is not configured" -ForegroundColor Yellow
}

elseif ($includescreenshot -eq '0')
{
write-host "Allow including screenshot with Office Feedback is disabled" -ForegroundColor Green
}
elseif ($includescreenshot -eq '1')
{
write-host "Allow including screenshot with Office Feedback is enabled" -ForegroundColor Red
}
else
{
write-host "Allow including screenshot with Office Feedback is set to an unknown configuration" -ForegroundColor Red
}



#Automatically receive small updates to improve reliability

$updatereliabilitydata = Get-ItemProperty -Path "Registry::HKEY_CURRENT_USER\software\policies\microsoft\office\$officeversion\common\" -Name updatereliabilitydata -ErrorAction SilentlyContinue|Select-Object -ExpandProperty updatereliabilitydata

if ($updatereliabilitydata -eq $null)
{
write-host "Automatically receive small updates to improve reliability is not configured" -ForegroundColor Yellow
}

elseif ($updatereliabilitydata -eq '0')
{
write-host "Automatically receive small updates to improve reliability is disabled" -ForegroundColor Green
}
elseif ($updatereliabilitydata -eq '1')
{
write-host "Automatically receive small updates to improve reliability is enabled" -ForegroundColor Red
}
else
{
write-host "Automatically receive small updates to improve reliability is set to an unknown configuration" -ForegroundColor Red
}


#Disable Opt-in Wizard on first run

$shownfirstrunoptin = Get-ItemProperty -Path "Registry::HKEY_CURRENT_USER\software\policies\microsoft\office\$officeversion\common\general" -Name shownfirstrunoptin -ErrorAction SilentlyContinue|Select-Object -ExpandProperty shownfirstrunoptin

if ($shownfirstrunoptin -eq $null)
{
write-host "Disable Opt-in Wizard on first run is not configured" -ForegroundColor Yellow
}

elseif ($shownfirstrunoptin -eq '1')
{
write-host "Disable Opt-in Wizard on first run is enabled" -ForegroundColor Green
}
elseif ($shownfirstrunoptin -eq '0')
{
write-host "Disable Opt-in Wizard on first run is disabled" -ForegroundColor Red
}
else
{
write-host "Disable Opt-in Wizard on first run is set to an unknown configuration" -ForegroundColor Red
}



#Enable Customer Experience Improvement Program

$qmenable = Get-ItemProperty -Path "Registry::HKEY_CURRENT_USER\software\policies\microsoft\office\$officeversion\common\" -Name qmenable -ErrorAction SilentlyContinue|Select-Object -ExpandProperty qmenable

if ($qmenable -eq $null)
{
write-host "Enable Customer Experience Improvement Program is not configured" -ForegroundColor Yellow
}

elseif ($qmenable -eq '0')
{
write-host "Enable Customer Experience Improvement Program is disabled" -ForegroundColor Green
}
elseif ($qmenable -eq '1')
{
write-host "Enable Customer Experience Improvement Program is enabled" -ForegroundColor Red
}
else
{
write-host "Enable Customer Experience Improvement Program is set to an unknown configuration" -ForegroundColor Red
}



#Send Office Feedback

$enabled = Get-ItemProperty -Path "Registry::HKEY_CURRENT_USER\software\policies\microsoft\office\$officeversion\common\feedback" -Name enabled -ErrorAction SilentlyContinue|Select-Object -ExpandProperty enabled

if ($enabled -eq $null)
{
write-host "Send Office Feedback is not configured" -ForegroundColor Yellow
}

elseif ($enabled -eq '0')
{
write-host "Send Office Feedback is disabled" -ForegroundColor Green
}
elseif ($enabled -eq '1')
{
write-host "Send Office Feedback is enabled" -ForegroundColor Red
}
else
{
write-host "Send Office Feedback is set to an unknown configuration" -ForegroundColor Red
}



#Send personal information

$sendcustomerdata = Get-ItemProperty -Path "Registry::HKEY_CURRENT_USER\software\policies\microsoft\office\$officeversion\common\" -Name sendcustomerdata -ErrorAction SilentlyContinue|Select-Object -ExpandProperty sendcustomerdata

if ($sendcustomerdata -eq $null)
{
write-host "Send personal information is not configured" -ForegroundColor Yellow
}

elseif ($sendcustomerdata -eq '0')
{
write-host "Send personal information is disabled" -ForegroundColor Green
}
elseif ($sendcustomerdata -eq '1')
{
write-host "Send personal information is enabled" -ForegroundColor Red
}
else
{
write-host "Send personal information is set to an unknown configuration" -ForegroundColor Red
}
