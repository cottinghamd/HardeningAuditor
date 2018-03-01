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

$officetemp = Get-OfficeVersion | select -ExpandProperty version
$officeversion = $officetemp.Substring(0,4)

#MS Word
$macroword = Get-ItemProperty -Path Registry::HKEY_CURRENT_USER\Software\Policies\Microsoft\Office\$officeversion\Word\Security -ErrorAction SilentlyContinue| Select-Object -ExpandProperty VBAWarnings -ErrorAction SilentlyContinue


#This is an example of querying a service status with powershell

If ($macroword -eq $null)
{
write-host "Macro settings have not been configured in Microsoft Word"
}
elseif ($macroword -eq "4"){
    write-host "Macros are disabled in Microsoft Word" -ForegroundColor Green
    }
    else
    {
    write-host "Macros are not disabled in Microsoft Word, value is $macroword" -ForegroundColor Red
      }
      if ($macroword -eq"1")
      {Write-Host "Enable all Macros"}
      elseif ($macroword -eq"2")
      {Write-Host "Disable all Macros with notification"}
      elseif ($macroword -eq"3")
      {Write-Host "Disable all Macros except those digitally signed"}

#MS PowerPoint
$macroppt = Get-ItemProperty -Path Registry::HKEY_CURRENT_USER\Software\Policies\Microsoft\Office\$officeversion\PowerPoint\Security -ErrorAction SilentlyContinue| Select-Object -ExpandProperty VBAWarnings -ErrorAction SilentlyContinue


If ($macroppt -eq $null)
{
write-host "Macro settings have not been configured in Microsoft PowerPoint"
}
elseif ($macroppt -eq "4"){
    write-host "Macros are disabled in Microsoft PowerPoint" -ForegroundColor Green
    }
    else
    {
    write-host "Macros are not disabled in Microsoft PowerPoint, value is $macroppt" -ForegroundColor Red
    }
       if ($macroppt -eq"1")
      {Write-Host "Enable all Macros"}
      elseif ($macroppt -eq"2")
      {Write-Host "Disable all Macros with notification"}
      elseif ($macroppt -eq"3")
      {Write-Host "Disable all Macros except those digitally signed"}
 
 #MS Excel
$macroexcel = Get-ItemProperty -Path Registry::HKEY_CURRENT_USER\Software\Policies\Microsoft\Office\$officeversion\Excel\Security -ErrorAction SilentlyContinue| Select-Object -ExpandProperty VBAWarnings -ErrorAction SilentlyContinue


If ($macroexcel -eq $null)
{
write-host "Macro settings have not been configured in Microsoft Excel"
}
elseif ($macroexcel -eq "4"){
    write-host "Macros are disabled in Microsoft Excel" -ForegroundColor Green
    }
    else
    {
    write-host "Macros are not disabled in Microsoft Excel, value is $macroexcel" -ForegroundColor Red
    }
       if ($macroexcel -eq"1")
      {Write-Host "Enable all Macros"}
      elseif ($macroexcel -eq"2")
      {Write-Host "Disable all Macros with notification"}
      elseif ($macroexcel -eq"3")
      {Write-Host "Disable all Macros except those digitally signed"}
   
     #MS Access
$macroaccess = Get-ItemProperty -Path Registry::HKEY_CURRENT_USER\Software\Policies\Microsoft\Office\$officeversion\access\Security -ErrorAction SilentlyContinue| Select-Object -ExpandProperty VBAWarnings -ErrorAction SilentlyContinue


If ($macroaccess -eq $null)
{
write-host "Macro settings have not been configured in Microsoft Access"
}
elseif ($macroaccess -eq "4"){
    write-host "Macros are disabled in Microsoft Access" -ForegroundColor Green
    }
    else
    {
    write-host "Macros are not disabled in Microsoft Access, value is $macroaccess" -ForegroundColor Red
    }
       if ($macroaccess -eq"1")
      {Write-Host "Enable all Macros"}
      elseif ($macroaccess -eq"2")
      {Write-Host "Disable all Macros with notification"}
      elseif ($macroaccess -eq"3")
      {Write-Host "Disable all Macros except those digitally signed"}

     #MS Outlook
$macrooutlook = Get-ItemProperty -Path Registry::HKEY_CURRENT_USER\Software\Policies\Microsoft\Office\$officeversion\outlook\Security -ErrorAction SilentlyContinue| Select-Object -ExpandProperty level -ErrorAction SilentlyContinue


If ($macrooutlook -eq $null)
{
write-host "Macro settings have not been configured in Microsoft Outlook"
}
elseif ($macrooutlook -eq "4"){
    write-host "Macros are disabled in Microsoft Outlook" -ForegroundColor Green
    }
    else
    {
    write-host "Macros are not disabled in Microsoft Outlook, value is $macrooutlook" -ForegroundColor Red
    }
   if ($macrooutlook -eq"1")
      {Write-Host "Enable all Macros"}
      elseif ($macrooutlook -eq"2")
      {Write-Host "Disable all Macros with notification"}
      elseif ($macrooutlook -eq"3")
      {Write-Host "Disable all Macros except those digitally signed"}

     #MS Project
$macroproject = Get-ItemProperty -Path "Registry::HKEY_CURRENT_USER\Software\Policies\Microsoft\Office\$officeversion\ms project\Security" -ErrorAction SilentlyContinue| Select-Object -ExpandProperty VBAWarnings -ErrorAction SilentlyContinue


If ($macroproject -eq $null)
{
write-host "Macro settings have not been configured in Microsoft Project"
}
elseif ($macroproject -eq "4"){
    write-host "Macros are disabled in Microsoft Project" -ForegroundColor Green
    }
    else
    {
    write-host "Macros are not disabled in Microsoft Project, value is $macroproject" -ForegroundColor Red
    }
       if ($macroproject -eq"1")
      {Write-Host "Enable all Macros"}
      elseif ($macroproject -eq"2")
      {Write-Host "Disable all Macros with notification"}
      elseif ($macroproject -eq"3")
      {Write-Host "Disable all Macros except those digitally signed"}
     
     #MS Publisher
$macropublisher = Get-ItemProperty -Path Registry::HKEY_CURRENT_USER\Software\Policies\Microsoft\Office\$officeversion\publisher\Security -ErrorAction SilentlyContinue| Select-Object -ExpandProperty VBAWarnings -ErrorAction SilentlyContinue


If ($macropublisher -eq $null)
{
write-host "Macro settings have not been configured in Microsoft Publisher"
}
elseif ($macropublisher -eq "4"){
    write-host "Macros are disabled in Microsoft Publisher" -ForegroundColor Green
    }
    else
    {
    write-host "Macros are not disabled in Microsoft Publisher, value is $macropublisher" -ForegroundColor Red
    }
   if ($macropublisher -eq"1")
      {Write-Host "Enable all Macros"}
      elseif ($macropublisher -eq"2")
      {Write-Host "Disable all Macros with notification"}
      elseif ($macropublisher -eq"3")
      {Write-Host "Disable all Macros except those digitally signed"}

         #MS Visio
$macrovisio = Get-ItemProperty -Path Registry::HKEY_CURRENT_USER\Software\Policies\Microsoft\Office\$officeversion\visio\Security -ErrorAction SilentlyContinue| Select-Object -ExpandProperty VBAWarnings -ErrorAction SilentlyContinue


If ($macrovisio -eq $null)
{
write-host "Macro settings have not been configured in Microsoft Visio"
}
elseif ($macrovisio -eq "4"){
    write-host "Macros are disabled in Microsoft Visio" -ForegroundColor Green
    }
    else
    {
    write-host "Macros are not disabled in Microsoft Visio, value is $macrovisio" -ForegroundColor Red
    }
       if ($macrovisio -eq"1")
      {Write-Host "Enable all Macros"}
      elseif ($macrovisio -eq"2")
      {Write-Host "Disable all Macros with notification"}
      elseif ($macrovisio -eq"3")
      {Write-Host "Disable all Macros except those digitally signed"}


#returns trusted locations

#MS Word
Write-Host "Trusted locations for Microsoft Word:" -ForegroundColor Yellow
Get-ChildItem "Registry::HKEY_CURRENT_USER\Software\Policies\Microsoft\Office\$officeversion\Word\Security\Trusted Locations" -ErrorAction SilentlyContinue
$locationWord = Get-ChildItem "Registry::HKEY_CURRENT_USER\Software\Policies\Microsoft\Office\$officeversion\Word\Security\Trusted Locations" -ErrorAction SilentlyContinue
If ($locationword -eq $null)
{write-host "No trusted locations have been configured in Microsoft Word"}



#MS Powerpoint
Write-Host "Trusted locations for Microsoft Powerpoint:" -ForegroundColor Yellow
Get-ChildItem "Registry::HKEY_CURRENT_USER\Software\Policies\Microsoft\Office\$officeversion\Powerpoint\Security\Trusted Locations" -ErrorAction SilentlyContinue
$locationppt = Get-ChildItem "Registry::HKEY_CURRENT_USER\Software\Policies\Microsoft\Office\$officeversion\Powerpoint\Security\Trusted Locations" -ErrorAction SilentlyContinue
If ($locationppt -eq $null)
{write-host "No trusted locations have been configured in Microsoft Powerpoint"}



#MS Excel
Write-Host "Trusted locations for Microsoft Excel:" -ForegroundColor Yellow
Get-ChildItem "Registry::HKEY_CURRENT_USER\Software\Policies\Microsoft\Office\$officeversion\Excel\Security\Trusted Locations" -ErrorAction SilentlyContinue
$locationexcel = Get-ChildItem "Registry::HKEY_CURRENT_USER\Software\Policies\Microsoft\Office\$officeversion\Excel\Security\Trusted Locations" -ErrorAction SilentlyContinue
If ($locationexcel -eq $null)
{write-host "No trusted locations have been configured in Microsoft Excel"}


#MS Access
Write-Host "Trusted locations for Microsoft Access:" -ForegroundColor Yellow
Get-ChildItem "Registry::HKEY_CURRENT_USER\Software\Policies\Microsoft\Office\$officeversion\Access\Security\Trusted Locations" -ErrorAction SilentlyContinue
$locationaccess = Get-ChildItem "Registry::HKEY_CURRENT_USER\Software\Policies\Microsoft\Office\$officeversion\Access\Security\Trusted Locations" -ErrorAction SilentlyContinue
If ($locationaccess -eq $null)
{write-host "No trusted locations have been configured in Microsoft Access"}


#MS Outlook
Write-Host "Trusted locations for Microsoft Outlook:" -ForegroundColor Yellow
Get-ChildItem "Registry::HKEY_CURRENT_USER\Software\Policies\Microsoft\Office\$officeversion\Outlook\Security\Trusted Locations" -ErrorAction SilentlyContinue
$locationoutlook = Get-ChildItem "Registry::HKEY_CURRENT_USER\Software\Policies\Microsoft\Office\$officeversion\Outlook\Security\Trusted Locations" -ErrorAction SilentlyContinue
If ($locationoutlook -eq $null)
{write-host "No trusted locations have been configured in Microsoft Outlook"}


#MS Project
Write-Host "Trusted locations for Microsoft Project:" -ForegroundColor Yellow
Get-ChildItem "Registry::HKEY_CURRENT_USER\Software\Policies\Microsoft\Office\$officeversion\MS Project\Security\Trusted Locations" -ErrorAction SilentlyContinue
$locationproject = Get-ChildItem "Registry::HKEY_CURRENT_USER\Software\Policies\Microsoft\Office\$officeversion\MS Project\Security\Trusted Locations" -ErrorAction SilentlyContinue
If ($locationproject -eq $null)
{write-host "No trusted locations have been configured in Microsoft Project"}


#MS Publisher
Write-Host "Trusted locations for Microsoft Publisher:" -ForegroundColor Yellow
Get-ChildItem "Registry::HKEY_CURRENT_USER\Software\Policies\Microsoft\Office\$officeversion\Publisher\Security\Trusted Locations" -ErrorAction SilentlyContinue
$locationpublisher = Get-ChildItem "Registry::HKEY_CURRENT_USER\Software\Policies\Microsoft\Office\$officeversion\Publisher\Security\Trusted Locations"-ErrorAction SilentlyContinue
If ($locationpublisher -eq $null)
{write-host "No trusted locations have been configured in Microsoft Publisher"}


#MS Visio
Write-Host "Trusted locations for Microsoft Visio:" -ForegroundColor Yellow
Get-ChildItem "Registry::HKEY_CURRENT_USER\Software\Policies\Microsoft\Office\$officeversion\Visio\Security\Trusted Locations" -ErrorAction SilentlyContinue
$locationvisio = Get-ChildItem "Registry::HKEY_CURRENT_USER\Software\Policies\Microsoft\Office\$officeversion\Visio\Security\Trusted Locations" -ErrorAction SilentlyContinue
If ($locationvisio -eq $null)
{write-host "No trusted locations have been configured in Microsoft Visio"}
else {write-host $locationvisio}
