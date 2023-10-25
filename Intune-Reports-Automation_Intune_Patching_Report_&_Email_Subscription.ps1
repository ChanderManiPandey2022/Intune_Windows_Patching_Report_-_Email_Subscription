<# 
.SYNOPSIS
<Intune_Patch_Compliance_Calculation_&_Email_Subscription Using PowerShell>
#>

<# 
.DESCRIPTION
<Intune_Patch_Compliance_Calculation_&_Email_Subscription Using PowerShell>
#>

<# 
.Demo
<YouTube video link--> https://www.youtube.com/watch?v=hAVgNvEAdKc
#>

<# 
.INPUTS
<Provide all required inforamtion in User Input Section-line No 45-55 >
#>

<# 
.OUTPUTS
<You will get Intune_Patch_Compliance_Calculation_&_Email_Subscription + report in CSV>
#>

<# 
.NOTES
Version:         1.1
Author:          Chander Mani Pandey
Creation Date:   16 July 2023
Find Author on 
Youtube:-        https://www.youtube.com/@chandermanipandey8763
Twitter:-        https://twitter.com/Mani_CMPandey
Facebook:-       https://www.facebook.com/profile.php?id=100087275409143&mibextid=ZbWKwL
LinkedIn:-       https://www.linkedin.com/in/chandermanipandey
Reddit:-         https://www.reddit.com/u/ChanderManiPandey 
#>

$error.clear() ## this is the clear error history 
cls
Set-ExecutionPolicy -ExecutionPolicy 'ByPass' -Scope 'Process' -Force -ErrorAction 'Stop' 
$ErrorActionPreference = 'SilentlyContinue';

#--------------------------------  User Input Section Start -----------------------------------------------------------------

$WorkingFolder = "C:\TEMP\MicroSoftPatchList\LatestCumulativeUpdateList" # Location where you want to create reporting folders
$From          = "123@abc.com"
$To            = "abc@abc.com"
$CC            = "def@abc.com"
$SmtpServer    = "smtp.gmail.com"
$Port          = '587'
$tenant = “abc.com”
$clientId = “xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx”
$clientSecret = “yyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyy”
# If you want to change the Subject Line then change the line number 475

#-------------------------------- User Input Section End---------------------------------------------------------------------

$startTime = Get-Date
Write-Host "===============================Phase-1 (Exporting Intune Device Dump) ======================================================(Started)" -ForegroundColor Green
$error.clear() ## this is the clear error history 
$Path ="$WorkingFolder\PD_Dump\"
New-Item -ItemType Directory -Path "$WorkingFolder\PD_Dump\" -Force | Out-Null
$MGIModule = Get-module -Name "Microsoft.Graph.Intune" -ListAvailable
Write-Host "Checking Microsoft.Graph.Intune is Installed or Not"
    If ($MGIModule -eq $null) 
    {
        Write-Host "Microsoft.Graph.Intune module is not Installed"
        Write-Host "Installing Microsoft.Graph.Intune module"
        Install-Module -Name Microsoft.Graph.Intune -Force
        Write-Host "Importing Microsoft.Graph.Intune module"
        Import-Module Microsoft.Graph.Intune -Force
    }
    ELSE 
    {   Write-Host "Microsoft.Graph.Intune is Installed"
        Write-Host "Importing Microsoft.Graph.Intune module"
        Import-Module Microsoft.Graph.Intune -Force
    }
$tenant = $tenant
$authority = “https://login.windows.net/$tenant”
$clientId = $clientId  
$clientSecret = $clientSecret
Update-MSGraphEnvironment -AppId $clientId -Quiet
Update-MSGraphEnvironment -AuthUrl $authority -Quiet
Connect-MSGraph -ClientSecret $ClientSecret -Quiet
Update-MSGraphEnvironment -SchemaVersion "Beta" -Quiet

#============Create Request Body==========================================================================================================================================================
$postBody = @{
 'reportName' = "DevicesWithInventory"
 'filter' = "(DeviceType eq '1') "
 'select' =  ("DeviceId"),("SerialNumber"), ("DeviceName") ,("ownerType") ,("OSVersion"),("UPN"),("LastContact"),("JoinType"),("Manufacturer"),("Model"),("ManagementAgent"),("SkuFamily"),("StorageTotal") ,("StorageFree")
  }
#=========== MakeRequest ==================================================================================================================================================================
$exportJob = Invoke-MSGraphRequest -HttpMethod POST -Url "DeviceManagement/reports/exportJobs" -Content $postBody
Write-Host "Export Job initiated for $ReportName Report "
#====================================Checking Report Ready status==========================================================================================================================
do{ 
$exportJob = Invoke-MSGraphRequest -HttpMethod Get -Url "DeviceManagement/reports/exportJobs('$($exportJob.id)')" -InformationAction SilentlyContinue
    Start-sleep -second 2
    Write-Host -NoNewline '...........'
  } while ($exportJob.status -eq 'inprogress')
  Write-Host 'DevicesWithInventory Report is in Ready(Completed) status for Downloading'
  If ($exportJob.status -eq 'completed') 
  { $fileName = (Split-path -Path $exportJob.url -Leaf).split('?')[0]
  Write-host "DevicesWithInventory Report Export Job completed.Writing File $fileName to Disk..."
  Invoke-WebRequest -Uri $exportJob.url -Method Get -OutFile $fileName
  Remove-Item –path $path* -include *.csv
  Expand-Archive -Path $fileName -DestinationPath $Path 
  $FileName = Get-ChildItem -Path $Path* -Include *.csv | Where {! $_.PSIsContainer } 
  $DevicesInfos = import-csv -Path $FileName.fullName
  
  }
Write-Host "===============================Phase-1 (Exporting Intune Device Dump) ====================================================(Completed)" -ForegroundColor Green
Write-Host "" 
Write-Host "===============================Phase-2 (Downloading and Creating MS Patch List) ============================================(Started)" -ForegroundColor Green
#-________________________________________________________Latest Patch List___________________________________________________________________________________

$Date = Get-Date -Format "MMMMMMMM dd, yyyy";
$OutFileMP = "$WorkingFolder\MicrosoftPatchList.csv";
$OutFileLP = "$WorkingFolder\MicrosoftLatestPatchList.csv";
$MergeOverallFile = "$WorkingFolder\MergeOverallFile.csv";
$Final_Patching_Report = "$WorkingFolder\Final_Patching_Report.csv";

$PatchingMonth = "";
$PatchReleaseDays = 0;
# Create an empty array of PSObject objects
$buildInfoArray = @()
#============================================================================================================================================================
#Creating working Folder

# Add each Build and Operating System to the array
"22623,Windows 11 22H2","22621,Windows 11 22H2 B1","22471,Windows 11 21H2","22468,Windows 11 21H2 B6","22463,Windows 11 21H2 B5",
"22458,Windows 11 21H2 B4","22454,Windows 11 21H2 B3","22449,Windows 11 21H2 B2","22000,Windows 11 21H2 B1","21996,Windows 11 Dev",
"19045,Windows 10 22H2","19044,Windows 10 21H2","19043,Windows 10 21H1","19042,Windows 10 20H2","19041,Windows 10 2004","19008,Windows 10 20H1",
"18363,Windows 10 1909","18362,Windows 10 1903","17763,Windows 10 1809","17134,Windows 10 1803","16299,Windows 10 1709 FC","15254,Windows 10 1709",
"15063,Windows 10 1703","14393,Windows 10 1607","10586,Windows 10 1511","10240,Windows 10 1507","9600,Windows 8.1",
"7601,Windows 7" | ForEach-Object {
    # Create a new PSObject object
    $buildInfo = New-Object -TypeName PSObject
    # Add the Build and Operating System properties to the object
    $buildInfo | Add-Member -MemberType NoteProperty -Name "Build" -Value ($_ -split ",")[0]
    $buildInfo | Add-Member -MemberType NoteProperty -Name "OperatingSystem" -Value ($_ -split ",")[1]
    # Add the object to the array
    $buildInfoArray += $buildInfo
}

#==============================================================================================================================================================================

$CollectedData = $BuildDetails = $PatchDetails = $MajorBuilds = $LatestPatches = @();
$BuildDetails = $buildInfoArray
#Download Windows Master Patch List
Write-Host "Downoading Patch List from Microsoft"-ForegroundColor yellow
$URI = "https://aka.ms/Windows11UpdateHistory";
$CollectedData += (Invoke-WebRequest -Uri $URI -UseBasicParsing -ErrorAction Continue).Links;
#$URI = "https://aka.ms/WindowsUpdateHistory";
$URI = "https://support.microsoft.com/en-us/help/4043454";
$CollectedData += (Invoke-WebRequest -Uri $URI -UseBasicParsing -ErrorAction Continue).Links;

#Filter Windows Master Patch List

Write-Host "Filtering Patch List"-ForegroundColor yellow
IF ($CollectedData) 
{   $CollectedDataAll = ($CollectedData | Where-Object {$_.class -eq "supLeftNavLink" -and $_.outerHTML -notmatch "mobile"}).outerHTML
    $CollectedData =  ($CollectedData | Where-Object {$_.class -eq "supLeftNavLink" -and $_.outerHTML -match "KB" -and $_.outerHTML -notmatch "out-of-band" -and $_.outerHTML -notmatch "Preview" -and $_.outerHTML -notmatch "mobile"}).outerHTML
    $PatchTuesdayOSBuilds = $CollectedData | ForEach-Object {if ($_ -match 'OS Build (\d+\.\d+)') {$matches[1] }}
    $CollectedDataPreview = $CollectedDataAll | Select-String -Pattern '(?<=<a class="supLeftNavLink" data-bi-slot="\d+" href="\/en-us\/help\/\d+">).*?(?=<\/a>)' | ForEach-Object {
                           if ($_ -match 'KB' -and $_ -notmatch 'out-of-band' -and $_ -match 'Preview' -and $_ -notmatch 'mobile') {  $_ }}
    $OSPreBuilds = [regex]::Matches($CollectedDataPreview, '\d+\.\d+').Value | Sort-Object -Unique
    $PreviewOSBuilds = $OSPreBuilds -join "`n"
    $CollectedDataOutofBand = $CollectedDataAll | Select-String -Pattern '(?<=<a class="supLeftNavLink" data-bi-slot="\d+" href="\/en-us\/help\/\d+">).*?(?=<\/a>)' | ForEach-Object {
                            if ($_ -match 'KB' -and $_ -match 'out-of-band' -and $_ -notmatch 'Preview' -and $_ -notmatch 'mobile') {  $_ }}
    $OSOOBBuilds = [regex]::Matches($CollectedDataOutofBand, '\d+\.\d+').Value | Sort-Object -Unique
    $OutofBandOSBuilds = $OSOOBBuilds -join "`n"
    }
   write-Host "All found Update Count= " $CollectedDataAll.Count
   write-Host "All found Patch Tuesday Update Count= " $CollectedData.count
   write-Host "All found Preview Update Count= " $CollectedDataPreview.count
   write-Host "All found Out of Band Update Count= " $CollectedDataOutofBand.count
   #$CollectedData.count
   #$PatchTuesdayOSBuilds.count
   #$CollectedDataPreview.count
   #$PreviewOSBuilds.count
   #$CollectedDataOutofBand.count
   #$OutofBandOSBuilds.count  
#Consolidate the Master Patch and Format the output
Write-Host "Consolidating Patch List"-ForegroundColor yellow
Foreach ($Line in $CollectedData) {
	$ReleaseDate = $PatchID = ""; $Builds = @();	
    $ReleaseDate = (($Line.Split(">")[1]).Split("&”)[0]).trim();
        IF ($ReleaseDate -match "build") {$ReleaseDate = ($ReleaseDate.split("-")[0]).trim();}
	$PatchID = ($Line.Split(" ;-") | Where-Object {$_ -match "KB"}).trim();
    $Builds = ($Line.Split(",) ") | Where-Object {$_ -like "*.*"}).trim();
	Foreach ($BLD in $Builds) {
		$MjBld = $MnBld = ""; $MjBld = $BLD.Split(".")[0]; $MnBld = $BLD.Split(".")[1];
            Foreach ($Line1 in $BuildDetails) {
                $BldNo = $OS = ""; $BldNo = $Line1.Build; $OS = $Line1.OperatingSystem; $MajorBuilds += $BldNo;
                IF ($MjBld -eq $BldNo) {Break;}
                ELSE {$OS = "Unknown";}
            }
            $PatchDetails += [PSCustomObject] @{OperatingSystem = $OS; Build = $BLD; MajorBuild = $MjBld; MinorBuild = $MnBld; PatchID = $PatchID; ReleaseDate = $ReleaseDate;}
       }
}
$MajorBuilds = $MajorBuilds | Select-Object -Unique | Sort-Object -Descending;
$PatchDetails = $PatchDetails | Select-Object OperatingSystem, Build, MajorBuild, MinorBuild, PatchID, ReleaseDate -Unique | Sort-Object MajorBuild,PatchID -Descending;
$PatchDetails | Export-Csv -Path $OutFileMP -NoTypeInformation;
Write-Host "Finalizing Patch List" -ForegroundColor yellow
IF ($PatchingMonth) {
	Foreach ($Bld in $MajorBuilds) {$LatestPatches += $PatchDetails | Where-Object {$_.MajorBuild -eq $Bld -and 
        $_.ReleaseDate -match $PatchingMonth.Year -and $_.ReleaseDate -match $PatchingMonth.Month} | Sort-Object PatchID -Descending | Select-Object -First 1;}
}
ELSE {
    $Today = Get-Date; $LatestDate = ($PatchDetails | Select-Object -First 1).ReleaseDate; $DiffDays = ([datetime]$Today - [datetime]$LatestDate).Days;
    IF ([int]$DiffDays -gt [int]$PatchReleaseDays) {
        Foreach ($Bld in $MajorBuilds) {$LatestPatches += $PatchDetails | Where-Object {$_.MajorBuild -eq $Bld} | Sort-Object PatchID -Descending | Select-Object -First 1;}
    }
    ELSE {
        $Month = ""; $Month = ((Get-Date).AddMonths(-1)).ToString("MMMMMMMM dd, yyyy").Split(" ,") | Select-Object -First 1;
        $Year = ""; $Year = ((Get-Date).AddMonths(-1)).ToString("MMMMMMMM dd, yyyy").Split(" ,") | Select-Object -Last 1;
        $PatchingMonth = [PSCustomObject]@{Month = $Month; Year = $Year;};
        Foreach ($Bld in $MajorBuilds) {$LatestPatches += $PatchDetails | Where-Object {$_.MajorBuild -eq $Bld -and 
            $_.ReleaseDate -match $PatchingMonth.Year -and $_.ReleaseDate -match $PatchingMonth.Month} | Sort-Object PatchID -Descending | Select-Object -First 1;}
 
	    #Adding Latest Patches for Other Builds Missing above
        $M = ""; $M = ((Get-Date).ToString("MMMMMMMM dd, yyyy")).split(" ,") | Select-Object -First 1;
        $Y = ""; $Y = ((Get-Date).ToString("MMMMMMMM dd, yyyy")).split(" ,") | Select-Object -Last 1;
        Foreach ($Bld1 in $MajorBuilds) {
            $Found = 0; Foreach ($Line in $LatestPatches) {$Bld2 = ""; $Bld2 = $Line.MajorBuild; IF ($Bld1 -eq $Bld2) {$Found = 1; Break;}}
            IF ($Found -eq 0) {$LatestPatches += $PatchDetails | Where-Object {$_.MajorBuild -eq $Bld1 -and 
                               $_.ReleaseDate -notlike "$M*$Y"} | Sort-Object PatchID -Descending | Select-Object -First 1;}
        }
    }
}
$LatestPatches = $LatestPatches | Select-Object OperatingSystem,Build,MajorBuild,MinorBuild,PatchID,ReleaseDate, @{Name="OSVersion"; Expression={"10.0.$($_.Build)"}}
$LatestPatches| Export-csv -Path $OutFileLP -NoTypeInformation
# Sort the patches by ReleaseDate in descending order and retrieve the most recent date
$mostRecentDate = $LatestPatches | Sort-Object -Property ReleaseDate -Descending | Select-Object -first 1 
# Store the most recent ReleaseDate in $patchtuesday variable
$patchtuesday = $mostRecentDate.ReleaseDate
$AllReleasedPatchs = import-csv -Path $OutFileMP 
$AllReleasedPatchs = $AllReleasedPatchs | Select-Object OperatingSystem,Build,MajorBuild,MinorBuild,PatchID,ReleaseDate, @{Name="OSVersion"; Expression={"10.0.$($_.Build)"}}
$AllReleasedPatchs | Export-csv -Path $OutFileMP -NoTypeInformation
$IntuneDeviceHardwareInfo = @()
foreach ($AllReleasedPatch in $AllReleasedPatchs) 
{ 
  $IntuneDeviceHSProps = [ordered]   @{
  OperatingSystem = $AllReleasedPatch.OperatingSystem
  OSVersion =$AllReleasedPatch.osversion 
  Build= $AllReleasedPatch.Build
  MajorBuild=$AllReleasedPatch.MajorBuild
  MinorBuild=$AllReleasedPatch.MinorBuild
  PatchID=$AllReleasedPatch.PatchID
  ReleaseDate=$AllReleasedPatch.ReleaseDate 
  PatchStatus = $status =if ($LatestPatches.OSversion  -contains $AllReleasedPatch.osversion ) {"Compliant"} else {"Non-Compliant"}
  NPS = $timeSpan = (Get-Date).Subtract([DateTime]::ParseExact($AllReleasedPatch.ReleaseDate, "MMMM d, yyyy", [CultureInfo]::InvariantCulture))
  NotPatchSince =   if ($status -eq 'Compliant') {"Compliant"} else {$timeSpan.Days.ToString() + " days"}
  RequiredPatch = if ($PatchStatus -eq "Compliant") { "Compliant" } else {
            $matchingPatch = $LatestPatches | Where-Object { $_.MajorBuild -eq $AllReleasedPatch.MajorBuild -and $_.OSVersion -eq $AllReleasedPatch.OSVersion }
            if ($matchingPatch) { "Compliant" } else {
                $latestMajorBuildPatches = $LatestPatches | Where-Object { $_.MajorBuild -eq $AllReleasedPatch.MajorBuild }
                if ($latestMajorBuildPatches) {
                    $latestMajorBuildPatches.PatchID -join ", "
                } else {"BNE"}
            }
        }
  RequiredPatchRD =     if ($PatchStatus -eq "Compliant") { "Compliant" } else {
            $matchingPatch = $LatestPatches | Where-Object { $_.MajorBuild -eq $AllReleasedPatch.MajorBuild -and $_.OSVersion -eq $AllReleasedPatch.OSVersion }
            if ($matchingPatch) { "Compliant" } else {
                $latestMajorBuildPatches = $LatestPatches | Where-Object { $_.MajorBuild -eq $AllReleasedPatch.MajorBuild }
                if ($latestMajorBuildPatches) {
                    $latestMajorBuildPatches.ReleaseDate -join ", "
                } else {"BNE"}
            }
        }
}
  $IntuneDeviceHSobject = New-Object -Type PSObject -Property $IntuneDeviceHSProps
    $IntuneDeviceHardwareInfo += $IntuneDeviceHSobject
 }
$FinalReport = $IntuneDeviceHardwareInfo | Select-Object OperatingSystem,OSVersion,Build,MajorBuild,MinorBuild,PatchID,ReleaseDate,PatchStatus,NotPatchSince,RequiredPatch,RequiredPatchRD
$FinalReport | Export-Csv -path $MergeOverallFile -NoTypeInformation
#invoke-item -path $MergeOverallFile
Write-Host "===============================Phase-2 (Downloading and Creating MS Patch List) ==========================================(Completed)" -ForegroundColor Green
Write-Host "" 
Write-Host "===============================Phase-3 (Generating Windows Patching Compliance Report) =====================================(Started)" -ForegroundColor Green
Write-Host "Calculating Patch compliance status against each device and other Inforamtion like DevcieNotPatchSince_InDays, Latest_RequiredPatch, RequiredPatchRD"  -ForegroundColor yellow
Write-Host "Other Inforamtion like DevcieNotPatchSince_InDays, Latest_RequiredPatch, RequiredPatchRD"  -ForegroundColor yellow

$compliantCount = 0 ; $manualCheckCount = 0 ;$nonCompliantCount = 0 ;$Lastcheckin_Indays=0 ;
$complianceReport = @()
$totalDevices = $DevicesInfos.Count
$progress = 0
foreach ($device in $DevicesInfos) {
    $deviceName = $device."device name"
    $Serialnumber = $device."Serial Number"
    $Manufacturer  = $device.Manufacturer 
    $Model            = $device.Model  
    $Managedby       = $device."Managed by"
    $SkuFamily        = $device.SkuFamily 
    $Totalstorage    = ($device."Total storage"  / 1024).ToString("N2")
    $Freestorage     = ($device."Free storage" / 1024).ToString("N2")
    $deviceOSVersion = $device."os version"
    $OSVersion =$deviceOSVersion.Split(".")[2]
    $OSVersionV =  If ($OSVersion -eq '10240') {'Win10-1507'} ElseIf ($OSVersion -eq  "10586") {"Win10-1511"} ElseIf ($OSVersion -eq "14393") {"Win10-1607"} ElseIf ($OSVersion -eq "15063") {
           "Win10-1703"} ElseIf ($OSVersion -eq "16299") {"Win10-1709"} ElseIf ($OSVersion -eq "17134") {"Win10-1803"} ElseIf ($OSVersion -eq "17763") {'Win10-1809'} ElseIf ($OSVersion -eq '18362') {
           "Win10-1903"} ElseIf ($OSVersion -eq "18363") {"Win10-1909"} ElseIf ($OSVersion -eq "19041") {"Win10-2004"} ElseIf ($OSVersion -eq "19042") {"Win10-20H2"} ElseIf ($OSVersion -eq "19043") {
           "Win10-21H1"} ElseIf ($OSVersion -eq "19044") {"Win10-21H2"} ElseIf ($OSVersion -eq "19045") {"Win10-22H2"} ElseIf ($OSVersion -eq "22000") {"Win11-21H2"} ElseIf ($OSVersion -eq "22621") {
           "Win11-22H2"}ElseIf ($OSVersion -eq "0") {"0.0.0.0"}ElseIf ($OSVersion -eq "7601") {"Win7-Or-Server"}ElseIf ($OSVersion -eq $null) {"No OS version"}Else {$deviceOSVersion }
    $Ownership = $device.Ownership
    #$Lastcheckin =([DateTime]$device."Last check-in").ToString("dd-MMM-yy")
    $Lastcheckin = ""
   if ([string]::IsNullOrEmpty($device."Last check-in")) {$Lastcheckin = "No_CheckIndate"} else {$Lastcheckin = ([DateTime]::ParseExact($device."Last check-in", "yyyy-MM-dd HH:mm:ss.fffffff", $null)).ToString("dd-MMM-yy")}
    $Lastcheckin1 =$device."Last check-in"
    $LastcheckinDate = [DateTime]::ParseExact($Lastcheckin1, "yyyy-MM-dd HH:mm:ss.fffffff", [System.Globalization.CultureInfo]::InvariantCulture)
    $today = Get-Date -Format "yyyy-MM-dd"
    $Lastcheckin_Indays =  (Get-Date $today) - $LastcheckinDate 
    $deviceCategory = "" 
    if ($Lastcheckin_Indays -eq "No_CheckIndate" ) {$deviceCategory = "No date"} 
    elseif ($Lastcheckin_Indays.Days -ge 0 -and $Lastcheckin_Indays.Days -le 1) { $deviceCategory = "0-1 days"} 
    elseif ($Lastcheckin_Indays.Days -ge 2 -and $Lastcheckin_Indays.Days -le 5) {$deviceCategory = "2-5 days" } 
    elseif ($Lastcheckin_Indays.Days -ge 6 -and $Lastcheckin_Indays.Days -le 10) {$deviceCategory = "5-10 days"} 
    elseif ($Lastcheckin_Indays.Days -ge 11 -and $Lastcheckin_Indays.Days -le 20) {$deviceCategory = "11-20 days"} 
    elseif ($Lastcheckin_Indays.Days -ge 21 -and $Lastcheckin_Indays.Days -le 30) {$deviceCategory = "21-30 days"} 
    elseif ($Lastcheckin_Indays.Days -ge 31 -and $Lastcheckin_Indays.Days -le 60) {$deviceCategory = "31-60 days"} 
    elseif ($Lastcheckin_Indays.Days -ge 61 -and $Lastcheckin_Indays.Days -le 90) {$deviceCategory = "61-90 days"} 
    elseif ($Lastcheckin_Indays.Days -gt 90) {$deviceCategory = "above 90 days"} else {$deviceCategory = "Check_Date"}
    $JoinType = $device.JoinType
    $PrimaryUserUPN =$device."Primary user UPN"
    $matchingPatch = $LatestPatches | Where-Object { $_.OSVersion -eq $deviceOSVersion }
    #$complianceStatus = if ($matchingPatch.OSVersion -ge $deviceOSVersion) { "Compliant" } else { "Non-Compliant" }
    $KBNumber = if ($complianceStatus -eq "Compliant") { $matchingPatch.PatchID } else { $null }
    $kbReleasedate = $matchingPatch.ReleaseDate
    $notPatchSince = $FinalReport | Where-Object { $_.OSVersion -eq $deviceOSVersion } | Select-Object -ExpandProperty NotPatchSince
    $notPatchSince = if ([string]::IsNullOrWhiteSpace($notPatchSince)) { "Manually Check" } else { $notPatchSince }
    $PatchID = $FinalReport | Where-Object { $_.OSVersion -eq $deviceOSVersion } | Select-Object -ExpandProperty PatchID
    $PatchID = if ([string]::IsNullOrWhiteSpace($PatchID)) { "Manually Check Installed KB" } else { $PatchID }
    $complianceStatus = if ($matchingPatch.OSVersion -ge $deviceOSVersion) { "Compliant" } elseif ($PatchID -eq  "Manually Check Compliance " ){ "Manually Check Installed KB"} else { "Non-Compliant" }
    $OSMajorMinorversion = $deviceOSVersion -replace '^.+?\..+?\.(.*)$', '$1'
    $PatchType = if ($PatchID  -ne "Manually Check Installed KB"){ "PatchTuesdayUpdate" } elseif ( $OSPreBuilds  -contains  $OSMajorMinorversion ){"Preview Update" }  elseif ( $OSOOBBuilds -contains  $OSMajorMinorversion ){ "OOB Update" } else {"Manually Check"}
    $ReleaseDate = $FinalReport | Where-Object { $_.OSVersion -eq $deviceOSVersion } | Select-Object -ExpandProperty ReleaseDate
    $ReleaseDate  = if ([string]::IsNullOrWhiteSpace($ReleaseDate )) { "Manually Check  Release Date" } else { $ReleaseDate }
    $RequiredPatch = $FinalReport | Where-Object { $_.OSVersion -eq $deviceOSVersion } | Select-Object -ExpandProperty RequiredPatch
    $RequiredPatch  = if ([string]::IsNullOrWhiteSpace($RequiredPatch )) { "Manually Check Required Patch" } else { $RequiredPatch}
    $RequiredPatchRD = $FinalReport | Where-Object { $_.OSVersion -eq $deviceOSVersion } | Select-Object -ExpandProperty RequiredPatchRD
    $RequiredPatchRD  = if ([string]::IsNullOrWhiteSpace($RequiredPatchRD  )) { "Manually Check Required Patch" } else { $RequiredPatchRD}
  $count = if ($RequiredPatch -eq "Compliant") {$compliantCount++} elseif($RequiredPatch -eq "Manually Check Required Patch") {$manualCheckCount++} else{$nonCompliantCount++}
  $reportRow = [PSCustomObject] @{
        "DeviceName" = $deviceName
        "Serialnumber"=  $Serialnumber
        "PrimaryUserUPN" = $PrimaryuserUPN
        "Ownership" = $Ownership
        "JoinType" = $JoinType
        "Manufacturer"= $Manufacturer 
        "Model"  = $Model  
        "Managedby" = $Managedby
        "Totalstorage (GB)" = $Totalstorage 
        "Freestorage (GB)"  = $Freestorage
        "Lastcheckin" = $Lastcheckin 
        "Lastcheckin_Indays" = $Lastcheckin_Indays.Days.ToString() + " days"
        "Lastcheckin_InBetween" = $deviceCategory
        "SkuFamily" = $SkuFamily 
        "OSVersion"=$OSVersionV
        "OS" = $deviceOSVersion
        "InstalledKB" = $PatchID
        "PatchType" = $PatchType 
        "InstalledKB_ReleaseDate" =$ReleaseDate
        "PatchingStatus" = $complianceStatus
        "DevcieNotPatchSince_InDays" = $notPatchSince
        "Latest_RequiredPatch" = $RequiredPatch
        "RequiredPatchRD " = $RequiredPatchRD 
        
          }
    $complianceReport += $reportRow
    $progress++
    $percentComplete = [String]::Format("{0:0.00}", ($progress / $totalDevices) * 100)
    Write-Progress -Activity "Generating Windows Patching Compliance Report" -Status "Progress: $percentComplete% Complete" -PercentComplete $percentComplete
    }
$complianceReport| Export-Csv -Path $Final_Patching_Report -NoTypeInformation
$totalCount = $compliantCount + $manualCheckCount + $nonCompliantCount
$compliancePercentage = "{0:N2}" -f ($compliantCount / $totalCount * 100)
Write-Host ""                                                             
Write-Host "Total Device Count: $totalCount"
Write-Host "Total Compliant Device Count: $compliantCount"
Write-Host "Total Manually Check Required Patch Device Count: $manualCheckCount" -ForegroundColor Yellow
Write-Host "Total Non-Compliant Device Count: $nonCompliantCount" -ForegroundColor Red
Write-Host ""
Write-Host "Patching Compliance Percentage: $compliancePercentage%" -ForegroundColor Green
Write-Host "===============================Phase-3 (Generating Windows Patching Compliance Report) ===================================(Completed)" -ForegroundColor Green
Write-Host ""
Write-Host "===============================Phase-4 (Sending E-mail(Windows Patching Compliance Report) =================================(Started)" -ForegroundColor Green

$EmailBody1 = @" 

<p><span style="font-weight:bold;">Hello All</span></p>
 <p></p>
 <p><span style="font-weight:bold;">Please Find Windows 10/11 Intune patching compliance report against $LatestDate  (Patch Tuesday).</span></p>

<head>
	<style> 	table, th,
                td {border: 3px solid black;}
	</style>
</head>

<table style="width: 68%" style="border-collapse: collapse; border: 1px solid #008080;">

 <tr>
    <td colspan="2" bgcolor="#71B2EE" style="background-color:Tan; font-size: large; height: 35px;">
        <b>Intune - Windows Patching Compliance Dashboard</b>   
    </td>
 </tr>


 <!----For Total Devices-------------------------------------------------------------------------------------->
<tr <tr style="background-color:lightgrey">
    <td style="width: 201px; height: 35px">&nbsp;Total Devices</td>
    <th style="height: 35px; width: 233px;">
    <b>VarTotal</b></td>
 </tr>
 
  <! --- For Compliance Devcies -------------------------------------------------------------------------------->
 <tr <tr style="background-color:MediumAquaMarine">
    <td style="width: 201px; height: 35px">&nbsp;Compliant Devices</td>
    <th style="height: 35px; width: 233px;">
    <b>Varsuccess</b></td>
 </tr>

<!----For Non_Compliant Devcies-------------------------------------------------------------------------------------->
<tr <tr style="background-color:Salmon">
    <td style="width: 201px; height: 35px">&nbsp;Non-Compliant Devices</td>
    <th style="height: 35px; width: 233px;">
    <b>Varfailure</b></td>
 </tr>

<!----For NeedToCheck Devcies-------------------------------------------------------------------------------------->
<tr <tr style="background-color:LightGoldenrodYellow">
    <td style="width: 201px; height: 35px">&nbsp;Need To Check Device(Manual Check)</td>
    <th style="height: 35px; width: 233px;">
    <b>VarNeedToCheck</b></td>
 </tr>


  <!----For Over All Compliance---------------------------------------------------------------------------------------->
<tr <tr style="background-color:lightgreen">
    <td style="width: 201px; height: 35px">&nbsp;<b>Compliance(%)</b></td>
    <th style="height: 35px; width: 233px;"> 
    <b>VarCompliance%</b></th>
 </tr>


</table>

<p><span style="font-weight:bold;">Note</span></p>

<p><small>Compliant Devices:- The device is compliant with the latest available Patch Tuesday($LatestDate) Updates</small></p>
<p><small>Non-Compliant Devices :- The device is not compliant with the latest available Patch Tuesday($LatestDate) updates </small></p>
<p><small>Need To Check Device (Manual Check):- The device has predominantly installed preview updates and out-of-band (OOB) updates </small></p>


<p><span style="font-weight:bold;">Regards</span></p>
<p><span style="font-weight:bold;">Patch Management Team</span></p>
 
"@
$EmailBody1= $EmailBody1.Replace("VarTotal",$totalCount)
$EmailBody1= $EmailBody1.Replace("Varsuccess",$compliantCount)
$EmailBody1= $EmailBody1.Replace("Varfailure",$nonCompliantCount)
$EmailBody1= $EmailBody1.Replace("VarCompliance",$compliancePercentage)
$EmailBody1= $EmailBody1.Replace("VarNeedToCheck",$manualCheckCount)

#___________________________________________________________________________________________________________________________________________________________

#----------------------------------------------SENDING EMAIL -----------------------------------------------------------------------------------------------
#___________________________________________________________________________________________________________________________________________________________
Write-Host "Sending Mail to $to,$cc" -ForegroundColor yellow

$Subject     = "Intune Patching- Windows Patching Compliance Report against $LatestDate (Patch Tuesday)" # $LatestDate is variable to pull latest patch tuesday date

Send-MailMessage -From $From -to $To -CC $CC -Subject $Subject  -Body $EmailBody1 -BodyAsHtml  -SmtpServer $SMTPServer -Port $Port -Attachments $Final_Patching_Report

Write-Host "Mail successfully sent to $to,$cc" -ForegroundColor Green
Write-Host "$LatestDate, Intune Patching compliance Report is avaialbe at this location:-" $Final_Patching_Report -ForegroundColor Green
Write-Host "===============================Phase-4 E-mailSent (Windows Patching Compliance Report)====================================(Completed)" -ForegroundColor Green

$endTime = Get-Date
$duration = $endTime - $startTime

Write-Host "Time duation to successfully excute this script is:- $duration" -ForegroundColor Green
#Invoke-Item -Path $Final_Patching_Report
