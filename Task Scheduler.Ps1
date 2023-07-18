<# 
.SYNOPSIS
<Intune_Patch_Compliance_Calculation_&_Email_Subscription Using PowerShell>
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
<Provide all required inforamtion in User Input Section-line No 46-56 >
#>
<# 
.OUTPUTS
<You will get Intune_Patch_Compliance_Calculation_&_Email_Subscription + report in CSV>
#>
<# 
.NOTES
Version:         1.0
Author:          Chander Mani Pandey
Creation Date:   16 July 2023
Find Author on 
Youtube:-        https://www.youtube.com/@chandermanipandey8763
Twitter:-        https://twitter.com/Mani_CMPandey
Facebook:-       https://www.facebook.com/profile.php?id=100087275409143&mibextid=ZbWKwL
LinkedIn:-       https://www.linkedin.com/in/chandermanipandey
Reddit:-         https://www.reddit.com/u/ChanderManiPandey 
#>
  
Set-ExecutionPolicy -ExecutionPolicy 'ByPass' -Scope 'Process' -Force -ErrorAction 'Stop' 

$error.clear() ## this is the clear error history 
cls
$ErrorActionPreference = 'SilentlyContinue';
  
 #------------------------------ User Input Section Start --------------------------------------------------

    $TaskName = "Intune_Patch_Compliance_Reporting"
    $Description = "Intune_Patch_Compliance_Reporting"
    $ScriptPath = "C:\Downloads\LatestCumulativeUpdateList\Intune-Reports-Automation_Intune_Patching_Report_&_Email_Subscription.ps1"
    $ScheduleTime = "9am"
    
 #------------------------------ User Input Section End -----------------------------------------------------

    # Create task scheduled  action
    $action = New-ScheduledTaskAction -Execute "Powershell.exe" -Argument "-NoProfile -ExecutionPolicy bypass -WindowStyle Hidden -File $ScriptPath"

    # Create trigger for scheduled task 
    $timespan = New-Timespan -minutes 5
    $triggers = @()
    $triggers += New-ScheduledTaskTrigger -Daily -At $ScheduleTime
   
    # Register scheduled task
    Register-ScheduledTask -User SYSTEM -Action $action -Trigger $triggers -TaskName "$TaskName" -Description "$Description" -Force
    Write-Host "$TaskName Succesfully created" -ForegroundColor Green