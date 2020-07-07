Param(
    [string]$wks
)

# Get XML config

$confPath = "$env:SCRIPTS\security\logonStatusTracker\config.xml"

[xml]$conf = Get-Content $confPath

# Send-MailMessage parameters

$mailEnc= [System.Text.Encoding]::UTF8
$mailSrv = $conf.Settings.Email.Server
$mailFrom = $conf.Settings.Email.From
$mailTo = $conf.Settings.Email.To
$mailSubject = "$wks local user logon report"
$mailPort = $conf.Settings.Email.Port

Function Get-CurrentUser{

    Param(
    [string]$wks  
    )

    try{
    
    $getCrntUser = $((gwmi win32_computersystem -cn $wks -ErrorAction Stop -ErrorVariable $err).Username).Replace('MAIN\','')

    return $getCrntUser

    } catch {
    
    Write-Debug "$err"
    
    }
    
}

Function Get-ProfileData{


    Param(
    [string]$wks,
    [string]$usr,
    [string]$type
    )

        try{

            switch($type){
    
            'WMI'{$usrPrf = gwmi win32_userprofile -cn $wks -ErrorAction Stop -ErrorVariable $err | ? {$_.localpath -like "*$usr*"}} 

            'AD'{[Microsoft.ActiveDirectory.Management.ADAccount]$usrPrf = Get-ADUser -Identity “$usr” -Properties LastBadPasswordAttempt, LastLogonDate, PasswordLastSet -ErrorAction Stop -ErrorVariable $err}

            }

            return $usrPrf
        
        }
    
        catch{ 
            
            Write-Debug "$err"

        }
               
}

$echo = Test-Connection -ComputerName $wks -Count 3 -Quiet

if($echo) {

    Write-Debug "Host $wks is UP"

    $currentUser = Get-CurrentUser -wks $wks

    $adUsrData = Get-ProfileData -wks $wks -usr $currentUser -type AD

    $wmiUsrProfile = Get-ProfileData -wks $wks -usr $currentUser -type WMI


} else {

    Write-Debug "Host $wks is down"

    Exit

}

# mail message body

$mailMsg = $null 
$mailMsg += "`nHostname: $wks"
$mailMsg += "`n"
$mailMsg += "`nUser: $currentUser"
$mailMsg += "`n"
$mailMsg += "`nAD Last logon date: $($adUsrData.LastLogonDate)"
$mailMsg += "`n"
$mailMsg += "`nAD bad password attempt: $($adUsrData.LastBadPasswordAttempt)"
$mailMsg += "`n"
$mailMsg += "`nAD password last set: $($adUsrData.PasswordLastSet)"
$mailMsg += "`n"
$mailMsg += "`nLocal profile last use time: $(([WMI]'').ConvertToDateTime($wmiUsrProfile.LastUseTime))"
$mailMsg += "`n"
$mailMsg += "`nIs logged in: $($wmiUsrProfile.Loaded)"
$mailMsg += "`n"
$mailMsg += "`nLocal user folder path: $($wmiUsrProfile.LocalPath)"

# send report by email

Send-MailMessage -SmtpServer $mailSrv `
-Port $mailPort `
-From $mailFrom `
-To $mailTo `
-Subject $mailSubject `
-Body $mailMsg `
-Encoding $mailEnc

Write-Debug "Send-MailMessage status: $?"