# PowerShell 2 only
$PSScriptRoot = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition

$msg = [String] $(Get-Content $PSScriptRoot\main.html) # Main message body
$pwdMsg = [String] $(Get-Content $PSScriptRoot\password.html) # Password message body
$msgSubject = "Some Subject"
$pwdMsgSubject = "Password Message Subject"
$cc = "xyz@abc.com"
$workingFolder = "$HOME\Desktop\Temp"
$zipName = "SomeName.zip"
$attachment = "$workingFolder\$zipName"

### Password generation ###

Function GET-Temppassword() {
    Param(
        [int]$length=10,
        [string[]]$sourcedata
    )

    For ($loop=1; $loop –le $length; $loop++) {
        $TempPassword+=($sourcedata | GET-RANDOM)
    }
    return $TempPassword
}

$ascii=$NULL;
For ($a=48;$a –le 122;$a++) {$ascii+=,[char][byte]$a }
$password = GET-Temppassword –length 9 –sourcedata $ascii
$password | clip

### Zip creation ###

echo "Creating ZIP file: $attachment"
echo "With Password: $password"
echo ""
zip.exe -j -P $password $attachment "$workingFolder\*.*"
echo ""

### Main message generation ###

$outlook = New-Object -ComObject Outlook.Application 
$mailMsg = $outlook.CreateItem(0)
$inspector = $mailMsg.GetInspector
$inspector.Activate()

$mailMsg.CC = $cc
$mailMsg.Subject = $msgSubject
$signature = $mailMsg.HTMLBody
$mailMsg.HTMLBody = $msg + $signature
$mailMsg.Attachments.Add($attachment) | out-null

### Password message generation ###

$pwdMailMsg = $outlook.CreateItem(0)
$inspector = $pwdMailMsg.GetInspector
$inspector.Activate() | out-null

$pwdMailMsg.CC = $cc
$pwdMailMsg.Subject = $pwdMsgSubject
$signature = $pwdMailMsg.HTMLBody
$pwdMailMsg.HTMLBody = $pwdMsg + $password + $signature

### Clean-up ###

echo "Clearing folder $workingFolder"
Remove-Item "$workingFolder\*.*" -Force