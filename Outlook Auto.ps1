Add-Type -assembly "Microsoft.Office.Interop.Outlook"
Add-Type -assembly "System.Runtime.Interopservices"
try {
	$Outlook = [Runtime.Interopservices.Marshal]::GetActiveObject('Outlook.Application')
    $outlookWasAlreadyRunning = $true
}
catch {
	try
    {
        $Outlook = New-Object -comobject Outlook.Application
        $outlookWasAlreadyRunning = $false
    }
    catch
    {
        write-host "Something went wrong. please exit Outlook."
        exit
    }
}

# get email template
$template = get-childitem "./8H.oft"
if ((Test-Path $template.FullName) -eq $false) {
    write-host "Could not find template. do you have a template named '8H.oft' in the folder?"
    exit
}

# create appointment item from template
$Appointment = $Outlook.CreateItemFromTemplate($template.FullName.ToString())

# set appointment properties
$Appointment.Subject = "8H for Oscar Gr$([char]0x00F8)nb$([char]0x00E6)k Kronbak"
$startTime = (Get-Date).Date.AddHours(9).AddMinutes(15)
$Appointment.Start = $startTime

$endTime = (Get-Date).Date.AddHours(12).AddMinutes(5)
$Appointment.End = $endTime
$Appointment.Location = "B2.15"
$Appointment.Body = "Discuss project status with me."

# save and send the appointment :)
$Appointment.Save()