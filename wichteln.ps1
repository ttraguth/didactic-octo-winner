#requires -version 2
<#
.SYNOPSIS
  Selects random pairs of names and sends email to inform user of result.
  Good for Wichteln
.DESCRIPTION
  
.INPUTS
  
.OUTPUTS
  <Outputs if any, otherwise state None - example: Log file stored in C:\Windows\Temp\<name>.log>
.NOTES
  Version:        1.0
  Author:         Tobias Traguth
  Creation Date:  23.11.2018
  Purpose/Change: Initial script development
  
.EXAMPLE
  
#>
$random = $null
$selected = New-Object System.Collections.ArrayList
# hashtable of names and email
$names = @{Name1="Name1@email.de";Name2="Name2@email.de"}
# unallowed combination
$nogo = @{Name1="Name2"}

# email credentials
$secpasswd = ConvertTo-SecureString $emailpassword -AsPlainText -Force
$mycreds = New-Object System.Management.Automation.PSCredential ($emailuser, $secpasswd)

# algorithm
foreach($name in $names.Keys) 
{
    $found = $false
    while (-not $found) 
    {
        $random = Get-Random -InputObject ((Compare-Object -ReferenceObject ($names.Keys |ForEach-Object ToString) -DifferenceObject $selected -PassThru) |ForEach-Object ToString)        

        if ($name -ne $random -and $selected -notcontains $random -and ($nogo[$name] -ne $random ))
        {
            $selected += $random
            $found = $true
            Write-Host Ein Geschenk von $name an $random
            $Body = "<p>Hallo $name,</p><p>Du darfst ein Geschenk an <b>$random</b> schenken"
            Send-MailMessage -To $names[$name] -Subject $ -Body $Body -SmtpServer $smtpserver -From $from -DeliveryNotificationOption OnFailure, OnSuccess -BodyAsHtml -Credential $mycreds -UseSsl -Port $smtpport 
            break
        }
        else {
            Write-Host "Skip Name $name Random $random"
            continue
        }
    }
}