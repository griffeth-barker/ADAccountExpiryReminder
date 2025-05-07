<#
.SYNOPSIS
  This script sends reminder notifications to the managers of users whose account will soon expire.
.DESCRIPTION
  This script gets a list of user objects from Active Directory whose AccountExpirationDate is within the specified TimeSpan and
  sends an email notification to the manager with the account information.
.PARAMETER TimeSpan
  An integer indicating the time span for user password expirations to notify.
  This parameter is required and does not accept pipeline input.
.INPUTS
  None
.OUTPUTS
  None
.NOTES
  Updated by      : Griff Barker (github@griff.systems)
  Change Date     : 2025-01-14
  Purpose/Change  : Initial development

  This script is intended to be run via the Windows Task Scheduled on a server.
  This script requires the ActiveDirectory PowerShell module and permissions to query Active Directory.
.EXAMPLE
  # Send email notifications to the managers of accounts that will expired within the next 30 days
 .\Invoke-ADAccountExpiryReminder.ps1 -TimeSpan 30
#>

[CmdletBinding()]
Param (
  [Parameter()]
  [ValidateRange(0, 30)]
  [int]$TimeSpan
)

Begin {
  ## MAINTENANCE BLOCK ####################################
  # Update these variables to fit your organization's needs
  $orgSmtpServer = "smtp.domain.tld"
  $orgHelpdeskEmail = "helpdesk@domain.tld"
  $logDir = "D:\Tasks\ADAccountExpiryReminder\log"
  ## END MAINTENANCE BLOCK ###############################

  try {
    $logFile = "$($MyInvocation.MyCommand.Name.Replace(".ps1","_"))" + "$(Get-Date -Format "yyyyMMddmmss").log"
    if (-not (Test-Path "$logDir")) {
      New-Item -Path "$logDir" -ItemType Directory -Confirm:$false | Out-Null
    }

    Start-Transcript -Path "$logDir\$logFile" -Force

    if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
      throw "The ActiveDirectory Powershell module is not available. Please install it before running this script."
    }

    try {
      Import-Module ActiveDirectory
      Write-Output "Successfully imported the ActiveDirectory PowerShell module."
    }
    catch {
      Write-Error $_.Exception
    }

    $stagingTable = New-Object System.Data.DataTable
    [void]$stagingTable.Columns.Add('Username')
    [void]$stagingTable.Columns.Add('Email')
    [void]$stagingTable.Columns.Add('Expires In')
    [void]$stagingTable.Columns.Add('ManagerEmail')

    Set-Content -Path ".\statusCode" -Value "0"
  }
  catch {
    Set-Content -Path ".\statusCode" -Value "1"
  }
}
Process {
  try {
    try {
      $userList = Search-ADAccount -AccountExpiring -UsersOnly -TimeSpan "$TimeSpan" | Where-Object { $_.DistinguishedName -like "*Contractors*" }
      Write-Output "Found $($userList.Count) users with accounts expiring either $TimeSpan days out or 7 days out."
    }
    catch {
      Write-Error $_.Exception
    }

    foreach ($u in $userList) {
      $uMail = Get-ADUser -Identity $u.samAccountName -Properties mail | Select-Object -ExpandProperty mail
      $managerMail = Get-ADUser -Identity $(Get-ADUser -Identity $u.samAccountName -Properties Manager, mail | Select-Object -ExpandProperty Manager) -Properties mail | Select-Object -ExpandProperty mail
      $expiration = $u.AccountExpirationDate
      $expdays = (New-Timespan -Start (Get-Date) -End $expiration).Days

      # Provide a one-time notification at the 30 day mark, as well as daily notifications once the account will expire in 7 days.
      if ($expdays -eq $TimeSpan -or $expdays -lt 8) {
        [void]$stagingTable.Rows.Add($u.SamAccountName.ToLower(), $uMail, $expdays, $managerMail)
      }
    }

    [array]$recipients = $stagingTable.ManagerEmail | Select-Object -Unique

    foreach ($recipient in $recipients) {
      $contractors = $stagingTable.Rows | Where-Object { $_.ManagerEmail -eq $recipient }
      $sorted = $contractors | Sort-Object -Property ExpirationDays
      $msg = @"
<body>
  <p style="font-family: 'Tahoma', sans-serif, font-size: 8.5pt;"><b>The listed contractors are about to expire.</b> Account expiration typically happens at the end of the governing contract.</p>
  <p style="font-family: 'Tahoma', sans-serif, font-size: 8.5pt;">Please take one of the following actions:</p>
  <ul>
    <li style="font-family: 'Tahoma', sans-serif, font-size: 8.5pt;">If you intend to renew this contract/contractor, please work with your Contract Admin as soon as possible and complete the process described on the intranet site.</li>
    <li style="font-family: 'Tahoma', sans-serif, font-size: 8.5pt;">If you do not wish to extend or the contractor has been terminated early or is no longer with the company, please let your Contract Admin know as soon as possible to terminate access.</li>
  </ul>
  <p style="font-family: 'Tahoma', sans-serif, font-size: 8.5pt;">If you have already contacted your Contract Admin about the listed contractors, you may ignore this notification.</p>
  <p style="font-family: 'Tahoma', sans-serif, font-size: 8.5pt;">If you do no know who your Contract Admin is, please email <a href='mailto:ContractAdmins@domain.tld'>ContractAdmins@domain.tld</a>.</p>
  <table style="border: 1px solid black; border-collapse: collapse; font-family: 'Tahoma', sans-serif, font-size: 8.5pt, text-align: center;">
    <thead>
      <tr>
        <th style="border: 1px solid black; border-collapse: collapse; text-align: center; padding: 8px;">Username</th>
        <th style="border: 1px solid black; border-collapse: collapse; text-align: center; padding: 8px;">Email Address</th>
        <th style="border: 1px solid black; border-collapse: collapse; text-align: center; padding: 8px;">Expires In</th>
      </tr>
    </thead>
    <tbody>
"@

      foreach ($s in $sorted) {
        $ddisplay = if ($s.'Expires In' -eq 1) {
          "day"
        }
        else {
          "days"
        }

        $msg += "<tr>"
        $msg += "<td style='border: 1px solid black; border-collapse: collapse; text-align: center; padding: 8px;'>$($s.Username)</td>"
        $msg += "<td style='border: 1px solid black; border-collapse: collapse; text-align: center; padding: 8px;'>$($s.Email)</td>"
        $msg += "<td style='border: 1px solid black; border-collapse: collapse; text-align: center; padding: 8px;'>$($s.'Expires In') $($ddisplay)</td>"
        $msg += "</tr>"
      }

      $msg += @"
    </tbody>
  </table>
</body>
"@

        $msgParams = @{
          To         = "$recipient"            # Prod
          #To        = "admin@domain.tld"      # Debug
          From       = "$orgHelpdeskEmail"
          Subject    = "[ACTION REQUIRED] The listed contractors are about to expire"
          Body       = $msg
          BodyAsHtml = $true
          SmtpServer = "$orgSmtpServer"
        }
      try {
        Send-MailMessage @msgParams
        Write-Output "Sent notification to $($s.ManagerEmail) about $($sorted.Username.Count) upcoming contractor account expiries."
      }
      catch {
        Write-Error $_.Exception
      }
    }
    Set-Content -Path ".\statusCode" -Value "0"
  }
  catch {
    Set-Content -Path ".\statusCode" -Value "1"
  }
}

End {
  Get-ChildItem -Path "$logDir" -Filter $("$($MyInvocation.MyCommand.Name.Replace(".ps1","_"))" + "*.log") -Recurse | Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-7) } | Remove-Item -Confirm:$false -Verbose
  Stop-Transcript
}
