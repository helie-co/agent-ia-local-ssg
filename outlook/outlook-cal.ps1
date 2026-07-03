param(
  # --- Mode: 'list' (default) or 'create' ---
  [string]$Action         = 'list',

  # --- list parameters ---
  [string]$StartDate      = '2023-09-01',
  [string]$EndDate        = '',
  [switch]$IncludeEntryId,

  # --- create parameters ---
  [string]$Subject        = '',
  [string]$Start          = '',   # e.g. '2026-07-06 09:00'
  [string]$End            = '',   # e.g. '2026-07-06 09:30'
  [string]$Location       = '',
  [string]$Body           = '',
  [string]$Required       = '',   # comma-separated attendee emails (required)
  [string]$Optional       = '',   # comma-separated attendee emails (optional)
  [switch]$Teams                  # create as online Teams meeting
)

$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

function Write-Json {
  param(
    [Parameter(ValueFromPipeline = $true)]
    $InputObject,

    [int]$Depth = 10
  )

  process {
    $json = $InputObject | ConvertTo-Json -Depth $Depth -Compress
    [Console]::Out.WriteLine($json)
  }
}

function Normalize-Text {
  param([string]$Text)

  if ([string]::IsNullOrWhiteSpace($Text)) { return '' }
  return ($Text -replace '\s+', ' ').Trim()
}

function Get-BodyPreview {
  param([string]$Body)

  $text = Normalize-Text $Body
  if ($text.Length -le 1800) { return $text }
  return $text.Substring(0, 1800)
}

function Convert-ToSlug {
  param([string]$Text)

  $s = $Text.ToLowerInvariant()

  # Transliterate common accented characters
  $map = @{
    'à'='a'; 'â'='a'; 'ä'='a'; 'á'='a'; 'ã'='a';
    'è'='e'; 'é'='e'; 'ê'='e'; 'ë'='e';
    'î'='i'; 'ï'='i'; 'í'='i'; 'ì'='i';
    'ô'='o'; 'ö'='o'; 'ó'='o'; 'ò'='o'; 'õ'='o';
    'ù'='u'; 'û'='u'; 'ü'='u'; 'ú'='u';
    'ç'='c'; 'ñ'='n'; 'œ'='oe'; 'æ'='ae';
  }
  foreach ($key in $map.Keys) { $s = $s.Replace($key, $map[$key]) }

  # Keep only alphanumeric and hyphens
  $s = $s -replace '[^a-z0-9]+', '-'
  $s = $s.Trim('-')
  if ($s.Length -gt 70) { $s = $s.Substring(0, 70).TrimEnd('-') }
  return $s
}

function Get-SmtpAddress {
  param($Recipient)

  try {
    $entry = $Recipient.AddressEntry
    if ($entry) {
      if ($entry.Type -eq 'EX') {
        $exchangeUser = $entry.GetExchangeUser()
        if ($exchangeUser -and -not [string]::IsNullOrWhiteSpace($exchangeUser.PrimarySmtpAddress)) {
          return @{
            name  = Normalize-Text $exchangeUser.Name
            email = $exchangeUser.PrimarySmtpAddress
          }
        }

        $exchangeList = $entry.GetExchangeDistributionList()
        if ($exchangeList -and -not [string]::IsNullOrWhiteSpace($exchangeList.PrimarySmtpAddress)) {
          return @{
            name  = Normalize-Text $exchangeList.Name
            email = $exchangeList.PrimarySmtpAddress
          }
        }
      }

      if (-not [string]::IsNullOrWhiteSpace($entry.Address) -and $entry.Address -like '*@*') {
        return @{
          name  = Normalize-Text $Recipient.Name
          email = $entry.Address
        }
      }
    }
  } catch {
    # Fall through
  }

  if (-not [string]::IsNullOrWhiteSpace($Recipient.Address) -and $Recipient.Address -like '*@*') {
    return @{
      name  = Normalize-Text $Recipient.Name
      email = $Recipient.Address
    }
  }

  return @{
    name  = Normalize-Text $Recipient.Name
    email = $null
  }
}

function Format-Participant {
  param($p)

  if ([string]::IsNullOrWhiteSpace($p.email)) {
    return $p.name
  }
  if ([string]::IsNullOrWhiteSpace($p.name)) {
    return $p.email
  }
  return "$($p.name) <$($p.email)>"
}

function Test-IsTeams {
  param(
    [string]$Location,
    [string]$Body
  )

  $haystack = ($Location + ' ' + $Body).ToLowerInvariant()
  return $haystack -match 'teams\.microsoft\.com|meet\.teams\.microsoft\.com|microsoftteams'
}

function Convert-EventToObject {
  param($Item, [switch]$IncludeEntryId)

  # --- titre ---
  $subject = Normalize-Text $Item.Subject

  # --- dates ---
  $startDt = $null
  $endDt   = $null
  try { $startDt = [datetime]$Item.Start } catch {}
  try { $endDt   = [datetime]$Item.End   } catch {}

  $startStr = if ($startDt) { $startDt.ToString('s') } else { $null }
  $endStr   = if ($endDt)   { $endDt.ToString('s')   } else { $null }

  # --- organisateur ---
  $organizerName  = Normalize-Text $Item.Organizer
  $organizerEmail = $null
  try {
    $sender = $Item.GetOrganizer()
    if ($sender) {
      if ($sender.Type -eq 'EX') {
        $ex = $sender.GetExchangeUser()
        if ($ex) { $organizerEmail = $ex.PrimarySmtpAddress }
      } elseif ($sender.Address -like '*@*') {
        $organizerEmail = $sender.Address
      }
    }
  } catch {}

  $organizerStr = if ($organizerEmail) { "$organizerName <$organizerEmail>" } else { $organizerName }

  # --- participants requis ---
  $required = New-Object System.Collections.Generic.List[string]
  $optional = New-Object System.Collections.Generic.List[string]

  foreach ($recipient in $Item.Recipients) {
    $p = Get-SmtpAddress $recipient
    $display = Format-Participant $p

    # RecipientType: 1=Required, 2=Optional, 3=Resource
    try {
      $type = $recipient.Type
    } catch {
      $type = 1
    }

    if ($type -eq 2) {
      if (-not $optional.Contains($display)) { $optional.Add($display) }
    } else {
      if (-not $required.Contains($display)) { $required.Add($display) }
    }
  }

  # --- lieu ---
  $location = Normalize-Text $Item.Location

  # --- description ---
  $body = ''
  try { $body = Get-BodyPreview $Item.Body } catch {}

  # --- Teams vs présentielle ---
  $isTeams = Test-IsTeams -Location $location -Body $body
  $format  = if ($isTeams) { 'Teams' } else { if ([string]::IsNullOrWhiteSpace($location)) { 'Inconnu' } else { 'Presentielle' } }

  # --- statut annulé ---
  $cancelled = $false
  try {
    # olMeetingCanceled = 5, olMeetingReceived = 1
    $meetingStatus = $Item.MeetingStatus
    $cancelled = ($meetingStatus -eq 5)
  } catch {}

  # --- slug pour le nom de fichier ---
  # Format: yyyy-MM-dd-HHhMM-sujet  (ex: 2026-06-19-14h00-reunion-equipe)
  $slug = ''
  if ($startDt) {
    $datePart = $startDt.ToString('yyyy-MM-dd') + '-' + $startDt.ToString('HH') + 'h' + $startDt.ToString('mm')
    $slug = $datePart + '-' + (Convert-ToSlug $subject)
  } else {
    $slug = 'sans-date-' + (Convert-ToSlug $subject)
  }

  $result = [ordered]@{
    slug             = $slug
    subject          = $subject
    start            = $startStr
    end              = $endStr
    organizer        = $organizerStr
    required         = @($required.ToArray())
    optional         = @($optional.ToArray())
    location         = $location
    format           = $format
    isTeams          = $isTeams
    cancelled        = $cancelled
    bodyPreview      = $body
  }

  if ($IncludeEntryId) {
    $entryId = $null
    try { $entryId = $Item.EntryID } catch {}
    $result['entryId'] = $entryId
  }

  return [pscustomobject]$result
}

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

$outlook   = New-Object -ComObject Outlook.Application
$namespace = $outlook.GetNamespace('MAPI')
$calendar  = $namespace.GetDefaultFolder(9)   # olFolderCalendar = 9

# ---------------------------------------------------------------------------
# Action: create
# ---------------------------------------------------------------------------

if ($Action -eq 'create') {
  if ([string]::IsNullOrWhiteSpace($Subject)) { throw '-Subject is required for action create.' }
  if ([string]::IsNullOrWhiteSpace($Start))   { throw '-Start is required for action create.' }
  if ([string]::IsNullOrWhiteSpace($End))     { throw '-End is required for action create.' }

  $appt = $outlook.CreateItem(1)   # 1 = olAppointmentItem
  $appt.Subject = $Subject
  $appt.Start   = [datetime]::Parse($Start)
  $appt.End     = [datetime]::Parse($End)

  if (-not [string]::IsNullOrWhiteSpace($Location)) { $appt.Location = $Location }
  if (-not [string]::IsNullOrWhiteSpace($Body))     { $appt.Body     = $Body     }

  $hasAttendees = (-not [string]::IsNullOrWhiteSpace($Required)) -or (-not [string]::IsNullOrWhiteSpace($Optional))
  if ($hasAttendees) {
    $appt.MeetingStatus = 1   # 1 = olMeeting

    foreach ($email in ($Required -split ',')) {
      $email = $email.Trim()
      if ([string]::IsNullOrWhiteSpace($email)) { continue }
      $recip = $appt.Recipients.Add($email)
      $recip.Type = 1   # olRequired
    }

    foreach ($email in ($Optional -split ',')) {
      $email = $email.Trim()
      if ([string]::IsNullOrWhiteSpace($email)) { continue }
      $recip = $appt.Recipients.Add($email)
      $recip.Type = 2   # olOptional
    }

    $appt.Recipients.ResolveAll() | Out-Null
  }

  if ($Teams) {
    # --- Création via Microsoft Graph pour obtenir un vrai lien Teams ---
    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
    Import-Module Microsoft.Graph.Calendar -ErrorAction Stop
    Connect-MgGraph -Scopes 'Calendars.ReadWrite' -NoWelcome

    $graphContext = Get-MgContext
    $userId = $graphContext.Account

    $startDt = [datetime]::Parse($Start)
    $endDt   = [datetime]::Parse($End)

    $attendeesList = [System.Collections.Generic.List[object]]::new()
    foreach ($email in ($Required -split ',')) {
      $email = $email.Trim()
      if ([string]::IsNullOrWhiteSpace($email)) { continue }
      $attendeesList.Add(@{ emailAddress = @{ address = $email }; type = 'required' })
    }
    foreach ($email in ($Optional -split ',')) {
      $email = $email.Trim()
      if ([string]::IsNullOrWhiteSpace($email)) { continue }
      $attendeesList.Add(@{ emailAddress = @{ address = $email }; type = 'optional' })
    }

    $graphBody = @{
      subject               = $Subject
      start                 = @{ dateTime = $startDt.ToString('s'); timeZone = 'Romance Standard Time' }
      end                   = @{ dateTime = $endDt.ToString('s');   timeZone = 'Romance Standard Time' }
      isOnlineMeeting       = $true
      onlineMeetingProvider = 'teamsForBusiness'
      attendees             = $attendeesList.ToArray()
    }
    if (-not [string]::IsNullOrWhiteSpace($Location)) { $graphBody['location'] = @{ displayName = $Location } }
    if (-not [string]::IsNullOrWhiteSpace($Body))     { $graphBody['body']     = @{ contentType = 'text'; content = $Body } }

    $event = New-MgUserEvent -UserId $userId -BodyParameter $graphBody

    [pscustomobject]@{
      status      = 'created'
      entryId     = $event.Id
      subject     = $event.Subject
      start       = $event.Start.DateTime
      end         = $event.End.DateTime
      location    = $event.Location.DisplayName
      isMeeting   = ($attendeesList.Count -gt 0)
      isTeams     = $true
      teamsLink   = $event.OnlineMeeting.JoinUrl
    } | Write-Json
    exit 0
  }

  $appt.Save()

  [pscustomobject]@{
    status    = 'created'
    entryId   = $appt.EntryID
    subject   = $appt.Subject
    start     = ([datetime]$appt.Start).ToString('s')
    end       = ([datetime]$appt.End).ToString('s')
    location  = $appt.Location
    isMeeting = $hasAttendees
    isTeams   = $false
  } | Write-Json
  exit 0
}

# ---------------------------------------------------------------------------
# Action: list (default)
# ---------------------------------------------------------------------------

$dtStart = [datetime]::Parse($StartDate).Date
$dtEnd   = if ([string]::IsNullOrWhiteSpace($EndDate)) { $dtStart.AddDays(1) } else { [datetime]::Parse($EndDate).Date.AddDays(1) }

$items = $calendar.Items
$items.IncludeRecurrences = $true
$items.Sort('[Start]')

# Outlook Restrict uses locale-sensitive date parsing.
# Use the current culture format to avoid month/day swaps.
$filterStart = $dtStart.ToString('g')
$filterEnd   = $dtEnd.ToString('g')
$filter = "[Start] >= '$filterStart' AND [Start] < '$filterEnd'"

$restricted = $items.Restrict($filter)

$events = New-Object System.Collections.Generic.List[object]
foreach ($item in $restricted) {
  # Skip non-appointment items (e.g. tasks that leaked into calendar view)
  $class = ''
  try { $class = $item.MessageClass } catch {}
  if ($class -notlike 'IPM.Appointment*' -and $class -ne '') { continue }

  $obj = Convert-EventToObject $item -IncludeEntryId:$IncludeEntryId
  $events.Add($obj)
}

$eventArray = @($events.ToArray())

[pscustomobject]@{
  generatedAt  = (Get-Date).ToString('s')
  source       = 'Outlook Calendar'
  startDate    = $dtStart.ToString('s')
  endDate      = $dtEnd.AddDays(-1).ToString('s')
  count        = $eventArray.Count
  events       = $eventArray
} | Write-Json -Depth 10
