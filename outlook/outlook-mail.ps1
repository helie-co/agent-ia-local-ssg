param(
  [ValidateSet('list', 'read', 'search', 'resolve', 'count-older-than', 'list-older-than', 'archive', 'delete', 'draft-new', 'draft-reply', 'send-new', 'mark-read')]
  [string]$Action = 'list',

  [int]$Limit = 10,

  [int]$Days = 30,

  [string]$EntryId = '',

  [string]$BodyPath = '',

  [string]$Query = '',

  [ValidateSet('inbox', 'sent', 'archive', 'deleted', 'all')]
  [string]$Folder = 'inbox',

  [string]$To = '',

  [string]$Cc = '',

  [string]$Bcc = '',

  [string]$Subject = ''
)

$ErrorActionPreference = 'Stop'

function Write-Json {
  param(
    [Parameter(ValueFromPipeline = $true)]
    $InputObject,

    [int]$Depth = 8
  )

  process {
    $json = $InputObject | ConvertTo-Json -Depth $Depth -Compress
    [Console]::Out.WriteLine($json)
  }
}

function Get-OutlookNamespace {
  $outlook = New-Object -ComObject Outlook.Application
  return $outlook.GetNamespace('MAPI')
}

function Resolve-FullPath {
  param([string]$Path)

  $clean = $Path.Trim('"')
  if ([System.IO.Path]::IsPathRooted($clean)) {
    return [System.IO.Path]::GetFullPath($clean)
  }

  return [System.IO.Path]::GetFullPath((Join-Path (Get-Location) $clean))
}

function Get-MailItemByEntryId {
  param(
    $Namespace,
    [string]$Id
  )

  if ([string]::IsNullOrWhiteSpace($Id)) {
    throw 'EntryId obligatoire.'
  }

  return $Namespace.GetItemFromID($Id)
}

function Get-SmtpAddress {
  param($Sender)

  try {
    if ($Sender -and $Sender.Type -eq 'EX') {
      $exchangeUser = $Sender.GetExchangeUser()
      if ($exchangeUser -and -not [string]::IsNullOrWhiteSpace($exchangeUser.PrimarySmtpAddress)) {
        return $exchangeUser.PrimarySmtpAddress
      }
    }

    if ($Sender -and -not [string]::IsNullOrWhiteSpace($Sender.Address) -and $Sender.Address -like '*@*') {
      return $Sender.Address
    }
  } catch {
    # Fall back to SenderEmailAddress below.
  }

  return $null
}

function Normalize-Text {
  param([string]$Text)

  if ([string]::IsNullOrWhiteSpace($Text)) { return '' }
  return ($Text -replace '\s+', ' ').Trim()
}

function Get-BodyPreview {
  param([string]$Body)

  $text = Normalize-Text $Body
  $text = [System.Text.RegularExpressions.Regex]::Replace($text, '[\x00-\x08\x0B\x0C\x0E-\x1F]', '')
  if ($text.Length -le 1800) { return $text }
  return $text.Substring(0, 1800)
}

function Test-MeetingLikeMail {
  param($Mail)

  $subject = Normalize-Text $Mail.Subject
  $messageClass = Normalize-Text $Mail.MessageClass
  $body = Normalize-Text $Mail.Body
  $haystack = ($subject + ' ' + $messageClass + ' ' + $body).ToLowerInvariant()

  if ($messageClass -like 'IPM.Schedule.Meeting*') { return $true }
  if ($haystack -match '\b(teams|reunion|réunion|meeting|invitation|calendrier|calendar|webex|zoom)\b') { return $true }
  return $false
}

function Convert-MailToObject {
  param($Mail)

  $senderSmtp = Get-SmtpAddress $Mail.Sender
  if ([string]::IsNullOrWhiteSpace($senderSmtp)) {
    $senderSmtp = $Mail.SenderEmailAddress
  }

  return [pscustomobject]@{
    id = $Mail.EntryID
    subject = $Mail.Subject
    senderName = $Mail.SenderName
    senderEmail = $senderSmtp
    receivedTime = if ($Mail.ReceivedTime) { $Mail.ReceivedTime.ToString('s') } else { $null }
    to = $Mail.To
    cc = $Mail.CC
    unread = [bool]$Mail.UnRead
    hasAttachments = [bool]($Mail.Attachments.Count -gt 0)
    messageClass = $Mail.MessageClass
    isMeetingLike = [bool](Test-MeetingLikeMail $Mail)
    bodyPreview = Get-BodyPreview $Mail.Body
  }
}

function Test-MailLikeItem {
  param($Item)

  if ($null -eq $Item) { return $false }
  return ($Item.MessageClass -like 'IPM.Note*' -or $Item.MessageClass -like 'IPM.Schedule.Meeting*')
}

function Test-OlderThanCutoff {
  param(
    $Item,
    [datetime]$Cutoff
  )

  if ($null -eq $Item.ReceivedTime) { return $false }
  return ([datetime]$Item.ReceivedTime -lt $Cutoff)
}

function Get-ArchiveFolder {
  param($Inbox)

  $store = $Inbox.Store
  try {
    $archive = $store.GetDefaultFolder(35)
    if ($archive) { return $archive }
  } catch {
    # Some Outlook profiles do not expose olFolderArchive.
  }

  $root = $store.GetRootFolder()
  foreach ($folder in $root.Folders) {
    if ($folder.Name -match '^(Archive|Archives)$') {
      return $folder
    }
  }

  return $Inbox.Folders.Add('Archive')
}

function Get-MailFolder {
  param(
    $Namespace,
    [string]$Name
  )

  switch ($Name) {
    'inbox' { return $Namespace.GetDefaultFolder(6) }
    'sent' { return $Namespace.GetDefaultFolder(5) }
    'deleted' { return $Namespace.GetDefaultFolder(3) }
    'archive' {
      $inbox = $Namespace.GetDefaultFolder(6)
      return Get-ArchiveFolder $inbox
    }
    default { throw "Dossier non supporte: $Name" }
  }
}

function Get-MailFoldersForSearch {
  param(
    $Namespace,
    [string]$Name
  )

  if ($Name -ne 'all') {
    return @((Get-MailFolder -Namespace $Namespace -Name $Name))
  }

  $folders = New-Object System.Collections.Generic.List[object]
  foreach ($folderName in @('inbox', 'sent', 'archive', 'deleted')) {
    try {
      $folder = Get-MailFolder -Namespace $Namespace -Name $folderName
      if ($folder) { $folders.Add($folder) }
    } catch {
      # Ignore unavailable optional folders such as Archive.
    }
  }

  return @($folders.ToArray())
}

function Convert-MailToDetailObject {
  param($Mail)

  $obj = Convert-MailToObject $Mail
  $obj | Add-Member -NotePropertyName body -NotePropertyValue (Normalize-Text $Mail.Body)
  $obj | Add-Member -NotePropertyName importance -NotePropertyValue $Mail.Importance
  $obj | Add-Member -NotePropertyName size -NotePropertyValue $Mail.Size
  return $obj
}

function Test-MailMatchesQuery {
  param(
    $Mail,
    [string]$Text
  )

  if ([string]::IsNullOrWhiteSpace($Text)) { return $true }

  $needle = $Text.ToLowerInvariant()
  $senderSmtp = Get-SmtpAddress $Mail.Sender
  if ([string]::IsNullOrWhiteSpace($senderSmtp)) { $senderSmtp = $Mail.SenderEmailAddress }
  $haystack = @(
    $Mail.Subject,
    $Mail.SenderName,
    $senderSmtp,
    $Mail.To,
    $Mail.CC,
    $Mail.Body
  ) -join ' '

  return $haystack.ToLowerInvariant().Contains($needle)
}

$namespace = Get-OutlookNamespace

switch ($Action) {
  'list' {
    $inbox = $namespace.GetDefaultFolder(6)
    $items = $inbox.Items
    $items.Sort('[ReceivedTime]', $true)

    $mails = New-Object System.Collections.Generic.List[object]
    foreach ($item in $items) {
      if ($mails.Count -ge $Limit) { break }
      if ($item.MessageClass -notlike 'IPM.Note*' -and $item.MessageClass -notlike 'IPM.Schedule.Meeting*') { continue }
      $mails.Add((Convert-MailToObject $item))
    }

    $mailArray = @($mails.ToArray())

    [pscustomobject]@{
      generatedAt = (Get-Date).ToString('s')
      source = 'Outlook Inbox'
      limit = $Limit
      count = $mailArray.Count
      mails = $mailArray
    } | Write-Json -Depth 8
  }

  'read' {
    $mail = Get-MailItemByEntryId -Namespace $namespace -Id $EntryId

    [pscustomobject]@{
      generatedAt = (Get-Date).ToString('s')
      source = 'Outlook Mail'
      mail = Convert-MailToDetailObject $mail
    } | Write-Json -Depth 8
  }

  'search' {
    $folders = Get-MailFoldersForSearch -Namespace $namespace -Name $Folder
    $mails = New-Object System.Collections.Generic.List[object]

    foreach ($mailFolder in $folders) {
      $items = $mailFolder.Items
      try { $items.Sort('[ReceivedTime]', $true) } catch { try { $items.Sort('[SentOn]', $true) } catch {} }

      foreach ($item in $items) {
        if ($mails.Count -ge $Limit) { break }
        if (-not (Test-MailLikeItem $item)) { continue }
        if (-not (Test-MailMatchesQuery -Mail $item -Text $Query)) { continue }

        $obj = Convert-MailToObject $item
        $obj | Add-Member -NotePropertyName folder -NotePropertyValue $mailFolder.Name
        $mails.Add($obj)
      }

      if ($mails.Count -ge $Limit) { break }
    }

    $mailArray = @($mails.ToArray())

    [pscustomobject]@{
      generatedAt = (Get-Date).ToString('s')
      source = 'Outlook Search'
      query = $Query
      folder = $Folder
      limit = $Limit
      count = $mailArray.Count
      mails = $mailArray
    } | Write-Json -Depth 8
  }

  'resolve' {
    if ([string]::IsNullOrWhiteSpace($Query)) {
      throw 'Query obligatoire pour resolve. Utiliser -Query "Nom à chercher".'
    }

    $results = New-Object System.Collections.Generic.List[object]
    $needle = $Query.ToLowerInvariant()

    foreach ($list in $namespace.AddressLists) {
      if ($list.Name -match '^(Global Address List|GAL|Contacts|CarNet d.adresses)$') {
        foreach ($entry in $list.AddressEntries) {
          try {
            $name = Normalize-Text $entry.Name
            $addr = ''
            if ($entry.Type -eq 'EX') {
              $exUser = $entry.GetExchangeUser()
              if ($exUser) {
                $addr = $exUser.PrimarySmtpAddress
                $dept = Normalize-Text $exUser.Department
                $title = Normalize-Text $exUser.JobTitle
              }
            } elseif ($entry.Address -like '*@*') {
              $addr = $entry.Address
            }

            if ($name.ToLowerInvariant().Contains($needle) -or $addr.ToLowerInvariant().Contains($needle)) {
              $results.Add([pscustomobject]@{
                name = $name
                email = $addr
                type = $entry.Type
                department = if ($dept) { $dept } else { $null }
                title = if ($title) { $title } else { $null }
              })
            }
          } catch {
            # Skip entries that error (distribution lists, invalid entries, etc.)
          }
        }
      }
    }

    [pscustomobject]@{
      generatedAt = (Get-Date).ToString('s')
      source = 'Outlook Address Book'
      query = $Query
      count = $results.Count
      contacts = @($results.ToArray())
    } | Write-Json -Depth 6
  }

  'count-older-than' {
    $inbox = $namespace.GetDefaultFolder(6)
    $items = $inbox.Items
    $cutoff = (Get-Date).AddDays(-1 * $Days)
    $count = 0

    foreach ($item in $items) {
      if (-not (Test-MailLikeItem $item)) { continue }
      if (Test-OlderThanCutoff -Item $item -Cutoff $cutoff) { $count++ }
    }

    [pscustomobject]@{
      generatedAt = (Get-Date).ToString('s')
      source = 'Outlook Inbox'
      days = $Days
      cutoff = $cutoff.ToString('s')
      count = $count
    } | Write-Json -Depth 4
  }

  'list-older-than' {
    $inbox = $namespace.GetDefaultFolder(6)
    $items = $inbox.Items
    $items.Sort('[ReceivedTime]', $false)
    $cutoff = (Get-Date).AddDays(-1 * $Days)

    $mails = New-Object System.Collections.Generic.List[object]
    foreach ($item in $items) {
      if ($mails.Count -ge $Limit) { break }
      if (-not (Test-MailLikeItem $item)) { continue }
      if (-not (Test-OlderThanCutoff -Item $item -Cutoff $cutoff)) { continue }
      $mails.Add((Convert-MailToObject $item))
    }

    $mailArray = @($mails.ToArray())

    [pscustomobject]@{
      generatedAt = (Get-Date).ToString('s')
      source = 'Outlook Inbox'
      days = $Days
      cutoff = $cutoff.ToString('s')
      limit = $Limit
      count = $mailArray.Count
      mails = $mailArray
    } | Write-Json -Depth 8
  }

  'archive' {
    $mail = Get-MailItemByEntryId -Namespace $namespace -Id $EntryId
    $inbox = $namespace.GetDefaultFolder(6)
    $archive = Get-ArchiveFolder $inbox
    $moved = $mail.Move($archive)
    Write-Host "Archive: $($moved.Subject)"
  }

  'delete' {
    $mail = Get-MailItemByEntryId -Namespace $namespace -Id $EntryId
    $subject = $mail.Subject
    $mail.Delete()
    Write-Host "Supprime vers Elements supprimes: $subject"
  }

  'draft-new' {
    if ([string]::IsNullOrWhiteSpace($To)) {
      throw 'To obligatoire pour draft-new.'
    }
    if ([string]::IsNullOrWhiteSpace($Subject)) {
      throw 'Subject obligatoire pour draft-new.'
    }
    if ([string]::IsNullOrWhiteSpace($BodyPath)) {
      throw 'BodyPath obligatoire pour draft-new.'
    }

    $bodyFile = Resolve-FullPath $BodyPath
    if (-not (Test-Path -LiteralPath $bodyFile)) {
      throw "Fichier de corps introuvable: $bodyFile"
    }

    $body = Get-Content -LiteralPath $bodyFile -Raw -Encoding UTF8
    $outlook = New-Object -ComObject Outlook.Application
    $mail = $outlook.CreateItem(0)
    $mail.To = $To
    if (-not [string]::IsNullOrWhiteSpace($Cc)) { $mail.CC = $Cc }
    if (-not [string]::IsNullOrWhiteSpace($Bcc)) { $mail.BCC = $Bcc }
    $mail.Subject = $Subject
    $mail.Body = $body.Trim()
    $mail.Save()
    $mail.Display($false)

    Write-Host "Brouillon ouvert: $Subject -> $To"
  }

  'draft-reply' {
    if ([string]::IsNullOrWhiteSpace($BodyPath)) {
      throw 'BodyPath obligatoire pour draft-reply.'
    }

    $bodyFile = Resolve-FullPath $BodyPath
    if (-not (Test-Path -LiteralPath $bodyFile)) {
      throw "Fichier de reponse introuvable: $bodyFile"
    }

    $body = Get-Content -LiteralPath $bodyFile -Raw -Encoding UTF8
    $mail = Get-MailItemByEntryId -Namespace $namespace -Id $EntryId
    $reply = $mail.Reply()
    $reply.Body = $body.Trim() + "`r`n`r`n" + $reply.Body
    $reply.Save()
    $reply.Display($false)
    Write-Host "Brouillon ouvert: $($reply.Subject)"
  }

  'send-new' {
    if ([string]::IsNullOrWhiteSpace($To)) {
      throw 'To obligatoire pour send-new.'
    }
    if ([string]::IsNullOrWhiteSpace($Subject)) {
      throw 'Subject obligatoire pour send-new.'
    }
    if ([string]::IsNullOrWhiteSpace($BodyPath)) {
      throw 'BodyPath obligatoire pour send-new.'
    }

    $bodyFile = Resolve-FullPath $BodyPath
    if (-not (Test-Path -LiteralPath $bodyFile)) {
      throw "Fichier de corps introuvable: $bodyFile"
    }

    $body = Get-Content -LiteralPath $bodyFile -Raw -Encoding UTF8
    $outlook = New-Object -ComObject Outlook.Application
    $mail = $outlook.CreateItem(0)
    $mail.To = $To
    if (-not [string]::IsNullOrWhiteSpace($Cc)) { $mail.CC = $Cc }
    if (-not [string]::IsNullOrWhiteSpace($Bcc)) { $mail.BCC = $Bcc }
    $mail.Subject = $Subject
    $mail.Body = $body.Trim()
    $mail.Send()

    Write-Host "Mail envoye: $Subject -> $To"
  }

  'mark-read' {
    $mail = Get-MailItemByEntryId -Namespace $namespace -Id $EntryId
    $mail.UnRead = $false
    $mail.Save()
    Write-Host "Marque comme lu: $($mail.Subject)"
  }
}
