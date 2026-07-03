param(
  [ValidateSet('list', 'search', 'complete')]
  [string]$Action = 'list',

  [ValidateSet('all', 'active', 'completed', 'overdue')]
  [string]$Status = 'active',

  [int]$Limit = 50,

  [string]$Query = '',

  [string]$EntryId = ''
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

function Normalize-Text {
  param([string]$Text)

  if ([string]::IsNullOrWhiteSpace($Text)) { return '' }
  return ($Text -replace '\s+', ' ').Trim()
}

function Get-BodyPreview {
  param([string]$Body)

  $text = Normalize-Text $Body
  $text = [System.Text.RegularExpressions.Regex]::Replace($text, '[\x00-\x08\x0B\x0C\x0E-\x1F]', '')
  if ($text.Length -le 1200) { return $text }
  return $text.Substring(0, 1200)
}

function Convert-TaskStatus {
  param($Value)

  switch ([int]$Value) {
    0 { return 'not-started' }
    1 { return 'in-progress' }
    2 { return 'completed' }
    3 { return 'waiting' }
    4 { return 'deferred' }
    default { return "unknown-$Value" }
  }
}

function Convert-OutlookDate {
  param($Value)

  try {
    if ($null -eq $Value) { return $null }
    $dt = [datetime]$Value
    if ($dt.Year -le 1900) { return $null }
    return $dt.ToString('s')
  } catch {
    return $null
  }
}

function Test-TaskLikeItem {
  param($Item)

  if ($null -eq $Item) { return $false }
  return ($Item.MessageClass -like 'IPM.Task*')
}

function Test-TaskMatchesStatus {
  param(
    $Task,
    [string]$WantedStatus
  )

  $isComplete = $false
  try { $isComplete = ([int]$Task.Status -eq 2) } catch {}

  $dueDate = $null
  try { $dueDate = [datetime]$Task.DueDate } catch {}
  $hasDueDate = ($dueDate -and $dueDate.Year -gt 1900)
  $isOverdue = (-not $isComplete -and $hasDueDate -and $dueDate.Date -lt (Get-Date).Date)

  switch ($WantedStatus) {
    'all' { return $true }
    'active' { return (-not $isComplete) }
    'completed' { return $isComplete }
    'overdue' { return $isOverdue }
    default { return $true }
  }
}

function Test-TaskMatchesQuery {
  param(
    $Task,
    [string]$Text
  )

  if ([string]::IsNullOrWhiteSpace($Text)) { return $true }

  $needle = $Text.ToLowerInvariant()
  $haystack = @(
    $Task.Subject,
    $Task.Body,
    $Task.Categories,
    $Task.Owner,
    $Task.Companies
  ) -join ' '

  return $haystack.ToLowerInvariant().Contains($needle)
}

function Convert-TaskToObject {
  param($Task)

  $status = $null
  try { $status = Convert-TaskStatus $Task.Status } catch { $status = 'unknown' }

  $complete = $false
  try { $complete = ([int]$Task.Status -eq 2) } catch {}

  $dueDateRaw = $null
  try { $dueDateRaw = [datetime]$Task.DueDate } catch {}
  $dueDate = Convert-OutlookDate $dueDateRaw
  $overdue = $false
  if (-not $complete -and $dueDateRaw -and $dueDateRaw.Year -gt 1900) {
    $overdue = ($dueDateRaw.Date -lt (Get-Date).Date)
  }

  return [pscustomobject]@{
    id = $Task.EntryID
    subject = Normalize-Text $Task.Subject
    status = $status
    complete = $complete
    percentComplete = $Task.PercentComplete
    importance = $Task.Importance
    startDate = Convert-OutlookDate $Task.StartDate
    dueDate = $dueDate
    dateCompleted = Convert-OutlookDate $Task.DateCompleted
    creationTime = Convert-OutlookDate $Task.CreationTime
    lastModificationTime = Convert-OutlookDate $Task.LastModificationTime
    owner = Normalize-Text $Task.Owner
    categories = Normalize-Text $Task.Categories
    overdue = $overdue
    bodyPreview = Get-BodyPreview $Task.Body
  }
}

$outlook = New-Object -ComObject Outlook.Application
$namespace = $outlook.GetNamespace('MAPI')
$tasksFolder = $namespace.GetDefaultFolder(13)
$items = $tasksFolder.Items

# Action: complete a single task by EntryID
if ($Action -eq 'complete') {
  if ([string]::IsNullOrWhiteSpace($EntryId)) {
    [pscustomobject]@{ status = 'error'; message = 'EntryId is required for complete action' } | Write-Json
    exit 1
  }
  $found = $null
  foreach ($item in $items) {
    if (-not (Test-TaskLikeItem $item)) { continue }
    if ($item.EntryID -eq $EntryId) { $found = $item; break }
  }
  if ($null -eq $found) {
    [pscustomobject]@{ status = 'error'; message = "Task not found: $EntryId" } | Write-Json
    exit 1
  }
  $found.Status = 2
  $found.PercentComplete = 100
  $found.DateCompleted = (Get-Date)
  $found.Save()
  [pscustomobject]@{ status = 'ok'; subject = (Normalize-Text $found.Subject); entryId = $EntryId } | Write-Json
  exit 0
}

try { $items.Sort('[DueDate]', $false) } catch {}

$tasks = New-Object System.Collections.Generic.List[object]
foreach ($item in $items) {
  if ($tasks.Count -ge $Limit) { break }
  if (-not (Test-TaskLikeItem $item)) { continue }
  if (-not (Test-TaskMatchesStatus -Task $item -WantedStatus $Status)) { continue }
  if ($Action -eq 'search' -and -not (Test-TaskMatchesQuery -Task $item -Text $Query)) { continue }

  $tasks.Add((Convert-TaskToObject $item))
}

$taskArray = @($tasks.ToArray())

[pscustomobject]@{
  generatedAt = (Get-Date).ToString('s')
  source = 'Outlook Tasks'
  action = $Action
  status = $Status
  query = $Query
  limit = $Limit
  count = $taskArray.Count
  tasks = $taskArray
} | Write-Json -Depth 8
