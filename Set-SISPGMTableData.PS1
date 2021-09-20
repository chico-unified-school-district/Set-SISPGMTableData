<#
.SYNOPSIS
 Insert data into SIS pertaining to student quarantine status and parent preference
.DESCRIPTION
Read various CSV files produced each day by staff and automated processes, compare data to detemine
valid references/data, and insert valid records
into the SIS database.
.EXAMPLE
.\Set-SISPGMTableData.PS1 -SQLServer sisServer -SQLDatabase sisDatabase -SQLCredential $SQLCred -CSVRoot '\\server\share' -FileSystemCredential $fscred
.EXAMPLE
.\Set-SISPGMTableData.PS1 -SQLServer sisServer -SQLDatabase sisDatabase -SQLCredential $SQLCred -CSVRoot '\\server\share' -FileSystemCredential $fscred -WhatIf
.EXAMPLE
.\Set-SISPGMTableData.PS1 -SQLServer sisServer -SQLDatabase sisDatabase -SQLCredential $SQLCred -CSVRoot '\\server\share' -FileSystemCredential $fscred -Verbose -WhatIf
.INPUTS
SQL Server name
SQL Server Database
SQL Server Credential
Folder Path where csv files are located
Credential with permission to read filesystem of csv folder
.OUTPUTS
Database records are added
INSERT statements are output to console for logging
Reference entry count
Notification entry counts
Records processed count
.NOTES
Something like functional programming was attempted in this process.
Certain logic was better suited to imperative code and so that's what was used.
#>

[cmdletbinding()]
param (
 [Parameter(Mandatory = $true)]
 [Alias('Server')]
 [ValidateScript( { Test-Connection -ComputerName $_ -Quiet -Count 2 })]
 [string]$SQLServer,
 [Parameter(Mandatory = $true)]
 [Alias('Database')]
 [string]$SQLDatabase,
 [Parameter(Mandatory = $true)]
 [Alias('DBCred')]
 [System.Management.Automation.PSCredential]$SQLCredential,
 [Parameter(Mandatory = $true)]
 [Alias('FSCred')]
 [System.Management.Automation.PSCredential]$FileSystemCredential,
 [Parameter(Mandatory = $true)]
 [string]$CSVRoot,
 [SWITCH]$WhatIf
)

# Variable
$sqlParams = @{
 ServerInstance = $SQLServer
 Database       = $SQLDatabase
 Credential     = $SQLCredential
}

if (-not(Get-Module -Name SqlServer -ListAvailable)) {
 Install-Module -Name SqlServer -Scope CurrentUser -Confirm:$false -AllowClobber -Force
}
Import-Module -Name SqlServer -Cmdlet 'Invoke-SqlCmd' -ErrorAction Stop | Out-Null

# if (Get-PSDrive -Name csvshare | Remove-PSDrive -Confirm:$false -Force -ErrorAction SilentlyContinue -ErrorVariable derp
Write-Verbose "Adding PSDrive"
New-PSDrive -name csvshare -Root $CSVRoot -PSProvider FileSystem -Credential $FileSystemCredential | Out-Null

function checkData {
 process {
  $id = $_.'Student ID'
  $refDate = getCCRefDate $ccRefList $id | Select-Object -First 1
  if (($refDate) -and ($refDate -ne 'NULL') -and ([int]$id -match "\d{5,6}")) { $_ }
 }
}
function getFullName {
 process {
  ( Get-ChildItem -Path csvshare: -Filter $_ ).Fullname
 }
}
function formatDate ($someDate, $days) {
 process {
  if (($null -eq $someDate) -or ('' -eq $someDate) -or ('NULL' -eq $someDate)) { 'NULL' }
  else { Get-Date -Date ((Get-Date -Date $someDate).AddDays($days)) -Format yyyy-MM-dd }
 }
}
function formatDupeCheckSQL ([int]$id, $code, $startDate) {
 "SELECT pid FROM PGM WHERE pid = $id AND CD = `'$code`' AND PSD = `'$startDate`';"
}
function formatSQL ([int]$scl, [int]$id, [int]$sq, $code, $startDate, $endDate) {
 "INSERT INTO PGM(SCL,PID, SQ, CD, ESD, PSD, PED) VALUES ($scl,$id,$sq,`'$code`',`'$startDate`',`'$startDate`',`'$endDate`');"
}
function getSCL ([int]$id) {
 $sql = "SELECT TOP 1 SC FROM STU WHERE ID = $id ORDER BY SC ASC"
 (Invoke-Sqlcmd @sqlParams -Query $sql).SC | Select-Object -First 1
}
function getCCRefDate ($refList, $id) {
 ($refList.Where( { $_.pid -eq $id | Where-Object { ($null -ne $_.psd) -or ($_.psd -ne 'null') } |
   Sort-Object -Property psd -Descending | Select-Object -First 1 })).psd
}
function getCode ($val) {
 switch ($val) {
  'yes' { 'MQ' }
  'no' { 'HQ' }
  default { $null }
 }
}
function getNewPGMSQ ($id) {
 # Iterate the sequence number (SQ) to ensure a unique insert
 $i = (Invoke-Sqlcmd @sqlParams -Query "Select TOP 1 SQ FROM PGM WHERE PID = $id ORDER BY SQ DESC").SQ
 if ((0 -eq $i) -or ($null -eq $i)) { 0 } else { $i++; $i }
}

function insertData {
 begin { $i = 0; 'begin processing' }
 process {
  $id = $_.'Student ID'
  $code = getCode $_.'Approved'
  Write-Verbose "$id $code"
  $startRefDate = getCCRefDate $ccRefList $id | Select-Object -First 1
  $startDate = formatDate $startRefDate
  $endDate = formatDate $startRefDate 10

  # check for duplicate record

  $dupeCheckSQL = formatDupeCheckSQL -id $id -code $code -startDate $startDate
  $dupeCheck = Invoke-Sqlcmd @sqlParams -Query $dupeCheckSQL
  if ($dupeCheck) {
   '{0} {1} {2} Possible Duplicate' -f $id, $code, $startDate
  }
  else {
   $sqNum = getNewPGMSQ -id $id
   $scl = getSCL -id $id
   ($insertSQL = formatSQL -scl $scl -id $id -sq $sqNum -code $code -startDate $startDate -endDate $endDate)
   if (-not($WhatIf)) { Invoke-Sqlcmd @sqlParams -Query $insertSQL }
  }
  $i++
 }
 end { 'Processed Records: ' + $i; 'End insertData' }
}
function testRun { if ($WhatIf) { '[TEST RUN]' } }
# Processing
testRun

# Create reference object from all CC*.csv files in csvshare
'Ref Lists'
$refPaths = 'CC_ELM.csv', 'CC_NOVAC.csv' | getFullName
$ccRefList = Import-CSV -Path $refPaths | Where-Object { ($_.psd -ne 'NULL') -and ($null -ne $_.psd) }
'Total Refs: ' + $ccRefList.count

# Process all entries from the notification file in cvsshare
'Process Response notifications file'
$responsePaths = '*covid-19*.csv' | getFullName
$responseList = Import-CSV -Path $responsePaths | Where-Object { ($_.'Approved' -eq 'yes') -or ($_.'Approved' -eq 'no') }
'Total Responses: ' + $responseList.count 
$responseList | checkData | insertData

testRun