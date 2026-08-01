param(
  [Parameter(Mandatory = $true)]
  [string]$FolderPath,

  [switch]$Recurse
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Get-DateFromFirstLine {
  param(
    [Parameter(Mandatory = $true)]
    [string]$FilePath
  )

  try {
    $firstLine = Get-Content `
      -LiteralPath $FilePath `
      -TotalCount 1 `
      -Encoding UTF8
  }
  catch {
    Write-Warning "ファイルを読み込めませんでした: $FilePath"
    Write-Warning $_.Exception.Message
    return $null
  }

  if ([string]::IsNullOrWhiteSpace($firstLine)) {
    return $null
  }

  # 対応例:
  # # 2026-08-01
  # # 2026/08/01
  # # 2026/08/01の日記
  # 2026-08-01
  if ($firstLine -match '(?<!\d)(?<Year>\d{4})[-/](?<Month>\d{1,2})[-/](?<Day>\d{2})(?!\d)') {
    $dateText = '{0}-{1}-{2}' -f `
      $Matches.Year,
      $Matches.Month,
      $Matches.Day

    $parsedDate = [datetime]::MinValue
    $isValidDate = [datetime]::TryParseExact(
      $dateText,
      "yyyy-MM-dd",
      [System.Globalization.CultureInfo]::InvariantCulture,
      [System.Globalization.DateTimeStyles]::None,
      [ref]$parsedDate
    )

    if ($isValidDate) {
      return $parsedDate.ToString("yyyy-MM-dd")
    }
  }

  return $null
}

if (-not (Test-Path -LiteralPath $FolderPath -PathType Container)) {
  throw "指定されたフォルダが存在しません: $FolderPath"
}

$searchParameters = @{
  LiteralPath = $FolderPath
  Filter      = "*.md"
  File        = $true
}

if ($Recurse) {
  $searchParameters.Recurse = $true
}

$markdownFiles = Get-ChildItem @searchParameters

$renamedCount = 0
$unchangedCount = 0
$skippedCount = 0
$errorCount = 0

foreach ($file in $markdownFiles) {
  $headingDate = Get-DateFromFirstLine -FilePath $file.FullName

  if ($null -eq $headingDate) {
    Write-Warning "1行目から有効な日付を取得できませんでした: $($file.FullName)"
    $skippedCount++
    continue
  }

  $expectedFileName = "$headingDate.md"

  if ($file.Name -ceq $expectedFileName) {
    $unchangedCount++
    continue
  }

  $destinationPath = Join-Path `
    -Path $file.DirectoryName `
    -ChildPath $expectedFileName

  Write-Host ""
  Write-Host "ファイル名と1行目の日付が一致していません。" `
    -ForegroundColor Yellow
  Write-Host "現在のファイル名 : $($file.Name)"
  Write-Host "1行目の日付     : $headingDate"
  Write-Host "変更後           : $expectedFileName"

  if (Test-Path -LiteralPath $destinationPath) {
    Write-Warning "リネーム先のファイルがすでに存在します。"
    Write-Host "既存ファイル: $destinationPath"

    while ($true) {
      $answer = Read-Host "このファイルをスキップして続行しますか？ [Y] 続行 / [N] 全処理を中断"

      switch ($answer.Trim().ToUpperInvariant()) {
        "Y" {
          Write-Host "リネームを行わず、次のファイルへ進みます。"
          $skippedCount++
          break
        }

        "N" {
          Write-Host "処理を中断しました。"
          Write-Host ""
          Write-Host "リネーム済み : $renamedCount 件"
          Write-Host "変更不要     : $unchangedCount 件"
          Write-Host "スキップ     : $skippedCount 件"
          Write-Host "エラー       : $errorCount 件"
          return
        }

        default {
          Write-Host "YまたはNを入力してください。" `
            -ForegroundColor Yellow
          continue
        }
      }

      break
    }

    continue
  }

  try {
    Rename-Item `
      -LiteralPath $file.FullName `
      -NewName $expectedFileName

    Write-Host "リネームしました。" -ForegroundColor Green
    $renamedCount++
  }
  catch {
    Write-Warning "リネームに失敗しました: $($file.FullName)"
    Write-Warning $_.Exception.Message
    $errorCount++
  }
}

Write-Host ""
Write-Host "処理が完了しました。" -ForegroundColor Cyan
Write-Host "リネーム済み : $renamedCount 件"
Write-Host "変更不要     : $unchangedCount 件"
Write-Host "スキップ     : $skippedCount 件"