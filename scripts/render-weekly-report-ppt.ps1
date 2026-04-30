param(
  [Parameter(Mandatory = $true)][string]$DraftPath,
  [Parameter(Mandatory = $true)][string]$OutputPath
)

$ErrorActionPreference = 'Stop'

$root = Split-Path -Parent $PSScriptRoot
$templatePath = Join-Path $root 'assets\weekly-report-template-v3.pptx'
$utf8 = New-Object System.Text.UTF8Encoding($false)
$draftJson = [System.IO.File]::ReadAllText($DraftPath, $utf8)
$draft = $draftJson | ConvertFrom-Json

function Find-ShapeByNameOrHeuristic {
  param(
    $slide,
    [string]$preferredName,
    [scriptblock]$predicate
  )

  try {
    return $slide.Shapes.Item($preferredName)
  } catch {}

  foreach ($shape in @($slide.Shapes)) {
    if (& $predicate $shape) {
      return $shape
    }
  }

  throw "Unable to find shape: $preferredName"
}

function Set-TextStyle {
  param(
    $range,
    [double]$size = 10,
    [bool]$bold = $false
  )

  $range.Font.Size = [math]::Max(10, $size)
  $range.Font.Bold = $(if ($bold) { -1 } else { 0 })
  $range.ParagraphFormat.SpaceBefore = 0
  $range.ParagraphFormat.SpaceAfter = 0
  $range.ParagraphFormat.SpaceWithin = 0.9
}

function Set-CellText {
  param(
    $cell,
    [string]$text,
    [double]$size = 10,
    [bool]$bold = $false
  )

  $shape = $cell.Shape
  $shape.TextFrame.MarginLeft = 6
  $shape.TextFrame.MarginRight = 6
  $shape.TextFrame.MarginTop = 3
  $shape.TextFrame.MarginBottom = 2
  $shape.TextFrame.TextRange.Text = $text
  Set-TextStyle -range $shape.TextFrame.TextRange -size $size -bold $bold
  return $shape.TextFrame.TextRange
}

function Clear-TableText {
  param($table)

  for ($row = 1; $row -le $table.Rows.Count; $row++) {
    for ($col = 1; $col -le $table.Columns.Count; $col++) {
      try {
        $cellShape = $table.Cell($row, $col).Shape
        $cellShape.TextFrame.TextRange.Text = ''
      } catch {}
    }
  }
}

function Add-TextBox {
  param(
    $slide,
    [double]$left,
    [double]$top,
    [double]$width,
    [double]$height,
    [string]$text,
    [double]$size = 10,
    [bool]$bold = $false,
    [int]$color = 0
  )

  $shape = $slide.Shapes.AddTextbox(1, $left, $top, $width, $height)
  $shape.TextFrame.TextRange.Text = $text
  $shape.TextFrame.MarginLeft = 4
  $shape.TextFrame.MarginRight = 4
  $shape.TextFrame.MarginTop = 2
  $shape.TextFrame.MarginBottom = 2
  Set-TextStyle -range $shape.TextFrame.TextRange -size $size -bold $bold
  $shape.TextFrame.TextRange.Font.Color.RGB = $color
  $shape.Line.Visible = 0
  return $shape
}

function Add-RoundedPanel {
  param(
    $slide,
    [double]$left,
    [double]$top,
    [double]$width,
    [double]$height,
    [int]$fillColor,
    [int]$lineColor
  )

  $shape = $slide.Shapes.AddShape(5, $left, $top, $width, $height)
  $shape.Fill.ForeColor.RGB = $fillColor
  $shape.Line.ForeColor.RGB = $lineColor
  $shape.Line.Weight = 1
  return $shape
}

function Add-BulletTextBlock {
  param(
    $slide,
    [double]$left,
    [double]$top,
    [double]$width,
    [double]$height,
    [object[]]$lines,
    [double]$size = 10
  )

  $content = ((@($lines) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }) | ForEach-Object { ". $_" }) -join "`r"
  return Add-TextBox -slide $slide -left $left -top $top -width $width -height $height -text $content -size $size -bold $false -color 0
}

function Apply-SubstringStyle {
  param(
    $range,
    [string]$needle,
    [Nullable[bool]]$bold = $null,
    [Nullable[int]]$rgb = $null
  )

  if ([string]::IsNullOrWhiteSpace($needle)) {
    return
  }

  $full = $range.Text
  $startIndex = 0

  while ($true) {
    $foundAt = $full.IndexOf($needle, $startIndex, [System.StringComparison]::Ordinal)
    if ($foundAt -lt 0) {
      break
    }

    $fragment = $range.Characters($foundAt + 1, $needle.Length)
    if ($null -ne $bold) {
      $fragment.Font.Bold = $(if ($bold.Value) { -1 } else { 0 })
    }
    if ($null -ne $rgb) {
      $fragment.Font.Color.RGB = $rgb.Value
    }

    $startIndex = $foundAt + $needle.Length
  }
}

function Format-WeeklyLineText {
  param([string]$line)

  $text = [string]$line
  if ([string]::IsNullOrWhiteSpace($text)) {
    return ''
  }

  $trimmed = $text.Trim()
  if ($trimmed.StartsWith('# ')) {
    return $trimmed.Substring(2).Trim()
  }

  $firstCode = 0
  if ($trimmed.Length -gt 0) {
    $firstCode = [int][char]$trimmed[0]
  }

  if ($firstCode -eq 0x2022) {
    return $trimmed
  }

  if ($firstCode -ge 0x2460 -and $firstCode -le 0x2468) {
    return $trimmed
  }

  if ($text -match '^\s+-\s+') {
    return "   - $($trimmed -replace '^\-\s+', '')"
  }

  if ($trimmed.StartsWith('- ')) {
    return "   - $($trimmed.Substring(2).Trim())"
  }

  return ". $trimmed"
}

function Remove-WeeklyMarkerPrefix {
  param([string]$text)

  $cleaned = ([string]$text).Trim()
  if ([string]::IsNullOrWhiteSpace($cleaned)) {
    return ''
  }

  $firstCode = [int][char]$cleaned[0]
  if ($firstCode -eq 0x2022 -or ($firstCode -ge 0x2460 -and $firstCode -le 0x2468)) {
    return $cleaned.Substring(1).Trim()
  }

  return ($cleaned -replace '^[\.\-\s]+', '').Trim()
}

function Apply-WeeklyHierarchyStyle {
  param($range)

  $range.ParagraphFormat.SpaceBefore = 0
  $range.ParagraphFormat.SpaceAfter = 0
  $range.ParagraphFormat.SpaceWithin = 0.82

  foreach ($line in @($range.Text -split "`r")) {
    $text = ([string]$line).Trim()
    if ([string]::IsNullOrWhiteSpace($text)) {
      continue
    }

    if ($text -match '^(Product|Technology|Commercialization|Funding|Company / Team)$') {
      Apply-SubstringStyle -range $range -needle $text -bold $true -rgb $accentBlue
      continue
    }

    $firstCode = [int][char]$text[0]
    if ($firstCode -ge 0x2460 -and $firstCode -le 0x2468) {
      $marker = $text.Substring(0, 1)
      Apply-SubstringStyle -range $range -needle $marker -bold $true -rgb $accentBlue
      continue
    }

    if ($firstCode -eq 0x2022) {
      $marker = $text.Substring(0, 1)
      Apply-SubstringStyle -range $range -needle $marker -bold $true
      continue
    }
  }
}

$accentBlue = 12808772
$powerPoint = New-Object -ComObject PowerPoint.Application
$powerPoint.Visible = 1
$presentation = $null

try {
  $presentation = $powerPoint.Presentations.Open($templatePath, $false, $false, $false)
  $slide = $presentation.Slides.Item(1)

  try {
    $slide.Shapes.Item('Picture 4').Delete()
  } catch {}

  $slide.Shapes.Item('Text Placeholder 7').TextFrame.TextRange.Text = $draft.title
  $slide.Shapes.Item('Text Placeholder 7').TextFrame.TextRange.Font.Size = 28

  $summaryShape = Find-ShapeByNameOrHeuristic -slide $slide -preferredName 'Table 3' -predicate {
    param($shape)
    $shape.Type -eq 19 -and $shape.Width -gt 900 -and $shape.Height -gt 450
  }
  $coreShape = Find-ShapeByNameOrHeuristic -slide $slide -preferredName '表格 11' -predicate {
    param($shape)
    $shape.Type -eq 19 -and $shape.Top -gt 140 -and $shape.Height -gt 250
  }

  $summaryTable = $summaryShape.Table
  $coreTable = $coreShape.Table

  Clear-TableText -table $summaryTable
  Clear-TableText -table $coreTable

  $meetingLine = "- date : $($draft.meetingInformation.date)`r- participants : $($draft.meetingInformation.participants)"
  $companyDescription = "$($draft.companyDescription.line1)`rFounded in $($draft.companyDescription.foundedYear); HQ in $($draft.companyDescription.hqCity)"
  $opinion = "Opinion for HMG Impact`r$($draft.opinion)"
  $nextStep = "Next Step`r$($draft.nextStep)"

  Set-CellText -cell $summaryTable.Cell(1, 1) -text "Meeting Information`r$meetingLine" | Out-Null
  Set-CellText -cell $summaryTable.Cell(1, 2) -text "Meeting Information`r$meetingLine" | Out-Null
  Set-CellText -cell $summaryTable.Cell(2, 1) -text "Company Description`r$companyDescription" | Out-Null
  Set-CellText -cell $summaryTable.Cell(2, 2) -text "Company Description`r$companyDescription" | Out-Null
  Set-CellText -cell $summaryTable.Cell(3, 1) -text "Category : $($draft.category)" -bold $true | Out-Null
  Set-CellText -cell $summaryTable.Cell(3, 2) -text "Sourced by / CRADLE contact : $($draft.sourceContact)" -bold $true | Out-Null
  Set-CellText -cell $summaryTable.Cell(5, 1) -text $opinion | Out-Null
  Set-CellText -cell $summaryTable.Cell(5, 2) -text $nextStep | Out-Null

  foreach ($entry in @(
    @{ Cell = $summaryTable.Cell(1, 1).Shape.TextFrame.TextRange; Needle = 'Meeting Information' },
    @{ Cell = $summaryTable.Cell(2, 1).Shape.TextFrame.TextRange; Needle = 'Company Description' },
    @{ Cell = $summaryTable.Cell(3, 1).Shape.TextFrame.TextRange; Needle = 'Category :' },
    @{ Cell = $summaryTable.Cell(3, 2).Shape.TextFrame.TextRange; Needle = 'Sourced by / CRADLE contact :' },
    @{ Cell = $summaryTable.Cell(5, 1).Shape.TextFrame.TextRange; Needle = 'Opinion for HMG Impact' },
    @{ Cell = $summaryTable.Cell(5, 2).Shape.TextFrame.TextRange; Needle = 'Next Step' }
  )) {
    Apply-SubstringStyle -range $entry.Cell -needle $entry.Needle -bold $true
  }

  $sections = @($draft.sections)
  for ($i = 0; $i -lt [Math]::Min(3, $sections.Count); $i++) {
    $row = $i + 1
    $section = $sections[$i]
    Set-CellText -cell $coreTable.Cell($row, 1) -text $section.label -bold $true | Out-Null
    $bodyLines = @($section.lines) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    $body = ($bodyLines | ForEach-Object { Format-WeeklyLineText $_ }) -join "`r"
    $range = Set-CellText -cell $coreTable.Cell($row, 2) -text $body
    Apply-WeeklyHierarchyStyle -range $range
    foreach ($line in @($bodyLines)) {
      $visibleLine = Format-WeeklyLineText $line
      if ($visibleLine -match ':') {
        $prefix = ($visibleLine -split ':')[0] + ':'
        $prefix = Remove-WeeklyMarkerPrefix $prefix
        if (-not $prefix.EndsWith(':')) {
          $prefix = "${prefix}:"
        }
        Apply-SubstringStyle -range $range -needle $prefix -bold $true -rgb $accentBlue
      }
    }
  }

  if ($draft.appendix) {
    $slide2 = $presentation.Slides.Add(2, 12)
    $slide2.FollowMasterBackground = -1

    Add-TextBox -slide $slide2 -left 18 -top 10 -width 420 -height 32 -text $draft.appendix.title -size 28 -bold $true -color 0 | Out-Null
    Add-TextBox -slide $slide2 -left 20 -top 43 -width 420 -height 16 -text $draft.appendix.subtitle -size 11 -bold $false -color 8421504 | Out-Null
    $line = $slide2.Shapes.AddLine(18, 62, 942, 62)
    $line.Line.ForeColor.RGB = 0
    $line.Line.Weight = 1.2

    $roadmapPanel = Add-RoundedPanel -slide $slide2 -left 24 -top 78 -width 920 -height 205 -fillColor 16777215 -lineColor 12632256
    Add-TextBox -slide $slide2 -left 38 -top 86 -width 260 -height 18 -text $draft.appendix.roadmap.title -size 12 -bold $true -color 0 | Out-Null
    Add-TextBox -slide $slide2 -left 38 -top 106 -width 860 -height 18 -text $draft.appendix.roadmap.subtitle -size 10 -bold $false -color $accentBlue | Out-Null
    $roadmapRange = Add-BulletTextBlock -slide $slide2 -left 38 -top 132 -width 880 -height 132 -lines @($draft.appendix.roadmap.lines) -size 10
    foreach ($lineText in @($draft.appendix.roadmap.lines)) {
      if ($lineText -match '(20\d{2}|pilot|launch|mass production|SOP|roadmap|timeline|milestone)') {
        $matches = [regex]::Matches($lineText, '(20\d{2}|pilot|launch|mass production|SOP|roadmap|timeline|milestone)', 'IgnoreCase')
        foreach ($m in $matches) {
          Apply-SubstringStyle -range $roadmapRange.TextFrame.TextRange -needle $m.Value -bold $true -rgb $accentBlue
        }
      }
    }

    $econPanel = Add-RoundedPanel -slide $slide2 -left 24 -top 300 -width 452 -height 220 -fillColor 16777215 -lineColor 12632256
    Add-TextBox -slide $slide2 -left 38 -top 308 -width 300 -height 18 -text $draft.appendix.economics.title -size 12 -bold $true -color 0 | Out-Null
    Add-TextBox -slide $slide2 -left 38 -top 328 -width 400 -height 24 -text $draft.appendix.economics.subtitle -size 10 -bold $false -color $accentBlue | Out-Null
    $econRange = Add-BulletTextBlock -slide $slide2 -left 38 -top 360 -width 408 -height 144 -lines @($draft.appendix.economics.lines) -size 10
    foreach ($lineText in @($draft.appendix.economics.lines)) {
      if ($lineText -match '(\$[\d\.]+[mb]?|RMB[\d,\.]+|\d+%|payback|margin|cost|revenue|valuation)') {
        $matches = [regex]::Matches($lineText, '(\$[\d\.]+[mb]?|RMB[\d,\.]+|\d+%|payback|margin|cost|revenue|valuation)', 'IgnoreCase')
        foreach ($m in $matches) {
          Apply-SubstringStyle -range $econRange.TextFrame.TextRange -needle $m.Value -bold $true -rgb $accentBlue
        }
      }
    }

    $fundingPanel = Add-RoundedPanel -slide $slide2 -left 492 -top 300 -width 452 -height 220 -fillColor 16777215 -lineColor 12632256
    Add-TextBox -slide $slide2 -left 506 -top 308 -width 220 -height 18 -text $draft.appendix.funding.title -size 12 -bold $true -color 0 | Out-Null
    Add-TextBox -slide $slide2 -left 506 -top 328 -width 410 -height 24 -text $draft.appendix.funding.subtitle -size 10 -bold $false -color $accentBlue | Out-Null

    $fundingTableShape = $slide2.Shapes.AddTable([Math]::Max(2, @($draft.appendix.funding.rows).Count + 1), 4, 506, 362, 410, 136)
    $fundingTable = $fundingTableShape.Table
    $headers = @('Round', 'Raised', 'Valuation', 'Key Shareholders')
    for ($col = 1; $col -le 4; $col++) {
      $headerRange = Set-CellText -cell $fundingTable.Cell(1, $col) -text $headers[$col - 1] -size 10 -bold $true
      $fundingTable.Cell(1, $col).Shape.Fill.ForeColor.RGB = 15925247
      $headerRange.Font.Color.RGB = 0
    }

    $rows = @($draft.appendix.funding.rows)
    for ($row = 0; $row -lt $rows.Count; $row++) {
      $targetRow = $row + 2
      $values = @($rows[$row])
      for ($col = 1; $col -le 4; $col++) {
        $value = if ($values.Count -ge $col) { [string]$values[$col - 1] } else { 'N/A' }
        $cellRange = Set-CellText -cell $fundingTable.Cell($targetRow, $col) -text $value -size 10 -bold $false
        if ($value -match '(\$[\d\.]+[mb]?|RMB[\d,\.]+|pre-ipo|ipo|series)') {
          Apply-SubstringStyle -range $cellRange -needle $matches[0] -bold $true -rgb $accentBlue
        }
      }
    }
  }

  $presentation.SaveAs($OutputPath)
  $presentation.Close()
}
finally {
  if ($null -ne $presentation) {
    try {
      [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($presentation)
    } catch {}
  }
  if ($null -ne $powerPoint) {
    try {
      $powerPoint.Quit()
    } catch {}
    try {
      [System.Runtime.Interopservices.Marshal]::ReleaseComObject($powerPoint) | Out-Null
    } catch {}
  }
}
