param(
  [Parameter(Mandatory = $true)][string]$DraftPath,
  [Parameter(Mandatory = $true)][string]$OutputPath
)

$ErrorActionPreference = 'Stop'

$utf8 = New-Object System.Text.UTF8Encoding($false)
$draftJson = [System.IO.File]::ReadAllText($DraftPath, $utf8)
$draft = $draftJson | ConvertFrom-Json

$ppLayoutBlank = 12
$msoTextOrientationHorizontal = 1
$msoShapeRectangle = 1
$msoFalse = 0
$msoTrue = -1

$ink = 0
$muted = 10066329
$line = 14079702
$accentSoft = 16382457

function Add-TextBox {
  param(
    $slide,
    [double]$left,
    [double]$top,
    [double]$width,
    [double]$height,
    [string]$text,
    [double]$size = 10.5,
    [bool]$bold = $false,
    [int]$color = 0,
    [bool]$italic = $false
  )

  $shape = $slide.Shapes.AddTextbox($msoTextOrientationHorizontal, $left, $top, $width, $height)
  $shape.TextFrame.TextRange.Text = $text
  $shape.TextFrame.MarginLeft = 0
  $shape.TextFrame.MarginRight = 0
  $shape.TextFrame.MarginTop = 0
  $shape.TextFrame.MarginBottom = 0
  $shape.TextFrame.TextRange.Font.Name = 'Arial Narrow'
  $shape.TextFrame.TextRange.Font.Size = [math]::Max(10.5, $size)
  $shape.TextFrame.TextRange.Font.Bold = $(if ($bold) { $msoTrue } else { $msoFalse })
  $shape.TextFrame.TextRange.Font.Italic = $(if ($italic) { $msoTrue } else { $msoFalse })
  $shape.TextFrame.TextRange.Font.Color.RGB = $color
  $shape.TextFrame.TextRange.ParagraphFormat.SpaceBefore = 0
  $shape.TextFrame.TextRange.ParagraphFormat.SpaceAfter = 0
  $shape.TextFrame.TextRange.ParagraphFormat.SpaceWithin = 0.92
  $shape.Line.Visible = $msoFalse
  return $shape
}

function Add-SectionBlock {
  param(
    $slide,
    [double]$left,
    [double]$top,
    [double]$width,
    [string]$label,
    [object[]]$lines
  )

  Add-TextBox -slide $slide -left $left -top $top -width 120 -height 14 -text $label -size 10.5 -bold $true -color $ink | Out-Null
  $accent = $slide.Shapes.AddShape($msoShapeRectangle, $left + $width - 18, $top + 3, 10, 4)
  $accent.Fill.ForeColor.RGB = $accentSoft
  $accent.Line.Visible = $msoFalse

  $body = ((@($lines) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }) | ForEach-Object { ". $_" }) -join "`r"
  Add-TextBox -slide $slide -left $left -top ($top + 20) -width $width -height 88 -text $body -size 10.5 -bold $false -color $ink | Out-Null
}

$powerPoint = New-Object -ComObject PowerPoint.Application
$powerPoint.Visible = 1
$presentation = $null

try {
  $presentation = $powerPoint.Presentations.Add($msoTrue)
  $presentation.PageSetup.SlideWidth = 540
  $presentation.PageSetup.SlideHeight = 780

  $slide = $presentation.Slides.Add(1, $ppLayoutBlank)
  $slide.FollowMasterBackground = $msoTrue

  Add-TextBox -slide $slide -left 34 -top 28 -width 440 -height 24 -text $draft.title -size 18 -bold $true -color $ink | Out-Null

  $panel = $slide.Shapes.AddShape($msoShapeRectangle, 34, 55, 480, 357)
  $panel.Fill.Visible = $msoFalse
  $panel.Line.ForeColor.RGB = $line
  $panel.Line.Weight = 1

  $sections = @($draft.sections)
  $sectionTop = 74
  $sectionGap = 126
  foreach ($section in $sections) {
    Add-SectionBlock -slide $slide -left 52 -top $sectionTop -width 430 -label $section.label -lines $section.lines
    $sectionTop += $sectionGap
  }

  Add-TextBox -slide $slide -left 34 -top 610 -width 460 -height 30 -text 'Prepared in OI News Report portrait format. Content intentionally occupies roughly the upper half to two-thirds of the page.' -size 10.5 -bold $false -color $muted -italic $true | Out-Null

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
