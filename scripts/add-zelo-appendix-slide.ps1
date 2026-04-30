$ErrorActionPreference = 'Stop'

$root = 'C:\Users\HMATC\Desktop\Dev FANG\meeting-automation-mvp'
$inputPath = Join-Path $root 'generated\zelo-weekly-report-v13-slide1-polished.pptx'
$outputPath = Join-Path $root 'generated\zelo-weekly-report-v13-two-slide.pptx'
$slide1PreviewPath = Join-Path $root 'generated\zelo-weekly-report-v13-slide1.png'
$slide2PreviewPath = Join-Path $root 'generated\zelo-weekly-report-v13-slide2.png'

$imgZ5 = Join-Path $root 'generated\zelo_product_manual_unzip\word\media\image24.png'
$imgL5 = Join-Path $root 'generated\zelo_product_manual_unzip\word\media\image31.png'
$imgZ8 = Join-Path $root 'generated\zelo_product_manual_unzip\word\media\image53.png'

$msoFalse = 0
$msoTrue = -1
$msoTextOrientationHorizontal = 1
$ppLayoutBlank = 12
$msoShapeRoundedRectangle = 5
$msoShapeRectangle = 1
$msoShapeChevron = 52
$msoShapeOval = 9
$msoShapeIsoscelesTriangle = 7

$accentBlue = 12808772
$lightBlueFill = 15925247
$lighterBlueFill = 16250620
$darkText = 0
$midGray = 8421504
$white = 16777215

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
    [int]$color = 0,
    [string]$fontName = '현대하모니 L'
  )

  $shape = $slide.Shapes.AddTextbox($msoTextOrientationHorizontal, $left, $top, $width, $height)
  $shape.TextFrame.TextRange.Text = $text
  $shape.TextFrame.MarginLeft = 2
  $shape.TextFrame.MarginRight = 2
  $shape.TextFrame.MarginTop = 1
  $shape.TextFrame.MarginBottom = 1
  $shape.TextFrame.TextRange.Font.Name = $fontName
  $shape.TextFrame.TextRange.Font.Size = $size
  $shape.TextFrame.TextRange.Font.Bold = $(if ($bold) { $msoTrue } else { $msoFalse })
  $shape.TextFrame.TextRange.Font.Color.RGB = $color
  $shape.TextFrame.TextRange.ParagraphFormat.SpaceBefore = 0
  $shape.TextFrame.TextRange.ParagraphFormat.SpaceAfter = 0
  $shape.TextFrame.TextRange.ParagraphFormat.SpaceWithin = 0.9
  $shape.Line.Visible = $msoFalse
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

  $shape = $slide.Shapes.AddShape($msoShapeRoundedRectangle, $left, $top, $width, $height)
  $shape.Fill.ForeColor.RGB = $fillColor
  $shape.Line.ForeColor.RGB = $lineColor
  $shape.Line.Weight = 1.2
  return $shape
}

function Add-RoadmapCard {
  param(
    $slide,
    [double]$left,
    [double]$top,
    [double]$width,
    [double]$height,
    [string]$yearTitle,
    [string]$bodyText,
    [string]$imagePath
  )

  $card = Add-RoundedPanel -slide $slide -left $left -top $top -width $width -height $height -fillColor $lightBlueFill -lineColor $accentBlue
  $header = $slide.Shapes.AddShape($msoShapeRoundedRectangle, $left, $top, $width, 27)
  $header.Fill.ForeColor.RGB = $accentBlue
  $header.Line.Visible = $msoFalse
  $header.Adjustments.Item(1) = 0.08
  $header.TextFrame.TextRange.Text = $yearTitle
  $header.TextFrame.TextRange.Font.Name = '현대하모니 M'
  $header.TextFrame.TextRange.Font.Size = 12
  $header.TextFrame.TextRange.Font.Bold = $msoTrue
  $header.TextFrame.TextRange.Font.Color.RGB = $white
  $header.TextFrame.TextRange.ParagraphFormat.Alignment = 2

  $img = $slide.Shapes.AddPicture($imagePath, $msoFalse, $msoTrue, $left + 12, $top + 36, $width - 24, 78)
  $img.LockAspectRatio = $msoTrue
  if ($img.Width -gt ($width - 24)) {
    $img.Width = $width - 24
  }
  $img.Left = $left + (($width - $img.Width) / 2)
  $img.Top = $top + 40

  $text = Add-TextBox -slide $slide -left ($left + 10) -top ($top + 120) -width ($width - 20) -height 62 -text $bodyText -size 8.8 -bold $false -color $darkText
  return @($card, $header, $img, $text)
}

function Add-UnitCard {
  param(
    $slide,
    [double]$left,
    [double]$top,
    [double]$width,
    [double]$height,
    [string]$title,
    [string]$imagePath,
    [string[]]$lines,
    [bool]$isTraditional = $false
  )

  $card = Add-RoundedPanel -slide $slide -left $left -top $top -width $width -height $height -fillColor $white -lineColor $accentBlue
  $titleShape = Add-TextBox -slide $slide -left ($left + 10) -top ($top + 6) -width ($width - 20) -height 18 -text $title -size 11 -bold $true -color $accentBlue

  if ($isTraditional) {
    $body = $slide.Shapes.AddShape($msoShapeRoundedRectangle, $left + 16, $top + 28, $width - 32, 46)
    $body.Fill.ForeColor.RGB = 15790320
    $body.Line.ForeColor.RGB = 13487565
    $body.Line.Weight = 0.8
    $body.Adjustments.Item(1) = 0.12

    $cab = $slide.Shapes.AddShape($msoShapeRoundedRectangle, $left + 118, $top + 24, 78, 34)
    $cab.Fill.ForeColor.RGB = 14877353
    $cab.Line.ForeColor.RGB = 9605778
    $cab.Line.Weight = 0.8
    $cab.Adjustments.Item(1) = 0.18

    $window = $slide.Shapes.AddShape($msoShapeRoundedRectangle, $left + 132, $top + 28, 24, 16)
    $window.Fill.ForeColor.RGB = 16775142
    $window.Line.Visible = $msoFalse
    $window.Adjustments.Item(1) = 0.15

    $driverHead = $slide.Shapes.AddShape($msoShapeOval, $left + 165, $top + 31, 8, 8)
    $driverHead.Fill.ForeColor.RGB = 16764057
    $driverHead.Line.Visible = $msoFalse

    $driverBody = $slide.Shapes.AddShape($msoShapeRoundedRectangle, $left + 162, $top + 39, 13, 11)
    $driverBody.Fill.ForeColor.RGB = 10027059
    $driverBody.Line.Visible = $msoFalse
    $driverBody.Adjustments.Item(1) = 0.2

    $cargoLine = $slide.Shapes.AddLine($left + 36, $top + 51, $left + 112, $top + 51)
    $cargoLine.Line.ForeColor.RGB = 13619151
    $cargoLine.Line.Weight = 1

    $doorLine = $slide.Shapes.AddLine($left + 92, $top + 34, $left + 92, $top + 58)
    $doorLine.Line.ForeColor.RGB = 13619151
    $doorLine.Line.Weight = 1

    foreach ($wheelLeft in @(($left + 42), ($left + 134))) {
      $wheel = $slide.Shapes.AddShape($msoShapeOval, $wheelLeft, $top + 61, 16, 16)
      $wheel.Fill.ForeColor.RGB = 3355443
      $wheel.Line.Visible = $msoFalse

      $wheelInner = $slide.Shapes.AddShape($msoShapeOval, $wheelLeft + 4.5, $top + 65.5, 7, 7)
      $wheelInner.Fill.ForeColor.RGB = 12632256
      $wheelInner.Line.Visible = $msoFalse
    }

    $driverTag = Add-TextBox -slide $slide -left ($left + $width - 94) -top ($top + 12) -width 80 -height 14 -text 'driver-led' -size 8.5 -bold $true -color $midGray
  } else {
    $img = $slide.Shapes.AddPicture($imagePath, $msoFalse, $msoTrue, $left + 24, $top + 24, $width - 48, 60)
    $img.LockAspectRatio = $msoTrue
    if ($img.Width -gt ($width - 48)) {
      $img.Width = $width - 48
    }
    $img.Left = $left + (($width - $img.Width) / 2)
    $img.Top = $top + 22
  }

  $rowTop = $top + 90
  foreach ($line in $lines) {
    $textShape = Add-TextBox -slide $slide -left ($left + 12) -top $rowTop -width ($width - 24) -height 15 -text $line -size 9.2 -bold $false -color $darkText
    $rowTop += 15
  }

  return $card
}

$powerPoint = New-Object -ComObject PowerPoint.Application
$powerPoint.Visible = 1

try {
  $presentation = $powerPoint.Presentations.Open($inputPath, $msoFalse, $msoFalse, $msoFalse)
  $slide = $presentation.Slides.Add(2, $ppLayoutBlank)
  $slide.FollowMasterBackground = $msoTrue

  Add-TextBox -slide $slide -left 18 -top 10 -width 380 -height 34 -text '[Appendix] ZELO' -size 28 -bold $true -color $darkText -fontName '현대하모니 M' | Out-Null
  Add-TextBox -slide $slide -left 19 -top 42 -width 420 -height 18 -text 'Product Roadmap & Unit Economics' -size 11 -bold $false -color $midGray | Out-Null

  $line = $slide.Shapes.AddLine(18, 62, 942, 62)
  $line.Line.ForeColor.RGB = $darkText
  $line.Line.Weight = 1.2

  $topPanel = Add-RoundedPanel -slide $slide -left 26 -top 74 -width 918 -height 208 -fillColor $white -lineColor 12632256
  Add-TextBox -slide $slide -left 36 -top 80 -width 240 -height 16 -text 'Product Roadmap' -size 12 -bold $true -color $darkText | Out-Null
  Add-TextBox -slide $slide -left 36 -top 98 -width 740 -height 18 -text 'Tech roadmap shifts from lidar-led toward vision-led perception, with lidar moving into an assistive role.' -size 10 -bold $false -color $accentBlue | Out-Null

  Add-RoadmapCard -slide $slide -left 38 -top 122 -width 278 -height 148 `
    -yearTitle '2023 | First Mass-Production L4 Product' `
    -bodyText "- First mass-produced autonomous urban-delivery product at 5m3.`r- Helped reduce industry cost by more than 50%." `
    -imagePath $imgZ5 | Out-Null

  Add-RoadmapCard -slide $slide -left 344 -top 122 -width 278 -height 148 `
    -yearTitle '2024 | Full-Scenario Urban L4 Lineup' `
    -bodyText "- Product coverage expanded to 2-10m3 across broader city scenarios.`r- Dual Orin stack delivers 500+ TOPS compute." `
    -imagePath $imgL5 | Out-Null

  Add-RoadmapCard -slide $slide -left 650 -top 122 -width 278 -height 148 `
    -yearTitle '2025 | 2nd-Gen Mass-Production Series' `
    -bodyText "- Solid-state lidar + camera replaces legacy mechanical lidar-heavy schemes.`r- Stereo 3D perception targets higher-precision environmental sensing." `
    -imagePath $imgZ8 | Out-Null

  $bottomPanel = Add-RoundedPanel -slide $slide -left 26 -top 296 -width 918 -height 224 -fillColor $white -lineColor 12632256
  Add-TextBox -slide $slide -left 36 -top 302 -width 270 -height 16 -text 'Unit Economics Illustration' -size 12 -bold $true -color $darkText | Out-Null
  Add-TextBox -slide $slide -left 36 -top 320 -width 840 -height 18 -text 'Built on safety, aiming to deliver a materially better cost-performance trade-off than traditional small urban delivery fleets.' -size 10 -bold $false -color $accentBlue | Out-Null

  $badge = $slide.Shapes.AddShape($msoShapeRoundedRectangle, 40, 360, 120, 120)
  $badge.Fill.ForeColor.RGB = $lighterBlueFill
  $badge.Line.Visible = $msoFalse
  $badge.Adjustments.Item(1) = 0.2
  Add-TextBox -slide $slide -left 56 -top 390 -width 88 -height 44 -text "Economic`rPracticality" -size 18 -bold $true -color $accentBlue -fontName '현대하모니 M' | Out-Null

  Add-UnitCard -slide $slide -left 188 -top 346 -width 318 -height 148 `
    -title 'ZELO | 5m3 model' `
    -imagePath $imgZ5 `
    -lines @(
      '- Production input: vehicle',
      '- Cost structure: vehicle + subscription service fee',
      '- Operating cost: ~RMB3,000 / month',
      '- Management load: simple; no people management required',
      '- Uptime: supports 365 x 24h autonomous operation service'
    ) | Out-Null

  Add-UnitCard -slide $slide -left 540 -top 346 -width 360 -height 148 `
    -title 'Traditional | 6-8m3 van + driver' `
    -imagePath '' `
    -lines @(
      '- Production input: vehicle + driver',
      '- Cost structure: vehicle + tax + insurance + maintenance + repair + labor',
      '- Operating cost: RMB10,000+ / month',
      '- Management load: harder; training, team management, assessment, and turnover',
      '- Utilization: more than 40% of time is non-operating / rest time'
    ) `
    -isTraditional $true | Out-Null

  $presentation.SaveAs($outputPath)
  $presentation.Slides.Item(1).Export($slide1PreviewPath, 'PNG', 1366, 768)
  $slide.Export($slide2PreviewPath, 'PNG', 1366, 768)
  $presentation.Close()
}
finally {
  if ($null -ne $powerPoint) {
    $powerPoint.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($powerPoint) | Out-Null
  }
}
