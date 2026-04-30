$ErrorActionPreference = 'Stop'

$root = 'C:\Users\HMATC\Desktop\Dev FANG\meeting-automation-mvp'
$templatePath = Join-Path $root 'assets\weekly-report-template-v3.pptx'
$outputPath = Join-Path $root 'generated\zelo-weekly-report-v13-slide1-polished.pptx'
$previewPath = Join-Path $root 'generated\zelo-weekly-report-v13-slide1-polished.png'

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
  $range.ParagraphFormat.SpaceWithin = 0.84
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

function Set-LabelCell {
  param(
    $cell,
    [string]$text
  )

  $range = Set-CellText -cell $cell -text $text -size 10 -bold $true
  $cell.Shape.TextFrame.MarginLeft = 5
  $cell.Shape.TextFrame.MarginRight = 4
  $cell.Shape.TextFrame.MarginTop = 3
  return $range
}

$accentBlue = 12808772

$powerPoint = New-Object -ComObject PowerPoint.Application
$powerPoint.Visible = 1

try {
  $presentation = $powerPoint.Presentations.Open($templatePath, $false, $false, $false)
  $slide = $presentation.Slides.Item(1)

  try {
    $slide.Shapes.Item('Picture 4').Delete()
  } catch {}

  $slide.Shapes.Item('Text Placeholder 7').TextFrame.TextRange.Text = '[Startup] ZELO'
  $slide.Shapes.Item('Text Placeholder 7').TextFrame.TextRange.Font.Size = 28

  $summaryShape = Find-ShapeByNameOrHeuristic -slide $slide -preferredName 'Table 3' -predicate {
    param($shape)
    $shape.Type -eq 19 -and $shape.Width -gt 900 -and $shape.Height -gt 450
  }
  $coreShape = Find-ShapeByNameOrHeuristic -slide $slide -preferredName '表格 11' -predicate {
    param($shape)
    $shape.Type -eq 19 -and $shape.Top -gt 180 -and $shape.Height -gt 250 -and $shape.Height -lt 320
  }

  $summaryTable = $summaryShape.Table
  $coreTable = $coreShape.Table

  $meetingInfo = "- date : '26.4.03`r- participants : Cradle internal review; FANG Wei; Flora Liu; ZELO Pre-IPO and collaboration desk review"
  $companyDescription = "L4 robovan company focused on scaled urban B2B logistics and autonomous delivery.`rFounded in 2021; HQ in Suzhou"
  $category = 'Category : Autonomous Logistics / Physical AI'
  $source = 'Sourced by / CRADLE contact : FANG Wei'
  $impact = "Opinion for HMG Impact`rHigh-priority opportunity to track for both growth exposure and cross-border collaboration optionality."
  $nextStep = "Next Step`rPressure-test IPO timing, neutral-platform strategy, and collaboration fit across Singapore, China, and overseas logistics partners."

  Set-CellText -cell $summaryTable.Cell(1, 1) -text "Meeting Information`r$meetingInfo" -size 10 -bold $false | Out-Null
  Set-CellText -cell $summaryTable.Cell(1, 2) -text "Meeting Information`r$meetingInfo" -size 10 -bold $false | Out-Null
  Set-CellText -cell $summaryTable.Cell(2, 1) -text "Company Description`r$companyDescription" -size 10 -bold $false | Out-Null
  Set-CellText -cell $summaryTable.Cell(2, 2) -text "Company Description`r$companyDescription" -size 10 -bold $false | Out-Null
  Set-CellText -cell $summaryTable.Cell(3, 1) -text $category -size 10 -bold $true | Out-Null
  Set-CellText -cell $summaryTable.Cell(3, 2) -text $source -size 10 -bold $true | Out-Null
  Set-CellText -cell $summaryTable.Cell(5, 1) -text $impact -size 10 -bold $false | Out-Null
  Set-CellText -cell $summaryTable.Cell(5, 2) -text $nextStep -size 10 -bold $false | Out-Null

  Apply-SubstringStyle -range $summaryTable.Cell(3, 1).Shape.TextFrame.TextRange -needle 'Category :' -bold $true
  Apply-SubstringStyle -range $summaryTable.Cell(3, 2).Shape.TextFrame.TextRange -needle 'Sourced by / CRADLE contact :' -bold $true
  Apply-SubstringStyle -range $summaryTable.Cell(5, 1).Shape.TextFrame.TextRange -needle 'Opinion for HMG Impact' -bold $true
  Apply-SubstringStyle -range $summaryTable.Cell(5, 2).Shape.TextFrame.TextRange -needle 'Next Step' -bold $true
  Apply-SubstringStyle -range $summaryTable.Cell(2, 1).Shape.TextFrame.TextRange -needle 'Company Description' -bold $true
  Apply-SubstringStyle -range $summaryTable.Cell(1, 1).Shape.TextFrame.TextRange -needle 'Meeting Information' -bold $true

  $coreTable.Rows.Item(1).Height = 95
  $coreTable.Rows.Item(2).Height = 109
  $coreTable.Rows.Item(3).Height = 83

  Set-LabelCell -cell $coreTable.Cell(1, 1) -text 'Product' | Out-Null
  Set-LabelCell -cell $coreTable.Cell(2, 1) -text 'GTM / Econ' | Out-Null
  Set-LabelCell -cell $coreTable.Cell(3, 1) -text 'Pre-IPO' | Out-Null

  $productText = @(
    '. Platform: full-stack robovan family already spans courier, cold-chain, chassis, security, and multi-compartment formats.'
    '. Validation: company materials cite 200+ patents and 40+ awards.'
    '. Rollout: Z-series in 2024.06, E-series in 2025.05, and L-series in 2025.08.'
    '. Anchor SKU: Z5 at 1,000kg payload, 5.5m3 cargo, 230/330km range, and 40km/h.'
    '. Product ladder: E6 / L5 / Z8 plus refrigerated variants across 500kg, 1,000kg, 1,500kg, and 1,800kg payload bands.'
    '. Operating spec: -30C to 60C, IP55, and 4.5m turning radius on key medium-format platforms.'
    '. Scenario fit: parcel, cold-chain, security patrol, park logistics, and open-road city distribution are already covered in the handbook.'
    '. Cost-down: targeting larger 15-20m3 trucks, 2 Orin, lidar cut from 4 to 2 by FY26, and BOM down from RMB60k toward RMB45k.'
  ) -join "`r"

  $gtmText = @(
    '. Market thesis: urban B2B distribution is the real target; express is only ~15% of the broader city-distribution opportunity.'
    '. Demand signal: >15k express vehicles by end-2025 and >20k more planned in the following year.'
    '. Flagged proof point: 7,000-unit China Post tender.'
    '. Unit economics: RMB2k-3k monthly vehicle cost and 18-24 month payback above 500 parcels/day.'
    '. Efficiency upside: savings improve materially above 2,000 parcels/day.'
    '. Sweet spot: 5-10m3 city-distribution vehicles between gray-zone three-wheelers and regulated light trucks.'
    '. Regulatory tailwind: secondary arterial roads opening at 40-50km/h should improve turnover.'
    '. Installed base / model: >10k cumulative deliveries, ~20k installed-base discussion, and retained fleet still only in the low thousands.'
    '. Commercial structure: postal may prefer rental / operating service, while most other customers still appear to buy vehicles outright.'
  ) -join "`r"

  $preIpoText = @(
    '. IPO path: red-chip structure, Hong Kong-only route, and potential 2H filing window.'
    '. Timing implication: this is why the name should be flagged now rather than later.'
    '. Financing ask: US$2.5bn valuation / US$2.0bn pre-money.'
    '. FY25 frame: ~RMB400m revenue, ~RMB100m loss, and 12k-13k unit scale.'
    '. Scale case: 40k-50k units and RMB2.5-3.0bn revenue.'
    '. Margin bridge: profitability still relies on BOM reduction, teleoperation leverage, and better mix beyond parcel delivery.'
    '. Strategic question: neutrality after Cainiao alignment with postal, SF, DHL, and FedEx-type relationships.'
    '. Diligence focus: backlog quality, real teleoperation ratio, and route-right durability should determine whether the IPO case is financeable.'
    '. Collaboration fit: Singapore for capital-market / regional structuring, China for commercialization and policy access, and overseas for neutral-channel expansion.'
  ) -join "`r"

  $productRange = Set-CellText -cell $coreTable.Cell(1, 2) -text $productText -size 10 -bold $false
  $gtmRange = Set-CellText -cell $coreTable.Cell(2, 2) -text $gtmText -size 10 -bold $false
  $preIpoRange = Set-CellText -cell $coreTable.Cell(3, 2) -text $preIpoText -size 10 -bold $false

  foreach ($needle in @('Platform:', 'Validation:', 'Rollout:', 'Anchor SKU:', 'Product ladder:', 'Operating spec:', 'Scenario fit:', 'Cost-down:', '200+ patents', '40+ awards', '2024.06', '2025.05', '2025.08', '1,000kg payload', '5.5m3 cargo', '230/330km range', '40km/h', 'E6 / L5 / Z8', '500kg', '1,000kg', '1,500kg', '1,800kg', '-30C to 60C', 'IP55', '4.5m turning radius', '15-20m3', '2 Orin', '4 to 2 by FY26', 'RMB60k', 'RMB45k')) {
    Apply-SubstringStyle -range $productRange -needle $needle -bold $true
  }
  foreach ($needle in @('Platform:', 'Validation:', 'Rollout:', 'Anchor SKU:', 'Product ladder:', 'Operating spec:', 'Scenario fit:', 'Cost-down:')) {
    Apply-SubstringStyle -range $productRange -needle $needle -rgb $accentBlue
  }

  foreach ($needle in @('Market thesis:', '~15%', 'Demand signal:', '>15k express vehicles by end-2025', '>20k', 'Flagged proof point:', '7,000-unit China Post tender', 'Unit economics:', 'RMB2k-3k', '18-24 month payback', '500 parcels/day', 'Efficiency upside:', '2,000 parcels/day', 'Sweet spot:', '5-10m3', 'Regulatory tailwind:', '40-50km/h', 'Installed base / model:', '>10k cumulative deliveries', '~20k installed-base discussion', 'Commercial structure:', 'rental / operating service')) {
    Apply-SubstringStyle -range $gtmRange -needle $needle -bold $true
  }
  foreach ($needle in @('Market thesis:', 'Demand signal:', 'Flagged proof point:', 'Unit economics:', 'Efficiency upside:', 'Sweet spot:', 'Regulatory tailwind:', 'Installed base / model:', 'Commercial structure:')) {
    Apply-SubstringStyle -range $gtmRange -needle $needle -rgb $accentBlue
  }

  foreach ($needle in @('IPO path:', 'Hong Kong-only', '2H filing window', 'Timing implication:', 'Financing ask:', 'US$2.5bn valuation / US$2.0bn pre-money', 'FY25 frame:', '~RMB400m revenue', '~RMB100m loss', '12k-13k unit scale', 'Scale case:', '40k-50k units', 'RMB2.5-3.0bn revenue', 'Margin bridge:', 'teleoperation leverage', 'Strategic question:', 'Cainiao alignment', 'postal, SF, DHL, and FedEx-type relationships', 'Diligence focus:', 'backlog quality', 'real teleoperation ratio', 'route-right durability', 'Collaboration fit:', 'Singapore for capital-market / regional structuring', 'China for commercialization and policy access', 'overseas for neutral-channel expansion')) {
    Apply-SubstringStyle -range $preIpoRange -needle $needle -bold $true
  }
  foreach ($needle in @('IPO path:', 'Timing implication:', 'Financing ask:', 'FY25 frame:', 'Scale case:', 'Margin bridge:', 'Strategic question:', 'Diligence focus:', 'Collaboration fit:')) {
    Apply-SubstringStyle -range $preIpoRange -needle $needle -rgb $accentBlue
  }

  foreach ($needle in @('200+ patents', '40+ awards', '7,000-unit China Post tender')) {
    Apply-SubstringStyle -range $gtmRange -needle $needle -rgb $accentBlue
    Apply-SubstringStyle -range $gtmRange -needle $needle -bold $true
  }

  foreach ($needle in @('200+ patents', '40+ awards', 'Hong Kong-only', '2H filing window', 'US$2.5bn valuation / US$2.0bn pre-money', 'Singapore for capital-market / regional structuring', 'China for commercialization and policy access', 'overseas for neutral-channel expansion')) {
    Apply-SubstringStyle -range $productRange -needle $needle -rgb $accentBlue
  }

  foreach ($needle in @('7,000-unit China Post tender', 'Hong Kong-only', '2H filing window', 'US$2.5bn valuation / US$2.0bn pre-money', 'Singapore for capital-market / regional structuring', 'China for commercialization and policy access', 'overseas for neutral-channel expansion')) {
    Apply-SubstringStyle -range $preIpoRange -needle $needle -rgb $accentBlue
    Apply-SubstringStyle -range $preIpoRange -needle $needle -bold $true
  }

  $presentation.SaveAs($outputPath)
  $slide.Export($previewPath, 'PNG', 1366, 768)
  $presentation.Close()
}
finally {
  if ($null -ne $powerPoint) {
    $powerPoint.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($powerPoint) | Out-Null
  }
}
