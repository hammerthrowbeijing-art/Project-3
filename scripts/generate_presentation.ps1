$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

$repoRoot = Split-Path -Parent $PSScriptRoot
$outputPath = Join-Path $repoRoot "Break50_Contested_Discourse_Presentation.pptx"
$previewDir = Join-Path $repoRoot "outputs\presentation_preview"
$sb = [char]11

$figStance = Join-Path $repoRoot "outputs\figures\stance_distribution.png"
$figFrame = Join-Path $repoRoot "outputs\figures\frame_distribution.png"
$figEngagement = Join-Path $repoRoot "outputs\figures\engagement_by_stance.png"
$figTopDecile = Join-Path $repoRoot "outputs\figures\top_decile_stance_shares.png"

foreach ($path in @($figStance, $figFrame, $figEngagement, $figTopDecile)) {
    if (-not (Test-Path $path)) {
        throw "Required figure not found: $path"
    }
}

if (Test-Path $outputPath) {
    Remove-Item -LiteralPath $outputPath -Force
}

if (Test-Path $previewDir) {
    Remove-Item -LiteralPath $previewDir -Recurse -Force
}
New-Item -ItemType Directory -Path $previewDir -Force | Out-Null

function Add-TextBox {
    param(
        [Parameter(Mandatory = $true)]$Slide,
        [Parameter(Mandatory = $true)][string]$Text,
        [Parameter(Mandatory = $true)][double]$Left,
        [Parameter(Mandatory = $true)][double]$Top,
        [Parameter(Mandatory = $true)][double]$Width,
        [Parameter(Mandatory = $true)][double]$Height,
        [int]$FontSize = 20,
        [string]$FontName = "Aptos",
        [string]$ColorHex = "203040",
        [switch]$Bold,
        [switch]$Italic,
        [int]$Align = 1,
        [switch]$NoFill,
        [string]$FillHex = "FFFFFF",
        [switch]$NoLine
    )

    $shape = $Slide.Shapes.AddTextbox(1, $Left, $Top, $Width, $Height)
    if ($NoFill) {
        $shape.Fill.Visible = 0
    } else {
        $shape.Fill.Visible = -1
        $shape.Fill.ForeColor.RGB = [Convert]::ToInt32($FillHex.Substring(4, 2) + $FillHex.Substring(2, 2) + $FillHex.Substring(0, 2), 16)
    }
    if ($NoLine) {
        $shape.Line.Visible = 0
    }
    $range = $shape.TextFrame.TextRange
    $range.Text = $Text
    $range.Font.Name = $FontName
    $range.Font.Size = $FontSize
    $range.Font.Bold = [int]$Bold.IsPresent
    $range.Font.Italic = [int]$Italic.IsPresent
    $range.Font.Color.RGB = [Convert]::ToInt32($ColorHex.Substring(4, 2) + $ColorHex.Substring(2, 2) + $ColorHex.Substring(0, 2), 16)
    $range.ParagraphFormat.Alignment = $Align
    $shape.TextFrame.WordWrap = -1
    $shape.TextFrame.AutoSize = 0
    $shape.TextFrame.MarginLeft = 2
    $shape.TextFrame.MarginRight = 2
    $shape.TextFrame.MarginTop = 2
    $shape.TextFrame.MarginBottom = 2
    return $shape
}

function Add-BulletBox {
    param(
        [Parameter(Mandatory = $true)]$Slide,
        [Parameter(Mandatory = $true)][string[]]$Bullets,
        [Parameter(Mandatory = $true)][double]$Left,
        [Parameter(Mandatory = $true)][double]$Top,
        [Parameter(Mandatory = $true)][double]$Width,
        [Parameter(Mandatory = $true)][double]$Height,
        [int]$FontSize = 20,
        [string]$ColorHex = "203040"
    )

    $shape = $Slide.Shapes.AddTextbox(1, $Left, $Top, $Width, $Height)
    $shape.Fill.Visible = 0
    $shape.Line.Visible = 0
    $range = $shape.TextFrame.TextRange
    $range.Text = [string]::Join("`r", $Bullets)
    $range.Font.Name = "Aptos"
    $range.Font.Size = $FontSize
    $range.Font.Color.RGB = [Convert]::ToInt32($ColorHex.Substring(4, 2) + $ColorHex.Substring(2, 2) + $ColorHex.Substring(0, 2), 16)
    $shape.TextFrame.WordWrap = -1
    $shape.TextFrame.AutoSize = 0
    $shape.TextFrame.MarginLeft = 6
    $shape.TextFrame.MarginRight = 2
    $shape.TextFrame.MarginTop = 2
    $shape.TextFrame.MarginBottom = 2
    for ($i = 1; $i -le $Bullets.Count; $i++) {
        $p = $range.Paragraphs($i, 1)
        $p.ParagraphFormat.Bullet.Visible = -1
        $p.ParagraphFormat.Bullet.Character = 8226
        $p.ParagraphFormat.Bullet.RelativeSize = 1
        $p.ParagraphFormat.SpaceAfter = 8
        $p.ParagraphFormat.SpaceWithin = 1.05
    }
    return $shape
}

function Add-SectionHeader {
    param(
        [Parameter(Mandatory = $true)]$Slide,
        [Parameter(Mandatory = $true)][string]$Title,
        [string]$Subtitle = ""
    )

    $band = $Slide.Shapes.AddShape(1, 0, 0, 960, 64)
    $band.Fill.ForeColor.RGB = 1066645
    $band.Line.Visible = 0
    $null = Add-TextBox -Slide $Slide -Text $Title -Left 34 -Top 12 -Width 860 -Height 36 -FontSize 24 -FontName "Aptos Display" -ColorHex "F7F3ED" -Bold -NoFill -NoLine
    if (-not [string]::IsNullOrWhiteSpace($Subtitle)) {
        $null = Add-TextBox -Slide $Slide -Text $Subtitle -Left 34 -Top 540 -Width 880 -Height 18 -FontSize 10 -ColorHex "6C7A89" -NoFill -NoLine
    }
}

function Add-ImageScaled {
    param(
        [Parameter(Mandatory = $true)]$Slide,
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][double]$Left,
        [Parameter(Mandatory = $true)][double]$Top,
        [Parameter(Mandatory = $true)][double]$MaxWidth,
        [Parameter(Mandatory = $true)][double]$MaxHeight
    )

    $img = $Slide.Shapes.AddPicture($Path, 0, -1, 0, 0)
    $img.LockAspectRatio = -1
    $scale = [Math]::Min($MaxWidth / [double]$img.Width, $MaxHeight / [double]$img.Height)
    $img.Width = [double]$img.Width * $scale
    $img.Left = $Left + (($MaxWidth - [double]$img.Width) / 2)
    $img.Top = $Top + (($MaxHeight - [double]$img.Height) / 2)
    return $img
}

function Add-AccentRule {
    param(
        [Parameter(Mandatory = $true)]$Slide,
        [double]$Left = 36,
        [double]$Top = 82,
        [double]$Width = 120,
        [double]$Height = 5
    )

    $shape = $Slide.Shapes.AddShape(1, $Left, $Top, $Width, $Height)
    $shape.Fill.ForeColor.RGB = 3242405
    $shape.Line.Visible = 0
}

$ppt = $null
$presentation = $null

try {
    $ppt = New-Object -ComObject PowerPoint.Application
    $ppt.Visible = -1
    $presentation = $ppt.Presentations.Add()
    $presentation.PageSetup.SlideSize = 15

    $slide = $presentation.Slides.Add(1, 12)
    $bg = $slide.Shapes.AddShape(1, 0, 0, 960, 540)
    $bg.Fill.ForeColor.RGB = 15724527
    $bg.Line.Visible = 0
    $corner = $slide.Shapes.AddShape(1, 682, 0, 278, 540)
    $corner.Fill.ForeColor.RGB = 4078157
    $corner.Line.Visible = 0
    $null = Add-TextBox -Slide $slide -Text "Break 50" -Left 56 -Top 72 -Width 420 -Height 40 -FontSize 32 -FontName "Aptos Display" -ColorHex "8B4A2A" -Bold -NoFill -NoLine
    $null = Add-TextBox -Slide $slide -Text "Contested Discourse in the Comment Field" -Left 56 -Top 118 -Width 520 -Height 90 -FontSize 28 -FontName "Aptos Display" -ColorHex "22304A" -Bold -NoFill -NoLine
    $null = Add-TextBox -Slide $slide -Text "Stance, Framing, and Visible Reaction" -Left 56 -Top 216 -Width 500 -Height 36 -FontSize 20 -ColorHex "44576B" -Italic -NoFill -NoLine
    $null = Add-BulletBox -Slide $slide -Bullets @(
        "Break 50 X/Twitter comments (n = 1,008)",
        ("Contested discourse focus:" + $sb + "framing, stance, and visible reaction")
    ) -Left 58 -Top 326 -Width 440 -Height 84 -FontSize 13
    $null = Add-TextBox -Slide $slide -Text "Larry Nie`rProject presentation`rApril 2026" -Left 710 -Top 392 -Width 190 -Height 80 -FontSize 16 -ColorHex "F7F3ED" -Bold -NoFill -NoLine

    $slide = $presentation.Slides.Add(2, 12)
    Add-SectionHeader -Slide $slide -Title "Why This Case Matters" -Subtitle "The project now treats Break 50 as a contested sport-political discourse space."
    Add-AccentRule -Slide $slide
    $null = Add-BulletBox -Slide $slide -Bullets @(
        ("Trump's appearance turns Break 50 into a" + $sb + "politically charged crossover moment."),
        ("The comments negotiate legitimacy, support," + $sb + "condemnation, and sport/politics boundaries."),
        ("The revised project studies stance and framing first," + $sb + "then uses engagement as a secondary outcome."),
        ("The same repo infrastructure now supports" + $sb + "a genuine P3 contested-discourse study.")
    ) -Left 54 -Top 118 -Width 760 -Height 280 -FontSize 16
    $quoteBox = $slide.Shapes.AddShape(1, 50, 420, 860, 80)
    $quoteBox.Fill.ForeColor.RGB = 15131095
    $quoteBox.Line.ForeColor.RGB = 3242405
    $null = Add-TextBox -Slide $slide -Text "Core idea: these comments are not just engagement datapoints.`nThey are instances of disagreement, ambiguity, and boundary-policing." -Left 72 -Top 434 -Width 790 -Height 52 -FontSize 18 -ColorHex "22304A" -Bold -NoFill -NoLine

    $slide = $presentation.Slides.Add(3, 12)
    Add-SectionHeader -Slide $slide -Title "Research Questions" -Subtitle "Three linked questions organize the revised report."
    Add-AccentRule -Slide $slide
    $null = Add-BulletBox -Slide $slide -Bullets @(
        ("RQ1: Is Trump's appearance framed as politics," + $sb + "sport, or a blend of both?"),
        ("RQ2: How are stance positions distributed across" + $sb + "supportive, oppositional, depoliticizing, ambiguous, and sport-centered comments?"),
        ("RQ3: Do likes, views, and retweet presence" + $sb + "cluster differently across those positions?"),
        ("Author's attention quantity is retained as a control," + $sb + "not as the main theoretical outcome.")
    ) -Left 56 -Top 126 -Width 760 -Height 280 -FontSize 15
    $null = Add-TextBox -Slide $slide -Text "Deliverable logic" -Left 58 -Top 430 -Width 180 -Height 26 -FontSize 16 -ColorHex "8B4A2A" -Bold -NoFill -NoLine
    $null = Add-BulletBox -Slide $slide -Bullets @(
        "Map the discourse field",
        "Compare stance groups",
        "Check whether visible reaction privileges some positions over others"
    ) -Left 58 -Top 456 -Width 360 -Height 90 -FontSize 18

    $slide = $presentation.Slides.Add(4, 12)
    Add-SectionHeader -Slide $slide -Title "Data Snapshot" -Subtitle "One local dataset, kept outside the public repository."
    Add-AccentRule -Slide $slide
    $null = Add-BulletBox -Slide $slide -Bullets @(
        "1,008 public comments collected from July 23 to August 1, 2024",
        "740 usernames, 845 unique comment IDs, and 4 source posts",
        "920 English-language comments",
        "770 Trump-related references",
        "30 moral-condemnation comments and 9 depoliticizing appeals",
        "Raw B50_X_COMMENT.xlsx stays local and outside the GitHub repo"
    ) -Left 66 -Top 132 -Width 760 -Height 300 -FontSize 16

    $slide = $presentation.Slides.Add(5, 12)
    Add-SectionHeader -Slide $slide -Title "Coding Strategy" -Subtitle "Transparent heuristic coding documented in the repo codebook."
    Add-AccentRule -Slide $slide
    $null = Add-BulletBox -Slide $slide -Bullets @(
        "Stance categories: pro-Trump, anti-Trump, depoliticizing, referential/unclear, sport-centered, other",
        ("Frame categories: political, blended sport-politics," + $sb + "sport, depoliticizing bridge, other"),
        "Additional indicators: moral condemnation and depoliticizing appeal",
        "Reaction metrics: likes, views, and retweet presence",
        ("Controls: attention quantity, verification, language," + $sb + "media outlets, and author posts")
    ) -Left 66 -Top 132 -Width 780 -Height 300 -FontSize 12

    $slide = $presentation.Slides.Add(6, 12)
    Add-SectionHeader -Slide $slide -Title "Finding 1: The Discussion Is Strongly Politicized" -Subtitle "Frame distribution shows that politics is central, not peripheral."
    Add-AccentRule -Slide $slide
    $null = Add-ImageScaled -Slide $slide -Path $figFrame -Left 510 -Top 116 -MaxWidth 340 -MaxHeight 300
    $null = Add-BulletBox -Slide $slide -Bullets @(
        "50.0% of comments are coded as political.",
        "25.6% are coded as blended sport-politics.",
        ("Only 4.1% are sport-centered," + $sb + "and 0.9% are depoliticizing bridge comments."),
        ("Trump's appearance is being processed as a political event," + $sb + "not just a golf-media event.")
    ) -Left 54 -Top 126 -Width 410 -Height 290 -FontSize 18

    $slide = $presentation.Slides.Add(7, 12)
    Add-SectionHeader -Slide $slide -Title "Finding 2: Support, Opposition, And Ambiguity" -Subtitle "The largest stance category is politically referential but unresolved."
    Add-AccentRule -Slide $slide
    $null = Add-ImageScaled -Slide $slide -Path $figStance -Left 500 -Top 118 -MaxWidth 340 -MaxHeight 300
    $null = Add-BulletBox -Slide $slide -Bullets @(
        "Trump-referential/unclear: 623 comments (61.8%).",
        "Pro-Trump/supportive: 108 comments (10.7%).",
        "Anti-Trump/oppositional: 30 comments (3.0%).",
        "Depoliticizing/bridge: 9 comments (0.9%).",
        ("Supportive comments outnumber oppositional comments" + $sb + "by more than three to one.")
    ) -Left 52 -Top 126 -Width 410 -Height 300 -FontSize 17

    $slide = $presentation.Slides.Add(8, 12)
    Add-SectionHeader -Slide $slide -Title "Finding 3: Visible Reaction Differs By Stance" -Subtitle "Engagement does not distribute evenly across discourse positions."
    Add-AccentRule -Slide $slide
    $null = Add-ImageScaled -Slide $slide -Path $figEngagement -Left 500 -Top 126 -MaxWidth 340 -MaxHeight 260
    $null = Add-BulletBox -Slide $slide -Bullets @(
        "Sport-centered comments have the highest median views: 915.",
        ("Supportive comments have median likes of 2," + $sb + "versus 1 for oppositional comments."),
        ("Oppositional comments remain present, but their median visibility" + $sb + "is much lower: 82 views."),
        ("The most visible comments are not necessarily" + $sb + "the most openly partisan ones.")
    ) -Left 48 -Top 126 -Width 420 -Height 300 -FontSize 17

    $slide = $presentation.Slides.Add(9, 12)
    Add-SectionHeader -Slide $slide -Title "Finding 4: High-Engagement Comment Mix" -Subtitle "Top-like and top-view comments are dominated by politically charged but not always explicit stance-taking."
    Add-AccentRule -Slide $slide
    $null = Add-ImageScaled -Slide $slide -Path $figTopDecile -Left 492 -Top 132 -MaxWidth 360 -MaxHeight 250
    $null = Add-BulletBox -Slide $slide -Bullets @(
        ("Referential/unclear comments are 61.8% overall," + $sb + "64.2% of the top-like decile, and 62.4% of the top-view decile."),
        ("Supportive comments rise from 10.7% overall" + $sb + "to 12.3% of the top-like decile."),
        ("Oppositional comments rise from 3.0% overall" + $sb + "to 5.7% of the top-like decile."),
        ("Explicit contestation is not numerically dominant," + $sb + "but it still captures attention.")
    ) -Left 44 -Top 120 -Width 420 -Height 304 -FontSize 16

    $slide = $presentation.Slides.Add(10, 12)
    Add-SectionHeader -Slide $slide -Title "Controlled Check: Attention Still Matters" -Subtitle "The discourse pivot does not erase author-side visibility advantages."
    Add-AccentRule -Slide $slide
    $card1 = $slide.Shapes.AddShape(1, 72, 160, 240, 130)
    $card1.Fill.ForeColor.RGB = 15594855
    $card1.Line.ForeColor.RGB = 3242405
    $null = Add-TextBox -Slide $slide -Text "Likes model" -Left 98 -Top 182 -Width 180 -Height 24 -FontSize 20 -ColorHex "8B4A2A" -Bold -NoFill -NoLine
    $null = Add-TextBox -Slide $slide -Text "b = 0.137`rp = 0.001" -Left 98 -Top 220 -Width 160 -Height 52 -FontSize 24 -FontName "Aptos Display" -ColorHex "22304A" -Bold -NoFill -NoLine
    $card2 = $slide.Shapes.AddShape(1, 350, 160, 240, 130)
    $card2.Fill.ForeColor.RGB = 15594855
    $card2.Line.ForeColor.RGB = 3242405
    $null = Add-TextBox -Slide $slide -Text "Views model" -Left 376 -Top 182 -Width 180 -Height 24 -FontSize 20 -ColorHex "8B4A2A" -Bold -NoFill -NoLine
    $null = Add-TextBox -Slide $slide -Text "b = 0.239`rp < 0.001" -Left 376 -Top 220 -Width 180 -Height 52 -FontSize 24 -FontName "Aptos Display" -ColorHex "22304A" -Bold -NoFill -NoLine
    $null = Add-BulletBox -Slide $slide -Bullets @(
        ("After controlling for verification, language, media outlets," + $sb + "and author posts, attention quantity still predicts likes and views."),
        ("Stance coefficients are less stable because" + $sb + "anti-Trump and depoliticizing categories are small.")
    ) -Left 82 -Top 336 -Width 720 -Height 176 -FontSize 12

    $slide = $presentation.Slides.Add(11, 12)
    Add-SectionHeader -Slide $slide -Title "Takeaways And Limits" -Subtitle "What the revised P3 study contributes."
    Add-AccentRule -Slide $slide
    $null = Add-BulletBox -Slide $slide -Bullets @(
        ("Break 50 is best understood as a politically saturated" + $sb + "discourse space where golf and politics are repeatedly forced together."),
        ("The largest category is not explicit support or opposition," + $sb + "but a broad referential middle zone."),
        ("Depoliticizing comments are rare," + $sb + "so the conversation rarely escapes the politics-sport overlap."),
        "The coding is heuristic: a transparent first pass, not final manual classification.",
        "Future improvement: add manual coding or intercoder review."
    ) -Left 54 -Top 126 -Width 760 -Height 300 -FontSize 13
    $endBox = $slide.Shapes.AddShape(1, 52, 430, 856, 62)
    $endBox.Fill.ForeColor.RGB = 16314870
    $endBox.Line.Visible = 0
    $null = Add-TextBox -Slide $slide -Text "Bottom line: the repo now supports a P3 argument about contested discourse, not just a generic engagement-correlation study." -Left 74 -Top 448 -Width 812 -Height 28 -FontSize 19 -ColorHex "22304A" -Bold -NoFill -NoLine

    $presentation.SaveAs($outputPath)
    $presentation.SaveAs($previewDir, 18)
    $presentation.Close()
    $ppt.Quit()

    Write-Output "Created: $outputPath"
    Write-Output "Preview: $previewDir"
}
finally {
    if ($presentation -ne $null) {
        try { $presentation.Close() } catch {}
    }
    if ($ppt -ne $null) {
        try { $ppt.Quit() } catch {}
    }
}
