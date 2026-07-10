param(
    [string]$OutputDirectory = $PSScriptRoot,
    [switch]$ExportPreview
)

$ErrorActionPreference = "Stop"

function Get-Rgb([int]$Red, [int]$Green, [int]$Blue) {
    return $Red + (256 * $Green) + (65536 * $Blue)
}

$script:Color = @{
    Sidebar = Get-Rgb 20 41 35
    SidebarSoft = Get-Rgb 29 55 48
    Primary = Get-Rgb 23 108 88
    PrimaryDark = Get-Rgb 17 82 68
    Mint = Get-Rgb 228 242 237
    MintStrong = Get-Rgb 205 233 223
    Background = Get-Rgb 244 246 244
    Surface = Get-Rgb 255 255 255
    SurfaceSoft = Get-Rgb 248 250 248
    Text = Get-Rgb 23 33 30
    Muted = Get-Rgb 96 113 106
    Border = Get-Rgb 220 228 223
    Amber = Get-Rgb 184 112 18
    AmberSoft = Get-Rgb 255 242 217
    Red = Get-Rgb 179 61 69
    RedSoft = Get-Rgb 251 234 236
    Blue = Get-Rgb 40 106 155
    BlueSoft = Get-Rgb 231 242 250
}

$script:SlideWidth = 960
$script:SlideHeight = 540
$script:FontBody = "Aptos"
$script:FontDisplay = "Aptos Display"

function Add-Box {
    param(
        $Slide,
        [double]$Left,
        [double]$Top,
        [double]$Width,
        [double]$Height,
        [int]$Fill,
        [int]$Line = 0,
        [double]$Radius = 5,
        [double]$Transparency = 0
    )

    $shapeType = if ($Radius -gt 0) { 5 } else { 1 }
    $shape = $Slide.Shapes.AddShape($shapeType, $Left, $Top, $Width, $Height)
    $shape.Fill.Solid()
    $shape.Fill.ForeColor.RGB = $Fill
    $shape.Fill.Transparency = $Transparency
    if ($Line -eq 0) {
        $shape.Line.Visible = 0
    } else {
        $shape.Line.Visible = -1
        $shape.Line.ForeColor.RGB = $Line
        $shape.Line.Weight = 0.8
    }
    return $shape
}

function Add-Text {
    param(
        $Slide,
        [string]$Text,
        [double]$Left,
        [double]$Top,
        [double]$Width,
        [double]$Height,
        [double]$Size = 16,
        [int]$Color = $script:Color.Text,
        [bool]$Bold = $false,
        [int]$Align = 1,
        [string]$Font = $script:FontBody,
        [int]$Vertical = 1
    )

    $shape = $Slide.Shapes.AddTextbox(1, $Left, $Top, $Width, $Height)
    $shape.TextFrame2.MarginLeft = 0
    $shape.TextFrame2.MarginRight = 0
    $shape.TextFrame2.MarginTop = 0
    $shape.TextFrame2.MarginBottom = 0
    $shape.TextFrame2.WordWrap = -1
    $shape.TextFrame2.VerticalAnchor = $Vertical
    $range = $shape.TextFrame2.TextRange
    $range.Text = $Text
    $range.Font.Name = $Font
    $range.Font.Size = $Size
    $range.Font.Bold = if ($Bold) { -1 } else { 0 }
    $range.Font.Fill.ForeColor.RGB = $Color
    $range.ParagraphFormat.Alignment = $Align
    $range.ParagraphFormat.SpaceAfter = 0
    return $shape
}

function Add-Line {
    param($Slide, [double]$X1, [double]$Y1, [double]$X2, [double]$Y2, [int]$Color, [double]$Weight = 1.2)
    $line = $Slide.Shapes.AddLine($X1, $Y1, $X2, $Y2)
    $line.Line.ForeColor.RGB = $Color
    $line.Line.Weight = $Weight
    return $line
}

function Add-TextCard {
    param(
        $Slide,
        [string]$Text,
        [double]$Left,
        [double]$Top,
        [double]$Width,
        [double]$Height,
        [int]$Fill,
        [int]$TextColor,
        [double]$Size,
        [bool]$Bold = $false,
        [double]$Radius = 0,
        [double]$Padding = 0
    )

    $shape = Add-Box $Slide $Left $Top $Width $Height $Fill 0 $Radius
    $shape.TextFrame2.MarginLeft = $Padding
    $shape.TextFrame2.MarginRight = $Padding
    $shape.TextFrame2.MarginTop = 0
    $shape.TextFrame2.MarginBottom = 0
    $shape.TextFrame2.WordWrap = -1
    $shape.TextFrame2.VerticalAnchor = 3
    $shape.TextFrame2.TextRange.Text = $Text
    $shape.TextFrame2.TextRange.Font.Name = $script:FontBody
    $shape.TextFrame2.TextRange.Font.Size = $Size
    $shape.TextFrame2.TextRange.Font.Bold = if ($Bold) { -1 } else { 0 }
    $shape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = $TextColor
    $shape.TextFrame2.TextRange.ParagraphFormat.Alignment = 1
    return $shape
}

function Add-CircleLabel {
    param($Slide, [string]$Text, [double]$Left, [double]$Top, [double]$Size, [int]$Fill, [int]$TextColor = $script:Color.Surface, [double]$FontSize = 13)
    $circle = $Slide.Shapes.AddShape(9, $Left, $Top, $Size, $Size)
    $circle.Fill.Solid()
    $circle.Fill.ForeColor.RGB = $Fill
    $circle.Line.Visible = 0
    $circle.TextFrame2.MarginLeft = 0
    $circle.TextFrame2.MarginRight = 0
    $circle.TextFrame2.MarginTop = 0
    $circle.TextFrame2.MarginBottom = 0
    $circle.TextFrame2.VerticalAnchor = 3
    $circle.TextFrame2.TextRange.Text = $Text
    $circle.TextFrame2.TextRange.Font.Name = $script:FontBody
    $circle.TextFrame2.TextRange.Font.Size = $FontSize
    $circle.TextFrame2.TextRange.Font.Bold = -1
    $circle.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = $TextColor
    $circle.TextFrame2.TextRange.ParagraphFormat.Alignment = 2
    return $circle
}

function Add-SectionHeader {
    param($Slide, [string]$Number, [string]$Title, [string]$Subtitle)
    $null = Add-Text $Slide $Number 48 32 320 18 10 $script:Color.Primary $true 1 $script:FontBody
    $null = Add-Text $Slide $Title 48 58 800 42 28 $script:Color.Text $true 1 $script:FontDisplay
    if ($Subtitle) {
        $null = Add-Text $Slide $Subtitle 48 104 790 32 13 $script:Color.Muted $false
    }
}

function Add-Footer {
    param($Slide, [int]$SlideNumber, [bool]$Dark = $false)
    $lineColor = if ($Dark) { Get-Rgb 52 78 70 } else { $script:Color.Border }
    $textColor = if ($Dark) { Get-Rgb 151 173 165 } else { $script:Color.Muted }
    $null = Add-Line $Slide 48 508 912 508 $lineColor 0.7
    $null = Add-Text $Slide "DIGITAL ATTENDANCE | CLIENT PRESENTATION" 48 515 420 14 8 $textColor $true
    $null = Add-Text $Slide ([string]$SlideNumber).PadLeft(2, "0") 860 515 52 14 8 $textColor $true 3
}

function Add-FeatureCard {
    param($Slide, [double]$Left, [double]$Top, [double]$Width, [double]$Height, [string]$Number, [string]$Title, [string]$Body, [int]$Accent)
    $null = Add-Box $Slide $Left $Top $Width $Height $script:Color.Surface $script:Color.Border 5
    $null = Add-CircleLabel $Slide $Number ($Left + 18) ($Top + 17) 34 $Accent $script:Color.Surface 10
    $null = Add-Text $Slide $Title ($Left + 18) ($Top + 62) ($Width - 36) 25 14 $script:Color.Text $true
    $null = Add-Text $Slide $Body ($Left + 18) ($Top + 92) ($Width - 36) ($Height - 106) 10.5 $script:Color.Muted $false
}

function Add-Bullet {
    param($Slide, [double]$Left, [double]$Top, [double]$Width, [string]$Title, [string]$Body, [int]$Accent = $script:Color.Primary)
    $null = Add-CircleLabel $Slide "" $Left ($Top + 4) 10 $Accent $script:Color.Surface 6
    $null = Add-Text $Slide $Title ($Left + 22) $Top ($Width - 22) 22 13 $script:Color.Text $true
    $null = Add-Text $Slide $Body ($Left + 22) ($Top + 24) ($Width - 22) 42 10.5 $script:Color.Muted $false
}

New-Item -ItemType Directory -Force $OutputDirectory | Out-Null
$pptxPath = Join-Path $OutputDirectory "Digital_Attendance_Client_Presentation.pptx"
$pdfPath = Join-Path $OutputDirectory "Digital_Attendance_Client_Presentation.pdf"
$previewPath = Join-Path $OutputDirectory "presentation_preview"

$powerPoint = $null
$presentation = $null

try {
    $powerPoint = New-Object -ComObject PowerPoint.Application
    $presentation = $powerPoint.Presentations.Add()
    $presentation.PageSetup.SlideWidth = $script:SlideWidth
    $presentation.PageSetup.SlideHeight = $script:SlideHeight

    # Slide 1: Cover
    $slide = $presentation.Slides.Add(1, 12)
    $slide.FollowMasterBackground = 0
    $slide.Background.Fill.Solid()
    $slide.Background.Fill.ForeColor.RGB = $script:Color.Sidebar
    $accentPanel = Add-Box $slide 640 0 320 540 $script:Color.Primary 0 0
    $accentPanel.Fill.Transparency = 0.06
    $ring1 = $slide.Shapes.AddShape(9, 695, 75, 330, 330)
    $ring1.Fill.Visible = 0
    $ring1.Line.Visible = -1
    $ring1.Line.ForeColor.RGB = Get-Rgb 94 170 149
    $ring1.Line.Transparency = 0.45
    $ring1.Line.Weight = 1.5
    $ring2 = $slide.Shapes.AddShape(9, 760, 140, 200, 200)
    $ring2.Fill.Visible = 0
    $ring2.Line.Visible = -1
    $ring2.Line.ForeColor.RGB = Get-Rgb 178 223 210
    $ring2.Line.Transparency = 0.35
    $ring2.Line.Weight = 1.2
    $brand = Add-Box $slide 54 45 46 46 $script:Color.Primary 0 5
    $null = Add-Text $slide "FA" 54 45 46 46 12 $script:Color.Surface $true 2 $script:FontBody 3
    $null = Add-Text $slide "FACIAL ATTENDANCE" 112 49 260 18 10 (Get-Rgb 183 207 199) $true
    $null = Add-Text $slide "WORKFORCE OPERATIONS" 112 69 260 16 8 (Get-Rgb 117 152 141) $true
    $null = Add-Text $slide "Digital Facial`nAttendance System" 54 150 560 125 38 $script:Color.Surface $true 1 $script:FontDisplay
    $null = Add-Text $slide "A secure, camera-powered attendance workspace for employee enrollment, daily recognition, policy decisions, and controlled reporting." 56 294 510 72 15 (Get-Rgb 198 216 210) $false
    $null = Add-Box $slide 56 395 205 34 (Get-Rgb 33 62 53) (Get-Rgb 60 91 81) 5
    $null = Add-Text $slide "CLIENT SOLUTION OVERVIEW" 70 395 178 34 9 (Get-Rgb 208 228 221) $true 2 $script:FontBody 3
    $null = Add-Text $slide "Prepared for client review | July 2026" 56 477 360 18 9 (Get-Rgb 130 160 150) $false

    # Slide 2: Executive snapshot
    $slide = $presentation.Slides.Add(2, 12)
    $slide.Background.Fill.Solid(); $slide.Background.Fill.ForeColor.RGB = $script:Color.Background
    Add-SectionHeader $slide "01 / EXECUTIVE SNAPSHOT" "A simpler way to manage daily attendance" "One focused application brings recognition, records, policy rules, and admin decisions together."
    Add-FeatureCard $slide 48 158 204 280 "01" "Faster attendance" "Employees are recognized from a connected camera, reducing repetitive manual entry at the attendance point." $script:Color.Primary
    Add-FeatureCard $slide 268 158 204 280 "02" "Clear decisions" "Daily status, short-day review, and overtime are calculated from consistent, visible business rules." $script:Color.Blue
    Add-FeatureCard $slide 488 158 204 280 "03" "Controlled access" "Admin and standard-user roles separate daily scanning from sensitive employee and policy controls." $script:Color.Amber
    Add-FeatureCard $slide 708 158 204 280 "04" "Local data ownership" "Employees, users, settings, attendance, and decisions remain in a client-controlled Excel workbook." $script:Color.PrimaryDark
    Add-Footer $slide 2

    # Slide 3: End-to-end workflow
    $slide = $presentation.Slides.Add(3, 12)
    $slide.Background.Fill.Solid(); $slide.Background.Fill.ForeColor.RGB = $script:Color.Surface
    Add-SectionHeader $slide "02 / PRODUCT FLOW" "One system, from enrollment to final decision" "The workflow keeps daily operations simple while preserving admin control over exceptions."
    $steps = @(
        @{N="1"; T="Create employee"; B="Code, name, department, and active status"; C=$script:Color.Primary},
        @{N="2"; T="Register face"; B="Capture one clear face from the system camera"; C=$script:Color.Blue},
        @{N="3"; T="Recognize"; B="Continuously scan active registered employees"; C=$script:Color.Primary},
        @{N="4"; T="Calculate"; B="Duration, status, and overtime are derived daily"; C=$script:Color.Amber},
        @{N="5"; T="Review"; B="Admin finalizes short or incomplete days"; C=$script:Color.Red}
    )
    for ($i = 0; $i -lt $steps.Count; $i++) {
        $left = 48 + ($i * 174)
        if ($i -lt ($steps.Count - 1)) {
            $connector = Add-Line $slide ($left + 138) 256 ($left + 173) 256 $script:Color.Border 2
            $connector.Line.EndArrowheadStyle = 2
        }
        $null = Add-CircleLabel $slide $steps[$i].N ($left + 45) 188 56 $steps[$i].C $script:Color.Surface 16
        $null = Add-Text $slide $steps[$i].T $left 270 146 28 13 $script:Color.Text $true 2
        $null = Add-Text $slide $steps[$i].B $left 304 146 62 10 $script:Color.Muted $false 2
    }
    $null = Add-Box $slide 166 405 628 52 $script:Color.Mint 0 5
    $null = Add-Text $slide "Result: a traceable attendance day with clear ownership of every exception." 190 405 580 52 12 $script:Color.PrimaryDark $true 2 $script:FontBody 3
    Add-Footer $slide 3

    # Slide 4: Dashboard
    $slide = $presentation.Slides.Add(4, 12)
    $slide.Background.Fill.Solid(); $slide.Background.Fill.ForeColor.RGB = $script:Color.Background
    Add-SectionHeader $slide "03 / OPERATIONS DASHBOARD" "Designed for action, not just reporting" ""
    Add-Bullet $slide 50 170 286 "Immediate visibility" "Six concise indicators show workforce readiness, attendance outcomes, and pending review." $script:Color.Primary
    Add-Bullet $slide 50 260 286 "Fast access" "Primary scanner and employee-registration actions remain visible without searching through menus." $script:Color.Blue
    Add-Bullet $slide 50 350 286 "Exception focus" "Today summary and recent activity help admins identify missing scans and resolve short days." $script:Color.Amber
    $frame = Add-Box $slide 366 151 546 323 $script:Color.Surface $script:Color.Border 5
    $null = Add-Box $slide 366 151 92 323 $script:Color.Sidebar 0 5
    $null = Add-Box $slide 380 170 36 36 $script:Color.Primary 0 4
    $null = Add-Text $slide "FA" 380 170 36 36 9 $script:Color.Surface $true 2 $script:FontBody 3
    for ($i=0; $i -lt 5; $i++) {
        $fill = if ($i -eq 0) { $script:Color.Primary } else { $script:Color.SidebarSoft }
        $null = Add-Box $slide 380 (230 + $i*38) 64 28 $fill 0 3
        $null = Add-CircleLabel $slide ([string]($i+1)) 388 (235 + $i*38) 17 (Get-Rgb 72 111 99) $script:Color.Surface 6
    }
    $null = Add-Box $slide 458 151 454 42 $script:Color.SurfaceSoft 0 0
    $null = Add-Text $slide "Dashboard" 474 162 150 18 10 $script:Color.Text $true
    $cardColors = @($script:Color.Mint, $script:Color.BlueSoft, $script:Color.Mint, $script:Color.AmberSoft, $script:Color.RedSoft)
    for ($i=0; $i -lt 5; $i++) {
        $x = 474 + $i*84
        $metricValue = if ($i -eq 0) { 24 } elseif ($i -eq 1) { 21 } else { 3 }
        $null = Add-Box $slide $x 214 72 70 $script:Color.Surface $script:Color.Border 4
        $null = Add-Box $slide ($x+9) 223 18 18 $cardColors[$i] 0 3
        $null = Add-Text $slide ([string]$metricValue) ($x+10) 248 54 22 14 $script:Color.Text $true
    }
    $null = Add-Box $slide 474 301 270 143 $script:Color.Surface $script:Color.Border 4
    $null = Add-Text $slide "TODAY" 489 315 70 14 7 $script:Color.Primary $true
    $null = Add-Text $slide "Attendance summary" 489 333 160 20 11 $script:Color.Text $true
    for ($i=0; $i -lt 3; $i++) { $null = Add-Line $slide 489 (372+$i*22) 728 (372+$i*22) $script:Color.Border 0.7 }
    $null = Add-Box $slide 758 301 138 143 $script:Color.Surface $script:Color.Border 4
    $null = Add-Text $slide "RECENT ACTIVITY" 771 315 110 14 7 $script:Color.Primary $true
    for ($i=0; $i -lt 3; $i++) { $null = Add-CircleLabel $slide "" 772 (349+$i*27) 16 $script:Color.Mint $script:Color.Primary 5; $null = Add-Line $slide 797 (357+$i*27) 878 (357+$i*27) $script:Color.Border 1.5 }
    $null = Add-Text $slide "The home screen surfaces the workforce signals and next actions that matter today." 48 104 790 32 13 $script:Color.Muted $false
    Add-Footer $slide 4

    # Slide 5: Face enrollment
    $slide = $presentation.Slides.Add(5, 12)
    $slide.Background.Fill.Solid(); $slide.Background.Fill.ForeColor.RGB = $script:Color.Surface
    Add-SectionHeader $slide "04 / EMPLOYEE ENROLLMENT" "A guided two-step registration experience" "Employee records and biometric references stay connected through a clear, deliberate workflow."
    $null = Add-Box $slide 48 160 410 304 $script:Color.Background $script:Color.Border 5
    $null = Add-CircleLabel $slide "1" 74 185 38 $script:Color.Primary $script:Color.Surface 12
    $null = Add-Text $slide "Save employee details" 128 185 270 24 16 $script:Color.Text $true
    $null = Add-Text $slide "Capture the core workforce record first so every face reference is linked to a unique employee ID." 74 235 320 48 11 $script:Color.Muted $false
    $fields = @("Employee code", "Full name", "Department", "Status")
    for ($i=0; $i -lt 4; $i++) {
        $x = 74 + (($i % 2) * 174)
        $y = 305 + ([math]::Floor($i / 2) * 58)
        $null = Add-Text $slide $fields[$i] $x $y 152 14 8 $script:Color.Muted $true
        $null = Add-Box $slide $x ($y+18) 152 30 $script:Color.Surface $script:Color.Border 3
    }
    $null = Add-Box $slide 488 160 424 304 $script:Color.Sidebar 0 5
    $null = Add-CircleLabel $slide "2" 514 185 38 $script:Color.Primary $script:Color.Surface 12
    $null = Add-Text $slide "Capture one clear face" 568 185 280 24 16 $script:Color.Surface $true
    $camera = Add-Box $slide 526 235 210 166 (Get-Rgb 9 22 18) (Get-Rgb 64 99 87) 5
    $guide = $slide.Shapes.AddShape(9, 597, 263, 70, 100)
    $guide.Fill.Visible = 0; $guide.Line.Visible = -1; $guide.Line.ForeColor.RGB = Get-Rgb 177 224 210; $guide.Line.Weight = 1.3
    $null = Add-Box $slide 754 246 130 42 $script:Color.Primary 0 4
    $null = Add-Text $slide "START CAMERA" 754 246 130 42 9 $script:Color.Surface $true 2 $script:FontBody 3
    $null = Add-Box $slide 754 302 130 42 $script:Color.Surface 0 4
    $null = Add-Text $slide "CAPTURE FACE" 754 302 130 42 9 $script:Color.PrimaryDark $true 2 $script:FontBody 3
    $null = Add-Text $slide "One person in frame`nFront-facing and well lit`nBrowser camera permission required" 754 365 138 64 9 (Get-Rgb 182 204 196) $false
    Add-Footer $slide 5

    # Slide 6: Continuous attendance
    $slide = $presentation.Slides.Add(6, 12)
    $slide.Background.Fill.Solid(); $slide.Background.Fill.ForeColor.RGB = $script:Color.Background
    Add-SectionHeader $slide "05 / LIVE ATTENDANCE" "Built to stay on throughout the working day" "Continuous monitoring recognizes eligible employees while configurable cooldown prevents rapid duplicate scans."
    $null = Add-Box $slide 48 165 864 282 $script:Color.Surface $script:Color.Border 5
    $null = Add-CircleLabel $slide "A" 95 245 62 $script:Color.Primary $script:Color.Surface 18
    $null = Add-Text $slide "Camera feed" 70 322 112 22 13 $script:Color.Text $true 2
    $null = Add-Text $slide "Connected system camera" 63 348 126 34 9 $script:Color.Muted $false 2
    $null = Add-Line $slide 175 276 275 276 $script:Color.Border 2
    $null = Add-CircleLabel $slide "B" 292 245 62 $script:Color.Blue $script:Color.Surface 18
    $null = Add-Text $slide "Face match" 270 322 106 22 13 $script:Color.Text $true 2
    $null = Add-Text $slide "Active registered employees only" 255 348 138 34 9 $script:Color.Muted $false 2
    $null = Add-Line $slide 372 276 472 276 $script:Color.Border 2
    $null = Add-CircleLabel $slide "C" 489 245 62 $script:Color.Amber $script:Color.Surface 18
    $null = Add-Text $slide "Cooldown check" 464 322 114 22 13 $script:Color.Text $true 2
    $null = Add-Text $slide "Default 5 minutes, admin configurable" 451 348 142 34 9 $script:Color.Muted $false 2
    $null = Add-Line $slide 569 276 669 276 $script:Color.Border 2
    $null = Add-CircleLabel $slide "D" 686 245 62 $script:Color.PrimaryDark $script:Color.Surface 18
    $null = Add-Text $slide "Timestamp saved" 664 322 110 22 13 $script:Color.Text $true 2
    $null = Add-Text $slide "Event added to the Excel attendance sheet" 646 348 146 40 9 $script:Color.Muted $false 2
    $null = Add-Box $slide 801 224 78 104 $script:Color.Mint 0 5
    $null = Add-Text $slide "READY" 801 242 78 16 8 $script:Color.Primary $true 2
    $null = Add-Text $slide "24/7" 801 270 78 28 18 $script:Color.PrimaryDark $true 2
    Add-Footer $slide 6

    # Slide 7: Attendance policy
    $slide = $presentation.Slides.Add(7, 12)
    $slide.Background.Fill.Solid(); $slide.Background.Fill.ForeColor.RGB = $script:Color.Surface
    Add-SectionHeader $slide "06 / POLICY ENGINE" "Consistent rules with human control where needed" "Daily duration is measured from the first recognition timestamp to the last timestamp for that employee."
    $null = Add-Line $slide 100 233 860 233 $script:Color.Border 5
    $null = Add-Line $slide 100 233 480 233 $script:Color.Red 5
    $null = Add-Line $slide 480 233 720 233 $script:Color.Amber 5
    $null = Add-Line $slide 720 233 860 233 $script:Color.Primary 5
    $points = @(
        @{X=100; T="First scan"; B="Start of measured day"; C=$script:Color.Text},
        @{X=480; T="4 hours"; B="Half-day threshold"; C=$script:Color.Amber},
        @{X=720; T="8 hours"; B="Present threshold"; C=$script:Color.Primary},
        @{X=860; T="Last scan"; B="Extra time becomes OT"; C=$script:Color.Blue}
    )
    foreach ($point in $points) {
        $null = Add-CircleLabel $slide "" ($point.X-8) 225 16 $point.C $script:Color.Surface 6
        $null = Add-Text $slide $point.T ($point.X-55) 258 110 20 11 $script:Color.Text $true 2
        $null = Add-Text $slide $point.B ($point.X-65) 282 130 30 8.5 $script:Color.Muted $false 2
    }
    $null = Add-Box $slide 64 350 252 82 $script:Color.RedSoft 0 5
    $null = Add-Text $slide "BELOW 4 HOURS" 82 365 216 16 9 $script:Color.Red $true
    $null = Add-Text $slide "Sent to admin for final decision" 82 390 216 28 12 $script:Color.Text $true
    $null = Add-Box $slide 354 350 252 82 $script:Color.AmberSoft 0 5
    $null = Add-Text $slide "4 TO 8 HOURS" 372 365 216 16 9 $script:Color.Amber $true
    $null = Add-Text $slide "Automatically marked Half-day" 372 390 216 28 12 $script:Color.Text $true
    $null = Add-Box $slide 644 350 252 82 $script:Color.Mint 0 5
    $null = Add-Text $slide "8 HOURS OR MORE" 662 365 216 16 9 $script:Color.Primary $true
    $null = Add-Text $slide "Present, with excess counted as OT" 662 390 216 28 12 $script:Color.Text $true
    Add-Footer $slide 7

    # Slide 8: Admin controls
    $slide = $presentation.Slides.Add(8, 12)
    $slide.Background.Fill.Solid(); $slide.Background.Fill.ForeColor.RGB = $script:Color.Background
    Add-SectionHeader $slide "07 / ADMIN GOVERNANCE" "Exceptions remain visible and accountable" "Admin-only controls provide a final decision path without changing the underlying scan history."
    $null = Add-Box $slide 48 160 510 310 $script:Color.Surface $script:Color.Border 5
    $null = Add-Text $slide "FINAL DECISION QUEUE" 72 182 230 16 9 $script:Color.Primary $true
    $null = Add-Text $slide "Employee attendance review" 72 207 300 25 17 $script:Color.Text $true
    $null = Add-Box $slide 72 252 462 112 $script:Color.SurfaceSoft $script:Color.Border 4
    $null = Add-CircleLabel $slide "JS" 90 271 38 $script:Color.MintStrong $script:Color.PrimaryDark 10
    $null = Add-Text $slide "Jordan Smith" 143 270 170 18 11 $script:Color.Text $true
    $null = Add-Text $slide "EMP-014 | Operations" 143 292 170 16 8 $script:Color.Muted $false
    $null = Add-Box $slide 396 271 112 24 $script:Color.AmberSoft 0 5
    $null = Add-Text $slide "NEEDS REVIEW" 396 271 112 24 8 $script:Color.Amber $true 2 $script:FontBody 3
    $metrics = @("First 09:06", "Last 11:48", "Duration 2h 42m")
    for ($i=0; $i -lt 3; $i++) { $null = Add-Text $slide $metrics[$i] (90+$i*138) 330 120 15 8 $script:Color.Muted $true }
    $null = Add-Box $slide 72 386 205 40 $script:Color.Surface $script:Color.Border 4
    $null = Add-Text $slide "Final status: Absent" 86 386 175 40 9 $script:Color.Text $true 1 $script:FontBody 3
    $null = Add-Box $slide 293 386 126 40 $script:Color.Primary 0 4
    $null = Add-Text $slide "SAVE DECISION" 293 386 126 40 9 $script:Color.Surface $true 2 $script:FontBody 3
    Add-Bullet $slide 610 175 280 "Employee master" "Create, edit, activate, deactivate, delete, and maintain face registration." $script:Color.Primary
    Add-Bullet $slide 610 265 280 "User management" "Create software users and assign admin or standard access roles." $script:Color.Blue
    Add-Bullet $slide 610 355 280 "Settings control" "Adjust recognition cooldown while keeping attendance rules transparent." $script:Color.Amber
    Add-Footer $slide 8

    # Slide 9: Architecture
    $slide = $presentation.Slides.Add(9, 12)
    $slide.Background.Fill.Solid(); $slide.Background.Fill.ForeColor.RGB = $script:Color.Surface
    Add-SectionHeader $slide "08 / DATA ARCHITECTURE" "Excel-only persistence, by design" "The solution avoids an external database while keeping operational records structured and locally controlled."
    $null = Add-Box $slide 54 172 220 230 $script:Color.Sidebar 0 5
    $null = Add-Text $slide "LOCAL WEB APP" 78 194 170 16 9 (Get-Rgb 142 180 168) $true 2
    $null = Add-CircleLabel $slide "FA" 124 232 78 $script:Color.Primary $script:Color.Surface 20
    $null = Add-Text $slide "Flask application" 80 329 168 22 14 $script:Color.Surface $true 2
    $null = Add-Text $slide "Browser camera + face recognition" 74 359 180 32 9 (Get-Rgb 181 205 197) $false 2
    $connector = Add-Line $slide 275 286 372 286 $script:Color.Border 3
    $connector.Line.EndArrowheadStyle = 2
    $null = Add-Box $slide 382 156 272 262 $script:Color.Mint $script:Color.MintStrong 5
    $null = Add-Text $slide "ATTENDANCE_RECORDS.XLSX" 405 178 226 20 10 $script:Color.PrimaryDark $true 2
    $sheets = @("Employees", "Attendance", "Users", "Settings", "Decisions")
    for ($i=0; $i -lt $sheets.Count; $i++) {
        $null = Add-Box $slide 414 (218+$i*35) 208 27 $script:Color.Surface 0 3
        $null = Add-CircleLabel $slide ([string]($i+1)) 424 (224+$i*35) 15 $script:Color.Primary $script:Color.Surface 6
        $null = Add-Text $slide $sheets[$i] 449 (221+$i*35) 150 20 9 $script:Color.Text $true 1 $script:FontBody 3
    }
    $connector = Add-Line $slide 655 286 724 286 $script:Color.Border 3
    $connector.Line.EndArrowheadStyle = 2
    $null = Add-Box $slide 734 199 174 174 $script:Color.BlueSoft $script:Color.Border 5
    $null = Add-CircleLabel $slide "IMG" 788 226 66 $script:Color.Blue $script:Color.Surface 12
    $null = Add-Text $slide "Employee face images" 754 310 134 22 11 $script:Color.Text $true 2
    $null = Add-Text $slide "Stored in the local data folder" 754 340 134 28 8.5 $script:Color.Muted $false 2
    $null = Add-Text $slide "No external database | Simple backup | Client-controlled files" 160 449 640 20 11 $script:Color.PrimaryDark $true 2
    Add-Footer $slide 9

    # Slide 10: Responsive design
    $slide = $presentation.Slides.Add(10, 12)
    $slide.Background.Fill.Solid(); $slide.Background.Fill.ForeColor.RGB = $script:Color.Background
    Add-SectionHeader $slide "09 / RESPONSIVE EXPERIENCE" "One interface across operational screens" ""
    # Laptop
    $null = Add-Box $slide 62 177 470 246 (Get-Rgb 35 43 40) 0 5
    $null = Add-Box $slide 74 189 446 220 $script:Color.Surface 0 2
    $null = Add-Box $slide 74 189 75 220 $script:Color.Sidebar 0 0
    $null = Add-Box $slide 149 189 371 32 $script:Color.SurfaceSoft 0 0
    for ($i=0; $i -lt 5; $i++) { $null = Add-Box $slide (166+$i*66) 244 54 52 $script:Color.Surface $script:Color.Border 3 }
    $null = Add-Box $slide 166 313 210 74 $script:Color.Surface $script:Color.Border 3
    $null = Add-Box $slide 390 313 112 74 $script:Color.Surface $script:Color.Border 3
    $null = Add-Box $slide 42 423 510 13 (Get-Rgb 56 65 61) 0 3
    $null = Add-Text $slide "LAPTOP / DESKTOP" 200 449 190 16 9 $script:Color.Muted $true 2
    # Tablet
    $null = Add-Box $slide 594 185 190 246 (Get-Rgb 35 43 40) 0 5
    $null = Add-Box $slide 605 198 168 220 $script:Color.Surface 0 2
    $null = Add-Box $slide 605 198 168 28 $script:Color.SurfaceSoft 0 0
    for ($i=0; $i -lt 4; $i++) { $null = Add-Box $slide (618+($i%2)*72) (248+[math]::Floor($i/2)*62) 62 52 $script:Color.Surface $script:Color.Border 3 }
    $null = Add-Box $slide 618 374 134 26 $script:Color.Mint 0 3
    $null = Add-Text $slide "TABLET" 635 449 110 16 9 $script:Color.Muted $true 2
    # Mobile
    $null = Add-Box $slide 822 215 82 202 (Get-Rgb 35 43 40) 0 5
    $null = Add-Box $slide 828 226 70 178 $script:Color.Surface 0 3
    $null = Add-Box $slide 828 226 70 20 $script:Color.Sidebar 0 0
    for ($i=0; $i -lt 3; $i++) { $null = Add-Box $slide 836 (261+$i*38) 54 28 $script:Color.Surface $script:Color.Border 2 }
    $null = Add-Text $slide "MOBILE" 807 449 112 16 9 $script:Color.Muted $true 2
    $null = Add-TextCard $slide "The shared layout adapts from mobile attendance checks to desktop administration and large displays." 48 104 790 32 $script:Color.Background $script:Color.Muted 13 $false 0 0
    Add-Footer $slide 10

    # Slide 11: Security and operations
    $slide = $presentation.Slides.Add(11, 12)
    $slide.Background.Fill.Solid(); $slide.Background.Fill.ForeColor.RGB = $script:Color.Surface
    Add-SectionHeader $slide "10 / OPERATING CONTROLS" "Secure operation starts with clear ownership" "The application includes practical controls today, supported by a straightforward client go-live checklist."
    $null = Add-Box $slide 48 155 410 318 $script:Color.Mint $script:Color.MintStrong 5
    $null = Add-Text $slide "BUILT INTO THE APPLICATION" 72 179 300 17 9 $script:Color.Primary $true
    $builtIn = @(
        "Password hashes instead of stored plain-text passwords",
        "Admin-only employee, user, settings, and decision controls",
        "Active and inactive access status for employees and users",
        "Clean error handling for workbook locks and invalid actions",
        "Local storage of attendance records and face references"
    )
    for ($i=0; $i -lt $builtIn.Count; $i++) {
        $null = Add-CircleLabel $slide ([string]($i+1)) 72 (216+$i*45) 24 $script:Color.Primary $script:Color.Surface 8
        $null = Add-Text $slide $builtIn[$i] 108 (216+$i*45) 320 36 10.5 $script:Color.Text $false 1 $script:FontBody 3
    }
    $null = Add-Box $slide 486 155 426 318 $script:Color.Background $script:Color.Border 5
    $null = Add-Text $slide "CLIENT GO-LIVE CHECKLIST" 510 179 310 17 9 $script:Color.Blue $true
    $checklist = @(
        "Change the default admin password before use",
        "Define a regular workbook and face-image backup routine",
        "Keep the workbook closed while the application writes records",
        "Restrict workstation and folder access to authorized personnel",
        "Use HTTPS when exposing camera access beyond localhost"
    )
    for ($i=0; $i -lt $checklist.Count; $i++) {
        $null = Add-CircleLabel $slide "" 510 (216+$i*45) 24 $script:Color.BlueSoft $script:Color.Blue 8
        $null = Add-Text $slide "OK" 510 (216+$i*45) 24 24 7 $script:Color.Blue $true 2 $script:FontBody 3
        $null = Add-Text $slide $checklist[$i] 546 (216+$i*45) 330 36 10.5 $script:Color.Text $false 1 $script:FontBody 3
    }
    Add-Footer $slide 11

    # Slide 12: Rollout
    $slide = $presentation.Slides.Add(12, 12)
    $slide.Background.Fill.Solid(); $slide.Background.Fill.ForeColor.RGB = $script:Color.Surface
    $null = Add-Text $slide "11 / RECOMMENDED ROLLOUT" 54 40 300 18 10 (Get-Rgb 129 190 172) $true
    $null = Add-Text $slide "Start with policy confirmation, validate on a pilot workstation, then move into controlled daily use." 54 132 720 34 13 $script:Color.Muted $false
    $phases = @(
        @{N="01"; T="Confirm"; B="Attendance thresholds, cooldown, roles, and review ownership"},
        @{N="02"; T="Pilot"; B="Register a small employee group and validate camera positioning"},
        @{N="03"; T="Train"; B="Brief admins on enrollment, decisions, backup, and recovery"},
        @{N="04"; T="Go live"; B="Run continuously, monitor exceptions, and review adoption"}
    )
    for ($i=0; $i -lt $phases.Count; $i++) {
        $x = 54 + $i*219
        $null = Add-Box $slide $x 201 197 157 (Get-Rgb 29 55 48) (Get-Rgb 54 83 73) 5
        $null = Add-Text $slide $phases[$i].N ($x+18) 219 44 18 9 (Get-Rgb 126 187 169) $true
        $null = Add-Text $slide $phases[$i].T ($x+18) 252 160 25 16 $script:Color.Surface $true
        $null = Add-Text $slide $phases[$i].B ($x+18) 288 158 53 10 (Get-Rgb 178 202 194) $false
    }
    $null = Add-Box $slide 54 397 854 64 $script:Color.Primary 0 5
    $null = Add-Text $slide "Ready for a focused client pilot and policy review." 78 397 650 64 17 $script:Color.Surface $true 1 $script:FontDisplay 3
    $null = Add-Text $slide "Digital Facial Attendance" 734 397 150 64 9 (Get-Rgb 205 233 223) $true 3 $script:FontBody 3
    Add-Footer $slide 12
    $null = Add-TextCard $slide "A practical path to client adoption" 54 68 670 52 $script:Color.Surface $script:Color.Text 24 $true 5 18

    if (Test-Path $pptxPath) { Remove-Item -LiteralPath $pptxPath -Force }
    if (Test-Path $pdfPath) { Remove-Item -LiteralPath $pdfPath -Force }
    if (Test-Path $previewPath) { Remove-Item -LiteralPath $previewPath -Recurse -Force }

    $presentation.SaveAs($pptxPath, 24)
    $presentation.SaveAs($pdfPath, 32)
    if ($ExportPreview) {
        New-Item -ItemType Directory -Force $previewPath | Out-Null
        $presentation.Export($previewPath, "PNG", 1600, 900)
    }

    Write-Output "Created: $pptxPath"
    Write-Output "Created: $pdfPath"
    Write-Output "Slides: $($presentation.Slides.Count)"
}
finally {
    if ($presentation) {
        $presentation.Close()
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($presentation)
    }
    if ($powerPoint) {
        $powerPoint.Quit()
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($powerPoint)
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
