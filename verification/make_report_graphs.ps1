Add-Type -AssemblyName System.Drawing

$Root = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
$Assets = Join-Path $Root 'report_assets'

$Policies = @('FIFO', 'NUR (0,1)', 'NUR (1,0)', 'Second Chance', 'WSClock-Lite')
$Colors = @(
    [System.Drawing.Color]::FromArgb(77, 119, 204),
    [System.Drawing.Color]::FromArgb(91, 166, 89),
    [System.Drawing.Color]::FromArgb(240, 150, 62),
    [System.Drawing.Color]::FromArgb(168, 92, 198),
    [System.Drawing.Color]::FromArgb(211, 84, 78)
)

$Workloads = [ordered]@{
    Textbook = @{
        Reference = '123412512345'
        Frame4 = @(10, 6, 6, 7, 8)
        Frames = @(
            @(9, 10, 5),
            @(9, 6, 5),
            @(9, 6, 5),
            @(10, 7, 5),
            @(10, 8, 5)
        )
    }
    NURPriority = @{
        Reference = 'ABCDABEFABGHABCD'
        Frame4 = @(12, 10, 11, 10, 11)
        Frames = @(
            @(16, 12, 12),
            @(11, 10, 9),
            @(12, 11, 9),
            @(16, 10, 10),
            @(16, 11, 10)
        )
    }
    ClockSensitive = @{
        Reference = 'ABCDEABCDA'
        Frame4 = @(9, 8, 7, 9, 9)
        Frames = @(
            @(10, 9, 5),
            @(8, 8, 5),
            @(7, 7, 5),
            @(10, 9, 5),
            @(10, 9, 5)
        )
    }
    FrequencyBias = @{
        Reference = 'AAAABCAAADEFAAA'
        Frame4 = @(7, 6, 6, 6, 6)
        Frames = @(
            @(7, 7, 7),
            @(6, 6, 6),
            @(6, 6, 6),
            @(7, 6, 6),
            @(7, 6, 6)
        )
    }
    LocalityShift = @{
        Reference = 'ABCABCABCDEFDEFABC'
        Frame4 = @(9, 9, 10, 9, 9)
        Frames = @(
            @(9, 9, 9),
            @(11, 9, 7),
            @(10, 10, 7),
            @(9, 9, 9),
            @(9, 9, 8)
        )
    }
}

function New-Bitmap($Width, $Height) {
    $bmp = New-Object System.Drawing.Bitmap $Width, $Height
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $g.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::ClearTypeGridFit
    $g.Clear([System.Drawing.Color]::White)
    return @($bmp, $g)
}

function Draw-Text($g, $Text, $Font, $Brush, $X, $Y, $Width, $Height, $Align = 'Center') {
    $sf = New-Object System.Drawing.StringFormat
    $sf.Alignment = if ($Align -eq 'Left') { [System.Drawing.StringAlignment]::Near } else { [System.Drawing.StringAlignment]::Center }
    $sf.LineAlignment = [System.Drawing.StringAlignment]::Center
    $rect = New-Object System.Drawing.RectangleF $X, $Y, $Width, $Height
    $g.DrawString($Text, $Font, $Brush, $rect, $sf)
    $sf.Dispose()
}

function Save-BarChart($Name, $Title, $Reference, $Values) {
    $pair = New-Bitmap 1100 560
    $bmp = $pair[0]
    $g = $pair[1]
    $black = [System.Drawing.Brushes]::Black
    $grayPen = New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(210, 210, 210)), 2
    $axisPen = New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(80, 80, 80)), 3
    $titleFont = New-Object System.Drawing.Font 'Malgun Gothic', 24, ([System.Drawing.FontStyle]::Bold)
    $subFont = New-Object System.Drawing.Font 'Malgun Gothic', 15
    $axisFont = New-Object System.Drawing.Font 'Malgun Gothic', 15
    $labelFont = New-Object System.Drawing.Font 'Malgun Gothic', 14, ([System.Drawing.FontStyle]::Bold)

    Draw-Text $g $Title $titleFont $black 0 16 1100 40
    Draw-Text $g "Reference: $Reference / Frame=4" $subFont ([System.Drawing.Brushes]::DimGray) 0 58 1100 30

    $left = 100; $top = 105; $width = 920; $height = 330
    $max = [Math]::Max(1, (($Values | Measure-Object -Maximum).Maximum + 2))
    for ($i = 0; $i -le $max; $i += 2) {
        $y = $top + $height - ($i / $max) * $height
        $g.DrawLine($grayPen, $left, [float]$y, $left + $width, [float]$y)
        Draw-Text $g $i $axisFont $black 35 ($y - 14) 45 28
    }
    $g.DrawLine($axisPen, $left, $top, $left, $top + $height)
    $g.DrawLine($axisPen, $left, $top + $height, $left + $width, $top + $height)

    $barW = 112; $gap = 58
    for ($i = 0; $i -lt $Values.Count; $i++) {
        $x = $left + 35 + $i * ($barW + $gap)
        $h = ($Values[$i] / $max) * $height
        $y = $top + $height - $h
        $brush = New-Object System.Drawing.SolidBrush $Colors[$i]
        $g.FillRectangle($brush, [float]$x, [float]$y, $barW, [float]$h)
        $g.DrawRectangle([System.Drawing.Pens]::White, [int]$x, [int]$y, $barW, [int]$h)
        Draw-Text $g $Values[$i] $labelFont $black ($x - 8) ($y - 34) ($barW + 16) 30
        $label = $Policies[$i].Replace('Second Chance', "Second`nChance").Replace('WSClock-Lite', "WSClock`nLite")
        Draw-Text $g $label $axisFont $black ($x - 20) ($top + $height + 12) ($barW + 40) 58
        $brush.Dispose()
    }
    Draw-Text $g 'Page Fault Count' $axisFont $black 0 250 55 40

    $out = Join-Path $Assets $Name
    $bmp.Save($out, [System.Drawing.Imaging.ImageFormat]::Png)
    $titleFont.Dispose(); $subFont.Dispose(); $axisFont.Dispose(); $labelFont.Dispose()
    $grayPen.Dispose(); $axisPen.Dispose(); $g.Dispose(); $bmp.Dispose()
}

function Save-LineTrends() {
    $pair = New-Bitmap 1500 1020
    $bmp = $pair[0]
    $g = $pair[1]
    $titleFont = New-Object System.Drawing.Font 'Malgun Gothic', 26, ([System.Drawing.FontStyle]::Bold)
    $axisFont = New-Object System.Drawing.Font 'Malgun Gothic', 14
    $smallFont = New-Object System.Drawing.Font 'Malgun Gothic', 13
    Draw-Text $g 'Page Fault Trend by Frame Size' $titleFont ([System.Drawing.Brushes]::Black) 0 15 1500 50

    $keys = @($Workloads.Keys)
    for ($w = 0; $w -lt $keys.Count; $w++) {
        [int]$x0 = 65
        [int]$y0 = 95
        if ($w -eq 1) { $x0 = 790; $y0 = 95 }
        if ($w -eq 2) { $x0 = 65; $y0 = 410 }
        if ($w -eq 3) { $x0 = 790; $y0 = 410 }
        if ($w -eq 4) { $x0 = 430; $y0 = 725 }
        $plotW = 630; $plotH = 220
        Draw-Text $g $keys[$w] $axisFont ([System.Drawing.Brushes]::Black) $x0 ($y0 - 35) $plotW 30
        $max = 18
        $penGrid = New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(220,220,220)), 1
        $penAxis = New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(80,80,80)), 2
        for ($tick=0; $tick -le $max; $tick += 6) {
            $y = $y0 + $plotH - ($tick / $max) * $plotH
            $g.DrawLine($penGrid, $x0, [float]$y, $x0 + $plotW, [float]$y)
            Draw-Text $g $tick $smallFont ([System.Drawing.Brushes]::Black) ($x0 - 42) ($y - 12) 35 24
        }
        $g.DrawLine($penAxis, $x0, $y0, $x0, $y0 + $plotH)
        $g.DrawLine($penAxis, $x0, $y0 + $plotH, $x0 + $plotW, $y0 + $plotH)
        $xs = @(($x0 + 80), ($x0 + 315), ($x0 + 550))
        for ($i=0; $i -lt 3; $i++) { Draw-Text $g (3+$i) $smallFont ([System.Drawing.Brushes]::Black) ($xs[$i]-20) ($y0+$plotH+8) 40 24 }
        for ($p=0; $p -lt $Policies.Count; $p++) {
            $vals = $Workloads[$keys[$w]].Frames[$p]
            $pen = New-Object System.Drawing.Pen $Colors[$p], 4
            $points = @()
            for ($i=0; $i -lt 3; $i++) {
                $py = $y0 + $plotH - ($vals[$i] / $max) * $plotH
                $points += ,(New-Object System.Drawing.PointF ([float]$xs[$i]), ([float]$py))
            }
            $g.DrawLines($pen, $points)
            foreach ($pt in $points) { $g.FillEllipse((New-Object System.Drawing.SolidBrush $Colors[$p]), $pt.X-5, $pt.Y-5, 10, 10) }
            $pen.Dispose()
        }
        $penGrid.Dispose(); $penAxis.Dispose()
    }

    for ($p=0; $p -lt $Policies.Count; $p++) {
        $brush = New-Object System.Drawing.SolidBrush $Colors[$p]
        $lx = 160 + $p * 250
        $g.FillRectangle($brush, $lx, 970, 22, 14)
        Draw-Text $g $Policies[$p] $smallFont ([System.Drawing.Brushes]::Black) ($lx+28) 960 205 34 'Left'
        $brush.Dispose()
    }
    $bmp.Save((Join-Path $Assets 'frame_fault_trends.png'), [System.Drawing.Imaging.ImageFormat]::Png)
    $titleFont.Dispose(); $axisFont.Dispose(); $smallFont.Dispose(); $g.Dispose(); $bmp.Dispose()
}

function Save-ParamChart($Name, $Title, $Labels, $SeriesNames, $SeriesValues) {
    $pair = New-Bitmap 1000 520
    $bmp = $pair[0]
    $g = $pair[1]
    $titleFont = New-Object System.Drawing.Font 'Malgun Gothic', 23, ([System.Drawing.FontStyle]::Bold)
    $axisFont = New-Object System.Drawing.Font 'Malgun Gothic', 15
    $labelFont = New-Object System.Drawing.Font 'Malgun Gothic', 14, ([System.Drawing.FontStyle]::Bold)
    Draw-Text $g $Title $titleFont ([System.Drawing.Brushes]::Black) 0 18 1000 42
    $left=95; $top=90; $width=820; $height=300; $max=12
    $grid = New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(220,220,220)), 1
    $axis = New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(80,80,80)), 3
    for($tick=0; $tick -le $max; $tick+=2){
        $y=$top+$height-($tick/$max)*$height
        $g.DrawLine($grid,$left,[float]$y,$left+$width,[float]$y)
        Draw-Text $g $tick $axisFont ([System.Drawing.Brushes]::Black) 35 ($y-13) 40 26
    }
    $g.DrawLine($axis,$left,$top,$left,$top+$height)
    $g.DrawLine($axis,$left,$top+$height,$left+$width,$top+$height)
    $groupW=$width/$Labels.Count
    for($i=0; $i -lt $Labels.Count; $i++){
        Draw-Text $g $Labels[$i] $axisFont ([System.Drawing.Brushes]::Black) ($left+$i*$groupW) ($top+$height+12) $groupW 30
        for($s=0; $s -lt $SeriesNames.Count; $s++){
            $barW=[Math]::Min(58, ($groupW-26)/$SeriesNames.Count)
            $x=$left+$i*$groupW+18+$s*($barW+8)
            $val=$SeriesValues[$s][$i]
            $h=($val/$max)*$height
            $y=$top+$height-$h
            $brush=New-Object System.Drawing.SolidBrush $Colors[$s+1]
            $g.FillRectangle($brush,[float]$x,[float]$y,[float]$barW,[float]$h)
            Draw-Text $g $val $labelFont ([System.Drawing.Brushes]::Black) ($x-4) ($y-28) ($barW+8) 24
            $brush.Dispose()
        }
    }
    for($s=0; $s -lt $SeriesNames.Count; $s++){
        $brush=New-Object System.Drawing.SolidBrush $Colors[$s+1]
        $lx=250+$s*250
        $g.FillRectangle($brush,$lx,470,20,13)
        Draw-Text $g $SeriesNames[$s] $axisFont ([System.Drawing.Brushes]::Black) ($lx+28) 460 220 30 'Left'
        $brush.Dispose()
    }
    $bmp.Save((Join-Path $Assets $Name), [System.Drawing.Imaging.ImageFormat]::Png)
    $titleFont.Dispose(); $axisFont.Dispose(); $labelFont.Dispose(); $grid.Dispose(); $axis.Dispose(); $g.Dispose(); $bmp.Dispose()
}

function Save-WsClockDiagram() {
    $pair = New-Bitmap 760 235
    $bmp = $pair[0]
    $g = $pair[1]
    $titleFont = New-Object System.Drawing.Font 'Malgun Gothic', 20, ([System.Drawing.FontStyle]::Bold)
    $font = New-Object System.Drawing.Font 'Malgun Gothic', 13
    Draw-Text $g 'WSClock-Lite Operation Summary' $titleFont ([System.Drawing.Brushes]::Black) 0 12 760 34
    $items = @(
        '1. Inspect the page at the clock hand',
        '2. If R=1, clear R and refresh last-use',
        '3. If R=0, age>=threshold, M=0: select victim',
        '4. If old dirty page appears, clear M and continue'
    )
    for($i=0; $i -lt $items.Count; $i++){
        $y=58+$i*39
        $rect=New-Object System.Drawing.Rectangle 45,$y,670,30
        $brush=New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(241,246,252))
        $g.FillRectangle($brush,$rect)
        $g.DrawRectangle((New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(90,130,180),2)),$rect)
        Draw-Text $g $items[$i] $font ([System.Drawing.Brushes]::Black) 60 ($y+1) 640 28 'Left'
        $brush.Dispose()
    }
    $bmp.Save((Join-Path $Assets 'exe_wsclock_locality_crop.png'), [System.Drawing.Imaging.ImageFormat]::Png)
    $titleFont.Dispose(); $font.Dispose(); $g.Dispose(); $bmp.Dispose()
}

function Save-SecondChanceDiagram() {
    $pair = New-Bitmap 760 235
    $bmp = $pair[0]
    $g = $pair[1]
    $titleFont = New-Object System.Drawing.Font 'Malgun Gothic', 20, ([System.Drawing.FontStyle]::Bold)
    $font = New-Object System.Drawing.Font 'Malgun Gothic', 13
    Draw-Text $g 'Second Chance Clock Summary' $titleFont ([System.Drawing.Brushes]::Black) 0 12 760 34
    $items = @(
        '1. Clock hand always starts at frame 1',
        '2. Clock input resets R bits every N steps',
        '3. If R=1, clear R and move to next frame',
        '4. If R=0, select that page as victim'
    )
    for($i=0; $i -lt $items.Count; $i++){
        $y=58+$i*39
        $rect=New-Object System.Drawing.Rectangle 45,$y,670,30
        $brush=New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(246,242,250))
        $g.FillRectangle($brush,$rect)
        $g.DrawRectangle((New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(125,85,160),2)),$rect)
        Draw-Text $g $items[$i] $font ([System.Drawing.Brushes]::Black) 60 ($y+1) 640 28 'Left'
        $brush.Dispose()
    }
    $bmp.Save((Join-Path $Assets 'exe_second_clock_crop.png'), [System.Drawing.Imaging.ImageFormat]::Png)
    $titleFont.Dispose(); $font.Dispose(); $g.Dispose(); $bmp.Dispose()
}

foreach ($key in $Workloads.Keys) {
    Save-BarChart "aligned_$key.png" "$key workload: Page Fault by Policy" $Workloads[$key].Reference $Workloads[$key].Frame4
}
Save-LineTrends
Save-ParamChart 'aligned_clock_param.png' 'Second Chance: Clock Interval' @('2','3','4','6') @('Second Chance') @(@(9,9,9,9))
Save-ParamChart 'aligned_reset_param.png' 'NUR: Clock Interval' @('2','3','4','6') @('NUR (0,1)', 'NUR (1,0)') @(@(11,11,10,11), @(11,11,11,11))
Save-ParamChart 'aligned_modified_param.png' 'NUR: Modified Pages Sensitivity' @('A','AD','BDF','none') @('NUR (0,1)', 'NUR (1,0)') @(@(10,10,10,10), @(10,11,11,10))
Save-SecondChanceDiagram
Save-WsClockDiagram
