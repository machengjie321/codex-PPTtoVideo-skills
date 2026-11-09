[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$InputPptx,
    [Parameter(Mandatory)][string]$TimelineCsv,
    [Parameter(Mandatory)][string]$OutputDirectory,
    [ValidateRange(12,40)][double]$FontSize = 20,
    [ValidateRange(0.7,0.95)][double]$BandTopRatio = 0.9028,
    [ValidateRange(0.04,0.2)][double]$BandHeightRatio = 0.075,
    [ValidateRange(360,2160)][int]$VideoHeight = 1080,
    [ValidateRange(1,60)][int]$FramesPerSecond = 15,
    [ValidateRange(10,7200)][int]$ExportTimeoutSeconds = 600,
    [switch]$SkipVideo
)

$ErrorActionPreference = 'Stop'
if ($env:OS -ne 'Windows_NT') { throw 'Desktop Microsoft PowerPoint on Windows is required.' }
if ($BandTopRatio + $BandHeightRatio -gt 1) { throw 'The caption band extends below the slide.' }
$inputPath = (Resolve-Path -LiteralPath $InputPptx).Path
$timelinePath = (Resolve-Path -LiteralPath $TimelineCsv).Path
$out = [IO.Path]::GetFullPath($OutputDirectory)
$rows = @(Import-Csv -LiteralPath $timelinePath -Encoding utf8)
if ($rows.Count -eq 0) { throw 'The timeline is empty.' }
$required = @('slide_index','slide_number','slide_duration_s','caption_id','local_start_s','local_end_s','fade_in_s','fade_out_s','text')
foreach ($column in $required) {
    if ($column -notin $rows[0].PSObject.Properties.Name) { throw "Missing CSV column: $column" }
}
$culture = [Globalization.CultureInfo]::InvariantCulture
function Number([object]$value) {
    $parsed = [double]::Parse([string]$value, $culture)
    if ([double]::IsNaN($parsed) -or [double]::IsInfinity($parsed)) { throw 'Timeline numbers must be finite.' }
    return $parsed
}
$ids = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
foreach ($row in $rows) {
    $index = Number $row.slide_index
    if ($index -lt 1 -or $index -ne [Math]::Floor($index)) { throw 'slide_index must be a positive integer.' }
    if ($row.caption_id -notmatch '^CAPTION_[A-Za-z0-9_]+$' -or -not $ids.Add($row.caption_id)) {
        throw "Invalid or duplicated caption_id: $($row.caption_id)"
    }
    if ([string]::IsNullOrWhiteSpace($row.text)) { throw "Empty subtitle: $($row.caption_id)" }
    $start = Number $row.local_start_s; $end = Number $row.local_end_s
    $fi = Number $row.fade_in_s; $fo = Number $row.fade_out_s; $duration = Number $row.slide_duration_s
    if ($start -lt 0 -or $end -le $start -or $duration -le 0 -or $end -gt $duration -or
        $fi -le 0 -or $fo -le 0 -or $end-$start -lt $fi+$fo) {
        throw "Invalid cue bounds or fade duration: $($row.caption_id)"
    }
}
New-Item -ItemType Directory -Force -Path $out | Out-Null
$stem = [IO.Path]::GetFileNameWithoutExtension($inputPath)
$pptx = Join-Path $out ($stem + '_captioned.pptx')
$pdf = Join-Path $out ($stem + '_preview.pdf')
$mp4 = Join-Path $out ($stem + '_review.mp4')
$csv = Join-Path $out 'subtitle_timeline_FINAL.csv'
$reportPath = Join-Path $out 'automation_report.json'
$previewDir = Join-Path $out 'slides'
foreach ($path in @($pptx,$pdf,$mp4,$csv,$reportPath,$previewDir)) {
    if (Test-Path -LiteralPath $path) { throw "Output already exists; choose a fresh output directory: $path" }
}
$sourceHash = (Get-FileHash -LiteralPath $inputPath -Algorithm SHA256).Hash
$deck = $null; $saved = $false; $videoStatus = 'Not requested'
try {
    $app = New-Object -ComObject PowerPoint.Application
    # Normal open only. A repair request is a diagnostic issue, not a silent success.
    $deck = $app.Presentations.Open($inputPath, -1, 0, 0)
    $count = $deck.Slides.Count
    if (($rows | Where-Object { (Number $_.slide_index) -gt $count }).Count) { throw 'Timeline references a missing slide.' }
    $ordered = [Collections.Generic.List[object]]::new()
    $durations = [Collections.Generic.List[double]]::new()
    $globalOffset = 0.0
    # Check every slide before creating any caption objects.
    for ($i=1; $i -le $count; $i++) {
        $slide = $deck.Slides.Item($i)
        if ($slide.TimeLine.MainSequence.Count -gt 0 -or $slide.TimeLine.InteractiveSequences.Count -gt 0) {
            throw "Slide $i already has animations. Preserve and merge them deliberately; this helper will not clear them."
        }
        foreach ($shape in $slide.Shapes) {
            if ($shape.Name -like 'CAPTION_*') { throw "Slide $i already contains caption objects." }
        }
        $cues = @($rows | Where-Object { [int]$_.slide_index -eq $i } | Sort-Object { Number $_.local_start_s })
        if ($cues.Count -eq 0) { throw "No timeline cues for slide $i." }
        $duration = Number $cues[0].slide_duration_s
        $previousEnd = 0.0
        foreach ($cue in $cues) {
            if ((Number $cue.slide_duration_s) -ne $duration) { throw "Conflicting slide durations on slide $i." }
            $start = Number $cue.local_start_s; $end = Number $cue.local_end_s
            if ($start -lt $previousEnd - 0.000001) { throw "Overlapping cues on slide $i." }
            $ordered.Add([pscustomobject]@{
                slide_index=$i; slide_number=$cue.slide_number; slide_duration_s=$duration
                caption_id=$cue.caption_id; local_start_s=$start; local_end_s=$end
                global_start_s=$globalOffset+$start; global_end_s=$globalOffset+$end
                fade_in_s=(Number $cue.fade_in_s); fade_out_s=(Number $cue.fade_out_s); text=$cue.text
            })
            $previousEnd = $end
        }
        $durations.Add($duration); $globalOffset += $duration
    }
    $width = $deck.PageSetup.SlideWidth; $height = $deck.PageSetup.SlideHeight
    for ($i=1; $i -le $count; $i++) {
        $slide = $deck.Slides.Item($i); $previousEnd = 0.0
        foreach ($cue in @($ordered | Where-Object { $_.slide_index -eq $i })) {
            $shape = $slide.Shapes.AddTextbox(1, $width*0.03125, $height*$BandTopRatio, $width*0.9375, $height*$BandHeightRatio)
            $shape.Name = $cue.caption_id
            $shape.Line.Visible = 0
            $shape.Fill.Visible = -1; $shape.Fill.Solid()
            $shape.Fill.ForeColor.RGB = 0; $shape.Fill.Transparency = 0.4
            $shape.TextFrame.MarginLeft = 10; $shape.TextFrame.MarginRight = 10
            $shape.TextFrame.MarginTop = 2; $shape.TextFrame.MarginBottom = 2
            $shape.TextFrame.VerticalAnchor = 3
            $shape.TextFrame.WordWrap = -1
            $range = $shape.TextFrame.TextRange; $range.Text = $cue.text
            $range.Font.Name = 'Arial'; $range.Font.Size = $FontSize; $range.Font.Color.RGB = 16777215
            $range.ParagraphFormat.Alignment = 2

            # Demo-specific, fully editable diagram states, synchronized as one group.
            $names = [Collections.Generic.List[string]]::new()
            $names.Add($shape.Name)
            $state = [int]$cue.caption_id.Substring($cue.caption_id.Length-2)
            $sx=$width/1280; $sy=$height/720
            function DemoBox($x,$y,$w,$h,$rgb,$label,$font=18) {
                $q=$slide.Shapes.AddShape(1,$x*$sx,$y*$sy,$w*$sx,$h*$sy)
                $q.Fill.ForeColor.RGB=$rgb; $q.Line.Visible=0
                if($label) {
                    $q.TextFrame.TextRange.Text=$label
                    $q.TextFrame.TextRange.Font.Name='Arial'
                    $q.TextFrame.TextRange.Font.Size=$font*$sx
                    $q.TextFrame.TextRange.Font.Color.RGB=5067287
                    $q.TextFrame.VerticalAnchor=3
                    $q.TextFrame.TextRange.ParagraphFormat.Alignment=2
                }
                $names.Add($q.Name)
            }
            DemoBox 50 300 550 282 16446962 ''
            DemoBox 675 300 550 282 16446962 ''
            DemoBox 75 322 500 38 15460820 'Editable slide canvas' 20
            for($k=1;$k -le 3;$k++) {
                $label=if($k -le $state){@('Introduce the idea','Reveal the next step','Complete the story')[$k-1]}else{'...'}
                $color=if($k -le $state){@([int]16767166,[int]15134412,[int]15916004)[$k-1]}else{15987699}
                DemoBox 88 (378+($k-1)*55) 472 42 $color $label 21
            }
            DemoBox 700 322 500 143 5067287 ''
            for($k=1;$k -le $state;$k++) { DemoBox (724+($k-1)*155) 350 140 86 15134412 ('Scene '+$k) 20 }
            DemoBox 710 484 480 12 14540253 ''
            DemoBox 710 484 (160*$state) 12 9219856 ''
            for($k=1;$k -le 3;$k++) {
                $color=if($k -eq $state){15134412}else{15987699}
                DemoBox (710+($k-1)*163) 514 154 39 $color ('Cue '+$k) 20
            }
            $originalName=$shape.Name
            $shape=$slide.Shapes.Range([object[]]$names.ToArray()).Group()
            $shape.Name=$originalName

            # Effect times are relative to the completion of the preceding effect.
            $entrance = $slide.TimeLine.MainSequence.AddEffect($shape, 10, 0, 3)
            $entrance.Timing.Duration = $cue.fade_in_s
            $entrance.Timing.TriggerDelayTime = [Math]::Max(0,$cue.local_start_s-$previousEnd)
            $exit = $slide.TimeLine.MainSequence.AddEffect($shape, 10, 0, 3)
            $exit.Exit = -1
            $exit.Timing.Duration = $cue.fade_out_s
            $exit.Timing.TriggerDelayTime = [Math]::Max(0,$cue.local_end_s-$cue.local_start_s-$cue.fade_in_s-$cue.fade_out_s)
            $previousEnd = $cue.local_end_s
        }
        $slide.SlideShowTransition.AdvanceOnClick = 0
        $slide.SlideShowTransition.AdvanceOnTime = -1
        $slide.SlideShowTransition.AdvanceTime = $durations[$i-1]
    }
    $deck.SlideShowSettings.AdvanceMode = 2
    $deck.SlideShowSettings.ShowWithAnimation = -1
    $deck.SaveAs($pptx,24); $saved = $true
    $ordered | Export-Csv -LiteralPath $csv -NoTypeInformation -Encoding utf8
    New-Item -ItemType Directory -Path $previewDir | Out-Null
    for ($i=1; $i -le $count; $i++) {
        $cues = @($ordered | Where-Object { $_.slide_index -eq $i })
        for ($j=1; $j -lt $cues.Count; $j++) { $deck.Slides.Item($i).Shapes.Item($cues[$j].caption_id).Visible = 0 }
    }
    # SaveCopyAs PDF avoids PowerShell's optional-argument binding issue with ExportAsFixedFormat.
    $deck.SaveCopyAs($pdf,32)
    for ($i=1; $i -le $count; $i++) {
        $deck.Slides.Item($i).Export((Join-Path $previewDir ('Slide-{0:D2}.png' -f $i)), 'PNG', 1920, [int](1920*$height/$width))
    }
    foreach ($cue in $ordered) { $deck.Slides.Item($cue.slide_index).Shapes.Item($cue.caption_id).Visible = -1 }
    if (-not $SkipVideo) {
        $deck.CreateVideo($mp4,-1,5,$VideoHeight,$FramesPerSecond,75)
        $clock = [Diagnostics.Stopwatch]::StartNew()
        do {
            Start-Sleep -Milliseconds 500
            $status = [int]$deck.CreateVideoStatus
            if ($clock.Elapsed.TotalSeconds -gt $ExportTimeoutSeconds) { throw 'Video export timed out. Do not treat an incomplete MP4 as a verified output.' }
        } while ($status -eq 0 -or $status -eq 1 -or $status -eq 2)
        if ($status -ne 3) { throw "PowerPoint video export failed, status $status." }
        $videoStatus = 'Office export completed; decoded playback inspection still required'
    }
    $report = [ordered]@{
        source_pptx=$inputPath; source_sha256=$sourceHash; output_pptx=$pptx
        slides=$count; captions=$ordered.Count; native_fade_effects=$ordered.Count*2
        scheduled_duration_s=$globalOffset; native_office_export=$videoStatus
        actual_video_duration_s=$null; decoded_frames_inspected=$false
        note='Inspect actual video duration and decoded frames before claiming verified playback.'
    }
    $report | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $reportPath -Encoding utf8
    Write-Output ($report | ConvertTo-Json -Compress)
}
finally {
    if ($null -ne $deck) {
        # Discard temporary static-preview visibility changes. Never quit the user's application.
        $deck.Saved = -1
        $deck.Close()
    }
    if ((Get-FileHash -LiteralPath $inputPath -Algorithm SHA256).Hash -ne $sourceHash) { throw 'Source integrity check failed.' }
}
