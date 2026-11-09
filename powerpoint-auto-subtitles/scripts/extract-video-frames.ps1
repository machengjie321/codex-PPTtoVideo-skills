[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$VideoPath,
    [Parameter(Mandatory)][string]$TimelineCsv,
    [Parameter(Mandatory)][string]$OutputDirectory
)
$ErrorActionPreference = 'Stop'
if ([Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') { throw 'Run this helper with pwsh -Sta.' }
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
if (-not ('CaptionFramePump' -as [type])) {
    Add-Type -ReferencedAssemblies ([Windows.Threading.Dispatcher].Assembly.Location) -TypeDefinition @'
using System;
using System.Windows.Threading;
public class CaptionFramePump {
    public static void Wait(int ms) {
        var frame = new DispatcherFrame();
        var timer = new DispatcherTimer(DispatcherPriority.Background);
        timer.Interval = TimeSpan.FromMilliseconds(ms);
        timer.Tick += (sender,args) => { timer.Stop(); frame.Continue = false; };
        timer.Start(); Dispatcher.PushFrame(frame);
    }
}
'@
}
$movie = (Resolve-Path -LiteralPath $VideoPath).Path
$rows = @(Import-Csv -LiteralPath $TimelineCsv -Encoding utf8)
$out = [IO.Path]::GetFullPath($OutputDirectory)
if (Test-Path -LiteralPath $out) { throw 'Choose a new frame output directory.' }
New-Item -ItemType Directory -Path $out | Out-Null
$culture = [Globalization.CultureInfo]::InvariantCulture
$mp = New-Object Windows.Media.MediaPlayer
try {
    $mp.Volume = 0; $mp.ScrubbingEnabled = $true
    $mp.Open([Uri]$movie); $mp.Play(); [CaptionFramePump]::Wait(1000); $mp.Pause()
    if (-not $mp.NaturalDuration.HasTimeSpan) { throw 'The media decoder could not read the video duration.' }
    $duration = $mp.NaturalDuration.TimeSpan.TotalSeconds
    $samples = [Collections.Generic.List[object]]::new()
    foreach ($cue in $rows) {
        $start = [double]::Parse($cue.global_start_s,$culture)
        $end = [double]::Parse($cue.global_end_s,$culture)
        $time = ($start+$end)/2
        if ($time -lt 0 -or $time -gt $duration) { throw "Cue is outside the decoded video: $($cue.caption_id)" }
        $mp.Position = [TimeSpan]::FromSeconds($time)
        [CaptionFramePump]::Wait(500)
        $visual = New-Object Windows.Media.DrawingVisual
        $dc = $visual.RenderOpen()
        $dc.DrawVideo($mp,(New-Object Windows.Rect(0,0,1280,720))); $dc.Close()
        $bitmap = New-Object Windows.Media.Imaging.RenderTargetBitmap(1280,720,96,96,[Windows.Media.PixelFormats]::Pbgra32)
        $bitmap.Render($visual)
        $encoder = New-Object Windows.Media.Imaging.PngBitmapEncoder
        $encoder.Frames.Add([Windows.Media.Imaging.BitmapFrame]::Create($bitmap))
        $name = $cue.caption_id+'.png'
        $stream = [IO.File]::Create((Join-Path $out $name))
        try { $encoder.Save($stream) } finally { $stream.Close() }
        $samples.Add([pscustomobject]@{caption_id=$cue.caption_id; sample_s=$time; expected_text=$cue.text; image=$name})
    }
    [ordered]@{actual_video_duration_s=$duration; samples=$samples; visual_review_completed=$false} |
        ConvertTo-Json -Depth 6 | Set-Content -LiteralPath (Join-Path $out 'frame_manifest.json') -Encoding utf8
    Write-Output "Decoded duration: $duration seconds. Inspect the extracted frames; extraction alone is not visual validation."
}
finally { $mp.Close() }
