# codex-PPTtoVideo-skills

[English](README.md) | [简体中文](README.zh-CN.md)

**Turn an editable PowerPoint into an automatically playing captioned video with Codex.**

Codex prepares a subtitle timeline, adds native PowerPoint text boxes, applies Fade entrance and exit animations, sets automatic slide advance, and exports an MP4 through desktop Microsoft PowerPoint. The captions remain editable in the PPTX.

![Animated subtitle playback](media/subtitle-playback.gif)

*This GIF samples frames from the actual PowerPoint-exported demonstration video. It is a shortened visual preview, not a timing reference.*

## Final result

![PowerPoint subtitle result](media/result.png)

- [Captioned demo PPTX](examples/subtitle-demo_captioned.pptx)
- [Actual exported MP4](examples/subtitle-demo_review.mp4)
- [Static PDF preview](examples/subtitle-demo_preview.pdf)
- [Caption timeline CSV](examples/subtitle_timeline_FINAL.csv)
- [Demo verification notes](examples/VERIFICATION.md)

The public example contains three generic slides, nine captions, and 24 seconds of playback. It does not include research data or laboratory photographs.

## Two synchronized animated diagrams

The upper half now contains two changing diagrams: an editable slide builds up on the left, while a video player and cue timeline advance on the right. Both change with the bottom caption. The demo groups each pair of diagram states with its editable caption and applies native Fade entrance/exit effects to that group. This explicit scene choreography is separate from the general caption-only helper. No screen recording or pre-rendered GIF is embedded in the PPTX.

## How it works

```mermaid
flowchart LR
    A[Editable PPTX] --> B[Codex writes the caption timeline]
    B --> C[Native PowerPoint text boxes]
    C --> D[Fade in, hold, fade out]
    D --> E[Automatic slide advance]
    E --> F[PowerPoint exports MP4]
    F --> G[Inspect the actual video]
```

### 1. Prepare the slides and captions

Codex reads the presentation, keeps the existing editable content, and reserves a subtitle band below the main content. If there is no narration, the captions are written from the slide content. If narration exists, use the actual speech timings.

Each caption receives a unique shape name, slide index, local start/end time, and fade durations. The timeline also gives each slide a playback duration. Printed slide numbers can differ from their positions in the file.

### 2. Animate captions inside PowerPoint

Each subtitle is a separate editable text box. Its animation sequence is:

**Fade in → hold → fade out → next caption**

The helper uses PowerPoint's native `TimeLine.MainSequence.AddEffect` API and the Fade effect. It does not burn text into slide images or substitute an SRT file for PowerPoint animation.

For a caption starting at `s`, ending at `e`, with entrance duration `fi`, exit duration `fo`, and previous caption end `p`:

| Effect | Start mode | Delay | Duration |
|---|---|---|---|
| Entrance | After Previous | `s - p` | `fi` |
| Exit | After Previous | `e - s - fi - fo` | `fo` |

The first caption uses `p = 0`. With `fi = fo = 0.25 s`, a 4-second cue holds for 3.5 seconds. This makes the exit finish at the scheduled cue end. The helper also supports intentional gaps between cues.

### 3. Export with stored slide timings

The helper enables timed slide advance and disables click-to-advance. PowerPoint's `CreateVideo` exports the presentation using its stored timings and animations. No clicks or voice-over are required for a silent demonstration.

### 4. Verify playback

Check the duration of the real MP4 and inspect a decoded frame inside every caption interval. Also inspect selected fades and slide boundaries. Look for text that fails to disappear, overlapping captions, timing drift, and captions covering charts or video output.

An export-success status or a valid timeline CSV alone does not prove correct playback. The optional extraction helper reports media duration and produces frames for inspection; it does not automatically certify the result.

## Requirements

- Windows with desktop Microsoft PowerPoint installed.
- PowerShell 7 (`pwsh`) for the bundled helpers.
- An active desktop login session for PowerPoint COM automation.
- An execution environment authorized to control PowerPoint and write the requested outputs.
- For Codex use: a local editable PPTX and caption text or permission to draft captions from the slides.

## Install the Codex Skill

Clone this repository and copy the `powerpoint-auto-subtitles` folder into your Codex skills directory. The usual Windows location is `%USERPROFILE%\.codex\skills`; use your configured `CODEX_HOME` if it differs.

```powershell
git clone https://github.com/machengjie321/codex-PPTtoVideo-skills.git
Copy-Item -LiteralPath ./codex-PPTtoVideo-skills/powerpoint-auto-subtitles `
  -Destination "$env:USERPROFILE/.codex/skills/powerpoint-auto-subtitles" -Recurse
```

Refresh the available skills or start a new Codex chat if the new skill is not listed. Avoid overwriting an existing installation with local edits.

## Ask Codex to use it

```text
Use $powerpoint-auto-subtitles to add editable English captions to this
PowerPoint and export an autoplay MP4. Target 160 seconds. Build a timeline,
reserve a subtitle band, apply native Fade entrance/exit effects, and verify
the actual video. Preserve all measured values and editable diagrams.
Deliver the PPTX, MP4, PDF preview, final timing CSV, and a short QA report.
```

The example target of 160 seconds is configurable. The included public demo is 24 seconds.

## Run the helper directly

Prepare a timeline following [the CSV contract](powerpoint-auto-subtitles/references/timeline.md), then run:

```powershell
pwsh -NoProfile -File ./powerpoint-auto-subtitles/scripts/add-subtitles-and-export.ps1 `
  -InputPptx ./my-presentation.pptx `
  -TimelineCsv ./timeline.csv `
  -OutputDirectory ./captioned-output
```

Outputs include an animated PPTX, MP4, PDF, per-slide PNGs, final timeline CSV, and an automation report. The report deliberately leaves actual playback verification pending until the video has been inspected. The source deck is not overwritten.

The helper refuses existing animation sequences or caption objects. For an already animated deck, ask Codex to preserve and merge its timeline deliberately. Reserve the caption area before running the helper; it does not automatically rearrange existing charts and tables.

To extract verification frames:

```powershell
pwsh -NoProfile -Sta -File ./powerpoint-auto-subtitles/scripts/extract-video-frames.ps1 `
  -VideoPath ./captioned-output/my-presentation_review.mp4 `
  -TimelineCsv ./captioned-output/subtitle_timeline_FINAL.csv `
  -OutputDirectory ./video-check
```

## Further explanation

- [English walkthrough and short presentation script](powerpoint-auto-subtitles/references/workflow-walkthrough.md)
- [Timeline fields and timing formulas](powerpoint-auto-subtitles/references/timeline.md)
- [Office implementation, export and QA notes](powerpoint-auto-subtitles/references/office-workflow.md)
- [Skill instructions](powerpoint-auto-subtitles/SKILL.md)

Captions overlap in normal editing view because they share the same safe band. Their native animations control visibility during playback. The static PDF/PNGs show only the first caption on each slide.
