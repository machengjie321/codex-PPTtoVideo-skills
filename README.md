# codex-PPTtoVideo-skills

[English](README.md) | [简体中文](README.zh-CN.md)

## Step 1: Prepare the caption file first

Organize the captions before generating animation timings or video. You do not need to supply timestamps or a CSV yet.

1. Write the captions for each slide in the actual PPTX order.
2. Keep one idea per cue and list cues in playback order. The number of cues can vary by slide.
3. Put all slides in one **UTF-8 `captions.md` or `captions.txt`** using the format below.
4. For an English video, supply final English captions, or explicitly ask Codex to translate your draft before formatting it.

[Blank caption template](examples/captions.template.md) · [Completed English example](examples/captions.en.md) · [Chinese format example](examples/captions.zh-CN.md)

```text
# Captions

## Slide 1
- This slide introduces the research platform.
- The diagram shows its main components.

## Slide 2
- This slide compares the two processing paths.
- The following animation explains the data flow.
```

**Format rules:**

- `## Slide 1` refers to the first slide in the file; number consecutively from 1, independently of printed page labels.
- Every `- ` line is one complete cue. Do not insert a physical line break inside it.
- Put only visible caption text in the document. Send visual directions, comments, reading speed and target duration separately in the task message.
- Include at least one cue per slide. Add or remove template sections to match the deck and replace every bracketed placeholder.
- Leave timestamps out. Generate estimates after formatting, then review readability and total duration.

**If your draft is in Word, TXT or loosely organized notes, start with this request:**

```text
Use $powerpoint-auto-subtitles. Prepare the caption file only; do not export video yet.
PPTX: C:/work/my-presentation.pptx
Caption draft: C:/work/subtitle-draft.docx
Output: C:/work/captions.md (UTF-8)
Match the actual slide order. Use ## Slide N for each page and one - bullet per English cue.
Preserve meaning, values and units. Split long paragraphs at natural idea boundaries.
Translate a Chinese draft into English; do not invent content for missing pages.
Check slide correspondence and list missing or ambiguous mappings.
Deliver the normalized caption file and a short check report, without animations or MP4.
```

Replace the example paths with real files, or paste a draft labeled by slide. Resolve ambiguous page mappings first. Once the file is ready, use the export request later in this guide and explicitly name `captions.md` as the caption source.


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

## Write captions page by page

You can paste the captions into Codex or put all pages in one Markdown, TXT or Word document. Codex can also draft captions from your PPT: introduce the visual, explain the key point, and finish with the takeaway. Keep each cue focused on one idea and preserve the slide's facts and units.

```text
## Slide 1
- Introduce the main idea on this slide.
- Explain the next diagram state.

## Slide 2
- State the key comparison.
- Finish with the takeaway.
```

Each bullet is one subtitle. `Slide N` means its position in the PPTX. Word files are first converted to this UTF-8 Markdown/TXT format by Codex; the script does not directly parse Word files.

- [English caption template](examples/captions.en.md)
- [Chinese caption template](examples/captions.zh-CN.md)
- [Full writing, input-format and timing guide](powerpoint-auto-subtitles/references/caption-authoring.md)

### Automatic duration from text length

Defaults: English **2.5 words/s**, Chinese **4 characters/s**, plus **0.4 s** reading pause and **0.25 s** for each fade. A cue lasts at least **2.5 s**. Mixed-language text adds the two reading times. These are adjustable estimates; narration uses actual speech timing.

```text
cue duration = max(2.5, English words / 2.5 + Chinese characters / 4 + 0.4 + 0.5)
```

Longer text gets a longer hold time. The gap is 0.15 s between cues and the final slide hold is 0.5 s. No long caption is silently truncated; split dense captions for readability.

```powershell
python ./powerpoint-auto-subtitles/scripts/captions-to-timeline.py `
  ./examples/captions.en.md ./timeline-estimated.csv
```

Then pass `timeline-estimated.csv` to the PowerPoint helper below. Python 3 is required for this optional conversion step. Review the total duration and align any diagram animations to the generated timeline.

Example Codex request:

```text
Use $powerpoint-auto-subtitles.
PPTX: C:/work/my-presentation.pptx
Caption source: no existing script. Draft English captions only from this deck's
visible text, figures and speaker notes. Use one idea per cue; invent no results.
Narration: none. Output directory: C:/work/captioned-output
Save captions.md with ## Slide N headings and one bullet per caption.
Timing: estimate at 2.5 English words/s, with no fixed total duration; report the estimate first.
Generate the CSV and synchronize requested diagram changes with the final cues.
Export editable PPTX, MP4, PDF, timing CSV and QA notes; inspect actual playback.
```

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

Replace the example paths with actual existing files and a new output folder. This example uses a supplied caption document. For automatic drafting, use the earlier complete prompt. An attachment can replace a path if identified by its exact filename. If using audio, give its path and request alignment to that recording.

```text
Use $powerpoint-auto-subtitles.
PPTX: C:/work/my-presentation.pptx
Caption source: C:/work/captions.md (## Slide N headings; one bullet per cue)
Caption language: English. Preserve the supplied wording; flag any needed translation or rewriting.
Output directory: C:/work/captioned-output
Narration: none; create a silent captioned video.
Target total duration: 160 seconds.
Read the files, check slide-to-caption correspondence, and generate the final timeline first.
Estimate reading time from text length. If the captions cannot fit readably within
160 seconds, report the conflict rather than silently shortening text or speeding it up.
Apply native PowerPoint fades and automatic slide advance. Preserve all measured
values, charts and editable shapes. Do not overwrite the source.
Inspect the actual MP4 captions and duration. Deliver PPTX, MP4, PDF preview,
final timing CSV and a short QA report.
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
