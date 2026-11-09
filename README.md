# codex-PPTtoVideo-skills

[English](README.md) | [简体中文](README.zh-CN.md)

**Turn an editable PowerPoint into an automatically playing captioned video with Codex.**

Codex prepares a subtitle timeline, adds native PowerPoint text boxes, applies Fade entrance and exit animations, sets automatic slide advance, and exports an MP4 through desktop Microsoft PowerPoint. The captions remain editable in the PPTX.

## Demo and downloads

![Animated subtitle playback](media/subtitle-playback.gif)

*Above: an eight-second sampled preview of the exported video. Download the full 24-second video and source files below.*

- [Captioned demo PPTX](examples/subtitle-demo_captioned.pptx)
- [Actual exported MP4](examples/subtitle-demo_review.mp4)
- [Static PDF preview](examples/subtitle-demo_preview.pdf)
- [Caption timeline CSV](examples/subtitle_timeline_FINAL.csv)
- [Demo verification notes](examples/VERIFICATION.md)

The public example contains three generic slides, nine captions, and 24 seconds of playback. It does not include research data or laboratory photographs.

## Two synchronized animated diagrams

The upper half now contains two changing diagrams: an editable slide builds up on the left, while a video player and cue timeline advance on the right. Both change with the bottom caption. The demo groups each pair of diagram states with its editable caption and applies native Fade entrance/exit effects to that group. This explicit scene choreography is separate from the general caption-only helper. No screen recording or pre-rendered GIF is embedded in the PPTX.

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

## Complete workflow

**Prepare captions (choose: Agent drafts from PPTX / user supplies captions and Agent normalizes the format) → user reviews, edits and confirms → Agent receives final captions → estimated and checked timeline → native animation and automatic advance → MP4 export → playback QA and delivery.**

## Step 1: Prepare captions, then review and confirm

Choose one of two paths: have the Agent draft captions from the PPTX, or supply your own captions for format normalization. Both produce a standard `captions.md` for you to review, edit and confirm before video production.

### Path A: Agent drafts captions from the PPTX

Replace these example paths with real locations:

```text
Use $powerpoint-auto-subtitles to draft English captions automatically from my PPTX.
PPTX: C:/work/my-presentation.pptx
Caption output: C:/work/captions.md (UTF-8)
Review notes: C:/work/captions-review.md
Read each slide's visible text, figures and speaker notes, in actual slide order.
Write one idea per cue in presentation order. Preserve values, units and qualifications.
Do not invent unsupported claims or mechanically copy every chart label.
Use ## Slide N headings and one - bullet per complete English caption.
List ambiguous, unreadable or uncertain content in the separate review notes,
not in the visible captions. Do not fabricate captions for pages you cannot interpret.
For this stage, deliver only the draft and review notes. I will edit them manually
before asking you to create animations and export the video.
```

If you have speaker material, add “Reference script: C:/work/notes.docx”. Ask the Agent to flag conflicts with the deck in the review notes instead of choosing a value silently.

### Path B: Supply captions and let the Agent normalize the format

Provide your existing Word, TXT, Markdown or pasted captions together with the PPTX. The Agent organizes them without generating a replacement script.

```text
Use $powerpoint-auto-subtitles to format my supplied captions, not redraft them.
PPTX: C:/work/my-presentation.pptx
My captions: C:/work/my-subtitles.docx (or pasted text labeled by slide)
Normalized output: C:/work/captions.md (UTF-8; preserve the source file)
Review notes: C:/work/captions-review.md
Map captions to actual slide order using ## Slide N and one - bullet per cue.
Normalize formatting and check slide mapping only. Preserve my wording, order,
values and units. Do not translate, polish or add captions without my request.
If splitting a paragraph, retain the original text order. Suggest meaning-changing
edits only in the review notes rather than applying them.
Flag missing pages, ambiguous mappings and conflicts with the PPTX; do not guess.
Deliver the normalized captions and review notes. Wait for my manual review,
edits and confirmation before producing the video.
```

If you also want translation or polishing, specify the target language and permitted scope.

### Both paths: Review, edit and confirm the caption file


Use Notepad, VS Code or another text editor. Check slide correspondence, values and units; remove unwanted narration, add necessary explanations, polish the English and split long cues. Save as UTF-8 while retaining this structure:

```markdown
# Captions

## Slide 1
- This slide introduces the research platform.
- The diagram shows its main components.

## Slide 2
- This slide compares the two processing paths.
- The following animation explains the data flow.
```

- `## Slide N` is the actual one-based slide index, not its printed page label.
- Each `- ` line is one cue. Keep it on one physical line; multiple cues per slide are welcome.
- Only visible caption text belongs here; questions and suggestions belong in the review notes.
- Check all slides. The standard exporter requires at least one cue per slide; resolve incomplete or uncertain pages before export instead of using placeholders.
- No timestamps are needed yet. Timing will be estimated from the edited wording.

[Standard template](examples/captions.template.md) · [English example](examples/captions.en.md) · [Chinese format example](examples/captions.zh-CN.md) (format references, not forms you must fill from scratch).

### Return the confirmed captions to the Agent

After reviewing, editing and confirming, use the Step 2 request below and name `C:/work/captions.md` as the final caption source. The Agent should retain your edited wording and recalculate timing, not replace your script with a fresh draft from the deck. Explicitly request a one-pass workflow if you want to skip manual review.

## Step 2: Submit the complete production request

Replace the example paths with actual existing files and a new output folder. This example uses the caption document you manually edited in Step 1. An attachment can replace a path if identified by its exact filename. If using audio, give its path and request alignment to that recording.

```text
Use $powerpoint-auto-subtitles.
PPTX: C:/work/my-presentation.pptx
Caption source: C:/work/captions.md (my reviewed and confirmed final English captions)
Caption language: English. Preserve my edited wording; do not regenerate or overwrite the script.
Output directory: C:/work/captioned-output
Narration: none; create a silent captioned video.
Target total duration: 160 seconds.
Read the deck and final captions. Check slide-to-caption correspondence.
Calculate timings from the final text and save a copy in the output folder without overwriting the input.
Then continue to timeline generation, animation and video export; do not stop after formatting captions.
Estimate reading time from text length. If the captions cannot fit readably within
160 seconds, report the conflict rather than silently shortening text or speeding it up.
Apply native PowerPoint fades and automatic slide advance. Preserve all measured
values, charts and editable shapes. Do not overwrite the source.
Inspect the actual MP4 captions and duration. Deliver PPTX, MP4, PDF preview,
final timing CSV and a short QA report.
```

The example target of 160 seconds is configurable. The included public demo is 24 seconds.

Without a draft, replace the caption-source line with:

> Caption source: no existing script. Draft English captions only from this PPTX's text, figures and speaker notes, without inventing information.

Without a fixed duration requirement, replace the target with “Estimate from text length and report the final duration.”

## Step 3: Generate the caption timeline

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

## Step 4: Animate and export

Each subtitle is a separate editable text box. Its animation sequence is:

**Fade in → hold → fade out → next caption**

The helper uses PowerPoint's native `TimeLine.MainSequence.AddEffect` API and the Fade effect. It does not burn text into slide images or substitute an SRT file for PowerPoint animation.

For a caption starting at `s`, ending at `e`, with entrance duration `fi`, exit duration `fo`, and previous caption end `p`:

| Effect | Start mode | Delay | Duration |
|---|---|---|---|
| Entrance | After Previous | `s - p` | `fi` |
| Exit | After Previous | `e - s - fi - fo` | `fo` |

The first caption uses `p = 0`. With `fi = fo = 0.25 s`, a 4-second cue holds for 3.5 seconds. This makes the exit finish at the scheduled cue end. The helper also supports intentional gaps between cues.

### Export with stored slide timings

The helper enables timed slide advance and disables click-to-advance. PowerPoint's `CreateVideo` exports the presentation using its stored timings and animations. No clicks or voice-over are required for a silent demonstration.

## Step 5: Verify and deliver

Check the duration of the real MP4 and inspect a decoded frame inside every caption interval. Also inspect selected fades and slide boundaries. Look for text that fails to disappear, overlapping captions, timing drift, and captions covering charts or video output.

An export-success status or a valid timeline CSV alone does not prove correct playback. The optional extraction helper reports media duration and produces frames for inspection; it does not automatically certify the result.

Deliver the normalized `captions.md`, final timing CSV, editable animated PPTX, actual MP4, static PDF preview and short QA notes covering decoded duration, caption/visual synchronization and remaining limitations. The caption document is an intermediate artifact. A complete production request ends after video export and playback verification; stop after document preparation only when explicitly requested.

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

