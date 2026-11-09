# Desktop PowerPoint implementation and QA

## Native effects

The tested Windows workflow uses `PowerPoint.Application`, caption text boxes, `TimeLine.MainSequence.AddEffect`, native Fade effects, and timed slide transitions.

Relevant Office constants used by the helper:

- `msoAnimEffectFade = 10`
- `msoAnimTriggerAfterPrevious = 3`
- `ppSlideShowUseSlideTimings = 2`
- `ppSaveAsOpenXMLPresentation = 24`
- `ppSaveAsPDF = 32`
- `ppMediaTaskStatusDone = 3`; status 4 indicates failure

Use the actual slide width/height in points for caption geometry. A useful starting band is x = 3.125%, y = 90.28%, width = 93.75%, height = 7.5% of the slide. Adapt it when the source contains important bottom content. Start with white 20-pt Arial, centered, over a black fill with 40% transparency. Font size and band height still require visual inspection.

Keep the original presentation untouched. The helper refuses existing output names and does not quit the user's PowerPoint application. If export fails after saving the animated deck, retain that deck and explain which export did not complete.

## Existing animations

The bundled helper assumes no existing main or interactive sequences. Other workflows may preserve existing sequences and deliberately merge captions, but they must account for every preceding effect and trigger. Do not remove slide animations, media triggers, or user caption edits merely to satisfy the helper's preconditions.

Raw OpenXML animation authoring can create a file that passes structural checks but plays incorrectly or prompts repair. Prefer PowerPoint's native effect API when available. A repair operation is a diagnostic action, not an automatic success: compare all slide content and animation counts afterward and reopen the saved result normally. Never silently deliver a file that still requests repair.

## Video export

Use `CreateVideo` with `UseTimingsAndNarrations = true`. The fallback seconds-per-slide argument is not the intended timing source. Poll task status and enforce a timeout. A failed or timed-out export does not justify claiming a valid video.

Native caption objects can overlap in the editing canvas because they occupy the same location. Playback controls visibility. For PDF/PNG review, temporarily hide all except each slide's first cue in an unsaved export session; restore visibility before video export and close without saving the preview visibility changes.

## Validate the exported media

Decode the real MP4, rather than simulating frames from the CSV. Check:

1. Total media duration agrees with the sum of slide durations within an appropriate tolerance.
2. The correct slide and cue appear at the midpoint of every cue.
3. Adjacent cues do not leave residual text or overlapping backing bands.
4. Selected entrance/exit frames show the expected fading behavior.
5. First slide, last slide, slide transitions and any actual video segment render correctly.
6. Captions stay clear of charts, tables and application/video output.

The optional Windows frame-extraction helper sets `ScrubbingEnabled = true`. Without it, a paused decoder can return the same stale first frame after seeking. Check that frames actually differ before interpreting them as playback evidence.

Record the Office export result, observed media duration, sampled timestamps, and any limitations in a short QA report. Do not claim frame-accurate validation from a few midpoint samples.

Do not infer actual media duration solely from short or fractional slide timings. A small helper test scheduled for 6 seconds exported as 7 seconds in the tested Office setup; the 24-second public demonstration and the earlier 160-second presentation matched their scheduled totals. Measure each new export and investigate discrepancies rather than changing a report to match the intended duration.

## Missing media

When the user has not supplied a real demo clip, retain a neutral placeholder. Do not synthesize test results or substitute generic footage. Once the clip arrives, synchronize cue boundaries with actual visual events and recheck the complete export.
