---
name: powerpoint-auto-subtitles
description: Add editable timed subtitles to a PowerPoint deck and export an automatically advancing MP4 using desktop Microsoft PowerPoint on Windows. Use for slide-based captioned videos, including silent demonstrations and research presentations.
---

# PowerPoint automatic subtitles and video

Turn an existing local PPTX into an editable captioned presentation and a video. Keep the original deck intact. Subtitles are native PowerPoint text boxes with entrance and exit effects, rather than flattened slide images or an SRT-only deliverable.

## Inputs and decisions

Resolve a specific source PPTX (existing absolute path, exact attachment name, or explicitly identified open deck), the output directory, duration policy, caption language, and caption source before authoring. Caption sources can be a named document, explicitly supplied per-slide text, a named narration file, or an explicit request to draft from the deck. Verify that referenced files exist and that captions map to the actual slide order. If the source deck or caption source remains ambiguous in the conversation, ask only for that missing information; do not select an unrelated historical deck or invent a script source. A pasted example request with placeholder paths is not a real input deck.

Identify any real media to include. Reuse supplied assets and measured values. If no audio is supplied, write concise captions from the slide content; do not describe this as speech transcription. When audio exists, align captions to the actual narration.

For drafting captions from slides, accepting per-slide text or a document, and estimating timing from text length, read [caption authoring and timing](references/caption-authoring.md). The included `scripts/captions-to-timeline.py` converts normalized Markdown/TXT to the animation CSV with adjustable English-word and Chinese-character reading rates. Use actual speech timings when narration exists.

Produce a UTF-8 timeline before animating. Read [the timeline contract](references/timeline.md) for the schema and timing formulas. A slide's position in the file (`slide_index`) can differ from its printed page number (`slide_number`). Neither nine slides nor 160 seconds is a universal default.

## Caption-first intake

Support two caption-preparation paths: (A) draft from the specified PPTX; (B) normalize captions supplied by the user without redrafting them. For path B, preserve wording, order, values and units; translate or polish only when requested. Keep unresolved page mappings and proposed meaning-changing edits in separate review notes. Both paths can converge on user review, editing and confirmation before production when this staged workflow is requested. Support a draft–edit–produce workflow: when requested, automatically read the specified deck and save a standard per-slide caption draft plus separate review notes. Do not require the user to fill the template manually first. Flag unreadable content and uncertain claims in review notes rather than fabricating captions. When the user requests manual review before production, deliver the draft and wait for the edited file. In the later production stage, treat that edited file as the caption source of truth; preserve its wording, recalculate timing from it, and do not overwrite it with a new deck-derived script. This review stage is requested workflow scope, not a mandatory gate for explicitly authorized one-pass export.

Start a new subtitle workflow by helping the user provide captions in the standard per-slide format. Explain that `## Slide N` identifies the actual PPTX slide and every `- ` line is one cue. Accept a named caption file or pasted slide-labeled text. Normalize Word or unstructured drafts into a UTF-8 `captions.md` before computing timings. Keep visual directions and task instructions outside caption text. Check slide order, missing pages and unreplaced template placeholders. Do not invent missing captions when the user chose a supplied-script workflow; draft from the deck only when requested. If the user asks only to organize the caption document, deliver that document and mapping check without proceeding to animation/export. Reuse an already valid caption file instead of repeating intake. For an end-to-end video request, normalization is an intermediate step: continue through timing, native animation, export and actual playback QA without an unnecessary approval gate. Stop after the caption document only when that limited scope is explicitly requested, or when unresolved missing inputs prevent further work.

## Workflow

1. **Inspect and prepare the deck.** Preserve charts, tables, media, notes and existing design. Reserve a bottom caption band outside plot labels and video output. Reflow content first when needed. Adapt caption size to the actual slide dimensions. Keep captions to one or two lines.
2. **Write and check the timeline.** Give every cue a unique `CAPTION_…` shape name. Check local cue bounds, gaps, overlaps, slide durations and cumulative video time. Keep the final CSV with the outputs.
3. **Create native caption effects.** On Windows with desktop PowerPoint, use the supplied helper when its preconditions fit. It creates editable text boxes and native Fade entrance/exit effects through the PowerPoint object model. Its timing model uses automatic **After Previous** effects, which avoids ambiguity about absolute delays in With Previous chains.
4. **Set automatic slide advance.** Use the timeline duration for each slide, disable click-to-advance, and export using recorded slide timings and animations. Preserve unrelated presentation features.
5. **Export through PowerPoint.** Deliver the animated PPTX, actual MP4, final CSV, and requested PDF/PNGs. Static previews show the first cue on each slide; the PPTX retains every cue. A missing demonstration video stays a clearly labeled placeholder.
6. **Verify actual output.** Inspect the rendered slides and sample the MP4 inside every cue, across slide boundaries, and around selected fades. Confirm duration, readable captions, correct disappearance, and no leftover subtitle text. XML inspection, an SRT file, or an export-completed status alone does not establish correct playback. Report unverified items explicitly.

## Helpers

`scripts/add-subtitles-and-export.ps1` accepts a local PPTX and timeline CSV. It writes a copy, adds captions, exports PDF/PNGs and, unless requested otherwise, an MP4. It refuses existing animation sequences or caption objects because merging those sequences needs deliberate authoring. For an already animated source, preserve and adapt the existing timeline using the available presentation tools instead of clearing all effects.

```powershell
pwsh -NoProfile -File "<skill>/scripts/add-subtitles-and-export.ps1" `
  -InputPptx "C:/work/presentation.pptx" `
  -TimelineCsv "C:/work/subtitle_timeline.csv" `
  -OutputDirectory "C:/work/captioned-output"
```

Run only through an execution method permitted in the current environment. The helper requires a Windows desktop login session with installed Microsoft PowerPoint. A sandbox COM login-session error is an environment limitation: use the approved desktop execution path or native UI, and preserve any required approval boundary. Do not weaken security settings or bypass execution policy.

For video QA, prefer an available media decoder. The optional `scripts/extract-video-frames.ps1` uses Windows media decoding and needs PowerShell 7 in STA mode. Read [Office playback and export notes](references/office-workflow.md) when troubleshooting or adapting existing animations.

## Completion

Deliver concise links to the requested files. State whether the MP4's actual timing and caption changes were checked. Keep the qualification narrow: identify missing real media or untested playback without implying that a placeholder proves a live demonstration.

For an English explanation suitable for colleagues, use [the workflow walkthrough](references/workflow-walkthrough.md).
