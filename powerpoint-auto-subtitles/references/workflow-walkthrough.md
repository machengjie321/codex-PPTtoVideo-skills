# How we use Codex to add PowerPoint subtitles and export a video

Codex can automate the work between a finished presentation and a captioned video. We provide the PowerPoint deck, the intended running time, and the text or narration that the captions should follow. Codex then prepares a subtitle timeline, adds editable captions inside PowerPoint, assigns native animation timings, and exports the presentation as an MP4.

The process has six steps:

1. **Prepare the slides.** Codex checks the deck and reserves a consistent subtitle area. Charts, tables and diagrams remain editable. Existing slide content determines what the captions can claim.
2. **Build the subtitle timeline.** Each caption receives a slide index, a unique name, a start time, an end time, and fade durations. The timeline also defines how long each slide stays on screen.
3. **Add editable subtitle text boxes.** Codex creates one native PowerPoint text box per caption. This lets us revise the wording or styling later without rebuilding the whole slide.
4. **Apply native animation and automatic advance.** Each caption fades in, stays visible for its assigned interval, and fades out. PowerPoint advances to the next slide automatically. The sequence runs without clicks.
5. **Export the video through PowerPoint.** Codex uses the stored slide timings and animations to generate the MP4. It can also export a PDF and slide images for review.
6. **Check the actual video.** Codex checks the total duration and inspects frames from the exported file. This catches issues such as overlapping subtitles, captions that fail to disappear, incorrect slide timing, or text covering a graph.

In our example, nine slides displayed as pages 08–16 used 29 captions and ran for 160 seconds. The slide durations were 12, 16, 18, 25, 16, 22, 13, 18 and 20 seconds. Each caption used a 0.25-second entrance and exit fade. The actual PowerPoint-exported video was checked at a point inside every caption interval. These numbers describe that example; the same process supports different slide counts and running times.

The key lesson is that writing a subtitle file is only part of the work. Captions must exist as timed objects in PowerPoint, and the exported video must be checked to confirm that they play correctly.

## Example request to Codex

> Use $powerpoint-auto-subtitles to add editable English subtitles to this PowerPoint and export an automatically advancing MP4. Target a total duration of 160 seconds. Prepare a caption timeline, reserve a readable subtitle band, use native Fade entrance and exit effects, and verify the actual exported video. Preserve the measured values and existing editable content. Leave any missing demonstration clip as a clearly labeled placeholder. Deliver the animated PPTX, MP4, final timing CSV, PDF preview, and a short QA report.

## Short explanation for a presentation

> We use Codex to turn an existing PowerPoint into a captioned video. It first creates a subtitle timeline, then adds editable text boxes with native fade animations and automatic slide transitions. PowerPoint exports the timed presentation as an MP4. Finally, we inspect the real video to check subtitle changes, slide timing, and the total duration.
