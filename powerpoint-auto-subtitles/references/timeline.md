# Subtitle timeline contract

Use UTF-8 CSV. Required columns:

| Field | Meaning |
|---|---|
| `slide_index` | One-based position in the actual PPTX |
| `slide_number` | Printed page label, retained as metadata |
| `slide_duration_s` | Total time before automatic advance |
| `caption_id` | Unique PowerPoint shape name beginning with `CAPTION_` |
| `local_start_s` | Cue entrance start relative to the slide |
| `local_end_s` | Time by which the exit fade has finished |
| `fade_in_s` | Entrance duration, typically 0.25 s |
| `fade_out_s` | Exit duration, typically 0.25 s |
| `text` | Visible subtitle text |

The helper also writes `global_start_s` and `global_end_s`, calculated from preceding slide durations. These represent the scheduled times; verify the exported video separately.

```csv
slide_index,slide_number,slide_duration_s,caption_id,local_start_s,local_end_s,fade_in_s,fade_out_s,text
1,08,12,CAPTION_S08_01,0,4,0.25,0.25,This testbed runs CUDA-accelerated OCUDU on DGX Spark.
1,08,12,CAPTION_S08_02,4,8,0.25,0.25,An FRTek O-RU connects to a real iPhone.
1,08,12,CAPTION_S08_03,8,12,0.25,0.25,Cloud 5G Core supports end-to-end data traffic.
```

Every slide must have at least one cue for the supplied helper. For intentionally silent slides or slides with existing effects, adapt the authoring workflow rather than inventing caption text.

For cue start `s`, end `e`, entrance duration `fi`, exit duration `fo`, and previous cue end `p`:

- Entrance: After Previous, delay `s − p`, duration `fi`.
- Exit: After Previous, delay `e − s − fi − fo`, duration `fo`.
- For the first cue, use `p = 0`.
- Require `0 ≤ s < e ≤ slide_duration`, no overlapping cues, and `e − s ≥ fi + fo`.

This sequence makes the exit finish at `e`; the next entrance can start at the same timestamp or after an intentional gap. Do not assign an absolute per-slide time as the delay of each After Previous effect.

Quote CSV text containing commas or line breaks. Keep subtitle text concise enough for the reserved band and intended playback resolution. Display numbers are labels, not array indices.
