"""Convert a per-slide UTF-8 caption document into the animation helper's CSV."""
import argparse
import csv
import math
import re
import sys
from pathlib import Path

HEADER = re.compile(r'^## Slide ([1-9][0-9]*)(?:\s*\|\s*(.+))?$')
CJK = re.compile(r'[\u3400-\u4dbf\u4e00-\u9fff\uf900-\ufaff\U00020000-\U000323af]')
WORDS = re.compile(r"[^\W_]+(?:['’.-][^\W_]+)*", re.UNICODE)
FIELDS = ['slide_index','slide_number','slide_duration_s','caption_id',
          'local_start_s','local_end_s','fade_in_s','fade_out_s','text']


def parse_document(text):
    slides = []
    for line_no, raw in enumerate(text.splitlines(), 1):
        line = raw.strip()
        if not line or line.startswith('# '):
            continue
        match = HEADER.fullmatch(line)
        if match:
            index = int(match[1])
            if index != len(slides) + 1:
                raise ValueError(f'Line {line_no}: slide indexes must be consecutive, starting at 1.')
            slides.append((index, match[2] or str(index), []))
        elif line.startswith('- ') and slides:
            caption = line[2:].strip()
            if not caption:
                raise ValueError(f'Line {line_no}: empty caption.')
            slides[-1][2].append(caption)
        else:
            raise ValueError(f'Line {line_no}: expected ## Slide N or - caption; keep each cue on one line.')
    if not slides or any(not cues for _, _, cues in slides):
        raise ValueError('Every slide needs at least one caption.')
    return slides


def build_rows(slides, english_wps=2.5, chinese_cps=4.0, min_seconds=2.5,
               pause=0.4, fade=0.25, gap=0.15, tail=0.5):
    values = (english_wps, chinese_cps, min_seconds, pause, fade, gap, tail)
    if not all(math.isfinite(v) for v in values):
        raise ValueError('Timing settings must be finite.')
    if min(english_wps, chinese_cps, min_seconds, fade) <= 0 or min(pause, gap, tail) < 0:
        raise ValueError('Rates, minimum duration and fade must be positive; pause, gap and tail nonnegative.')
    rows = []
    for index, label, cues in slides:
        local = 0.0
        slide_rows = []
        for n, caption in enumerate(cues, 1):
            cjk_count = len(CJK.findall(caption))
            word_count = len(WORDS.findall(CJK.sub(' ', caption)))
            if not cjk_count and not word_count:
                raise ValueError(f'Slide {index}, cue {n}: caption has no readable words or Han characters.')
            seconds = round(max(min_seconds, cjk_count / chinese_cps + word_count / english_wps + pause + 2*fade), 3)
            if seconds > 12:
                print(f'Warning: slide {index}, cue {n} lasts {seconds}s; consider splitting it.', file=sys.stderr)
            end = round(local + seconds, 3)
            slide_rows.append(dict(slide_index=index, slide_number=label, caption_id=f'CAPTION_S{index:03d}_{n:03d}',
                                   local_start_s=local, local_end_s=end, fade_in_s=fade, fade_out_s=fade, text=caption))
            local = round(end + gap, 3)
        duration = round(slide_rows[-1]['local_end_s'] + tail, 3)
        for row in slide_rows:
            row['slide_duration_s'] = duration
        rows.extend(slide_rows)
    return rows


def main():
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument('input', type=Path, help='UTF-8 .md or .txt caption document')
    p.add_argument('output', type=Path, help='New timeline CSV (never overwrites)')
    for flag, default in [('english-wps',2.5),('chinese-cps',4.0),('min-seconds',2.5),
                          ('pause',0.4),('fade',0.25),('gap',0.15),('tail',0.5)]:
        p.add_argument('--'+flag, type=float, default=default)
    a = p.parse_args()
    try:
        if a.input.suffix.lower() not in ('.md','.txt'):
            raise ValueError('Convert Word/PDF/chat input to the documented UTF-8 format first.')
        slides = parse_document(a.input.read_text(encoding='utf-8-sig'))
        settings = {k:getattr(a,k) for k in ['english_wps','chinese_cps','min_seconds','pause','fade','gap','tail']}
        rows = build_rows(slides, **settings)
        with a.output.open('x', encoding='utf-8-sig', newline='') as f:
            w = csv.DictWriter(f, fieldnames=FIELDS)
            w.writeheader()
            w.writerows(rows)
        duration = sum(next(r['slide_duration_s'] for r in rows if r['slide_index']==i) for i,_,_ in slides)
        print(f'{len(slides)} slides, {len(rows)} captions; estimated total {duration:.3f}s. Review before export.')
    except (ValueError, OSError) as e:
        p.exit(2, f'Error: {e}\n')


if __name__ == '__main__':
    main()
