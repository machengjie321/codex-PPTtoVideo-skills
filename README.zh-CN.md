# 使用 Codex 自动给 PowerPoint 添加字幕并导出视频

[English](README.md) | [简体中文](README.zh-CN.md)

这个项目把“给 PPT 配字幕、设置自动播放、导出视频”的流程封装为一个 Codex Skill。

**字幕使用 PPT 内部原生文本框和动画实现，PPTX 中的文字可以继续编辑。** Codex 先编排时间表，再调用桌面版 PowerPoint 添加淡入淡出和自动翻页，最后由 PowerPoint 导出 MP4。

## 演示与下载

![字幕自动播放效果](media/subtitle-playback.gif)

*上方为实际导出视频的 8 秒抽帧预览，完整 24 秒视频及源文件见下方。*

- [可编辑字幕 PPTX](examples/subtitle-demo_captioned.pptx)
- [实际导出的视频 MP4](examples/subtitle-demo_review.mp4)
- [静态 PDF 预览](examples/subtitle-demo_preview.pdf)
- [字幕时间表 CSV](examples/subtitle_timeline_FINAL.csv)
- [示例验证记录](examples/VERIFICATION.md)

公开示例包含 3 页通用演示内容、9 条字幕，总时长 24 秒。示例未包含原始科研数据或实验室照片。

## 两幅同步变化的示意图

上方左图展示幻灯片内容逐步出现，右图展示视频播放与字幕时间轴推进；下方字幕同步切换。演示将每个时段的两幅图与可编辑字幕编为一组，再添加 PowerPoint 原生淡入/淡出动画。此演示使用专门的场景编排，通用字幕脚本仍只负责字幕，不会自动改写已有图形动画。PPTX 中未嵌入录屏或预渲染 GIF。

## 环境要求

- Windows 与已安装的桌面版 Microsoft PowerPoint。
- PowerShell 7（`pwsh`）。
- 可连接 PowerPoint 的桌面登录会话。
- 当前执行环境允许控制 PowerPoint 并写入指定输出位置。

## 安装 Skill

克隆仓库，将其中 `powerpoint-auto-subtitles` 文件夹放入 Codex 的 skills 目录。Windows 常见位置是 `%USERPROFILE%\.codex\skills`；如配置了其他 `CODEX_HOME`，使用对应位置。

```powershell
git clone https://github.com/machengjie321/codex-PPTtoVideo-skills.git
Copy-Item -LiteralPath ./codex-PPTtoVideo-skills/powerpoint-auto-subtitles `
  -Destination "$env:USERPROFILE/.codex/skills/powerpoint-auto-subtitles" -Recurse
```

如果当前聊天未发现新 Skill，可刷新可用技能或新建聊天。已有同名 Skill 且包含自定义修改时，不要直接覆盖。

## 完整流程

**字幕准备（二选一：Agent 根据 PPT 起草 / 用户提供字幕，Agent 整理格式）→ 用户检查、修改并确认 → Agent 统一版式并预留字幕区 → 用户检查确认 → 使用确认后的 PPT 与字幕 → 估算并检查时间表 → 添加原生动画与自动翻页 → 导出 MP4 → 检查并交付。**

## 第一步：准备字幕，再由用户检查确认

字幕准备有两种方式：没有现成字幕时，让 Agent 根据 PPT 起草；已有字幕时，提供给 Agent 整理格式。两条路径都输出标准的 `captions.md`，再由你检查、修改并确认，之后进入视频制作。

### 路径 A：由 Agent 根据 PPT 起草字幕

把下面的路径替换为实际文件：

```text
请使用 $powerpoint-auto-subtitles，根据 PPT 自动生成英文字幕初稿。
PPT 文件：C:/work/my-presentation.pptx
字幕输出：C:/work/captions.md（UTF-8）
核对记录：C:/work/captions-review.md
请读取每页可见文字、图表和演讲者备注，按实际页序起草字幕。
每条只讲一个意思，按讲解顺序排列；保留数值、单位和结论的限定条件，
不要添加 PPT 没有支持的信息，不要将每个图表标签机械地照抄成字幕。
使用 ## Slide N 标题，每个 - 列表项是一条完整英文字幕。
在独立的核对记录中列出含义不清、无法读取或需要我确认的内容，
不要把这些备注混入字幕正文，也不要为无法理解的页面编造字幕。
本次只生成初稿和核对记录，等我手动修改后，再制作动画和视频。
```

如果另有讲稿，可增加“参考讲稿：C:/work/notes.docx”。若讲稿与 PPT 冲突，要求 Agent 在核对记录中指出，不自行选择一个数值。

### 路径 B：用户提供字幕，Agent 整理格式

已有 Word、TXT、Markdown 或聊天中的字幕内容时，不必让 Agent 重新起草。提供字幕和对应 PPT，让它统一页序与格式，并保留原文。

```text
请使用 $powerpoint-auto-subtitles，整理我提供的字幕，不重新起草。
PPT 文件：C:/work/my-presentation.pptx
用户字幕：C:/work/my-subtitles.docx（也可直接粘贴带页码的字幕）
标准字幕输出：C:/work/captions.md（UTF-8，不覆盖原文件）
核对记录：C:/work/captions-review.md
按 PPT 实际页序整理为 ## Slide N，每条字幕使用一个 - 列表项。
仅统一格式并核对页面对应关系，保留我的措辞、顺序、数值和单位。
不要擅自翻译、润色或补充字幕；长段落需要拆分时保留原文顺序，
涉及含义变化的修改只在核对记录中建议，不直接改写。
缺页、页序不明确或与 PPT 内容冲突时，列出问题，不自行猜测。
交付标准字幕和核对记录，等我手动检查、修改并确认后再制作视频。
```

若希望同时翻译或润色，请在请求中明确说明目标语言和修改范围。

### 两条路径汇合：用户手动检查、修改并确认


可以用记事本、VS Code 或其他文本编辑器直接编辑。核对每页是否对应、数值和单位是否正确；删除不想讲的内容，补充必要说明，调整英文措辞，并把过长字幕拆成几条。保存为 UTF-8，保留下面的结构：

```markdown
# Captions

## Slide 1
- This slide introduces the research platform.
- The diagram shows its main components.

## Slide 2
- This slide compares the two processing paths.
- The following animation explains the data flow.
```

- `## Slide N` 是 PPT 中从 1 开始的实际页序，不是印刷页码。
- 每个 `- ` 行是一条字幕；同一条不要手动换行。一页可以有多条。
- 字幕文件只放要显示的文字；问题和修改建议放在独立核对记录里。
- 核对所有页面。标准导出脚本要求每页至少一条字幕；未完成或有疑问的页面先解决，不用占位文字直接导出。
- 此时不需要填写秒数，后续会根据修改后的文字估算时间。

[标准模板](examples/captions.template.md) · [英文示例](examples/captions.en.md) · [中文格式示例](examples/captions.zh-CN.md)（供检查格式使用，不要求手动从头填写。）

### 将确认后的字幕交回 Agent

检查、修改并确认后，先完成第二步的版式检查，再使用第三步的请求，明确指定 `C:/work/captions.md` 为最终字幕来源。Agent 应保留你修改后的措辞，重新计算时间，而不是重新从 PPT 覆盖字幕稿。若只想从 PPT 一次性生成视频、跳过手动审稿，需要明确提出这一要求。

## 第二步：Agent 统一版式并预留字幕区，用户检查确认

正式添加字幕前，先让 Agent 统一各页的字幕位置和底部留白。可先以页面底部约 **10%–15% 高度**作为字幕安全区的起点，再按字号、行数和实际内容调整；这不是固定比例。字幕区与正文之间应留出间距，不能遮住坐标轴、图注、页码或标志。

Agent 应尽量保留原有视觉风格，统一字幕区的高度、左右边距及文字样式。必要时移动或适度缩放正文，使各页形成一致的布局；不能通过白色矩形覆盖原内容来制造留白，也不能删改实验数值。已有页脚需要合理调整位置。复杂或拥挤页面应在检查记录中单独说明。

先提交这个排版请求：

```text
请使用 $powerpoint-auto-subtitles，先整理 PPT 版式和底部字幕区域。
PPT 文件：C:/work/my-presentation.pptx
最终字幕：C:/work/captions.md
排版输出目录：C:/work/layout-review
请按字幕实际长度，为所有页面预留一致的底部字幕安全区，
以底部约 10%–15% 高度为起点，按需要调整，并保留与正文的间距。
统一字幕区域的位置、边距和建议字号，尽量延续原有风格。
必要时调整正文位置和尺寸，保留实验数值、图表、可编辑图形及已有动画，
不要遮挡或删除坐标轴、图注、页码和标志，不覆盖原文件。
交付可编辑的 presentation-layout.pptx、PDF/逐页预览和版式检查记录。
预览中展示每页最长字幕的占位效果，以便检查两行字幕是否放得下；
预览占位不要遗留在待添加正式字幕的 PPTX 中。
先不要添加正式字幕动画或导出视频，等我检查并确认版式后继续。
```

**用户检查：**逐页确认内容没有被裁切或缩得难以阅读，图表标注完整，底部留白一致，最长字幕能清晰显示。需要调整时请 Agent 修改并更新预览。确认后，将排版版 `C:/work/layout-review/presentation-layout.pptx` 作为下一步的输入，避免再次使用未整理的原始 PPT。

## 第三步：提交完整制作请求

请把示例路径替换为实际存在的文件和新的输出目录。下面使用第一步经你手动修改后的标准字幕文件。也可以用附件的准确文件名指定输入。若有配音，请同时写明音频路径并要求按实际语音对齐。

```text
请使用 $powerpoint-auto-subtitles。
PPT 文件：C:/work/layout-review/presentation-layout.pptx（已确认版式及底部字幕区）
字幕来源：C:/work/captions.md（我已检查并确认的最终英文字幕）
字幕语言：英文；保留我修改后的措辞，不重新生成或覆盖字幕稿。
输出目录：C:/work/captioned-output
配音：无；制作无声字幕视频。
目标总时长：160 秒。
先读取 PPT 和最终字幕，核对页序及字幕对应关系。
使用最终字幕计算时间，并将副本保存到输出目录，不覆盖输入文件。
随后继续生成时间表、制作动画和导出视频，不在整理完字幕后结束任务。
按字数估算阅读时间；如果字幕无法在 160 秒内清晰展示，报告冲突，
不要擅自删减文字或强行加速。
用 PPT 原生淡入淡出动画和自动翻页播放，保留原有数值、图表和可编辑图形。
检查实际导出视频的字幕及总时长，不覆盖源文件。
交付 PPTX、MP4、PDF 预览、最终字幕 CSV 和简短 QA 记录。
```

160 秒只是示例目标，实际页数与时长可以调整。本仓库的公开效果演示为 24 秒。

如果没有字幕草稿，将“字幕来源”一行替换为：

> 字幕来源：没有现成字幕，请仅依据这个 PPT 的文字、图表和演讲者备注，按页起草英文字幕，不添加没有依据的信息。

若不需要固定总时长，将“目标总时长”改成“按字数估算，报告最终总时长”。160 秒不是强制默认值。

### 3.1 Agent 生成并检查时间表

提交第三步的完整请求后，以下制作与检查工作由 Agent 连续执行，你不需要逐条运行命令。

Agent 读取确认后的 `captions.md`，生成 `timeline-estimated.csv`，记录每条字幕的起止时间和每页播放时长。默认按英文每秒 2.5 词、中文每秒 4 字估算，另加阅读停顿和淡入淡出时间；这些速度可以调整。有配音时按实际语音对齐。

Agent 检查预计总时长、字幕是否过长以及文字是否放得下。指定 160 秒等目标时，需要核对可读性与时长是否兼容；不能靠截断文字或强行加速凑时长。具体公式见 [计时规则](powerpoint-auto-subtitles/references/caption-authoring.md)。

### 3.2 Agent 添加动画并导出

Agent 使用 [add-subtitles-and-export.ps1](powerpoint-auto-subtitles/scripts/add-subtitles-and-export.ps1)，依据时间表创建可编辑字幕文本框，添加原生 **淡入 → 停留 → 淡出** 动画，设置自动翻页，再由 PowerPoint 导出 MP4 和 PDF 预览。

原有图表、数值和可编辑图形需要保留。已有动画时，Agent 要有针对性地合并时间线；通用脚本不会自动合并已有动画，也不会自动重新排版图表。上方图形的同步变化需单独按字幕时间表编排。

### 3.3 Agent 检查视频并交付

Agent 解码实际 MP4，检查真实总时长、每条字幕、翻页及部分淡入淡出边界，确认没有残影、重叠或遮挡。“导出成功”不是播放检查完成；抽帧工具也不会自动判断画面是否合格。

| 交付文件 | 用途 |
|---|---|
| `captions.md` | 用户确认后的标准字幕稿 |
| `subtitle_timeline_FINAL.csv` | 最终字幕起止时间及每页时长 |
| 字幕版 `.pptx` | 可编辑字幕、原生动画及自动翻页 |
| `.mp4` | 实际自动播放视频 |
| `.pdf` | 静态预览，每页展示首条字幕 |
| QA 记录 | 实际时长、字幕与图形同步检查、已知限制 |

**使用 Codex 的完整流程到这里结束。下面的附录仅供希望自己运行程序的人使用，不是额外需要重复执行的步骤。**

## 可选附录：手动运行脚本

这是上述制作过程的命令行替代方式。先完成字幕与版式确认，确保下面的输入 PPT 已预留字幕区；在仓库根目录打开 PowerShell，准备 `C:/work/my-presentation.pptx` 和 `C:/work/captions.md`。需要 Python 3、PowerShell 7 和 Windows 桌面版 PowerPoint。以下输出路径必须尚未存在，重新运行时换用新路径。

### A. 字幕文件 → 时间表

```powershell
python ./powerpoint-auto-subtitles/scripts/captions-to-timeline.py C:/work/captions.md C:/work/timeline-estimated.csv
```

输入是最终字幕，输出是估算的 CSV。默认公式：

```text
字幕时长 = max(2.5, 英文词数 / 2.5 + 汉字数 / 4 + 0.4 + 0.25 + 0.25)
```

相邻字幕间隔 0.15 秒，每页末尾留 0.5 秒；超过 12 秒会提示拆分，不自动截断。需要放慢时，可在同一命令末尾添加 `--english-wps 2.0 --chinese-cps 3.2 --pause 0.7`。

检查生成的时间表后再进入 B。这个脚本只按字数估时，**不会自动把总时长调整为 160 秒**。有固定时长要求时，需要先合理调整文案或时间表并检查可读性。CSV 字段见 [格式说明](powerpoint-auto-subtitles/references/timeline.md)。

### B. PPT + 时间表 → 原生动画、MP4 和 PDF

```powershell
pwsh -NoProfile -File ./powerpoint-auto-subtitles/scripts/add-subtitles-and-export.ps1 -InputPptx C:/work/my-presentation.pptx -TimelineCsv C:/work/timeline-estimated.csv -OutputDirectory C:/work/captioned-output
```

这里直接使用 A 生成的同一个 CSV。脚本创建原生 Fade 动画和自动翻页，并输出 `my-presentation_captioned.pptx`、`my-presentation_review.mp4`、`my-presentation_preview.pdf`、逐页 PNG、`subtitle_timeline_FINAL.csv` 和执行记录。源 PPT 不会被覆盖。

脚本遇到已有动画或字幕对象会停止；这种情况需要有针对性地合并，不要清空原有动画。执行记录中的视频检查状态此时仍为待完成。

### C. 实际 MP4 → 检查帧与 QA 记录

```powershell
pwsh -NoProfile -Sta -File ./powerpoint-auto-subtitles/scripts/extract-video-frames.ps1 -VideoPath C:/work/captioned-output/my-presentation_review.mp4 -TimelineCsv C:/work/captioned-output/subtitle_timeline_FINAL.csv -OutputDirectory C:/work/video-check
```

这里使用 B 导出的视频及最终 CSV。打开抽取的画面逐条检查，再播放翻页和淡入淡出附近片段，记录实际总时长与发现的问题，完成 QA 记录。抽帧成功本身不等于视觉验证通过。

更多实现细节见 [动画计时公式](powerpoint-auto-subtitles/references/timeline.md)、[PowerPoint 实现与检查说明](powerpoint-auto-subtitles/references/office-workflow.md) 和 [Skill](powerpoint-auto-subtitles/SKILL.md)。编辑模式中字幕可能叠放，播放时由动画控制可见性。
