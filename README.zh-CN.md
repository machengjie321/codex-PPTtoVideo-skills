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

**字幕准备（二选一：Agent 根据 PPT 起草 / 用户提供字幕，Agent 整理格式）→ 用户检查、修改并确认 → 将最终字幕交回 Agent → 估算并检查时间表 → 添加原生动画与自动翻页 → 导出 MP4 → 检查并交付。**

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

检查、修改并确认后，使用下面第二步的请求，明确指定 `C:/work/captions.md` 为最终字幕来源。Agent 应保留你修改后的措辞，重新计算时间，而不是重新从 PPT 覆盖字幕稿。若只想从 PPT 一次性生成视频、跳过手动审稿，需要明确提出这一要求。

## 第二步：提交完整制作请求

请把示例路径替换为实际存在的文件和新的输出目录。下面使用第一步经你手动修改后的标准字幕文件。也可以用附件的准确文件名指定输入。若有配音，请同时写明音频路径并要求按实际语音对齐。

```text
请使用 $powerpoint-auto-subtitles。
PPT 文件：C:/work/my-presentation.pptx
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

## 第三步：生成字幕时间表

默认英文每秒 **2.5 个单词**，中文每秒 **4 个汉字**；中英混排时分别计算后相加。每条另加 **0.4 秒阅读停顿**及淡入、淡出各 **0.25 秒**，最短显示 **2.5 秒**。

```text
字幕时长 = max(2.5, 英文单词数 / 2.5 + 汉字数 / 4 + 0.4 + 0.5)
```

标点不计为汉字；英文缩写及数字按词计数。相邻字幕间隔默认 0.15 秒，每页最后留 0.5 秒。字多时延长的是停留时间，不是淡入淡出时间。超过 12 秒会提醒拆分，不会截断文字或强行加速。还需要检查实际文字是否能在字幕区域内清晰显示。

```powershell
python ./powerpoint-auto-subtitles/scripts/captions-to-timeline.py `
  ./examples/captions.zh-CN.md ./timeline-estimated.csv

# 技术内容可以放慢阅读速度：
python ./powerpoint-auto-subtitles/scripts/captions-to-timeline.py `
  ./examples/captions.zh-CN.md ./timeline-slower.csv `
  --english-wps 2.0 --chinese-cps 3.2 --pause 0.7
```

然后把生成的 CSV 传给 `add-subtitles-and-export.ps1` 的 `-TimelineCsv` 参数。文字转时间表这一步需要 Python 3，不能覆盖已存在的输出文件。示例字幕文档用于展示输入格式，并不是已发布 24 秒演示的原始字幕稿。

## 第四步：制作原生动画并导出

每条字幕都是独立文本框，动画顺序为：

**淡入 → 停留 → 淡出 → 下一条字幕**

脚本通过 PowerPoint 的 `TimeLine.MainSequence.AddEffect` 创建原生 Fade 动画。字幕可以在 PPT 中直接修改；单独生成 SRT 文件不能替代这个动画过程。

这里使用自动的 **After Previous（上一动画之后）** 顺序。设一条字幕开始时间为 `s`、结束时间为 `e`、淡入时长为 `fi`、淡出时长为 `fo`，上一条结束时间为 `p`：

| 动画 | 触发方式 | 延迟 | 持续时间 |
|---|---|---|---|
| 淡入 | 上一动画之后 | `s - p` | `fi` |
| 淡出 | 上一动画之后 | `e - s - fi - fo` | `fo` |

首条字幕使用 `p = 0`。例如一条字幕共 4 秒，淡入和淡出分别为 0.25 秒，则中间停留 3.5 秒。下一条可紧接着开始，也可根据时间表留出空档。

### 自动翻页并导出视频

脚本设置每页自动播放时长，关闭点击翻页，并让 PowerPoint 按已有计时和动画调用 `CreateVideo` 导出 MP4。因此，无声演示也可以自动完整播放。

## 第五步：检查并交付

确认视频真实总时长，在每条字幕区间抽取画面，并额外检查部分淡入淡出和翻页边界。重点检查字幕是否正确消失、有无残影或重叠、是否遮住图表，以及翻页是否产生时间偏差。

“时间表正确”或“导出状态成功”不代表播放效果一定正确。附带的抽帧脚本只负责解码、读取时长和生成待检查画面，不会自动判定视觉效果合格。

完整交付包括：

| 文件 | 用途 |
|---|---|
| `captions.md` | 标准逐页字幕稿 |
| `subtitle_timeline_FINAL.csv` | 最终字幕起止时间及翻页时间 |
| 字幕版 `.pptx` | 可编辑字幕、原生动画及自动翻页 |
| `.mp4` | 实际自动播放视频 |
| `.pdf` | 静态预览，每页展示首条字幕 |
| QA 记录 | 实际时长、字幕与图形同步检查、已知限制 |

字幕文件只是中间产物；完整制作任务以视频导出及实际播放检查完成为终点。只有明确要求“仅整理字幕”时，才在第一阶段结束。

## 直接运行脚本

按照 [CSV 格式说明](powerpoint-auto-subtitles/references/timeline.md) 准备时间表后运行：

```powershell
pwsh -NoProfile -File ./powerpoint-auto-subtitles/scripts/add-subtitles-and-export.ps1 `
  -InputPptx ./my-presentation.pptx `
  -TimelineCsv ./timeline.csv `
  -OutputDirectory ./captioned-output
```

脚本输出新 PPTX、MP4、PDF、逐页 PNG、最终 CSV 和执行记录，不覆盖原文件。执行记录会将实际视频检查标记为待完成，直到人工或 Codex 真正检查了解码结果。

该辅助脚本遇到已有动画或字幕对象会停止，避免清空原有播放逻辑。已有复杂动画时，请让 Codex 有针对性地合并时间线。字幕区域也需要先准备好，脚本不会自动重新排版原图表。

抽取视频检查帧：

```powershell
pwsh -NoProfile -Sta -File ./powerpoint-auto-subtitles/scripts/extract-video-frames.ps1 `
  -VideoPath ./captioned-output/my-presentation_review.mp4 `
  -TimelineCsv ./captioned-output/subtitle_timeline_FINAL.csv `
  -OutputDirectory ./video-check
```

更多内容见 [英文 Skill](powerpoint-auto-subtitles/SKILL.md)、[英文流程介绍](powerpoint-auto-subtitles/references/workflow-walkthrough.md) 和 [PowerPoint 实现与检查说明](powerpoint-auto-subtitles/references/office-workflow.md)。

编辑模式中，多条字幕文本框可能叠放在同一区域；播放时由动画控制出现与消失。静态 PDF 和 PNG 每页只展示首条字幕。

