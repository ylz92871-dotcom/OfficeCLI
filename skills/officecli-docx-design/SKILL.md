---
name: officecli-docx-design
description: "Use this skill when a .docx must look 'designed', not just structured — magazine/manual/handout layout with colored callout cards, KPI/cheat-sheet tables, pull quotes, section color-bands, kickers with big titles, and clean page breaks. Trigger for terms like 'design', '漂亮的', '宣传册/手册/讲义', 'callout', '卡片', '色块/色带', 'kicker', '杂志/手册版式'. All commands are officecli add/set/batch — NO python-docx, NO AppleScript. For plain reports/letters/memos route to the `word`/`officecli-docx` skill instead."
---

# OfficeCLI DOCX Design Skill — 杂志 / 手册 / 讲义式排版

This skill turns Word's XML (rendered as shaded/bordered "cards", single-cell tables, color-band rows,
kicker+title pairs) into documents that look designed — the layer the generic `word` skill skips. It **borrows** the design recipes from community docx-design skills (callout cards, KPI tables, pull quotes,
section color-bands, cheat-sheet tables) but rewrites every recipe as **officecli commands** driven through
`add`/`set`/`batch`, because this tool has one binary, no Python, no Office install.

**Before any command: confirm the property against `officecli help docx <element>`.** This skill teaches
*what looks good*, not every flag. When help and this skill disagree, help is authoritative.

Always set `FILE="target.docx"` first and use `"$FILE"` in every command.

---

## ⚠️ Pagination is the QA blind zone — read this FIRST

Word decides page breaks at render time. The CLI **cannot know** whether two blocks land on the same
page; the three bugs below are only visible in a screenshot, so the generic skill never catches them.
These are **static risks** the tool flags via `view issues` before you render:

- **`kicker_keep_next`** — a kicker (short lead line) directly above a title, with no `keepNext`,
  can be stranded on the prior page while the title jumps to the next. Fix: put `keepNext=true` on
  the kicker paragraph (or the title).
- **`card_split_risk`** — a shaded/bordered "card" paragraph without `keepLines` gets torn across a
  page boundary; a single-cell-table card needs `cantSplit=true` on its row. Fix: `keepLines=true`
  (+ `keepNext=true` when the card must stay glued to the following block).
- **`page_break_duplicate`** — a heading sets `pageBreakBefore=true` while its preceding block is
  already an explicit `pagebreak` (or another `pageBreakBefore`), so Word inserts a blank page.
  Fix: use **exactly one** break mechanism per logical boundary.

After each chapter/section, run the grid screenshot (below) so these break early instead of returning
after the whole document is written. **Never rely on `validate`** — these are render-time decisions and
`validate` stays silent on all three.

```bash
# Per-chapter QA (the agent report's recommendation: catch pagination per-chapter, not at the end)
officecli view "$FILE" issues --type format     # expect to be empty of the 3 pagination warnings
officecli view "$FILE" html --page <N>          # or a grid of pages after a multi-page chapter
officecli view "$FILE" screenshot -o "$CHAPTER_PNG" --grid 2   # 2×N contact sheet
```

---

## Design System (pick before authoring)

Pick **one** palette + one accent. Define the accent as hex once and reuse it. (There is no theme
system yet — colors are written into paragraph/table properties — so consistency lives in *you*
keeping the same hex throughout.)

- Kicker: small caps or distinct accent color, ~11–12pt, `spaceAfter=0`.
- Big title: H1 ≥ 20pt; subtitle ~14pt gray below.
- Body: one readable font (Calibri/mid-size CJK 宋体/微软雅黑), 12pt, `lineSpacing=1.3`.
- Cards: light gray `shd` fill + `pBoarder`, or dark accent fill + white text.
- Tables for KPI/cheat-sheet: header row with accent fill + white bold text; `wordWrap` on.

**Chinese documents**: never rely on the Latin default. Set an East Asia face on runs (see
Chinese Typography) or the CJK glyphs silently fall back to an unknown font. Use the
`typography.preset` presets added for exactly this.

---

## Recipes (all officecli)

### 1. Callout card (shaded box)

A compact box for a tip / warning / key point. A *single* shaded/bordered paragraph reads as a card
in Word. Because it can split across pages, set `keepLines=true` (and keepNext when it must glue to
the following text).

```bash
officecli add "$FILE" /body --type paragraph \
  --prop text="要点：这里是最关键的结论" \
  --prop shd=solid;FFF4E5 \        # VAL;FILL — a light warm fill
  --prop pbdr.all=single;1pt;E6A23C \ # STYLE;SIZE;COLOR — spans all 4 edges
  --prop keepLines=true
```

(Border size is in eighths-of-a-point or `1pt`/`2pt` units; `single;1pt;E6A23C` is a clean 1pt edge.)

Reset the card's "inside" rhythm with `spaceBefore`/`spaceAfter` so it doesn't hug neighbors.

### 2. Dark accent card / callout (inverted white text)

```bash
officecli add "$FILE" /body --type paragraph \
  --prop text="重要提醒" \
  --prop shd=solid;1F4E79 \        # dark blue fill
  --prop color=FFFFFF --prop size=14 --prop bold=true \
  --prop pbdr.all=single;2pt;1F4E79 \
  --prop keepLines=true
```

### 3. Kicker + big title pair (magazine)

Kicker = small accent line above the title. **The kicker must set `keepNext=true`** so it cannot be
orphaned away from the title (rule above).

```bash
officecli add "$FILE" /body --type paragraph \
  --prop text="PART 01 · 开篇" \
  --prop color=E6A23C --prop size=11 --prop caps=true --prop spaceAfter=0 --prop keepNext=true
officecli add "$FILE" /body --type paragraph \
  --prop text="把复杂讲成设计感" \
  --prop style=Heading1 --prop size=24
```

### 4. Section color-band (full-width horizontal rule / band)

A thin paragraph with a strong fill reads as a section divider. Use `spaceBefore` to separate it,
and `pageBreakBefore` on the section's first heading when you want each part on its own page.

```bash
officecli add "$FILE" /body --type paragraph \
  --prop text="" \
  --prop shd=solid;E6A23C --prop spaceBefore=240 --prop spaceAfter=0
```

(For a full-bleed band that must never splinter, treat it as a 1-col × 1-row table with the cell
`fill` set and the row `cantSplit=true` — see Recipe 6 for the table-card pattern.)

### 5. Pull quote

```bash
officecli add "$FILE" /body --type paragraph \
  --prop text="“设计感来自克制，而不是加法。”" \
  --prop italic=true --prop color=1F4E79 --prop size=16 \
  --prop pbdr.left=single;12pt;1F4E79 --prop spacing=240 --prop align=center
```

### 6. KPI / cheat-sheet table (color-band header)

A table where the header row carries the accent fill and white bold text; body rows stay clean.
**Multi-page tables repeat the header automatically by default; a one-cell "card" table needs
`cantSplit=true` on the row so it never tears.**

```bash
# Header row on an accent band, KPI cells below. Accent fill is set per CELL
# (tcPr), text per cell; the header row gets white bold on the accent fill.
officecli batch "$FILE" --commands '[
  {"command":"add","parent":"/body","type":"table","props":{"cols":"3","rows":"2"}},
  {"command":"set","path":"/body/t[1]/tr[1]/tc[1]","props":{"shd":"solid;1F4E79","text":"指标","bold":"true","color":"FFFFFF"}},
  {"command":"set","path":"/body/t[1]/tr[1]/tc[2]","props":{"shd":"solid;1F4E79","text":"数值","bold":"true","color":"FFFFFF"}},
  {"command":"set","path":"/body/t[1]/tr[1]/tc[3]","props":{"shd":"solid;1F4E79","text":"同比","bold":"true","color":"FFFFFF"}},
  {"command":"set","path":"/body/t[1]/tr[2]/tc[1]","props":{"text":"活跃用户"}}
]'
```

(A single-cell "card" table that must never split: add the row with `cantSplit=true` so it never tears
across a page.)

---

## Chinese Typography (中文排版)

- **Font fallback**: set an East Asia face. The tool's `view issues` reports `font_fallback_ea` when a
  CJK run has no eastAsia face and none is in `docDefaults`. Fix with the compound preset:

```bash
officecli set "$FILE" /body/p[N] --prop typography.preset=zh-body
# → SimSun/宋体 eastAsia on every run + mark, lineRule=atLeast single, spaceAfter=0
```

- **首行缩进 2 字符** (CJK first-line indent): use the char-relative rule, not a hard twip, so it
  doesn't drift with font size:

```bash
officecli set "$FILE" /body/p[N] --prop typography.preset=zh-first-indent   # w:firstLineChars=200 (2 chars)
```

- **禁则 (kinsoku/kicker)**: keep `kinsoku=true` (set on doc settings) so line-start closing punct is
  suppressed; the tool flags `kinsoku_violation` in `view issues`.
- **行距**: prefer `lineSpacing.preset=compact|normal|relaxed` or explicit `lineRule=atLeast` for CJK
  body so lines are never compressed below one full ascent.
- **pbdr multi-edge**: on a paragraph, author each edge separately (`pbdr.top=… / pbdr.bottom=…`) if
  you only want some sides; `pbdr.all` stretches to all four. The style-level multi-edge shorthand is
  not the same spelling — when writing a *paragraph* border, prefer the split `pbdr.<edge>=` form.

---

## Power layout (已内置，别再手搓)

这些能力工具**原生支持**，只是旧 skill 没教。设计文档优先用它们，而不是散落的硬编码 hex / 空段模拟。

### 主题一次换色（替代全文硬编码 hex）

`create` 的 docx 自带默认主题 part。换一套配色 = 改 theme 槽位，所有引用该槽的段落/表格跟着变：

```bash
officecli set "$FILE" / --prop theme.color.accent1=E6A23C    # 改主强调色
officecli set "$FILE" / --prop theme.color.accent2=1F4E79    # 改次强调色
officecli set "$FILE" / --prop theme.font.major.eastAsia=SimHei
officecli set "$FILE" / --prop theme.font.minor.latin=Georgia
```

可用槽位：`dk1 lt1 dk2 lt2 accent1..accent6 hlink folHlink`（颜色），`theme.font.major/minor.{latin,eastasia}`（字体）。
`get --json` 可回读 `theme.color.accent1` 等。段落显式 hex 优先于主题，因此**尽量让颜色来自主题槽位**而非逐段写死。

### 跨页表头重复 + 整行不裂（长表格）

```bash
officecli set "$FILE" /body/tbl[1]/tr[1] --prop header=true      # 每页重复表头行
officecli set "$FILE" /body/tbl[1]/tr[1] --prop cantSplit=true   # 该行不被拆到两页
```

### 竖排（段 / 单元格 / textbox）

```bash
officecli add "$FILE" /body --type section --prop textDirection=tbRl   # 整节竖排（tbRl = 竖排从右到左）
# 表格单元格 / textbox 内：--prop textDirection=<tbRl|lrTb|tbLr>
```

### 多栏（杂志排版）

```bash
officecli add "$FILE" /body --type section --prop columns=2          # 两栏
officecli set  "$FILE" /body/sectPr --prop columns.count=2 --prop columns.space=360
# 可选：columns.equalWidth=true/false、columns.separator=true（栏间竖线）、colWidths=4cm,3cm
```

### 文字环绕与锚定（浮动 textbox/shape）

`add type=shape`（或 `textbox`）产出 `<wp:anchor>` 浮动图形，可设绝对位置与环绕：

```bash
officecli add "$FILE" /body --type shape \
  --prop text="边栏" --prop wrap=square        # square|tight|topbottom|behind|infront|none
  --prop positionH=right --prop positionV=top   # 页边距锚定
  --prop behindDoc=true                          # 置于文字之后（水印/背景块）
```

`wrap=square/tight` 实现图文混排（文字绕排）；`behind`/`infront` 做叠层。

### 表格 → 图表（sourceTable 糖）

从已存在表格一键生成图表，免手写 `data=Series:1,2,3`：

```bash
officecli add "$FILE" /body --type table \
  --prop data="Quarter,Revenue;Q1,100;Q2,150;Q3,200;Q4,250" --prop header=true
officecli add "$FILE" /body --type chart \
  --prop sourceTable=/body/tbl[1] --prop chartType=column --prop title="季度营收"
```

约定：首行为表头 → 每列一条 series（列名作 series 名），首列数据行作 categories；
全数值表格（无表头）→ 各列名为 `Series N`。`data=`/`seriesN=` 显式数据优先于 `sourceTable`。

---

## Performance / long-document build (dump → batch → replay)

A single command line has a length cap, so a long doc is built in chunks. **Batch default is ATOMIC**:
one failed item discards the whole chunk (the file is untouched, your next chunk's indices still point
at the pre-chunk world). For step-by-step build that keeps successes and marks failures, use
`--best-effort --stop-on-error=false` (the autocommit mode):

```bash
officecli batch "$FILE" --input ops.json --best-effort --stop-on-error=false
# failing items are flagged with their index; earlier successes are retained (partialRetained)
```

Run `--dry-run` first against a throwaway copy to see which steps fail before committing the real
chunk. Reliable replay idiom: `close -> rm -> create -> batch -> close`.

---

## Delivery QA checklist

- `view "$FILE" issues --type format` — clean (no kinsoku/font_fallback/pagination warnings).
- `view "$FILE" outline` — Title → H1 → H2 hierarchy, not a flat list.
- Per-chapter `screenshot --grid` contact sheets — no torn cards, no kicker/title split, no double
  blank pages.
- `view "$FILE" html` — no `$xxx$` / `{var}` / `\t` / `\n` literals, no overflowed cells.