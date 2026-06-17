# Markdown Resume → Apple Pages

将 Markdown 简历转换为精美的 Word `.docx` 和 Apple Pages `.pages` 简历。

## 快速使用

```bash
# 1. 复制构建脚本到工作目录
cp /Users/mqs/.pi/agent/skills/md-resume-to-pages/scripts/build_pages_style_resume.py ./

# 2. 编辑脚本顶部的常量（路径、字号等）
#    SOURCE = Path("简历.md")
#    OUTPUT = Path("简历.docx")
#    PHOTO = Path("photo.png")

# 3. 生成 Word 文件
python3 build_pages_style_resume.py

# 4. 导出为 Pages
osascript /Users/mqs/.pi/agent/skills/md-resume-to-pages/scripts/export_to_pages09.scpt \
  /路径/到/简历.docx \
  /路径/到/简历.pages

# 5. 修复空白页（重要！）
python3 build_pages_style_resume.py fix-pages 简历.pages

# 6. 缩略图检查（可选）
unzip -p 简历.pages QuickLook/Thumbnail.png > /private/tmp/resume-thumb.png
open /private/tmp/resume-thumb.png
```

## 自定义

编辑脚本顶部的常量即可快速适配不同简历：

| 常量 | 说明 | 默认值 |
|---|---|---|
| `SOURCE` | 源 Markdown 文件 | `简历.md` |
| `OUTPUT` | 输出的 Word 文件 | `简历.docx` |
| `PHOTO` | 证件照路径 | `photo.png` |
| `BASE_SIZE` | 正文字号（pt） | `9.2` |
| `HEADER_FONT_SIZE` | 顶部身份区字号 | `9` |
| `SECTION_TITLE_SIZE` | 章节标题字号 | `10` |

## Markdown 简历格式要求

```markdown
# 个人简历

| 项目 | 信息 |
|---|---|
| 姓名 | XXX  |
| 手机 | 138... |
| 工作年限 | 12年 |
| 求职意向 | Java 高级开发工程师 |
| 城市 | 成都 |

## 个人优势
(纯文本段落)

## 专业技能
**后端开发：** Java、Spring Boot...   ← 标签：值 格式

## 工作经历
### 公司名称
**职位：** ...
**时间：** ...
**主要工作内容：**
- 职责1
- 职责2

## 项目经历
### 项目名称
**所属公司：** ...
**技术栈：** ...
**项目描述：** ...
**主要职责：** ...
**项目成果：** ...

## 教育经历
| 学校 | 学历 | 专业 | 时间 |
|---|---|---|---|

## 自我评价
(纯文本段落)
```

## 字段映射

脚本通过 `PROFILE_FIELD_ALIASES` 自动匹配简历表头的各种常见写法。
如果简历使用了新的字段名，在脚本中对应的列表里追加即可。

## 目录结构

```
md-resume-to-pages/
├── README.md                    ← 本文件（使用说明）
├── SKILL.md                     ← 技能描述（给 pi 用）
├── references/
│   └── layout.md                ← 布局参考文档
├── scripts/
│   ├── build_pages_style_resume.py   ← 构建脚本（主模板）
│   └── export_to_pages09.scpt        ← Pages 导出 AppleScript
└── agents/
    └── openai.yaml              ← 技能注册信息
```
