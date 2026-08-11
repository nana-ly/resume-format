# resume-format — 简历排版工具

按专属模板格式生成 .docx 简历。独立运行，不依赖特定平台。

## 快速使用

```bash
# 1. 安装依赖
pip install python-docx

# 2. 准备内容 JSON（参照 scripts/resume.example.json 的结构）

# 3. 生成 docx
python scripts/build_resume.py --content 你的内容.json --output 简历.docx
```

## 内容 JSON 结构

```json
{
  "name": "姓名",
  "basic_info": "性别：男/女 | 年龄：xx",
  "contact": "手机号：1xx-xxxx-xxxx | 邮箱：xxx@xxx.com",
  "job_intent": "求职意向：XX岗位",
  "sections": [
    {
      "title": "教育经历",
      "items": [
        { "type": "entry", "left": "XX大学  XX学院", "center": "", "right": "202X.09-202X.06" },
        { "type": "body", "text": "专业 本科（排名）" },
        { "type": "body", "text": "主修课程：xxx" },
        { "type": "body", "text": "荣誉奖项：xxx" },
        { "type": "body", "text": "资格证书：xxx" }
      ]
    },
    {
      "title": "实习经历",
      "items": [
        { "type": "entry", "left": "公司名", "center": "岗位名", "right": "202X.0X-202X.0X" },
        { "type": "label_body", "label": "核心工作：", "text": "工作内容描述..." }
      ]
    },
    {
      "title": "项目经历",
      "items": [
        { "type": "entry", "left": "项目名（技术方向）", "center": "项目角色", "right": "202X.0X-202X.0X" },
        { "type": "label_body", "label": "项目背景：", "text": "背景描述..." },
        { "type": "label_body", "label": "技术栈：", "text": "技术列表..." },
        { "type": "label_body", "label": "代码仓库：", "text": "仓库地址" },
        { "type": "label_body", "label": "主要职责：", "text": "" },
        { "type": "bullet", "text": "职责要点1" },
        { "type": "bullet", "text": "职责要点2" },
        { "type": "label_body", "label": "项目成果：", "text": "" },
        { "type": "bullet", "text": "成果要点1" }
      ]
    },
    {
      "title": "校园经历",
      "items": [
        { "type": "entry", "left": "组织名", "center": "职位", "right": "202X.09-202X.09" },
        { "type": "label_body", "label": "核心工作：", "text": "工作内容..." }
      ]
    },
    {
      "title": "个人技能",
      "items": [
        { "type": "label_body", "label": "1.专业技能：", "text": "技能描述..." },
        { "type": "label_body", "label": "2.综合素质：", "text": "素质描述..." }
      ]
    }
  ]
}
```

### item 类型说明

| type | 说明 | 字段 |
|---|---|---|
| `entry` | 条目行：左名称 + 居中角色 + 右时间 | `left`, `center`(可空), `right` |
| `body` | 普通正文（不加粗） | `text` |
| `label_body` | 标签加粗 + 正文不加粗 | `label`, `text` |
| `bullet` | 要点行（• 前缀，悬挂缩进，不加粗） | `text` |

## 格式规范

详见 `references/format-spec.md`。核心要点：

- 页面：A4，上下 1cm，左右 1.27cm
- 字体：全宋体（联系方式/求职意向西文用 Times New Roman）
- 字号：姓名 14pt / 标题 13pt / 条目 12pt / 正文 10.5pt
- 加粗：姓名、标题、条目名称、所有标签（核心工作/项目背景/技术栈/主要职责/项目成果/专业技能/综合素质）
- 不加粗：基本信息、时间、正文、要点
- 行距：单倍，Normal 样式设 snapToGrid=0
- 要点：• 前缀 + 悬挂缩进

## 给 AI 工具的提示

如果你是 AI 助手（Codex/Claude/ChatGPT 等），使用步骤：
1. 读取 `references/format-spec.md` 了解完整格式规范
2. 读取 `scripts/resume.example.json` 了解内容结构
3. 向用户收集简历内容
4. 按结构生成 JSON 文件
5. 运行 `python scripts/build_resume.py --content 内容.json --output 输出.docx`
6. 把生成的 docx 交付给用户
