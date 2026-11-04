# fetch2word

把**微信公众号推文批量转成统一版式的 Word 文档**。

给它一份链接清单，它抓取正文、下载图片、套用内置模板排版，输出 `yyyy.mm.dd(标题).docx`。装饰性 gif、低清小图标、空行全部清理干净，图片内嵌不留外链。

模板已内置，开箱即用。输出文件名固定为 `yyyy.mm.dd(Title).docx`。

## 特性

- **开箱即用**：模板随仓库分发（已脱敏），只需提供链接清单
- **版式稳定**：页边距、字号、加粗、图片版心宽度全部从模板读取，代码不写死数值
- **图片内嵌**：按原顺序插入，带重试与原图回退，丢图会报警
- **干净输出**：剔除 gif 装饰、低清图标、空段落；回收模板残留部件，体积通常减少 60% 以上
- **可移植**：零绝对路径，换机器只要保持目录结构
- **可核查**：产出附带 `_manifest.csv`，另有自检脚本一键验证合规性

## 安装

```bash
git clone <your-repo-url> fetch2word
```

把 `skills/fetch2word/` 整个目录复制到你的 Skill 目录：

| 作用域 | 目标路径 |
|---|---|
| 项目级 | `<项目>/.workbuddy/skills/fetch2word/` |
| 用户级（全局） | `~/.workbuddy/skills/fetch2word/` |

安装依赖：

```bash
python -m pip install python-docx beautifulsoup4 lxml pillow
```

## 使用

项目里只需要一个链接清单：

```
你的项目/
├── 链接/10月.txt          # 一行一个 mp.weixin.qq.com 链接，可含「排序：从旧到新」
└── 产物/10月/              # 输出
```

```bash
python .workbuddy/skills/fetch2word/scripts/wx2docx.py \
  --links "链接/10月.txt" \
  --out "产物/10月"
```

跑完自检：

```bash
python .workbuddy/skills/fetch2word/scripts/check_output.py --out "产物/10月"
```

要求 `gif=0`、`empty=0`、`dangling=0`、无低清图、首段为日期行；退出码非 0 说明有问题。

### 常用参数

| 参数 | 默认 | 说明 |
|---|---|---|
| `--column` | `学院动态` | 标题无栏目前缀时补的前缀，`""` 不补 |
| `--date-note` | `（发出日期）` | 日期后紧跟的标注，`""` 只留纯日期 |
| `--sort` | `auto` | 读清单里的排序提示；也可 `asc`/`desc`/`none` |
| `--keep-gif` | 关 | 保留 gif |
| `--min-px` | 200 | 最长边低于此值视为装饰图标 |
| `--min-area` | 40000 | 像素面积低于此值视为低清装饰图 |
| `--workers` | 8 | 并发数 |
| `--cache` | 系统临时目录 | 保留中间产物，便于复跑排错 |
| `--template` | 内置 | 特殊情况覆盖版式 |

## 输出命名（固定）

```
文件名：yyyy.mm.dd(Title).docx
```

| 部分 | 取值 |
|---|---|
| `yyyy.mm.dd` | 文章发布日期，取自页面 `createTime`，不是运行日期 |
| `Title` | 文章标题去掉栏目前缀（`公示||xxx` → `xxx`），半角括号包裹 |

标题中的 `\ / : * ? " < > |` 会被剔除，超过 60 字截断。规则写死在 `safe_name()`，不提供参数开关；`check_output.py` 会按 `^\d{4}\.\d{2}\.\d{2}\(.+\)\.docx$` 校验，不匹配即判不合规。

## 关于内置模板

`skills/fetch2word/assets/template.docx` 是脱敏版模板：

| 项目 | 处理方式 |
|---|---|
| 机构名 | 「某职业学院 / 某某学院」等泛化名称 |
| 人名 | 统一替换为「张某某」「李某某」式 |
| 地点 | 「某报告厅」等泛化表述 |
| 配图 | 替换为同尺寸灰阶占位图（10 张，像素尺寸与原件逐一对应） |

**版式零改动**：页面尺寸、页边距、段落数（27）、段落格式、字号字重、图片位置与显示尺寸，逐项与原始模板比对一致。改动前后的差异仅限文字内容与图片像素。

需要换成别的版式时，用 `--template <path>` 指向自己的 docx。

## 仓库结构

```
fetch2word/
├── README.md
└── skills/
    └── fetch2word/
        ├── SKILL.md                    # 核心定义：场景、参数、硬规则
        ├── assets/
        │   └── template.docx           # 内置脱敏模板
        ├── references/
        │   ├── ooxml-notes.md          # 模板还原与 docx 操作陷阱
        │   └── troubleshooting.md      # 抓取要点与排错对照
        └── scripts/
            ├── wx2docx.py              # 主流程
            └── check_output.py         # 交付前自检
```

## 工作原理

1. 读清单、按出现顺序去重，识别排序提示
2. 并发抓页面（`urllib` 优先，`curl` 兜底）
3. 解析 `og:title`、`var createTime`、`#js_content`
4. 并发下载图片（带 Referer，失败重试并回退原图）
5. 以模板为基底重建正文，逐段插入文字与图片
6. 回收孤儿部件、删空段，写 `_manifest.csv`

## 已知限制

- 只支持公众号图文页；音频、视频、投票等互动组件不会转换
- 尾部装饰图偶有永久失效，会在报告中标出，需人工决定是否忽略
- 依赖 Word 文档结构（OOXML），不处理 `.doc`、`.wps`

## 许可

未附许可文件。若要公开分发，建议自行添加（如 MIT）。
