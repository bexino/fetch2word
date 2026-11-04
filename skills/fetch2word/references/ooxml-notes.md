# 模板还原与 OOXML 操作要点

## 模板结构怎么读

`read_template_spec()` 只依赖模板本身，不假设任何固定数值：

| 项目 | 读法 |
|---|---|
| 版心宽度 | `page_width - left_margin - right_margin` |
| 标题字号 / 加粗 | 模板首个含文字的 run 的 `font.size` / `bold` |
| 图片宽度 | 模板 `inline_shapes` 最大宽度，且不超过版心宽度 |

典型模板：A4（8.27×11.69 in）、左右边距 1.25 in、上下 1 in、正文默认 11pt、段后 160、行距 278 auto、标题 14pt 加粗、正文无首行缩进。这些**不要**写进代码，实际值以模板为准。

## 重建正文的正确姿势

```python
doc = Document(template_path)
clear_body(doc)          # 删 body 全部子节点，唯独保留 w:sectPr
```

`sectPr` 是页面设置的唯一来源，删掉它 Word 会回退默认 Letter + 1 inch 边距，之前所有版式功夫白费。

段落样式**不要手动加 `pPr`**。模板 `pPrDefault` 已经定义好段距行距，裸段落自动继承，这才是"严格按模板"。只有日期行与标题行需要显式 `b` + `sz`，且 `rFonts` 要带 `w:hint="eastAsia"`，否则中文字形可能走错字体。

## 三个必须知道的坑

### 1. 同一张图复用同一 rId

python-docx 对内容相同的图片缓存 `ImagePart`，`relate_to` 返回**已存在的关系**。所以同一张图插入两次 = 两个 `<a:blip>` 指向同一个 `rId`。

后果：若"删段落 + `drop_rel(rid)`"，另一处引用立刻变成悬空引用，Word 打开显示红叉。

正确做法：删完统一按引用集合回收。

```python
used = set(re.findall(r'r:embed="(rId\d+)"', doc.element.body.xml))
for rid, rel in list(doc.part.rels.items()):
    if "image" in rel.reltype and rid not in used:
        doc.part.drop_rel(rid)
```

自检时必须验证 `used - set(doc.part.rels.keys())` 为空。

### 2. 清空正文不会带走模板图片

`Document(template)` 复制来的 media 部件挂在 document part 的 rels 上，删段落不会删关系。这些"孤儿部件"每篇白占 1~2 MB。

按上面的引用集合反查回收后，体积通常能砍掉 60%~90%（实测 40MB → 13MB）。

### 3. 格式判定不要靠后缀

CDN 可能把 gif 存成 `.jpg`。判 gif 读文件头：

```python
with open(path, "rb") as fh:
    is_gif = fh.read(4) == b"GIF8"
```

## 图片尺寸策略

- 宽度：`min(像素宽 × 9525 EMU, 版心宽度)`，即按 96dpi 还原自然尺寸，超宽才缩放
- 装饰图判定：`max(w, h) < MIN_PX` 或 `w * h < MIN_AREA`
- 默认 `MIN_PX=200`、`MIN_AREA=40000`

实测分布（某学院公众号 11 个月样本）：被判定为装饰的是 26×51 ~ 221×45 的图标与分割线；实拍图一律 1080 宽，两者之间有明显断层，阈值不敏感。

## 空段落

公众号正文常见 `<p><br></p>`、`<section>&nbsp;</section>`，解析后会变成一堆空段。两道防线：

1. 解析阶段：`text.strip()` 为空直接丢弃
2. 生成收尾：`strip_empty_paragraphs()` 删除无文字且无 `blip`/`drawing`/`pict` 的段落

注意：图片段落本身 text 为空但有 `drawing`，**不能**被删。

## 图片段落要不要居中

模板原图为左对齐、宽度等于版心宽度，视觉上就是满宽。跟随模板，别自作主张加 `jc=center`。
