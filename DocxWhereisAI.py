"""
DOCX 元数据查看与清洗器
======================
检查内容：
1. docProps/core.xml   -> 作者、最后修改人、创建/修改时间
2. docProps/app.xml    -> 应用版本、修订次数、总编辑时长、字数等
3. word/settings.xml   -> w:rsid 修订会话ID
4. word/document.xml   -> 修订记录残留 (w:ins / w:del)
5. word/comments.xml   -> 批注信息
6. docProps/custom.xml -> 自定义属性
7. ZIP 内部文件时间戳分布
8. word/styles.xml     -> 样式异常检测
9. word/document.xml   -> 语言标记 w:lang 分布
10. 实际字数统计与 app.xml 记录对比
11. word/media/ 中图片的 EXIF 信息

当前阶段：
- tkinter GUI 文件选择与分析报告展示
- 修改/清理功能保留入口，后续阶段实现
"""

import io
import zipfile
import re
import tempfile
import tkinter as tk
from contextlib import redirect_stdout
from pathlib import Path
from datetime import datetime, timezone
from collections import Counter
from tkinter import filedialog, messagebox
from tkinter.scrolledtext import ScrolledText
import xml.etree.ElementTree as ET

NS_CP = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
NS_DC = "http://purl.org/dc/elements/1.1/"
NS_DCTERMS = "http://purl.org/dc/terms/"
NS_EP = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS_VT = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
NS_CUSTOM = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"

NS = {
    "cp": NS_CP,
    "dc": NS_DC,
    "dcterms": NS_DCTERMS,
    "ep": NS_EP,
    "w": NS_W,
    "vt": NS_VT,
    "custom": NS_CUSTOM,
}


def section(title):
    print("\n" + "=" * 70)
    print(title)
    print("=" * 70)


def check_core_properties(z):
    section("1. 核心属性 (docProps/core.xml)")
    try:
        data = z.read("docProps/core.xml")
    except KeyError:
        print("  未找到 core.xml")
        return
    root = ET.fromstring(data)

    fields = {
        "创建者 (dc:creator)": ".//dc:creator",
        "最后修改人 (cp:lastModifiedBy)": ".//cp:lastModifiedBy",
        "创建时间 (dcterms:created)": ".//dcterms:created",
        "修改时间 (dcterms:modified)": ".//dcterms:modified",
        "标题 (dc:title)": ".//dc:title",
        "主题 (dc:subject)": ".//dc:subject",
        "关键词 (cp:keywords)": ".//cp:keywords",
        "类别 (cp:category)": ".//cp:category",
        "版本 (cp:version)": ".//cp:version",
    }
    for label, xpath in fields.items():
        el = root.find(xpath, NS)
        val = el.text if el is not None else None
        print(f"  {label}: {val}")

    creator_el = root.find(".//dc:creator", NS)
    modifier_el = root.find(".//cp:lastModifiedBy", NS)
    if creator_el is not None and modifier_el is not None:
        if creator_el.text and modifier_el.text and creator_el.text != modifier_el.text:
            print(
                f"\n  ⚠ 提示：创建者 ({creator_el.text}) 与最后修改人 "
                f"({modifier_el.text}) 不一致。"
            )


def check_app_properties(z):
    section("2. 应用属性 (docProps/app.xml)")
    try:
        data = z.read("docProps/app.xml")
    except KeyError:
        print("  未找到 app.xml")
        return
    root = ET.fromstring(data)

    fields = {
        "应用程序 (Application)": ".//ep:Application",
        "应用版本 (AppVersion)": ".//ep:AppVersion",
        "公司 (Company)": ".//ep:Company",
        "修订次数 (cp:revision)": ".//cp:revision",
        "总编辑时长 (TotalTime, 分钟)": ".//ep:TotalTime",
        "字数 (Words)": ".//ep:Words",
        "字符数 (Characters)": ".//ep:Characters",
        "段落数 (Paragraphs)": ".//ep:Paragraphs",
        "页数 (Pages)": ".//ep:Pages",
        "行数 (Lines)": ".//ep:Lines",
    }
    for label, xpath in fields.items():
        el = root.find(xpath, NS)
        val = el.text if el is not None else None
        print(f"  {label}: {val}")

    total_time_el = root.find(".//ep:TotalTime", NS)
    if total_time_el is not None and total_time_el.text:
        try:
            minutes = int(total_time_el.text)
            print(f"\n  换算：总编辑时长约 {minutes} 分钟 ≈ {minutes/60:.1f} 小时")
            print("  （注意：该值可能因文档另存或跨设备编辑而不准确，仅供参考）")
        except ValueError:
            pass


def check_rsid(z):
    section("3. 修订会话ID (rsid) 分析")
    try:
        data = z.read("word/settings.xml")
    except KeyError:
        print("  未找到 settings.xml")
        return
    root = ET.fromstring(data)

    rsids_root = root.find(".//w:rsids", NS)
    if rsids_root is not None:
        rsid_elements = rsids_root.findall("w:rsid", NS)
        total_rsids = len(rsid_elements)
        print(f"  settings.xml 中 <w:rsids> 内声明的 RSID 总数: {total_rsids}")
        if total_rsids <= 20:
            for elem in rsid_elements:
                rid = elem.get(f"{{{NS_W}}}val")
                print(f"    - {rid}")
        else:
            print("  (数量较多，不逐一列出)")
    else:
        print("  未找到 <w:rsids> 节点")
        total_rsids = 0

    try:
        doc_data = z.read("word/document.xml")
    except KeyError:
        doc_data = None
    if doc_data:
        doc_root = ET.fromstring(doc_data)
        rsid_attrs = set()
        for elem in doc_root.iter():
            for attr_name in (
                f"{{{NS_W}}}rsidR",
                f"{{{NS_W}}}rsidRPr",
                f"{{{NS_W}}}rsidRDefault",
                f"{{{NS_W}}}rsidP",
            ):
                val = elem.get(attr_name)
                if val:
                    rsid_attrs.add(val)
        print(f"  正文中出现的独立 RSID 值数量: {len(rsid_attrs)}")
        if len(rsid_attrs) > 0:
            if len(rsid_attrs) <= 15:
                print("  列举如下:")
                for val in sorted(rsid_attrs):
                    print(f"    - {val}")
            else:
                print("  (数量较多，不逐一列出)")
        print(
            "\n  解读：RSID 数量大致反映文档经历过的独立编辑会话次数。"
            "\n  如果数字极小（如 1~2），可能表明文档在极短时间内一次性生成；"
            "\n  但另存为/复制内容也可能重置或改变 RSID，不能作为唯一依据。"
        )
    else:
        print("  无法读取 document.xml")


def check_track_changes(z):
    section("4. 修订记录残留 (Track Changes)")
    try:
        doc_data = z.read("word/document.xml")
    except KeyError:
        print("  无法读取 document.xml")
        return
    root = ET.fromstring(doc_data)
    inserts = root.findall(f".//{{{NS_W}}}ins")
    deletes = root.findall(f".//{{{NS_W}}}del")
    total = len(inserts) + len(deletes)
    if total == 0:
        print("  未发现 <w:ins> 或 <w:del> 节点，无修订记录残留。")
        return

    print(f"  发现 {len(inserts)} 处插入标记, {len(deletes)} 处删除标记。")
    authors = Counter()
    dates = set()
    for tag, elem in (("插入", inserts), ("删除", deletes)):
        for el in elem:
            author = el.get(f"{{{NS_W}}}author")
            date = el.get(f"{{{NS_W}}}date")
            if author:
                authors[author] += 1
            if date:
                dates.add(date)
    if authors:
        print("  修订作者分布:")
        for author, count in authors.most_common():
            print(f"    {author}: {count} 处")
    if dates:
        print("  涉及的修订时间（去重）:")
        for d in sorted(dates):
            print(f"    {d}")
    print("  ⚠ 注意：如果文档未在审阅模式下接受所有修订，这些信息会直接暴露原始编辑者身份。")


def check_comments(z):
    section("5. 批注 (Comments) 信息")
    try:
        data = z.read("word/comments.xml")
    except KeyError:
        print("  未找到 comments.xml，无批注。")
        return
    root = ET.fromstring(data)
    comments = root.findall(f"{{{NS_W}}}comment")
    if not comments:
        print("  批注文件存在但无批注条目。")
        return
    print(f"  批注总数: {len(comments)}")
    for i, comment in enumerate(comments, 1):
        author = comment.get(f"{{{NS_W}}}author")
        date = comment.get(f"{{{NS_W}}}date")
        initials = comment.get(f"{{{NS_W}}}initials")
        texts = []
        for t in comment.iter(f"{{{NS_W}}}t"):
            if t.text:
                texts.append(t.text)
        content = "".join(texts)
        content_preview = content[:80] + ("..." if len(content) > 80 else "")
        print(f"\n  [{i}] 作者: {author}  日期: {date}  缩写: {initials}")
        print(f"      内容: {content_preview}")


def check_custom_properties(z):
    section("6. 自定义属性 (docProps/custom.xml)")
    try:
        data = z.read("docProps/custom.xml")
    except KeyError:
        print("  未找到 custom.xml")
        return
    root = ET.fromstring(data)
    props = root.findall(f"{{{NS_CUSTOM}}}property")
    if not props:
        print("  无自定义属性。")
        return
    print(f"  发现 {len(props)} 个自定义属性：")
    for prop in props:
        name = prop.get("name")
        fmtid = prop.get("fmtid")
        pid = prop.get("pid")
        value = None
        for child in prop:
            value = child.text
            break
        print(f"    - {name} = {value}  (fmtid={fmtid}, pid={pid})")


def check_zip_timestamps(z):
    section("7. ZIP 内部文件时间戳分布")
    infos = z.infolist()
    timestamps = []
    print(f"  {'文件名':<50} {'内部时间戳'}")
    print("  " + "-" * 65)
    for info in sorted(infos, key=lambda i: i.filename):
        ts = datetime(*info.date_time)
        timestamps.append(ts)
        if any(
            key in info.filename
            for key in [
                "core.xml",
                "app.xml",
                "settings.xml",
                "document.xml",
                "styles.xml",
                "comments",
                "header",
                "footer",
            ]
        ):
            print(f"  {info.filename:<50} {ts}")
    unique_ts = set(timestamps)
    print(f"\n  共 {len(timestamps)} 个内部文件，时间戳唯一值数量: {len(unique_ts)}")
    if len(unique_ts) <= 1:
        print("  ⚠ 提示：所有内部文件时间戳完全相同，常见于一次性生成/重新打包。")
    else:
        print("  时间戳存在差异，符合人工编辑的渐进特征。")


def check_styles_consistency(z):
    section("8. 样式表一致性 (styles.xml)")
    try:
        data = z.read("word/styles.xml")
    except KeyError:
        print("  未找到 styles.xml")
        return
    root = ET.fromstring(data)
    style_ids = []
    for style in root.findall(f"{{{NS_W}}}style"):
        sid = style.get(f"{{{NS_W}}}styleId")
        if sid:
            style_ids.append(sid)

    suspicious = [sid for sid in style_ids if re.search(r"[A-Za-z]+\d+$", sid)]
    if suspicious:
        print(f"  检测到 {len(suspicious)} 个可能因跨文档复制粘贴产生的异常样式ID：")
        for s in suspicious:
            print(f"    - {s}")
        print("  （这些带数字后缀的样式通常是粘贴内容时从其他文档带入的，可提示内容来源不单一）")
    else:
        print("  未发现明显异常样式ID（无数字后缀）。")
    print(f"  样式总数: {len(style_ids)}")


def check_lang_distribution(z):
    section("9. 语言标记 (w:lang) 分布")
    try:
        doc_data = z.read("word/document.xml")
    except KeyError:
        print("  无法读取 document.xml")
        return
    root = ET.fromstring(doc_data)
    langs = Counter()
    for lang_elem in root.iter(f"{{{NS_W}}}lang"):
        val = lang_elem.get(f"{{{NS_W}}}val")
        if val:
            langs[val] += 1
    if not langs:
        print("  未发现 <w:lang> 标记。")
        return
    print(f"  检测到以下语言标记（共 {len(langs)} 种）：")
    for lang, cnt in langs.most_common():
        print(f"    {lang}: {cnt} 次")
    if len(langs) > 1:
        print(
            "\n  ⚠ 提示：文档中混用了多种语言校对标记，可能是跨地区复制文本的痕迹。"
        )


def check_word_count(z):
    section("10. 字数统计交叉核对")
    try:
        app_data = z.read("docProps/app.xml")
        app_root = ET.fromstring(app_data)
        app_words = app_root.findtext(".//ep:Words", namespaces=NS)
        app_chars = app_root.findtext(".//ep:Characters", namespaces=NS)
    except Exception:
        print("  无法读取 app.xml 中的字数信息")
        return

    try:
        doc_data = z.read("word/document.xml")
    except KeyError:
        print("  无法读取 document.xml")
        return
    doc_root = ET.fromstring(doc_data)
    texts = [t.text for t in doc_root.iter(f"{{{NS_W}}}t") if t.text]
    full_text = "".join(texts)
    actual_chars = len(full_text)
    actual_words = len(full_text.split())

    print(f"  实际统计文本字符数 (含空格): {actual_chars}")
    print(f"  实际统计词数 (按空白分割): {actual_words}")
    print(f"  ---")
    print(f"  app.xml 记录 Words: {app_words}")
    print(f"  app.xml 记录 Characters: {app_chars}")
    if app_words:
        try:
            app_w = int(app_words)
            diff = abs(actual_words - app_w)
            print(f"  词数差异: {diff} (实际 {actual_words}, 记录 {app_w})")
            if diff > actual_words * 0.1:
                print("  ⚠ 词数存在较大差异，文档可能被裁剪或统计未更新。")
        except ValueError:
            pass
    if app_chars:
        try:
            app_c = int(app_chars)
            diff = abs(actual_chars - app_c)
            print(f"  字符数差异: {diff} (实际 {actual_chars}, 记录 {app_c})")
            if diff > 50:
                print("  ⚠ 字符数存在较大差异。")
        except ValueError:
            pass


def check_embedded_media_exif(z, tmp_dir):
    section("11. 嵌入图片 EXIF 信息")
    try:
        from PIL import Image
        from PIL.ExifTags import TAGS
    except ImportError:
        print("  未安装 Pillow，跳过 EXIF 检查。")
        return
    media_files = [n for n in z.namelist() if n.startswith("word/media/")]
    if not media_files:
        print("  未发现嵌入媒体文件。")
        return
    print(f"  共发现 {len(media_files)} 个媒体文件，检查 EXIF 中...\n")
    found_any = False
    for name in media_files:
        if not name.lower().endswith((".jpg", ".jpeg", ".png", ".tiff")):
            continue
        try:
            data = z.read(name)
            tmp_path = tmp_dir / Path(name).name
            tmp_path.write_bytes(data)
            img = Image.open(tmp_path)
            exif_data = img._getexif() if hasattr(img, "_getexif") else None
            if exif_data:
                found_any = True
                print(f"  文件: {name}")
                for tag_id, value in exif_data.items():
                    tag = TAGS.get(tag_id, tag_id)
                    if tag in ("Make", "Model", "DateTime", "Software", "GPSInfo",
                               "Artist", "Copyright"):
                        print(f"      {tag}: {value}")
                print()
        except Exception:
            continue
    if not found_any:
        print("  未在图片中发现可读 EXIF 信息。")


# ===== 阶段2完成：报告展示 =====
def _run_check(check_name, check_func, z, *args):
    try:
        check_func(z, *args)
    except Exception as exc:
        print(f"\n  ⚠ {check_name} 检查失败: {exc}")


def analyze_docx(path: Path) -> str:
    path = Path(path)
    output = io.StringIO()

    with redirect_stdout(output):
        print(f"正在分析文件: {path}")

        if not path.exists():
            print(f"错误：文件不存在: {path}")
            return output.getvalue()

        try:
            print(f"文件大小: {path.stat().st_size / 1024:.1f} KB")
        except OSError as exc:
            print(f"错误：无法读取文件信息: {exc}")
            return output.getvalue()

        try:
            with zipfile.ZipFile(path, "r") as z:
                names = set(z.namelist())
                required_parts = ["[Content_Types].xml", "word/document.xml"]
                missing_parts = [part for part in required_parts if part not in names]
                if missing_parts:
                    print(
                        "错误：文件不是合法 DOCX，缺少必要结构: "
                        + ", ".join(missing_parts)
                    )
                    return output.getvalue()

                checks = [
                    ("核心属性", check_core_properties, ()),
                    ("应用属性", check_app_properties, ()),
                    ("修订会话ID", check_rsid, ()),
                    ("修订记录残留", check_track_changes, ()),
                    ("批注信息", check_comments, ()),
                    ("自定义属性", check_custom_properties, ()),
                    ("ZIP 内部文件时间戳", check_zip_timestamps, ()),
                    ("样式表一致性", check_styles_consistency, ()),
                    ("语言标记分布", check_lang_distribution, ()),
                    ("字数统计交叉核对", check_word_count, ()),
                ]
                for check_name, check_func, args in checks:
                    _run_check(check_name, check_func, z, *args)

                with tempfile.TemporaryDirectory() as tmp_name:
                    _run_check(
                        "嵌入图片 EXIF 信息",
                        check_embedded_media_exif,
                        z,
                        Path(tmp_name),
                    )

                section("检查完毕")
                print("以上信息仅供自查参考，不构成任何结论性证据。\n")
        except zipfile.BadZipFile:
            print("错误：文件不是合法 ZIP/DOCX，无法打开。")
        except OSError as exc:
            print(f"错误：无法读取文件: {exc}")
        except Exception as exc:
            print(f"错误：分析过程中发生异常: {exc}")

    return output.getvalue()


def modify_timestamps(original_path):
    """
    修改时间戳：
    1. core.xml 的创建/修改时间改为当前 UTC 时间
    2. 所有 ZIP 条目的时间改为当前时间
    生成新文件：原文件名 + '_modified' 后缀
    """
    path = Path(original_path)
    new_path = path.with_stem(path.stem + "_modified")
    now = datetime.now(timezone.utc)
    new_time_str = now.strftime("%Y-%m-%dT%H:%M:%SZ")
    # ZIP 内部时间戳（不含时区，使用本地时间或者 UTC）
    new_zip_date = (now.year, now.month, now.day, now.hour, now.minute, now.second)

    # 读取原始文件全部内容到内存
    with zipfile.ZipFile(path, "r") as zin:
        items = {}
        for info in zin.infolist():
            data = zin.read(info.filename)
            items[info.filename] = (info, data)

    # 修改 core.xml 的内容（如果存在）
    core_key = "docProps/core.xml"
    if core_key in items:
        info, data = items[core_key]
        root = ET.fromstring(data)
        # 修改创建时间和修改时间
        for tag, xpath in [("创建时间", ".//dcterms:created"),
                           ("修改时间", ".//dcterms:modified")]:
            el = root.find(xpath, NS)
            if el is not None:
                el.text = new_time_str
        new_data = ET.tostring(root, encoding="UTF-8", xml_declaration=True)
        items[core_key] = (info, new_data)
        print(f"  已修改 core.xml 中的创建/修改时间为 {new_time_str}")
    else:
        print("  未找到 core.xml，跳过修改。")

    # 写入新 zip，并将所有条目时间设为当前时间
    with zipfile.ZipFile(new_path, "w", zipfile.ZIP_DEFLATED) as zout:
        for fname, (info, data) in items.items():
            new_info = zipfile.ZipInfo(filename=fname)
            # 设置新的时间戳
            new_info.date_time = new_zip_date
            zout.writestr(new_info, data)

    print(f"  新文件已生成: {new_path}")
    return new_path


# ===== 阶段1完成：GUI 基础框架 =====
class DocxMetadataApp:
    def __init__(self, root):
        self.root = root
        self.selected_path = None

        self.root.title("DOCX 元数据查看与清洗器")
        self.root.geometry("980x720")

        self._build_widgets()

    def _build_widgets(self):
        toolbar = tk.Frame(self.root, padx=10, pady=10)
        toolbar.pack(fill=tk.X)

        select_button = tk.Button(
            toolbar,
            text="选择 DOCX 文件",
            command=self.select_file,
            width=16,
        )
        select_button.pack(side=tk.LEFT, padx=(0, 8))

        analyze_button = tk.Button(
            toolbar,
            text="分析",
            command=self.analyze_selected_file,
            width=10,
        )
        analyze_button.pack(side=tk.LEFT, padx=(0, 8))

        metadata_button = tk.Button(
            toolbar,
            text="修改元数据",
            command=self.show_future_message,
            width=12,
        )
        metadata_button.pack(side=tk.LEFT, padx=(0, 8))

        clean_button = tk.Button(
            toolbar,
            text="一键清理",
            command=self.show_future_message,
            width=12,
        )
        clean_button.pack(side=tk.LEFT)

        self.report_text = ScrolledText(
            self.root,
            wrap=tk.WORD,
            font=("Menlo", 12),
            padx=10,
            pady=10,
        )
        self.report_text.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))
        self.report_text.configure(state=tk.DISABLED)

        self.status_label = tk.Label(
            self.root,
            text="请选择 DOCX 文件进行分析",
            anchor=tk.W,
            padx=10,
            pady=6,
            relief=tk.SUNKEN,
        )
        self.status_label.pack(fill=tk.X, side=tk.BOTTOM)

    def select_file(self):
        filename = filedialog.askopenfilename(
            title="选择 DOCX 文件",
            filetypes=[
                ("Word 文档", "*.docx"),
                ("所有文件", "*.*"),
            ],
        )
        if not filename:
            return

        self.selected_path = Path(filename)
        self.status_label.configure(text=f"当前文件: {self.selected_path}")

    def analyze_selected_file(self):
        if self.selected_path is None:
            messagebox.showwarning("未选择文件", "请先选择 DOCX 文件。")
            return

        self.status_label.configure(text=f"正在分析: {self.selected_path}")
        self.root.update_idletasks()

        report = analyze_docx(self.selected_path)
        self.set_report(report)

        if self.selected_path.exists():
            self.status_label.configure(text=f"分析完成: {self.selected_path}")
        else:
            self.status_label.configure(text=f"文件不存在: {self.selected_path}")

    def set_report(self, report):
        self.report_text.configure(state=tk.NORMAL)
        self.report_text.delete("1.0", tk.END)
        self.report_text.insert(tk.END, report)
        self.report_text.configure(state=tk.DISABLED)

    def show_future_message(self):
        messagebox.showinfo("后续阶段实现", "该功能将在后续阶段实现。")


def main():
    root = tk.Tk()
    DocxMetadataApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
