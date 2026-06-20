"""
PPTX 元数据 / 编辑痕迹检查脚本
=================================
检查内容：
1. docProps/core.xml   -> 作者、最后修改人、创建/修改时间
2. docProps/app.xml    -> 应用版本、修订次数(revision)、总编辑时长、幻灯片数等
3. zip 内部各文件的时间戳分布 -> 判断是否"一次性打包生成"
4. slideX.xml 中残留的时间戳 / 修订相关属性（PPTX 没有 Word 那种 rsid 体系，
   但部分版本会在 timing/动画节点或 cNvPr 里留一些 id，可作为辅助参考）
5. 嵌入图片的 EXIF 信息（如果图片是截图，通常没有 EXIF；如果是手机拍的，
   可能带设备型号、GPS 等——这与"作业截图来源"无关判断，仅供参考）

用法：
    python pptx_forensics.py
    （脚本内 FILE_PATH 改成你的文件路径）
"""

import zipfile
import shutil
from pathlib import Path
from datetime import datetime
import xml.etree.ElementTree as ET

# ====== 在这里改成你的文件路径 ======
FILE_PATH = "/Users/smter-mac/Downloads/463milestone1-repo/dist/oriental-trading-milestone1-milestone2-submission.pptx"
# ====================================

NS = {
    "cp": "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
    "dc": "http://purl.org/dc/elements/1.1/",
    "dcterms": "http://purl.org/dc/terms/",
    "ep": "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties",
}


def section(title):
    print("\n" + "=" * 70)
    print(title)
    print("=" * 70)


def check_core_properties(z):
    """docProps/core.xml: 作者、修改人、创建/修改时间"""
    section("1. 核心属性 (docProps/core.xml) —— 作者与时间")
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
        "最后修改时间 (dcterms:modified)": ".//dcterms:modified",
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
                f"({modifier_el.text}) 不一致，可能值得留意（也可能只是换了电脑/账号）。"
            )


def check_app_properties(z):
    """docProps/app.xml: 修订次数、总编辑时长、应用版本"""
    section("2. 应用属性 (docProps/app.xml) —— 修订与编辑统计")
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
        "修订次数 (RevisionNumber, 注意：PPTX 通常没有此字段，常见于 docx)": ".//cp:revision",
        "幻灯片数 (Slides)": ".//ep:Slides",
        "总编辑时长 (TotalTime, 单位:分钟)": ".//ep:TotalTime",
    }
    for label, xpath in fields.items():
        el = root.find(xpath, NS)
        val = el.text if el is not None else None
        print(f"  {label}: {val}")

    total_time_el = root.find(".//ep:TotalTime", NS)
    if total_time_el is not None and total_time_el.text:
        minutes = int(total_time_el.text)
        print(f"\n  换算：总编辑时长约 {minutes} 分钟 ≈ {minutes/60:.1f} 小时")
        print("  （注意：PowerPoint 的 TotalTime 经常不准确，文件被多次打开关闭、")
        print("   或跨设备编辑时该值可能被重置或异常偏低/偏高，仅供参考，不能单独作为证据）")


def check_zip_timestamps(z):
    """zip 内部各文件的修改时间分布"""
    section("3. ZIP 内部文件时间戳分布 —— 判断是否'一次性打包'")
    infos = z.infolist()
    timestamps = []
    print(f"  {'文件名':<45} {'内部时间戳'}")
    print("  " + "-" * 65)
    for info in sorted(infos, key=lambda i: i.filename):
        ts = datetime(*info.date_time)
        timestamps.append(ts)
        # 只打印关键文件，避免刷屏
        if any(
            key in info.filename
            for key in ["slide", "core.xml", "app.xml", "presentation.xml"]
        ) and "slideLayouts" not in info.filename and "slideMasters" not in info.filename:
            print(f"  {info.filename:<45} {ts}")

    unique_ts = set(timestamps)
    print(f"\n  共 {len(timestamps)} 个内部文件，时间戳唯一值数量: {len(unique_ts)}")
    if len(unique_ts) <= 1:
        print(
            "  ⚠ 提示：所有内部文件时间戳完全相同（常见于：用某些工具/脚本一次性批量生成，"
            "\n     或文件经过'重新打包'处理，如 LibreOffice 转存、某些 PPT 生成库导出）。"
            "\n     真实手工编辑的 PPT 往往各 slide 时间戳有细微差异（哪怕只差几分钟/几秒），"
            "\n     因为 PowerPoint 保存时通常只更新被修改部分的内部条目时间。"
            "\n     但要注意：很多正常另存为/压缩重打包操作也会导致时间戳统一，这不是决定性证据。"
        )
    else:
        print("  时间戳有差异，呈现出'渐进编辑'特征（更接近正常人工编辑过程）。")


def check_slide_count_vs_zip(z):
    """简单核对 slide 数量"""
    section("4. 幻灯片文件清单")
    slide_files = [n for n in z.namelist() if n.startswith("ppt/slides/slide") and n.endswith(".xml")]
    slide_files_sorted = sorted(
        slide_files, key=lambda x: int("".join(filter(str.isdigit, x.split("/")[-1])) or 0)
    )
    print(f"  共检测到 {len(slide_files_sorted)} 个 slide XML 文件")
    for f in slide_files_sorted:
        print(f"    - {f}")


def check_embedded_media_exif(z, tmp_dir):
    """尝试读取嵌入图片的 EXIF（如果安装了 PIL）"""
    section("5. 嵌入媒体文件检查（图片 EXIF，可选）")
    try:
        from PIL import Image
        from PIL.ExifTags import TAGS
    except ImportError:
        print("  未安装 Pillow，跳过 EXIF 检查。如需启用，运行：")
        print("    pip install Pillow --break-system-packages")
        return

    media_files = [n for n in z.namelist() if n.startswith("ppt/media/")]
    if not media_files:
        print("  未发现嵌入媒体文件。")
        return

    print(f"  共发现 {len(media_files)} 个嵌入媒体文件，检查其中可解析图片的 EXIF：\n")
    found_any_exif = False
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
                found_any_exif = True
                print(f"  文件: {name}")
                for tag_id, value in exif_data.items():
                    tag = TAGS.get(tag_id, tag_id)
                    if tag in ("Make", "Model", "DateTime", "Software", "GPSInfo"):
                        print(f"      {tag}: {value}")
                print()
        except Exception:
            continue
    if not found_any_exif:
        print("  未在图片中发现可读 EXIF 信息（截图类图片通常本来就没有 EXIF，属正常情况）。")


def main():
    path = Path(FILE_PATH)
    if not path.exists():
        print(f"文件不存在: {path}")
        return

    tmp_dir = Path("/tmp/pptx_forensics_tmp")
    if tmp_dir.exists():
        shutil.rmtree(tmp_dir)
    tmp_dir.mkdir()

    print(f"正在分析文件: {path}")
    print(f"文件大小: {path.stat().st_size / 1024:.1f} KB")

    with zipfile.ZipFile(path, "r") as z:
        check_core_properties(z)
        check_app_properties(z)
        check_zip_timestamps(z)
        check_slide_count_vs_zip(z)
        check_embedded_media_exif(z, tmp_dir)

    shutil.rmtree(tmp_dir, ignore_errors=True)

    section("检查完毕")
    print("以上信息仅供自查参考，不构成任何结论性证据。")
    print("如果你是想确认自己的提交文件'干不干净'，重点看：")
    print("  1) 第1节的创建者/修改人是否就是你自己")
    print("  2) 第2节的 TotalTime 是否符合你实际花费的时间数量级")
    print("  3) 第3节的时间戳是否有自然的渐进分布（而非全部完全相同）")


if __name__ == "__main__":
    main()