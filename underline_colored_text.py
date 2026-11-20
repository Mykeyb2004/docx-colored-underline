#!/usr/bin/env -S uv run --script
# /// script
# requires-python = ">=3.10"
# dependencies = []
# ///
"""Add double underlines (or a custom underline style) to all colored text in a .docx file."""

from __future__ import annotations

import argparse
from collections import Counter
from dataclasses import dataclass
from pathlib import Path
import sys
import xml.dom.minidom as minidom
import zipfile

# Elements that must appear BEFORE <w:u> inside <w:rPr> per ECMA-376 CT_RPr.
BEFORE_U_TAGS = {
    "w:rStyle",
    "w:rFonts",
    "w:b",
    "w:bCs",
    "w:i",
    "w:iCs",
    "w:caps",
    "w:smallCaps",
    "w:strike",
    "w:dstrike",
    "w:outline",
    "w:shadow",
    "w:emboss",
    "w:imprint",
    "w:noProof",
    "w:snapToGrid",
    "w:vanish",
    "w:webHidden",
    "w:color",
    "w:spacing",
    "w:w",
    "w:kern",
    "w:position",
    "w:sz",
    "w:szCs",
    "w:highlight",
}

# Elements that must appear BEFORE <w:color> inside <w:rPr> per ECMA-376 CT_RPr.
BEFORE_COLOR_TAGS = {
    "w:rStyle",
    "w:rFonts",
    "w:b",
    "w:bCs",
    "w:i",
    "w:iCs",
    "w:caps",
    "w:smallCaps",
    "w:strike",
    "w:dstrike",
    "w:outline",
    "w:shadow",
    "w:emboss",
    "w:imprint",
    "w:noProof",
    "w:snapToGrid",
    "w:vanish",
    "w:webHidden",
}

NON_COLORED_THEME_NAMES = {"", "dk1", "tx1", "text1"}

UNDERLINE_CHOICES = (
    "double",
    "single",
    "thick",
    "dotted",
    "dottedHeavy",
    "dash",
    "dashedHeavy",
    "dashLong",
    "dashLongHeavy",
    "dotDash",
    "dashDotHeavy",
    "dotDotDash",
    "dashDotDotHeavy",
    "wave",
    "wavyHeavy",
    "wavyDouble",
    "words",
)


@dataclass(frozen=True)
class ColorSpec:
    val: str
    theme_color: str

    @property
    def label(self) -> str:
        if self.val and self.theme_color:
            return f"#{self.val} ({self.theme_color})"
        if self.val:
            return f"#{self.val}"
        return f"theme:{self.theme_color}"


def direct_children(node: minidom.Element, tag: str) -> list[minidom.Element]:
    """Return direct child elements matching tag name."""
    return [
        child
        for child in node.childNodes
        if child.nodeType == child.ELEMENT_NODE and child.tagName == tag
    ]


def run_text(r_node: minidom.Element) -> str:
    """Extract concatenated text from <w:t> elements inside a <w:r>."""
    parts: list[str] = []
    for t_node in r_node.getElementsByTagName("w:t"):
        for child in t_node.childNodes:
            if child.nodeType in (child.TEXT_NODE, child.CDATA_SECTION_NODE):
                parts.append(child.data)
    return "".join(parts)


def parse_color_element(color_el: minidom.Element) -> ColorSpec:
    """Parse a <w:color> element into normalized uppercase hex and theme name."""
    val = color_el.getAttribute("w:val").strip().upper()
    theme = color_el.getAttribute("w:themeColor").strip().lower()
    return ColorSpec(val=val, theme_color=theme)


def load_style_colors(styles_xml: bytes | None) -> dict[str, ColorSpec]:
    """Resolve effective run color for each styleId in word/styles.xml."""
    if not styles_xml:
        return {}

    dom = minidom.parseString(styles_xml)
    raw_styles: dict[str, tuple[str | None, ColorSpec | None]] = {}

    for style_el in dom.getElementsByTagName("w:style"):
        style_id = style_el.getAttribute("w:styleId").strip()
        if not style_id:
            continue
        based_els = direct_children(style_el, "w:basedOn")
        based_on = based_els[0].getAttribute("w:val").strip() if based_els else None

        color_spec: ColorSpec | None = None
        rprs = direct_children(style_el, "w:rPr")
        if rprs:
            colors = direct_children(rprs[0], "w:color")
            if colors:
                color_spec = parse_color_element(colors[0])

        raw_styles[style_id] = (based_on, color_spec)

    resolved: dict[str, ColorSpec] = {}

    def resolve(style_id: str, seen: set[str]) -> ColorSpec | None:
        if style_id in resolved:
            return resolved[style_id]
        if style_id in seen or style_id not in raw_styles:
            return None
        seen.add(style_id)
        based_on, direct_color = raw_styles[style_id]
        color = direct_color if direct_color is not None else (resolve(based_on, seen) if based_on else None)
        if color is not None:
            resolved[style_id] = color
        return color

    for sid in raw_styles:
        resolve(sid, set())

    return resolved


def enclosing_paragraph_style(r_node: minidom.Element) -> str | None:
    """Return the enclosing paragraph's w:pStyle value, if any."""
    parent = r_node.parentNode
    while parent is not None and parent.nodeType == parent.ELEMENT_NODE:
        if parent.tagName == "w:p":
            pprs = direct_children(parent, "w:pPr")
            if pprs:
                pstyles = direct_children(pprs[0], "w:pStyle")
                if pstyles:
                    return pstyles[0].getAttribute("w:val").strip() or None
            return None
        parent = parent.parentNode
    return None


def effective_run_color(
    r_node: minidom.Element,
    style_colors: dict[str, ColorSpec],
) -> ColorSpec | None:
    """Determine effective <w:color> of a run (direct -> rStyle -> pStyle)."""
    rprs = direct_children(r_node, "w:rPr")
    if rprs:
        rpr = rprs[0]
        colors = direct_children(rpr, "w:color")
        if colors:
            return parse_color_element(colors[0])
        rstyles = direct_children(rpr, "w:rStyle")
        if rstyles:
            rs_id = rstyles[0].getAttribute("w:val").strip()
            if rs_id in style_colors:
                return style_colors[rs_id]

    p_style = enclosing_paragraph_style(r_node)
    if p_style and p_style in style_colors:
        return style_colors[p_style]
    return None


def is_colored(color: ColorSpec | None, ignored_vals: set[str]) -> bool:
    """Return True if the resolved ColorSpec represents non-black chromatic text."""
    if color is None:
        return False
    if color.val in ignored_vals:
        return False
    if not color.val and color.theme_color in NON_COLORED_THEME_NAMES:
        return False
    return True


def insert_rpr_child(rpr: minidom.Element, elem: minidom.Element, before_tags: set[str]) -> None:
    """Insert elem into <w:rPr> right after all elements in before_tags."""
    for child in rpr.childNodes:
        if child.nodeType == child.ELEMENT_NODE and child.tagName not in before_tags:
            rpr.insertBefore(elem, child)
            return
    rpr.appendChild(elem)


def ensure_rpr(dom: minidom.Document, r_node: minidom.Element) -> minidom.Element:
    """Return existing <w:rPr> on a run or create one as the first child."""
    rprs = direct_children(r_node, "w:rPr")
    if rprs:
        return rprs[0]
    rpr = dom.createElement("w:rPr")
    if r_node.firstChild is not None:
        r_node.insertBefore(rpr, r_node.firstChild)
    else:
        r_node.appendChild(rpr)
    return rpr


def set_run_black(dom: minidom.Document, rpr: minidom.Element) -> None:
    """Set run color explicitly to black (#000000), clearing theme attributes."""
    colors = direct_children(rpr, "w:color")
    if colors:
        c_el = colors[0]
        for attr in ("w:themeColor", "w:themeTint", "w:themeShade"):
            if c_el.hasAttribute(attr):
                c_el.removeAttribute(attr)
        c_el.setAttribute("w:val", "000000")
    else:
        c_el = dom.createElement("w:color")
        c_el.setAttribute("w:val", "000000")
        insert_rpr_child(rpr, c_el, BEFORE_COLOR_TAGS)


def process_xml_part(
    xml_bytes: bytes,
    style_colors: dict[str, ColorSpec],
    underline_style: str,
    ignored_vals: set[str],
    include_whitespace: bool,
    to_black: bool,
) -> tuple[bytes, dict[str, int], Counter[str]]:
    """Process one OOXML part (document.xml, header*.xml, etc.) and return updated XML."""
    dom = minidom.parseString(xml_bytes)
    stats = {"updated": 0, "already_target": 0, "skipped": 0}
    color_counts: Counter[str] = Counter()

    for r_node in dom.getElementsByTagName("w:r"):
        color = effective_run_color(r_node, style_colors)
        if not is_colored(color, ignored_vals):
            stats["skipped"] += 1
            continue

        txt = run_text(r_node)
        if not txt or (not include_whitespace and not txt.strip()):
            stats["skipped"] += 1
            continue

        assert color is not None
        color_counts[color.label] += 1
        rpr = ensure_rpr(dom, r_node)

        us = direct_children(rpr, "w:u")
        if us:
            if us[0].getAttribute("w:val") == underline_style and not to_black:
                stats["already_target"] += 1
            else:
                us[0].setAttribute("w:val", underline_style)
                stats["updated"] += 1
        else:
            u_elem = dom.createElement("w:u")
            u_elem.setAttribute("w:val", underline_style)
            insert_rpr_child(rpr, u_elem, BEFORE_U_TAGS)
            stats["updated"] += 1

        if to_black:
            set_run_black(dom, rpr)

    return dom.toxml(encoding="utf-8"), stats, color_counts


def select_target_parts(
    names: list[str],
    include_headers_footers: bool,
    include_notes: bool,
) -> set[str]:
    """Select which XML parts inside the .docx ZIP archive to process."""
    targets = {"word/document.xml"}
    for name in names:
        if include_headers_footers and (
            (name.startswith("word/header") or name.startswith("word/footer")) and name.endswith(".xml")
        ):
            targets.add(name)
        if include_notes and name in {"word/footnotes.xml", "word/endnotes.xml", "word/comments.xml"}:
            targets.add(name)
    return targets


def default_output_path(input_path: Path, underline_style: str) -> Path:
    """Build default output path in the same directory as input_path."""
    suffix_tag = "双下划线" if underline_style == "double" else underline_style
    return input_path.with_name(f"{input_path.stem}_{suffix_tag}{input_path.suffix}")


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="为 Word (.docx) 文档中的彩色文本添加双下划线（或指定下划线样式），并另存为新文件。",
    )
    parser.add_argument("input", type=Path, help="输入的 .docx 文件路径")
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        default=None,
        help="输出的 .docx 文件路径（默认保存到同级目录：<原文件名>_双下划线.docx）",
    )
    parser.add_argument(
        "-u",
        "--underline",
        default="double",
        choices=UNDERLINE_CHOICES,
        help="下划线样式（默认: double 双下划线）",
    )
    parser.add_argument(
        "--include-headers-footers",
        action="store_true",
        help="同时处理页眉和页脚中的彩色文本（默认仅处理正文）",
    )
    parser.add_argument(
        "--include-notes",
        action="store_true",
        help="同时处理脚注、尾注和批注中的彩色文本",
    )
    parser.add_argument(
        "--include-whitespace",
        action="store_true",
        help="也为空白字符（空格/空 run）添加下划线（默认跳过）",
    )
    parser.add_argument(
        "--to-black",
        action="store_true",
        help="加上下划线的同时，将彩色文字改为纯黑 (#000000)",
    )
    parser.add_argument(
        "--ignore-colors",
        default="000000,AUTO",
        help="逗号分隔的忽略颜色值（默认: 000000,AUTO；例如可追加 FFFFFF 忽略白色文字）",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="仅统计匹配的彩色文本数量，不写入输出文件",
    )
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    input_path: Path = args.input.expanduser()
    if not input_path.is_file():
        print(f"错误：找不到输入文件 {input_path}", file=sys.stderr)
        return 1

    output_path: Path = args.output.expanduser() if args.output else default_output_path(input_path, args.underline)
    if not args.dry_run and output_path.resolve() == input_path.resolve():
        print("错误：输出路径不能与输入路径相同，请指定不同的新文件名。", file=sys.stderr)
        return 1

    ignored_vals = {""} | {c.strip().lstrip("#").upper() for c in args.ignore_colors.split(",") if c.strip()}

    with zipfile.ZipFile(input_path, "r") as zin:
        names = zin.namelist()
        if "word/document.xml" not in names:
            print(f"错误：{input_path} 不是有效的 .docx 文件（缺少 word/document.xml）", file=sys.stderr)
            return 1

        styles_xml = zin.read("word/styles.xml") if "word/styles.xml" in names else None
        style_colors = load_style_colors(styles_xml)
        targets = select_target_parts(names, args.include_headers_footers, args.include_notes)

        total_updated = 0
        total_already = 0
        total_colors: Counter[str] = Counter()
        rewritten_parts: dict[str, bytes] = {}

        for part in sorted(targets):
            if part not in names:
                continue
            new_bytes, stats, c_counts = process_xml_part(
                zin.read(part),
                style_colors=style_colors,
                underline_style=args.underline,
                ignored_vals=ignored_vals,
                include_whitespace=args.include_whitespace,
                to_black=args.to_black,
            )
            rewritten_parts[part] = new_bytes
            total_updated += stats["updated"]
            total_already += stats["already_target"]
            total_colors.update(c_counts)

        if not args.dry_run:
            output_path.parent.mkdir(parents=True, exist_ok=True)
            with zipfile.ZipFile(output_path, "w") as zout:
                for item in zin.infolist():
                    data = rewritten_parts.get(item.filename)
                    if data is None:
                        data = zin.read(item.filename)
                    zout.writestr(item, data)

    mode = "[预览模式]" if args.dry_run else "完成"
    print(f"{mode} 处理文件: {input_path}")
    if not args.dry_run:
        print(f"输出文件: {output_path}")
    print(
        f"下划线样式: {args.underline} | 新增/修改: {total_updated} 处 | 原已匹配: {total_already} 处"
    )
    if total_colors:
        breakdown = ", ".join(f"{label}: {cnt}" for label, cnt in total_colors.most_common())
        print(f"颜色分布: {breakdown}")
    else:
        print("未检测到符合条件的彩色文本。")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
