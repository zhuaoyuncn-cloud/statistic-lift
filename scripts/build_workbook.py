"""从核实过的当周单表数据填充学院量化表。"""
import argparse
import json
import math
from collections import defaultdict
from copy import copy
from datetime import date
from decimal import Decimal, ROUND_HALF_UP
from io import BytesIO
from pathlib import Path
from statistics import mean
from xml.etree import ElementTree as ET
from zipfile import ZipFile

from openpyxl import load_workbook
from openpyxl.comments import Comment
from openpyxl.styles import Side
from openpyxl.utils import get_column_letter

from report_text import SUBJECTS, class_key, lines_of, name_token, render_events


TEMPLATE = Path(__file__).resolve().parents[1] / "assets" / "量化表模板.xlsx"
SHEETS = ("综合", *SUBJECTS)
# 表体首行、宽度、模板通报区首行（如有）
LAYOUT = {"综合": (3, 9, 18), "时政学习": (4, 14, 19), "晚自习": (3, 7, 18),
          "早操": (3, 7, 18), "宿舍卫生": (5, 10, None),
          "教室卫生": (3, 9, None), "宿舍安全": (5, 4, None)}


def numeric(value):
    if type(value) not in (int, float) or not math.isfinite(value):
        raise ValueError(f"分数必须是可核实的数值：{value}")
    return value


def chinese_week(n):
    digits = "零一二三四五六七八九"
    if type(n) is not int or not 1 <= n <= 99:
        raise ValueError("教学周次应为 1 到 99 的整数")
    if n < 10:
        return digits[n]
    return (digits[n // 10] if n >= 20 else "") + "十" + (digits[n % 10] if n % 10 else "")


def width(ws, column):
    for dim in ws.column_dimensions.values():
        if dim.min <= column <= dim.max:
            return dim.width
    return ws.sheet_format.defaultColWidth or 8.43


def alter(style, **changes):
    result = copy(style)
    for key, value in changes.items():
        setattr(result, key, value)
    return result


class Form:
    def __init__(self, ws):
        self.ws = ws
        self.first, self.cols, note = LAYOUT[ws.title]
        self.styles = [ws.cell(self.first, col)._style for col in range(1, self.cols + 1)]
        self.height = ws.row_dimensions[self.first].height or ws.sheet_format.defaultRowHeight
        self.note_style = ws.cell(note, 1)._style if note else None
        for region in list(ws.merged_cells.ranges):
            if region.min_row >= self.first:
                ws.unmerge_cells(str(region))
        ws.delete_rows(self.first, ws.max_row - self.first + 1)
        for number in list(ws.row_dimensions):
            if number >= self.first:
                del ws.row_dimensions[number]
        self.cache = {}

    def cell(self, row, col, value=None):
        cell = self.ws.cell(row, col)
        cell._style = copy(self.styles[col - 1])
        cell.font = alter(cell.font, name="仿宋_GB2312", scheme=None)
        dim = self.ws.row_dimensions[row]
        dim.height = max(self.height, dim.height or 0)
        if isinstance(value, str):
            # 中文姓名作为完整词排版；宽度留出不同字体渲染的余量。
            budget = max(2, int(width(self.ws, col) * 9 / (cell.font.sz or 12)))
            lines = lines_of(value, budget)
            value = "\n".join(lines)
            cell.alignment = alter(cell.alignment, wrap_text=len(lines) > 1, shrink_to_fit=len(lines) == 1)
            need = len(lines) * ((cell.font.sz or 12) * 1.4 + 2)
            if need > 409 or len(value) > 32767:
                raise ValueError(f"{self.ws.title} {cell.coordinate} 文字超出 Excel 行高或字数，请按原表拆分明细")
            dim.height = max(dim.height, need)
        cell.value = value
        if isinstance(value, str):
            cell.data_type = "s"
        return cell

    def formula(self, row, col, expression, value):
        cell = self.cell(row, col)
        cell.value = "=" + expression
        self.cache[cell.coordinate] = value
        return cell

    def band(self, row, first, last):
        cell = self.ws.cell(row, 1)
        cell._style = copy(self.note_style)
        cell.border = alter(cell.border, top=cell.border.top if first else Side(),
                            bottom=cell.border.bottom if last else Side())
        self.ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=self.cols)
        cell.font = alter(cell.font, name="仿宋_GB2312", scheme=None)
        return cell

    def notes(self, start, paragraphs):
        size = self.ws.cell(start, 1)
        size._style = copy(self.note_style)
        font_size = size.font.sz or 14
        budget = int(sum(width(self.ws, c) for c in range(1, self.cols + 1)) * 8.5 / font_size)
        lines = [line for para in paragraphs for line in lines_of(para, budget)]
        # 合并区内显式折行，分块承载长名单以避开 Excel 的行高上限。
        for offset in range(0, len(lines), 10):
            row = start + offset // 10
            ws = self.ws
            cell = self.band(row, offset == 0, offset + 10 >= len(lines) and ws.title != "综合")
            cell.alignment = alter(cell.alignment, horizontal="left", vertical="top",
                                   wrap_text=True, shrink_to_fit=False, indent=0)
            cell.number_format = "General"
            chunk = lines[offset:offset + 10]
            cell.value = "\n".join(chunk)
            cell.data_type = "s"
            ws.row_dimensions[row].height = (font_size * 1.5 + 2) * len(chunk) + 3
        return start + (len(lines) + 9) // 10


def score(form, row, col, entry, values, expression):
    """独立复算明细，并在分表采用分上注明需要追溯的差异。"""
    recalc = None
    if values:
        recalc = sum(values) if form.ws.title == "时政学习" else mean(values)

    adopted = entry.get("score")
    audit = entry.get("audit")
    registered, source, reason, pending = adopted, "未提供", "", False
    if audit is not None:
        if not isinstance(audit, dict):
            raise ValueError(f"{entry['class']} audit 须为对象")
        registered = audit["registered"]
        source = audit["source"]
        reason = audit.get("reason", "")
        pending = audit.get("pending", False)
        if not isinstance(source, str) or not source.strip() or not isinstance(reason, str) or type(pending) is not bool:
            raise ValueError(f"{entry['class']} audit 需填写来源单元格，reason 须为文字，pending 须为布尔值")
        source, reason = source.strip(), reason.strip()
        if pending and (adopted is None or not reason):
            raise ValueError(f"{entry['class']} 明细待核实，需注明原因并填写已确认的采用分")
        if registered is None and adopted is not None and not reason:
            raise ValueError(f"{entry['class']} 原表空白，需注明采用分的补分原因和出处")

    if registered is not None:
        registered = numeric(registered)
    differs = registered is not None and recalc is not None and not math.isclose(
        registered, recalc, rel_tol=0, abs_tol=1e-9)
    if adopted is not None:
        val = numeric(adopted)
        cell = form.cell(row, col, val)
    else:
        if recalc is None:
            raise ValueError(f"{form.ws.title} {entry['class']} 缺少分数依据")
        if differs and not reason:
            raise ValueError(f"{entry['class']} 登记分与复算分不同，需说明采用复算的原因或填写正式周分")
        val = recalc
        cell = form.formula(row, col, expression, val)

    override = registered is not None and not math.isclose(registered, val, rel_tol=0, abs_tol=1e-9)
    if differs or override or (audit is not None and (pending or reason or registered is None)):
        registered_text = "空白" if registered is None else f"{registered:.12g}"
        recalc_label = "复算分（按已确认明细暂算）" if pending else "复算分"
        recalc_text = "无可复算明细" if recalc is None else f"{recalc:.12g}"
        if not reason:
            reason = "原表周分空白，采用明细复算分" if registered is None else "差异原因待核实"
        note = (f"班级：{entry['class']}\n来源单元格：{source}\n登记分：{registered_text}\n"
                f"{recalc_label}：{recalc_text}\n原因：{reason}\n本表采用值：{val:.12g}")
        if pending:
            note += "\n核对状态：待核实"
        cell.comment = Comment(note, "量化分核对", width=380, height=185)
    return cell.coordinate, val


def fixed_rows(form, rows, roster):
    subject = form.ws.title
    result = {}
    rows = sorted(rows, key=lambda e: (bool(e.get("public")), class_key(e["class"])))
    for r, entry in enumerate(rows, form.first):
        cls = entry["class"]
        public = entry.get("public", False) if subject == "教室卫生" else False
        if cls in result and not public:
            raise ValueError(f"{subject} 班级重复：{cls}")
        label = f"公共教室{entry['room']}\n（{cls}）" if public else cls
        form.cell(r, 1, label)
        if subject == "时政学习":
            fields = entry["values"]
            if len(fields) != 12:
                raise ValueError("时政学习 values 须对应 B:M 十二列")
            for col, value in enumerate(fields, 2):
                if col == 3 and value:
                    value = name_token(value)
                form.cell(r, col, value)
            components = [fields[c] for c in (4, 6, 7, 8, 10, 11)]
            vals = [numeric(v) for v in components if v is not None]
            if len(vals) != 6:
                vals = []
            ref = score(form, r, 14, entry, vals, "+".join(f"{c}{r}" for c in "FHIJLM"))
        elif subject == "宿舍安全":
            ref = score(form, r, 2, entry, [], "")
            form.cell(r, 3, entry.get("check"))
            form.cell(r, 4, entry.get("hazard"))
        else:
            is_room = subject == "教室卫生"
            first = 3 if is_room else 2
            if is_room:
                form.cell(r, 2, entry.get("room"))
                form.cell(r, 9, entry.get("remarks"))
            days = entry["days"]
            if len(days) != 5:
                raise ValueError(f"{subject} {cls} 须按表头提供五个日分位置")
            vals = [numeric(v) for v in days if v is not None]
            for col, value in enumerate(days, first):
                form.cell(r, col, value)
            region = f"{get_column_letter(first)}{r}:{get_column_letter(first + 4)}{r}"
            ref = score(form, r, first + 5, entry, vals, f"AVERAGE({region})")
        if not public:
            result[cls] = ref
    if set(result) != set(roster):
        raise ValueError(f"{subject} 普通班成绩与名册不符：{set(roster) ^ set(result)}")
    return result


def dormitory(form, rows, roster):
    groups = defaultdict(list)
    for e in rows:
        groups[e["class"]].append(e)
    if set(groups) != set(roster):
        raise ValueError(f"宿舍卫生 班级与名册不符：{set(roster) ^ set(groups)}")
    # 两栏顺读；按房间数选择高度最接近的班级分界，整班不跨栏。
    total = len(rows)
    sizes = [len(groups[c]) for c in roster]
    cut = min(range(1, len(roster)), key=lambda i: abs(total - 2 * sum(sizes[:i]))) if len(roster) > 1 else 1
    result = {}
    for col, classes in [(1, roster[:cut]), (6, roster[cut:])]:
        r = form.first
        for cls in classes:
            start = r
            points, official = [], set()
            audit = None
            for entry in groups[cls]:
                form.cell(r, col)
                room = form.cell(r, col + 1, entry.get("room"))
                if entry.get("mixed"):
                    room.font = alter(room.font, color="FF000000")
                form.cell(r, col + 2, entry.get("condition"))
                value = entry.get("points")
                if value is not None:
                    points.append(numeric(value))
                form.cell(r, col + 3, value)
                form.cell(r, col + 4)
                if entry.get("score") is not None:
                    official.add(numeric(entry["score"]))
                if entry.get("audit") is not None:
                    if audit is not None:
                        raise ValueError(f"宿舍卫生 {cls} 的 audit 只需填写一次")
                    audit = entry["audit"]
                r += 1
            if len(official) > 1:
                raise ValueError(f"宿舍卫生 {cls} 有多个班级周分")
            for c in [col, col + 4]:
                if r - start > 1:
                    form.ws.merge_cells(start_row=start, start_column=c, end_row=r-1, end_column=c)
            form.cell(start, col, cls)
            letter = get_column_letter(col + 3)
            entry = {"class": cls, "score": next(iter(official), None), "audit": audit}
            result[cls] = score(form, start, col + 4, entry, points, f"AVERAGE({letter}{start}:{letter}{r-1})")
    mixed = list(dict.fromkeys(e["room"] for e in rows if e.get("mixed")))
    if mixed:
        row = form.ws.max_row + 1
        form.ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=10)
        note = form.ws.cell(row, 1, "注：混合宿舍为" + "、".join(mixed) + "。")
        note.alignment = alter(note.alignment, wrap_text=True)
        budget = int(sum(width(form.ws, c) for c in range(1, 11)) * 8.5 / (note.font.sz or 12))
        form.ws.row_dimensions[row].height = max(18, len(lines_of(note.value, budget)) * 18)
    return result


def set_headers(wb, grade, week, issued, sections):
    cn = chinese_week(week)
    titles = (f"工程机械学院{grade}级第{cn}周量化表",
              f"工程机械学院团委第{cn}周{grade}级政治学习纪实分表",
              f"工程机械学院纪检部第{week}周{grade % 100:02}级晚自习表",
              f"工程机械学院体育部第{cn}周{grade}级早操情况统计表",
              f"{grade}级工程机械学院普通本科宿舍卫生检查表",
              f"工程机械学院{grade % 100:02}级教室卫生平时卫生检查分数统计表",
              f"{grade}级第{cn}周安管部班级量化表")
    for tab, title in zip(("综合", *SUBJECTS), titles):
        wb[tab]["A1"] = title
    wb["宿舍卫生"]["A3"] = f"第{week}周"
    for tab, addr in [("宿舍卫生", "F3"), ("教室卫生", "I2")]:
        stamp = sections[tab].get("header_date")
        if stamp:
            d = date.fromisoformat(stamp)
            wb[tab][addr] = f"{d.year}.{d.month}.{d.day}" if tab == "宿舍卫生" else f"备注：{d.month}.{d.day}"
    return f"{issued.year}年{issued.month}月{issued.day}日"


def get_order(data, field, allowed, *, complete=True):
    order = data.get(field, list(allowed))
    if (not isinstance(order, list) or any(name not in allowed for name in order)
            or len(order) != len(set(order)) or (complete and len(order) != len(allowed))):
        raise ValueError(f"{field} 须按需排列以下名称，每项最多出现一次：{'、'.join(allowed)}")
    return order


def build(data):
    grade, week = data["grade"], data["week"]
    if type(grade) is not int or not 1000 <= grade <= 9999:
        raise ValueError("grade 请填写四位入学年级")
    roster = sorted(data["classes"], key=class_key)
    if not roster or len(set(roster)) != len(roster):
        raise ValueError("班级名册为空或有重复")
    sections = data["sections"]
    if set(sections) != set(SUBJECTS):
        raise ValueError(f"六项单表未对齐：{set(sections) ^ set(SUBJECTS)}")
    sheet_order = get_order(data, "sheet_order", SHEETS)
    score_order = get_order(data, "score_order", SUBJECTS)
    report_subjects = get_order(data, "report_subjects", SUBJECTS, complete=False)
    wb = load_workbook(TEMPLATE)
    for index, name in enumerate(sheet_order):
        wb.move_sheet(name, offset=index - wb.sheetnames.index(name))
    wb.active = wb["综合"]
    for col, subject in enumerate(score_order, 2):
        wb["综合"].cell(2, col, subject)
    forms = {ws.title: Form(ws) for ws in wb}
    issued = set_headers(wb, grade, week, date.fromisoformat(data["issued"]), sections)
    info, reports = {}, {}
    for subject in SUBJECTS:
        entry, form = sections[subject], forms[subject]
        try:
            rows = entry["rows"]
            if any(row["class"] not in roster for row in rows):
                raise ValueError("成绩表中有名册以外的班级")
            handler = dormitory if subject == "宿舍卫生" else fixed_rows
            info[subject] = handler(form, rows, roster)
            events = entry.get("events")
            if events is None and subject in report_subjects:
                raise ValueError("缺少综合通报所需记录，确认本周无问题时填写空数组")
            body = render_events(subject, events, roster, entry.get("problem_groups", [])) if events is not None else []
            notes = entry.get("notes", [])
            if form.note_style is not None and (body or notes):
                form.notes(form.first + len(rows), ["备注：", *body, *notes])
            reports[subject] = body
        except (KeyError, TypeError, ValueError) as error:
            raise ValueError(f"{subject}：{error}") from error

    summary = forms["综合"]
    scores, totals = {}, {}
    for cls in roster:
        scores[cls] = [Decimal(str(info[subject][cls][1])).quantize(Decimal("0.1"), rounding=ROUND_HALF_UP)
                       for subject in score_order]
        totals[cls] = sum(scores[cls])

    # Python 排序是稳定的，同分时沿用名册班序。
    ranked_classes = sorted(roster, key=totals.__getitem__, reverse=True)
    last = summary.first + len(ranked_classes) - 1
    previous_total = None
    for r, cls in enumerate(ranked_classes, summary.first):
        summary.cell(r, 1, cls)
        for col, (subject, rounded) in enumerate(zip(score_order, scores[cls]), 2):
            addr, _ = info[subject][cls]
            summary.formula(r, col, f"ROUND('{subject}'!{addr},1)", float(rounded))
        total = totals[cls]
        summary.formula(r, 8, f"ROUND(SUM(B{r}:G{r}),1)", float(total))
        if total != previous_total:
            rank = r - summary.first + 1
            previous_total = total
        summary.formula(r, 9, f"RANK(H{r},$H${summary.first}:$H${last},0)", rank)
    footer = last + 1
    if report_subjects:
        paragraphs = ["通报内容："]
        for subject in report_subjects:
            paragraphs.extend([subject + "：", *reports[subject]])
        footer = summary.notes(footer, paragraphs)
    for r, label in enumerate(["工程机械学院", issued], footer):
        cell = summary.band(r, False, r == footer + 1)
        cell.alignment = alter(cell.alignment, horizontal="right", vertical="center")
        cell.value = label
        summary.ws.row_dimensions[r].height = 28
    for form in forms.values():
        ws = form.ws
        ws.print_area = f"A1:{get_column_letter(form.cols)}{ws.max_row}"
    wb.calculation.fullCalcOnLoad = True
    return wb, forms


def save(wb, forms, path):
    # 由 openpyxl 写出原生公式，补入同一批数据算出的缓存供文件预览使用。
    raw, finished = BytesIO(), BytesIO()
    wb.save(raw)
    ns = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
    sheets = {f"xl/worksheets/sheet{i}.xml": forms[t].cache for i, t in enumerate(wb.sheetnames, 1)}
    with ZipFile(raw) as src, ZipFile(finished, "w") as dst:
        for part in src.infolist():
            content = src.read(part)
            if cache := sheets.get(part.filename):
                xml = ET.fromstring(content)
                for cell in xml.iter(ns + "c"):
                    if cell.get("r") in cache:
                        cell.find(ns + "v").text = str(cache[cell.get("r")])
                content = ET.tostring(xml, encoding="utf-8")
            dst.writestr(part, content)
    with path.open("xb") as out:
        out.write(finished.getbuffer())


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("data", type=Path)
    parser.add_argument("output", type=Path, help="新 .xlsx 文件路径")
    args = parser.parse_args()
    if args.output.suffix.lower() != ".xlsx":
        parser.error("输出文件需要 .xlsx 扩展名")
    try:
        data = json.loads(args.data.read_text(encoding="utf-8-sig"))
        wb, forms = build(data)
        save(wb, forms, args.output)
    except (KeyError, ValueError, TypeError, OSError) as error:
        parser.exit(2, f"未生成量化表：{error}\n")


if __name__ == "__main__":
    main()
