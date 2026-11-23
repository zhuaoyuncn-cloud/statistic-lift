"""量化表中按姓名边界折行的文字和六项通报。"""
import re
import unicodedata
from collections import defaultdict
from datetime import date
from itertools import groupby


SUBJECTS = ("时政学习", "晚自习", "早操", "宿舍卫生", "教室卫生", "宿舍安全")
MAJORS = ("机械", "机工", "智造", "成型", "型职贯", "机职贯")


def class_key(label):
    match = re.fullmatch(r"(机械|机工|智造|成型|型职贯|机职贯)([0-9]+)", label)
    if not match:
        raise ValueError(f"班级格式不清楚：{label}")
    return MAJORS.index(match[1]), int(match[2])


def name_token(name):
    if not isinstance(name, str):
        raise ValueError(f"姓名应为文本：{name}")
    name = re.sub(r"\s+", "", name)
    if not name or any(c in name for c in "{}"):
        raise ValueError(f"无效姓名：{name}")
    return "{" + name + "}"


def lines_of(text, budget):
    """把带 {姓名} 标记的文字拆成物理行，预算单位为半角字符。"""
    lines, line, used = [], "", 0
    for token in re.findall(r"\{[^{}\n]+\}|[A-Za-z0-9#]+|\n|.", text):
        if token == "\n":
            lines.append(line.rstrip())
            line, used = "", 0
            continue
        if token.startswith("{") and token.endswith("}"):
            token = name_token(token[1:-1])[1:-1]
            if len(token) == 2:
                token = "　".join(token)
        elif token in "{}":
            raise ValueError(f"未配对的姓名标记：{text}")
        width = sum(2 if unicodedata.east_asian_width(c) in "FW" else 1 for c in token)
        if line and used + width > budget:
            lines.append(line.rstrip())
            line, used = "", 0
        if token == " " and not line:
            continue
        line += token
        used += width
    lines.append(line.rstrip())
    return lines


def month_day(value):
    if not value:
        return ""
    day = date.fromisoformat(value)
    return f"{day.month}月{day.day}日"


def class_label(classes):
    parts = []
    for label in classes:
        major = MAJORS[class_key(label)[0]]
        parts.append(("　".join(major) if len(major) == 2 else major) + label[len(major):])
    return "、".join(parts)


def render_events(subject, events, roster, problem_groups=()):
    if events == []:
        return ["无"]
    if not isinstance(events, list):
        raise ValueError(f"{subject}：请提供逐条记录或确认无问题的空数组")
    records = []
    for event in events:
        labels = event["class"]
        labels = [labels] if isinstance(labels, str) else labels
        if not labels or set(labels) - set(roster):
            raise ValueError(f"{subject}：通报班级与名册不符 {labels}")
        classes = tuple(sorted(labels, key=class_key))
        day, issue = event.get("date", ""), event["issue"].strip()
        month_day(day)
        if not issue or (subject == "晚自习" and not day):
            raise ValueError(f"{subject}缺少问题或检查日期：{event}")
        names = event.get("names", [])
        if not isinstance(names, list):
            raise ValueError("names 须为逐人数组")
        who = " ".join(name_token(n) for n in names)
        place = " ".join(filter(None, (event.get("room"), event.get("bed"))))
        records.append((day, classes, issue, place, who))

    def by_class(record):
        return tuple(class_key(c) for c in record[1])

    records.sort(key=by_class)
    if subject == "时政学习":
        return ["；".join(f"{class_label(cs)}：" + " ".join(filter(None, (month_day(d), p, n, i)))
                         for d, cs, i, p, n in records)]
    out = []
    if subject in ("晚自习", "早操"):
        for day, daily in groupby(sorted(records, key=lambda r: r[0]), key=lambda r: r[0]):
            if day:
                out.append(month_day(day))
            buckets = defaultdict(list)
            for record in daily:
                issue = record[2]
                key = issue if subject == "早操" or issue in problem_groups else ""
                buckets[key].append(record)
            for key in sorted(buckets, key=lambda k: ({"旷操": 0, "迟到": 1}.get(k, 2), k)):
                if key:
                    out.append(key + "：")
                for cs, items in groupby(buckets[key], key=lambda r: r[1]):
                    details = [" ".join(filter(None, (p, n, "" if key else i))) for _, _, i, p, n in items]
                    if key:
                        out.append(class_label(cs) + "：")
                        if any(details):
                            out.append(" ".join(filter(None, details)))
                    else:
                        out.append(class_label(cs) + "：" + "；".join(details))
    elif subject == "宿舍卫生":
        for cs, items in groupby(records, key=lambda r: r[1]):
            for index, (day, _, issue, place, names) in enumerate(items):
                prefix = class_label(cs) + "：" if index == 0 else ""
                out.append(prefix + " ".join(filter(None, (month_day(day), place, names, issue))))
    else:
        for cs, items in groupby(records, key=lambda r: r[1]):
            out.append(class_label(cs) + "：")
            for room, located in groupby(sorted(items, key=lambda r: r[3]), key=lambda r: r[3]):
                if room:
                    out.append(room)
                for d, _, issue, _, names in sorted(located, key=lambda r: r[0]):
                    out.append(" ".join(filter(None, (month_day(d), names, issue))))
    return out
