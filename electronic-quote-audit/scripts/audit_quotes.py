#!/usr/bin/env python3
"""Local, deterministic quote checks. Python 3.12+, standard library only."""
from __future__ import annotations

import argparse
import csv
import hashlib
import html
import io
import json
import re
import stat
import sys
import zipfile
from datetime import datetime, time, timezone
from decimal import Decimal, InvalidOperation, localcontext, Inexact
from fractions import Fraction
from pathlib import Path, PurePosixPath
from xml.etree import ElementTree as ET
from zoneinfo import ZoneInfo, ZoneInfoNotFoundError

RULES_VERSION = "electronic-quote-audit-rules-v0.1"
SCHEMA_VERSION = "electronic-quote-audit-config-v1"
MAX_FILE = 10 * 1024 * 1024
MAX_EXPANDED = 40 * 1024 * 1024
MAX_XML = 12 * 1024 * 1024
MAX_ROWS, MAX_COLS, MAX_CELLS = 20000, 256, 200000
FIELDS = {
    "manufacturer", "mpn", "internal_id", "quantity", "quantity_unit", "required_receipt",
    "unit_price", "price_quantity", "price_unit", "moq", "order_multiple",
    "minimum_release", "package_quantity", "tier_min", "tier_max", "tier_basis",
    "currency", "tax_basis", "freight", "freight_currency", "freight_scope", "freight_tax_basis",
    "valid_from", "valid_until", "valid_until_alternate", "stock",
    "distributor_receipt", "manufacturer_lead_time", "customer_receipt",
    "customer_receipt_alternate", "notes",
}
NS = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
RID = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"


class InputError(ValueError):
    """A structural or safety error; no result is written."""


def digest(data):
    return hashlib.sha256(data).hexdigest()


def jdump(value):
    return json.dumps(value, ensure_ascii=False, indent=2, sort_keys=True) + "\n"


def strict_json(data):
    def pairs(items):
        result = {}
        for key, value in items:
            if key in result:
                raise InputError(f"JSON重复键：{key}")
            result[key] = value
        return result
    try:
        return json.loads(data.decode("utf-8-sig"), object_pairs_hook=pairs,
                          parse_constant=lambda v: (_ for _ in ()).throw(InputError(f"非法JSON数值：{v}")))
    except (UnicodeError, json.JSONDecodeError) as exc:
        raise InputError(f"配置不是有效UTF-8 JSON：{exc}") from exc


def confined(root, relative, *, exists=True):
    if not isinstance(relative, str) or not relative or "\\" in relative:
        raise InputError("路径必须是根目录下的非空相对路径，使用/分隔")
    p = PurePosixPath(relative)
    if p.is_absolute() or any(x in {"..", "."} for x in p.parts) or ":" in relative:
        raise InputError(f"拒绝绝对路径或路径遍历：{relative}")
    candidate = root
    for part in p.parts:
        candidate = candidate / part
        if candidate.is_symlink():
            raise InputError(f"拒绝符号链接：{relative}")
    try:
        candidate.resolve().relative_to(root)
    except ValueError as exc:
        raise InputError("路径越过指定根目录") from exc
    if exists and (not candidate.is_file() or not stat.S_ISREG(candidate.stat().st_mode)):
        raise InputError(f"输入不是普通文件：{relative}")
    return candidate


def read_bytes(root, relative):
    p = confined(root, relative)
    if p.stat().st_size > MAX_FILE:
        raise InputError(f"文件超过10MiB：{relative}")
    data = p.read_bytes()
    if len(data) > MAX_FILE:
        raise InputError(f"文件超过10MiB：{relative}")
    return data


def colname(index):
    result = ""
    while index:
        index, rem = divmod(index - 1, 26)
        result = chr(65 + rem) + result
    return result


def colindex(name):
    result = 0
    for letter in name:
        result = result * 26 + ord(letter) - 64
    return result


def xml(data, label):
    if len(data) > MAX_XML:
        raise InputError(f"XML超过限制：{label}")
    try:
        decoded = data.decode("utf-8-sig")
    except UnicodeError as exc:
        raise InputError(f"仅支持UTF-8 XLSX XML：{label}") from exc
    declared = re.search(r"<\?xml[^>]*encoding\s*=\s*['\"]([^'\"]+)['\"]", decoded, re.I)
    if "\x00" in decoded or declared and declared[1].lower().replace("-", "") not in {"utf8", "ascii", "usascii"}:
        raise InputError(f"仅支持UTF-8 XLSX XML编码：{label}")
    if re.search(r"<!\s*(DOCTYPE|ENTITY)", decoded, re.I):
        raise InputError(f"拒绝DTD或实体声明：{label}")
    try:
        return ET.fromstring(decoded)
    except ET.ParseError as exc:
        raise InputError(f"无效XML：{label}：{exc}") from exc


def read_xlsx(data, sheet_name):
    try:
        z = zipfile.ZipFile(io.BytesIO(data))
        infos = z.infolist()
        if len(infos) > 1000 or sum(i.file_size for i in infos) > MAX_EXPANDED:
            raise InputError("XLSX条目数或解压总量超限")
        names = set()
        trees = {}
        for item in infos:
            path = PurePosixPath(item.filename)
            if (item.filename in names or path.is_absolute() or ".." in path.parts
                    or "\\" in item.filename or ":" in item.filename
                    or stat.S_ISLNK(item.external_attr >> 16)):
                raise InputError("XLSX重复/危险路径或ZIP符号链接")
            names.add(item.filename)
            if item.flag_bits & 1:
                raise InputError("不支持加密XLSX")
            if item.file_size > MAX_XML or (item.file_size > 1024 and item.file_size > max(1, item.compress_size) * 200):
                raise InputError("XLSX条目大小或压缩比超限")
            lower = item.filename.lower()
            if "vbaproject" in lower or "externallinks/" in lower:
                raise InputError("拒绝含宏或外部工作簿链接的XLSX")
            if lower.endswith((".xml", ".rels")):
                tree = xml(z.read(item), item.filename)
                trees[item.filename] = tree
                if any(n.tag.rsplit("}", 1)[-1] == "f" for n in tree.iter()):
                    raise InputError(f"拒绝含公式的整个XLSX（即使存在缓存）：{item.filename}")
                for n in tree.iter():
                    if n.tag.rsplit("}", 1)[-1] == "Relationship" and n.get("TargetMode", "").lower() == "external":
                        raise InputError("拒绝含外部关系/外链的整个XLSX")
                    if "macroenabled" in n.get("ContentType", "").lower():
                        raise InputError("拒绝宏格式工作簿")
        workbook = trees.get("xl/workbook.xml")
        rels = trees.get("xl/_rels/workbook.xml.rels")
        if workbook is None or rels is None:
            raise InputError("XLSX缺少工作簿结构")
        matched = [s for s in workbook.findall("s:sheets/s:sheet", NS) if s.get("name") == sheet_name]
        if len(matched) != 1:
            raise InputError(f"找不到唯一表：{sheet_name}")
        rel = [r for r in rels if r.get("Id") == matched[0].get(RID)]
        if len(rel) != 1:
            raise InputError("XLSX工作表关系无效")
        target = rel[0].get("Target", "")
        if ".." in PurePosixPath(target).parts or "\\" in target or ":" in target:
            raise InputError("XLSX工作表关系路径无效")
        target = target.lstrip("/") if target.startswith("/") else "xl/" + target
        tree = trees.get(target)
        if tree is None:
            raise InputError("XLSX工作表XML缺失")
        strings = []
        if "xl/sharedStrings.xml" in trees:
            for si in trees["xl/sharedStrings.xml"].findall("s:si", NS):
                strings.append("".join(t.text or "" for t in si.findall(".//s:t", NS)))
        rows, count = {}, 0
        for row in tree.findall("s:sheetData/s:row", NS):
            rn = int(row.get("r", "0"))
            if rn < 1 or rn > MAX_ROWS or rn in rows:
                raise InputError("XLSX行号超限或重复")
            cells = {}
            for cell in row.findall("s:c", NS):
                address = cell.get("r", "")
                m = re.fullmatch(r"([A-Z]+)([1-9][0-9]*)", address)
                if not m or int(m[2]) != rn:
                    raise InputError("XLSX单元格地址无效")
                cn = colindex(m[1])
                if cn > MAX_COLS or cn in cells:
                    raise InputError("XLSX列号超限或重复")
                kind = cell.get("t", "n")
                v = cell.find("s:v", NS)
                raw = "" if v is None else (v.text or "")
                storage = raw
                if kind == "s":
                    try:
                        raw = strings[int(raw)]
                    except (ValueError, IndexError) as exc:
                        raise InputError("共享字符串索引无效") from exc
                    kind = "text"
                elif kind == "inlineStr":
                    raw = "".join(t.text or "" for t in cell.findall(".//s:t", NS))
                    kind = "text"
                elif kind == "str":
                    kind = "text"
                elif kind == "n":
                    kind = "number"
                else:
                    kind = "unsupported:" + kind
                if len(raw) > 10000:
                    raise InputError("XLSX单元格文本过长")
                cells[cn] = {"raw": raw, "type": kind, "storage_raw": storage, "style_index": cell.get("s")}
                count += 1
                if count > MAX_CELLS:
                    raise InputError("XLSX单元格总量超限")
            rows[rn] = cells
        return rows
    except (zipfile.BadZipFile, RuntimeError, ValueError, KeyError) as exc:
        if isinstance(exc, InputError):
            raise
        raise InputError(f"无法安全读取XLSX：{exc}") from exc


def read_csv(data):
    try:
        rows = {}
        count = 0
        for rn, row in enumerate(csv.reader(io.StringIO(data.decode("utf-8-sig"), newline=""), strict=True), 1):
            if rn > MAX_ROWS or len(row) > MAX_COLS or any(len(v) > 10000 for v in row):
                raise InputError("CSV行、列或单元格大小超限")
            count += len(row)
            if count > MAX_CELLS:
                raise InputError("CSV单元格总量超限")
            rows[rn] = {i: {"raw": v, "type": "text"} for i, v in enumerate(row, 1)}
        return rows
    except (UnicodeError, csv.Error) as exc:
        raise InputError(f"CSV须为有效UTF-8且引号结构完整：{exc}") from exc


def load_source(root, spec, config_source, pointer, role):
    if not isinstance(spec, dict):
        raise InputError("文件配置须为对象")
    path = spec.get("file")
    data = read_bytes(root, path)
    suffix = Path(path).suffix.lower()
    sheet = spec.get("sheet", "CSV") if suffix == ".xlsx" else "CSV"
    if suffix == ".csv":
        raw_rows = read_csv(data)
    elif suffix == ".xlsx":
        if not isinstance(sheet, str) or not sheet:
            raise InputError("XLSX必须明确sheet")
        raw_rows = read_xlsx(data, sheet)
    else:
        raise InputError("仅支持.csv或纯值.xlsx；不支持.xls/.xlsm/附件")
    header_row = spec.get("header_row", 1)
    if type(header_row) is not int or header_row not in raw_rows:
        raise InputError("header_row必须是存在的正整数行号")
    headers = {}
    for cn, cell in raw_rows[header_row].items():
        if cell["raw"]:
            if cell["type"] != "text" or cell["raw"] in headers:
                raise InputError("表头必须为唯一文本，不自动改写或去空格")
            headers[cell["raw"]] = cn
    columns, constants = spec.get("columns", {}), spec.get("constants", {})
    if not isinstance(columns, dict) or not isinstance(constants, dict):
        raise InputError("columns/constants必须为对象")
    if set(columns) & set(constants) or (set(columns) | set(constants)) - FIELDS:
        raise InputError("字段名不支持或同一字段同时配置列与常量")
    for field, header in columns.items():
        if not isinstance(header, str) or header not in headers:
            raise InputError(f"{path}列映射未找到：{field} -> {header}")
    if any(not isinstance(v, str) for v in constants.values()):
        raise InputError("所有配置常量须是字符串；明确零写为\"0\"")
    metadata = {"file": path, "sha256": digest(data), "sheet": sheet,
                "original_name": spec.get("original_name", path), "role": role}
    records = []
    all_rows = []
    for rn, cells in sorted(raw_rows.items()):
        all_rows.append({"row": rn, "cells": {colname(cn) + str(rn): c for cn, c in cells.items()}})
        if rn <= header_row or not any(c["raw"] != "" for c in cells.values()):
            continue
        fields = {}
        for field in FIELDS:
            if field in columns:
                cn = headers[columns[field]]
                cell = cells.get(cn, {"raw": "", "type": "missing"})
                fields[field] = {**metadata, "row": rn, "cell": colname(cn) + str(rn), **cell,
                                 "mapping_evidence": {**config_source, "pointer": f"{pointer}/columns/{field}"}}
            elif field in constants:
                fields[field] = {**config_source, "kind": "config_constant", "pointer": f"{pointer}/constants/{field}",
                                 "row": None, "cell": None, "raw": constants[field], "type": "text"}
            else:
                fields[field] = {**metadata, "row": rn, "cell": None, "raw": None, "type": "missing"}
        records.append({"row": rn, "file": path, "sheet": sheet, "fields": fields})
    return {**metadata, "rows": all_rows, "records": records, "spec": spec}


def raw(record, field):
    return record["fields"][field]["raw"]


def refs(records, fields):
    return [r["fields"][f] for r in records for f in fields]


def decimal_value(value):
    if not isinstance(value, str) or not re.fullmatch(r"[+-]?(?:\d+(?:\.\d*)?|\.\d+)", value.strip()):
        raise ValueError("须为十进制数字文本，无千位分隔/科学计数法/货币符号")
    text = value.strip()
    if len(text) > 60 or len(text.lstrip("+-").split(".")[0]) > 24 or ("." in text and len(text.split(".")[1]) > 24):
        raise ValueError("数字位数超出安全范围（整数和小数各最多24位）")
    d = Decimal(text)
    if not d.is_finite():
        raise ValueError("数值必须有限")
    return d


def ds(value):
    if isinstance(value, Fraction):
        return str(value.numerator) if value.denominator == 1 else f"{value.numerator}/{value.denominator}"
    return format(value, "f") if isinstance(value, Decimal) else value


def tax_is_explicit(value):
    if value == "excluded":
        return True
    if not isinstance(value, str):
        return False
    m = re.fullmatch(r"included_(\d+(?:\.\d+)?)pct", value)
    return bool(m and Decimal(m[1]) <= 100)


def instant(text, date_tz=None, end=False):
    if not isinstance(text, str) or not text:
        raise ValueError("缺少日期")
    if re.fullmatch(r"\d{4}-\d{2}-\d{2}", text):
        if not date_tz:
            raise ValueError("日期无时区，须配置date_timezone")
        tz = ZoneInfo(date_tz)
        day = datetime.strptime(text, "%Y-%m-%d").date()
        return datetime.combine(day, time.max if end else time.min, tzinfo=tz)
    result = datetime.fromisoformat(text.replace("Z", "+00:00"))
    if result.tzinfo is None or result.utcoffset() is None:
        raise ValueError("时间戳缺少时区")
    return result


def safe_text(value):
    if value is None:
        return ""
    s = str(value)
    # Controls are replaced only in presentation exports; original evidence stays intact.
    dangerous = bool(s and (s[0] in "=+-@" or ord(s[0]) < 32 or ord(s[0]) == 127))
    stripped = s.lstrip(" \t\r\n\v\f\ufeff")
    dangerous |= bool(stripped and stripped[0] in "=+-@")
    s = "".join(f"\\u{ord(c):04x}" if ord(c) < 32 or ord(c) == 127 else c for c in s)
    return "'" + s if dangerous else s


class Audit:
    def __init__(self, config, config_source, sources):
        self.config, self.config_source, self.sources = config, config_source, sources
        self.issues, self.results, self.matches = [], [], []
        self.as_of = instant(config["as_of"])

    def issue(self, code, message, evidence, demand_row=None, supplier=None, blocking=True):
        item = {"id": f"Q{len(self.issues)+1:04d}", "code": code, "message": message,
                "demand_row": demand_row, "supplier": supplier, "blocking": blocking, "evidence": evidence}
        self.issues.append(item)
        return item["id"]

    def identity(self, record, supplier=None):
        good = True
        for f in ("manufacturer", "mpn"):
            field = record["fields"][f]
            if field["raw"] is None or field["raw"] == "":
                good = False
            elif field["type"] != "text":
                self.issue("IDENTITY_NOT_TEXT", f"{f}是数值或不支持的单元格类型，不能恢复前导零或丢失精度；请改成原始文本后重试", [field], supplier=supplier)
                good = False
        return (raw(record, "manufacturer"), raw(record, "mpn")) if good else None

    def number(self, records, field, demand_row, label, *, minimum=None, positive=False, required=True):
        values = [raw(r, field) for r in records]
        if all(v in (None, "") for v in values):
            if required:
                self.issue("MISSING_" + field.upper(), f"{field}缺失，不能补零或套用包装量", refs(records, [field]), demand_row, label)
            return None
        if len(set(values)) != 1:
            self.issue("CONFLICT_" + field.upper(), f"{field}存在矛盾或部分缺失，须保留所有来源确认", refs(records, [field]), demand_row, label)
            return None
        try:
            if any(r["fields"][field]["type"].startswith("unsupported:") for r in records):
                raise ValueError("不支持该单元格类型（布尔、错误值或日期序列）")
            value = decimal_value(values[0])
            if positive and value <= 0 or minimum is not None and value < minimum:
                raise ValueError("必须为正数" if positive else f"不能小于{minimum}")
            return value
        except ValueError as exc:
            self.issue("INVALID_" + field.upper(), f"{field}非法：{exc}", refs(records, [field]), demand_row, label)
            return None

    def text(self, records, field, demand_row, label, *, required=True):
        values = [raw(r, field) for r in records]
        if len(set(values)) != 1:
            self.issue("CONFLICT_" + field.upper(), f"{field}来源相互冲突或部分缺失", refs(records, [field]), demand_row, label)
            return None
        value = values[0]
        if value in (None, ""):
            if required:
                self.issue("MISSING_" + field.upper(), f"{field}未明确，待客户确认", refs(records, [field]), demand_row, label)
            return None
        if value.strip().lower() in {"unknown", "n/a", "na", "null", "none", "tbd", "?", "待定", "待确认", "未知", "不详"}:
            self.issue("UNCONFIRMED_" + field.upper(), f"{field}为未知占位文本，不能作为明确业务口径", refs(records, [field]), demand_row, label)
            return None
        return value

    def factor(self, from_unit, to_unit, demand_row, label, evidence):
        if from_unit == to_unit:
            return Decimal(1), []
        selected = [r for r in self.config.get("unit_conversions", [])
                    if r.get("from") == from_unit and r.get("to") == to_unit and label in r.get("suppliers", [])]
        if len(selected) != 1:
            self.issue("UNIT_CONVERSION_REQUIRED", f"缺少唯一已授权单位关系：1 {from_unit}等于多少{to_unit}；不由包装量推断", evidence, demand_row, label)
            return None, []
        rule = selected[0]
        try:
            factor = decimal_value(rule.get("factor"))
            if factor <= 0 or not rule.get("source") or not rule.get("version") or rule.get("customer_confirmed") is not True:
                raise ValueError("单位关系须正数及客户确认、version和source")
        except ValueError as exc:
            self.issue("INVALID_UNIT_RULE", str(exc), [self.config_source, rule], demand_row, label)
            return None, []
        return factor, [{**rule, "config_evidence": {**self.config_source, "pointer": f"/unit_conversions/{self.config['unit_conversions'].index(rule)}"}}]

    def mapped_identities(self, demand, exact):
        identities = {} if exact is None else {exact: {"kind": "exact", "chain": []}}
        maps = self.config.get("customer_mappings")
        if not maps:
            return identities
        valid = (isinstance(maps, dict) and maps.get("customer_confirmed") is True and maps.get("version")
                 and maps.get("authorization") and maps.get("scope") == "this_config")
        if not valid:
            self.issue("MAPPING_NOT_AUTHORIZED", "客户映射缺少确认、版本、授权依据或本批次范围；不应用映射", [self.config_source], demand["row"])
            return identities
        for entry_index, entry in enumerate(maps.get("entries", [])):
            if entry.get("demand_row") != demand["row"] or entry.get("internal_id") != raw(demand, "internal_id"):
                continue
            chain = {"version": maps["version"], "authorization": maps["authorization"], "scope": "this_config", "entry": entry,
                     "config_evidence": {**self.config_source, "pointer": f"/customer_mappings/entries/{entry_index}"}, "original_demand": refs([demand], ["manufacturer", "mpn", "internal_id"])}
            candidates = [(entry, "customer_mapping")] + [(alt, "authorized_alternative") for alt in entry.get("alternatives", [])]
            for candidate, kind in candidates:
                key = (candidate.get("manufacturer"), candidate.get("mpn"))
                if not all(isinstance(v, str) and v != "" for v in key):
                    self.issue("INVALID_MAPPING", "映射目标身份须为完整非空文本", [self.config_source, entry], demand["row"])
                    continue
                if key not in identities:
                    identities[key] = {"kind": kind, "chain": [chain]}
        return identities

    def validity(self, rows, demand_row, label, spec):
        before = len(self.issues)
        until = self.text(rows, "valid_until", demand_row, label)
        start = self.text(rows, "valid_from", demand_row, label, required=False)
        state = "unconfirmed"
        if until:
            try:
                end = instant(until, self.config.get("date_timezone"), end=True)
                begin = instant(start, self.config.get("date_timezone")) if start else None
                if begin and begin > end:
                    raise ValueError("起始日期晚于截至日期")
                state = "expired" if self.as_of > end else "not_yet_valid" if begin and self.as_of < begin else "current"
                if state != "current":
                    self.issue("QUOTE_" + state.upper(), "报价已过期或尚未生效；仅展示原始货值，不参与当前可比金额", refs(rows, ["valid_from", "valid_until"]), demand_row, label)
            except (ValueError, ZoneInfoNotFoundError) as exc:
                self.issue("VALIDITY_UNCONFIRMED", f"有效期不能确定：{exc}", refs(rows, ["valid_from", "valid_until"]), demand_row, label)
        alternates = [r for r in rows if raw(r, "valid_until_alternate") not in (None, "")]
        if alternates:
            try:
                conflict = any(instant(raw(r, "valid_until_alternate"), self.config.get("date_timezone"), end=True)
                               != instant(until, self.config.get("date_timezone"), end=True) for r in alternates)
            except (ValueError, TypeError, ZoneInfoNotFoundError):
                conflict = True
            if conflict:
                state = "unconfirmed"
                self.issue("VALIDITY_CONFLICT", "主有效期与其他来源矛盾；不能挑选较晚日期", refs(rows, ["valid_until", "valid_until_alternate", "notes"]), demand_row, label)
        notes = [r for r in rows if raw(r, "notes") not in (None, "")]
        if notes and spec.get("notes_reviewed") is not True:
            state = "unconfirmed"
            self.issue("NOTES_REQUIRE_REVIEW", "备注含未确认条件。请逐项核对交期、有效期、价格及替代声明；仅客户已审阅且无未决冲突时才能设置notes_reviewed=true", refs(notes, ["notes"]), demand_row, label)
        return state if len(self.issues) == before or state in {"expired", "not_yet_valid"} else "unconfirmed"

    def delivery(self, rows, demand, label):
        values = {}
        for field in ("stock", "distributor_receipt", "manufacturer_lead_time", "customer_receipt", "customer_receipt_alternate"):
            values[field] = [{"value": raw(r, field), "evidence": r["fields"][field]} for r in rows]
        for field in ("customer_receipt", "distributor_receipt", "manufacturer_lead_time"):
            nonempty = {raw(r, field) for r in rows if raw(r, field) not in (None, "")}
            if len(nonempty) > 1:
                self.issue("DELIVERY_CONFLICT", f"{field}存在不同说法，不能作按期履约结论", refs(rows, [field, "notes"]), demand["row"], label, False)
        primaries = {raw(r, "customer_receipt") for r in rows if raw(r, "customer_receipt") not in (None, "")}
        alternates = {raw(r, "customer_receipt_alternate") for r in rows if raw(r, "customer_receipt_alternate") not in (None, "")}
        if alternates and alternates != primaries:
            self.issue("CUSTOMER_RECEIPT_CONFLICT", "客户收货承诺来源冲突，保留双方说法并人工确认", refs(rows, ["customer_receipt", "customer_receipt_alternate", "notes"]), demand["row"], label, False)
        if raw(demand, "required_receipt") not in (None, "") and not primaries:
            self.issue("CUSTOMER_RECEIPT_MISSING", "缺少客户收货承诺；库存/到库/厂商交期不能代替收货时间", refs([demand], ["required_receipt"]) + refs(rows, ["stock", "distributor_receipt", "manufacturer_lead_time", "customer_receipt"]), demand["row"], label, False)
        values["fulfilment_conclusion"] = "未作按期履约判断"
        return values

    def fx(self, amount, from_currency, target, demand_row, label, evidence):
        if from_currency == target:
            return amount, []
        rules = [r for r in self.config.get("fx", []) if r.get("from") == from_currency
                 and r.get("to") == target and label in r.get("suppliers", [])]
        if len(rules) != 1:
            self.issue("FX_REQUIRED", f"缺少唯一客户授权汇率：{from_currency}到{target}；不联网补汇率", evidence + [self.config_source], demand_row, label)
            return None, []
        r = rules[0]
        try:
            rate = decimal_value(r.get("rate"))
            if (rate <= 0 or r.get("direction") != "multiply" or r.get("customer_confirmed") is not True
                    or not r.get("source") or not r.get("version") or r.get("scope") != "this_config"):
                raise ValueError("汇率须正数、方向multiply（目标币/原币）、客户确认、source/version和本配置范围")
        except ValueError as exc:
            self.issue("INVALID_FX", str(exc), [self.config_source, r], demand_row, label)
            return None, []
        return amount * rate, [{"expression": f"{ds(amount)} {from_currency} * {ds(rate)} {target}/{from_currency}", "rule": r,
                                "config_evidence": {**self.config_source, "pointer": f"/fx/{self.config['fx'].index(r)}"}}]

    def calculate(self, demand, rows, supplier, match):
        label, rn = supplier["spec"]["label"], demand["row"]
        first_issue = len(self.issues)
        result = {"demand_row": rn, "supplier": label, "manufacturer": raw(rows[0], "manufacturer"),
                  "mpn": raw(rows[0], "mpn"), "match_kind": match["kind"], "mapping_chain": match["chain"],
                  "original_demand_identity": {k: raw(demand, k) for k in ("manufacturer", "mpn", "internal_id")},
                  "demand_quantity": raw(demand, "quantity"), "demand_unit": raw(demand, "quantity_unit"),
                  "quote_rows": [r["row"] for r in rows], "calculation": [], "evidence": refs([demand], ["manufacturer", "mpn", "internal_id", "quantity", "quantity_unit", "notes"])
                  + refs(rows, sorted(FIELDS)), "order_quantity": None, "extra_quantity": None, "goods_amount": None,
                  "currency": None, "comparable_amount": None, "comparison_currency": self.config.get("comparison", {}).get("currency"),
                  "calculation_status": "blocked", "comparability": "pending", "tax_processing": "未执行税费转换"}
        self.results.append(result)
        result["delivery"] = self.delivery(rows, demand, label)
        result["validity"] = self.validity(rows, rn, label, supplier["spec"])
        demand_notes_reviewed = (raw(demand, "notes") in (None, "")
                                 or self.sources[0]["spec"].get("notes_reviewed") is True)
        result["demand_notes_review"] = "confirmed_or_empty" if demand_notes_reviewed else "pending"
        if not demand_notes_reviewed:
            self.issue("NOTES_REQUIRE_REVIEW", "需求文件备注含未确认条件，保留货值但不形成完整可比金额；仅客户真实复核后可在demand配置设置notes_reviewed=true，源单元格或供应商声明不构成确认",
                       refs([demand], ["notes"]) + [{**self.config_source, "pointer": "/demand/notes_reviewed",
                                                    "raw": self.sources[0]["spec"].get("notes_reviewed")}], rn, label)
        demand_qty = self.number([demand], "quantity", rn, label, positive=True)
        demand_unit = self.text([demand], "quantity_unit", rn, label)
        quote_unit = self.text(rows, "quantity_unit", rn, label)
        moq = self.number(rows, "moq", rn, label, minimum=0)
        multiple = self.number(rows, "order_multiple", rn, label, positive=True)
        result["constraints"] = {f: [raw(r, f) for r in rows] for f in ("moq", "order_multiple", "minimum_release", "package_quantity")}
        qty_factor = None
        if demand_unit and quote_unit:
            qty_factor, unit_rules = self.factor(quote_unit, demand_unit, rn, label, refs(rows, ["quantity_unit"]) + refs([demand], ["quantity_unit"]))
            result["unit_rules"] = unit_rules
        order = None
        if all(x is not None for x in (demand_qty, qty_factor, moq, multiple)):
            try:
                requested_in_quote = Fraction(demand_qty) / Fraction(qty_factor)
                steps = max(requested_in_quote, Fraction(moq)) / Fraction(multiple)
                order = Decimal(-(-steps.numerator // steps.denominator)) * multiple
                result.update(order_quantity=ds(order), order_unit=quote_unit,
                              extra_quantity=ds(order * qty_factor - demand_qty), extra_unit=demand_unit)
                result["calculation"].append({"step": "order_quantity", "expression": f"ceil(max({ds(demand_qty)}/{ds(qty_factor)}, {ds(moq)})/{ds(multiple)})*{ds(multiple)} = {ds(order)} {quote_unit}",
                                               "evidence": refs([demand], ["quantity", "quantity_unit"]) + refs(rows, ["quantity_unit", "moq", "order_multiple"])})
            except (InvalidOperation, Inexact):
                self.issue("NONTERMINATING_QUANTITY", "单位换算出现非终止小数；v0不隐式舍入，需客户提供可精确表达的规则", refs(rows, ["quantity_unit", "order_multiple"]), rn, label)
        tier_basis = self.text(rows, "tier_basis", rn, label)
        if tier_basis not in (None, "order_quantity", "demand_quantity"):
            self.issue("UNSUPPORTED_TIER_BASIS", "tier_basis仅支持order_quantity或demand_quantity，单位为报价quantity_unit", refs(rows, ["tier_basis"]), rn, label)
        tiers, tiers_valid = [], True
        for row in rows:
            lo = self.number([row], "tier_min", rn, label, minimum=0)
            hi_raw = raw(row, "tier_max")
            hi = None if hi_raw == "unbounded" else self.number([row], "tier_max", rn, label, minimum=0)
            price = self.number([row], "unit_price", rn, label, minimum=0)
            denominator = self.number([row], "price_quantity", rn, label, positive=True)
            price_unit = self.text([row], "price_unit", rn, label)
            if lo is None or hi_raw != "unbounded" and hi is None or price is None or denominator is None or price_unit is None:
                tiers_valid = False
            if lo is not None and hi is not None and hi < lo:
                tiers_valid = False
                self.issue("INVALID_TIER_RANGE", "阶梯上界小于下界", refs([row], ["tier_min", "tier_max"]), rn, label)
            tiers.append({"row": row, "min": lo, "max": hi, "price": price, "denominator": denominator, "price_unit": price_unit})
        if tiers_valid:
            for i, t in enumerate(tiers):
                for u in tiers[i+1:]:
                    if (t["max"] is None or u["min"] <= t["max"]) and (u["max"] is None or t["min"] <= u["max"]):
                        tiers_valid = False
                        self.issue("TIER_OVERLAP", "阶梯重叠/重复断点，须确认版本与区间；不选择有利价格", refs([t["row"], u["row"]], ["tier_min", "tier_max", "unit_price", "price_quantity", "price_unit"]), rn, label)
        chosen = None
        if tiers_valid and order is not None and tier_basis in ("order_quantity", "demand_quantity"):
            basis = order if tier_basis == "order_quantity" else Fraction(demand_qty) / Fraction(qty_factor)
            selected = [t for t in tiers if basis >= t["min"] and (t["max"] is None or basis <= t["max"])]
            if len(selected) == 1:
                chosen = selected[0]
                result["selected_tier"] = {"row": chosen["row"]["row"], "min": ds(chosen["min"]), "max": ds(chosen["max"]) or "unbounded", "basis": tier_basis, "basis_quantity": ds(basis)}
            else:
                self.issue("NO_APPLICABLE_TIER", "已得订购量无唯一适用阶梯；不为更低价主动超购", refs(rows, ["tier_min", "tier_max", "tier_basis"]), rn, label)
        release = self.number([chosen["row"]], "minimum_release", rn, label, minimum=0) if chosen else None
        result["selected_minimum_release"] = ds(release)
        eligible_release = release is not None and order is not None and order >= release
        if release is not None and order is not None and order < release:
            self.issue("MINIMUM_RELEASE_NOT_MET", "订购量未满足所报价Minimum Release；不将其当MOQ自动超购", refs(rows, ["minimum_release"]), rn, label)
        currency = self.text(rows, "currency", rn, label)
        result["currency"] = currency
        goods = None
        if chosen and eligible_release:
            price_factor, price_rules = self.factor(chosen["price_unit"], quote_unit, rn, label,
                                                    refs([chosen["row"]], ["price_unit", "price_quantity", "quantity_unit"]))
            if price_factor is not None:
                try:
                    effective_denominator = chosen["denominator"] * price_factor
                    goods = order * chosen["price"] / effective_denominator
                    result.update(goods_amount=ds(goods), calculation_status="calculated", unit_price=ds(chosen["price"]),
                                  price_quantity=ds(chosen["denominator"]), price_unit=chosen["price_unit"], effective_price_denominator=ds(effective_denominator), price_unit_rules=price_rules)
                    result["calculation"].append({"step": "goods_amount", "expression": f"{ds(order)} {quote_unit} * {ds(chosen['price'])} / ({ds(chosen['denominator'])} * {ds(price_factor)}) = {ds(goods)} {currency or '币种待确认'}", "evidence": refs([chosen["row"]], ["unit_price", "price_quantity", "price_unit", "currency"])})
                except (Inexact, InvalidOperation):
                    self.issue("NONTERMINATING_AMOUNT", "金额为非终止小数或超出精度上限；不隐式舍入，需客户明确可用规则", refs([chosen["row"]], ["unit_price", "price_quantity"]), rn, label)
        tax = self.text(rows, "tax_basis", rn, label)
        result["tax_basis"] = tax
        freight = self.number(rows, "freight", rn, label, minimum=0)
        freight_currency = self.text(rows, "freight_currency", rn, label)
        scope = self.text(rows, "freight_scope", rn, label)
        freight_tax = self.text(rows, "freight_tax_basis", rn, label, required=freight is not None and freight > 0)
        result.update(freight_amount=ds(freight), freight_currency=freight_currency, freight_scope=scope, freight_tax_basis=freight_tax)
        comparison = self.config.get("comparison", {})
        target, target_tax, target_freight = (comparison.get(k) for k in ("currency", "tax_basis", "freight_basis"))
        complete = demand_notes_reviewed
        if not all(isinstance(v, str) and v for v in (target, target_tax, target_freight)) or target_freight != "goods_plus_freight":
            complete = False
            self.issue("COMPARISON_BASIS_REQUIRED", "目标currency/tax_basis/freight_basis未明确；v0只支持goods_plus_freight且同税口径", [self.config_source], rn, label)
        if not tax_is_explicit(tax) or not tax_is_explicit(target_tax) or tax != target_tax:
            complete = False
            self.issue("TAX_BASIS_UNCONFIRMED", "税口径与目标不同或缺失；保留原货值，v0不自动转换税费", refs(rows, ["tax_basis"]) + [self.config_source], rn, label)
        demand_count = len(self.sources[0]["records"])
        if scope not in ("line", "included", "order") or scope == "order" and demand_count != 1:
            complete = False
            self.issue("FREIGHT_SCOPE_UNCONFIRMED", "运费范围须明确；订单级运费在多需求行时不自动分摊", refs(rows, ["freight", "freight_scope"]), rn, label)
        if scope == "included" and freight != Decimal(0):
            complete = False
            self.issue("FREIGHT_INCLUDED_CONFLICT", "freight_scope=included时运费须明确0，不能二次加费", refs(rows, ["freight", "freight_scope"]), rn, label)
        if freight is None or not freight_currency or not currency:
            complete = False
        if freight is not None and freight > 0 and (not tax_is_explicit(freight_tax) or freight_tax != target_tax):
            complete = False
            self.issue("FREIGHT_TAX_BASIS_UNCONFIRMED", "非零运费的税口径未明确或与目标不一致，不与货值混加", refs(rows, ["freight", "freight_tax_basis"]) + [self.config_source], rn, label)
        if goods is not None and currency and target and freight is not None and freight_currency:
            converted_goods, goods_fx = self.fx(goods, currency, target, rn, label, refs(rows, ["currency", "unit_price"]))
            converted_freight, freight_fx = self.fx(freight, freight_currency, target, rn, label, refs(rows, ["freight", "freight_currency"]))
            result["fx_calculation"] = goods_fx + freight_fx
            if converted_goods is not None:
                result["target_currency_goods_amount"] = ds(converted_goods)
            if complete and converted_goods is not None and converted_freight is not None and result["validity"] == "current":
                total = converted_goods + converted_freight
                result.update(comparable_amount=ds(total), comparability="comparable")
                result["calculation"].append({"step": "comparable_amount", "expression": f"{ds(converted_goods)} + {ds(converted_freight)} = {ds(total)} {target}；税口径 {tax}；未执行税费转换", "evidence": refs(rows, ["tax_basis", "freight", "freight_currency", "freight_scope", "valid_until"]) + [self.config_source]})
        result["clarification_ids"] = [i["id"] for i in self.issues[first_issue:]]
        return result

    def run(self):
        demand_source = self.sources[0]
        supplier_index = []
        for supplier in self.sources[1:]:
            index = {}
            for record in supplier["records"]:
                identity = self.identity(record, supplier["spec"]["label"])
                if identity:
                    index.setdefault(identity, []).append(record)
                    if len(index[identity]) > 100:
                        raise InputError("同一供应商同一严格身份最多100条阶梯记录；请先人工清理重复/冲突版本")
                else:
                    self.issue("SUPPLIER_IDENTITY_UNUSABLE", "供应商身份缺失或非文本，该行不参与匹配；不能用备注恢复", refs([record], ["manufacturer", "mpn", "notes"]), supplier=supplier["spec"]["label"])
            supplier_index.append((supplier, index))
        for demand in demand_source["records"]:
            identity = self.identity(demand)
            identities = self.mapped_identities(demand, identity)
            if not identities:
                self.issue("DEMAND_MAPPING_REQUIRED", "需求制造商/完整料号缺失；客户内部号、替代料与认可厂商须客户授权配置，供应商声明无效", refs([demand], ["manufacturer", "mpn", "internal_id"]), demand["row"])
            for supplier, index in supplier_index:
                matched = [(key, index[key], mapping) for key, mapping in identities.items() if key in index]
                if not matched:
                    scope = {k: supplier[k] for k in ("file", "sha256", "sheet")}
                    scope.update(kind="searched_scope", pointer=f"/inputs/{self.sources.index(supplier)}/records", fields=["manufacturer", "mpn", "notes"])
                    self.matches.append({"demand_row": demand["row"], "supplier": supplier["spec"]["label"], "status": "no_match", "demand_identity": refs([demand], ["manufacturer", "mpn", "internal_id"]), "supplier_identity_evidence": [scope]})
                    self.issue("NO_AUTHORIZED_MATCH", "无严格身份或客户明确授权映射匹配；不推断相似/兼容/替代", refs([demand], ["manufacturer", "mpn", "internal_id"]) + [scope], demand["row"], supplier["spec"]["label"])
                    continue
                for key, rows, mapping in matched:
                    self.matches.append({"demand_row": demand["row"], "supplier": supplier["spec"]["label"], "status": "matched", "match_kind": mapping["kind"], "quote_rows": [r["row"] for r in rows], "identity_evidence": refs([demand] + rows, ["manufacturer", "mpn", "internal_id"]), "mapping_chain": mapping["chain"]})
                    self.calculate(demand, rows, supplier, mapping)
        used = {(m["supplier"], r) for m in self.matches if m["status"] == "matched" for r in m["quote_rows"]}
        exclusions = []
        for supplier, _ in supplier_index:
            label = supplier["spec"]["label"]
            for record in supplier["records"]:
                if (label, record["row"]) not in used:
                    exclusions.append({"supplier": label, "row": record["row"], "reason": "未获严格身份或客户授权匹配；未参与金额", "evidence": refs([record], ["manufacturer", "mpn", "notes"])})
        return {"schema_version": "electronic-quote-audit-result-v1", "rules_version": RULES_VERSION,
                "config_version": self.config["config_version"], "config_evidence": self.config_source,
                "as_of": self.config["as_of"], "date_timezone": self.config.get("date_timezone"),
                "date_boundary_rule": "ISO日期使用显式date_timezone，当地自然日截至23:59:59.999999；含时区时间戳截至端点包含。无隐式系统时间。",
                "decimal_policy": "十进制定点；不隐式舍入；非终止小数阻止该计算。整数和小数各最多24位。",
                "ranking": "不排名、不作最低价或采购建议；不同采购行不合计为订单总价",
                "configuration": self.config, "inputs": [{k: v for k, v in source.items() if k != "spec"} for source in self.sources],
                "matches": self.matches, "excluded_quotes": exclusions,
                "comparison": self.results, "clarifications": self.issues,
                "summary": {"demand_lines": len(demand_source["records"]), "supplier_count": len(self.sources) - 1,
                            "calculated_candidates": sum(r["goods_amount"] is not None for r in self.results),
                            "comparable_candidates": sum(r["comparable_amount"] is not None for r in self.results),
                            "clarification_count": len(self.issues)}}


def validate_config(config):
    if not isinstance(config, dict) or config.get("schema_version") != SCHEMA_VERSION:
        raise InputError(f"schema_version须为{SCHEMA_VERSION}")
    if not isinstance(config.get("config_version"), str) or not config["config_version"]:
        raise InputError("必须提供config_version")
    if config.get("mapping_confirmed") is not True:
        raise InputError("人工待配置：先确认文件/列映射和单位语义，再设置mapping_confirmed=true")
    try:
        instant(config.get("as_of"))
    except (ValueError, TypeError, ZoneInfoNotFoundError) as exc:
        raise InputError(f"as_of必须是显式含时区时间戳：{exc}") from exc
    if not re.search(r"T", config.get("as_of", "")):
        raise InputError("as_of须明确时间及偏移，不接受纯日期")
    if "date_timezone" in config:
        try:
            ZoneInfo(config["date_timezone"])
        except (ZoneInfoNotFoundError, TypeError) as exc:
            raise InputError("date_timezone不是可用IANA时区") from exc
    suppliers = config.get("suppliers")
    if not isinstance(suppliers, list) or not 2 <= len(suppliers) <= 5:
        raise InputError("须显式配置2–5家供应商")
    labels = []
    for spec in suppliers:
        label = spec.get("label") if isinstance(spec, dict) else None
        if not isinstance(label, str) or not label.strip() or len(label) > 80:
            raise InputError("供应商label须为独立配置的1–80字符文本")
        labels.append(label)
    if len(set(labels)) != len(labels):
        raise InputError("供应商label不能重复")
    for key in ("unit_conversions", "fx"):
        if key in config and (not isinstance(config[key], list) or not all(isinstance(x, dict) for x in config[key])):
            raise InputError(f"{key}须为对象数组")
    if config.get("rounding") not in (None, "exact"):
        raise InputError("v0仅支持rounding=exact；不自动应用其他舍入策略")
    maps = config.get("customer_mappings")
    if maps is not None:
        if not isinstance(maps, dict) or not isinstance(maps.get("entries", []), list):
            raise InputError("customer_mappings必须是对象，entries必须为数组")
        seen = set()
        for entry in maps.get("entries", []):
            if (not isinstance(entry, dict) or type(entry.get("demand_row")) is not int or entry["demand_row"] < 2
                    or not isinstance(entry.get("internal_id"), str) or not entry["internal_id"]
                    or not isinstance(entry.get("alternatives", []), list)
                    or not all(isinstance(a, dict) for a in entry.get("alternatives", []))):
                raise InputError("映射entry须明确demand_row、非空internal_id，alternatives须对象数组")
            if entry["demand_row"] in seen:
                raise InputError("同一需求行不能有多份映射entry，先解决授权冲突")
            seen.add(entry["demand_row"])


def export_csv(path, headers, rows):
    with path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.writer(handle)
        writer.writerow([safe_text(h) for h in headers])
        for row in rows:
            writer.writerow([safe_text(row.get(h)) for h in headers])


def output(root, relative, audit):
    path = confined(root, relative, exists=False)
    if path.exists():
        raise InputError("输出目录必须全新且不存在，避免覆盖")
    # Make only the requested fresh leaf. Parent must already exist and be confined.
    if not path.parent.is_dir():
        raise InputError("输出目录的父目录须已存在；请显式创建")
    path.mkdir(exist_ok=False)
    (path / "audit.json").write_text(jdump(audit), encoding="utf-8")
    headers = ["demand_row", "supplier", "manufacturer", "mpn", "match_kind", "demand_quantity", "demand_unit",
               "order_quantity", "order_unit", "extra_quantity", "extra_unit", "selected_tier", "unit_price",
               "price_quantity", "price_unit", "goods_amount", "currency", "tax_basis", "freight_amount",
               "freight_currency", "freight_scope", "validity", "comparable_amount", "comparison_currency",
               "comparability", "calculation", "delivery", "clarification_ids"]
    rows = []
    for result in audit["comparison"]:
        rows.append({k: json.dumps(v, ensure_ascii=False, sort_keys=True) if isinstance(v, (list, dict)) else v for k, v in result.items()})
    export_csv(path / "comparison.csv", headers, rows)
    issue_rows = [{**i, "evidence": json.dumps(i["evidence"], ensure_ascii=False, sort_keys=True)} for i in audit["clarifications"]]
    export_csv(path / "clarifications.csv", ["id", "demand_row", "supplier", "code", "blocking", "message", "evidence"], issue_rows)
    e = lambda value: html.escape(str(value if value is not None else "待确认"), quote=True)
    parts = ['<!doctype html><html lang="zh-CN"><meta charset="utf-8">',
             '<meta http-equiv="Content-Security-Policy" content="default-src \'none\'; style-src \'unsafe-inline\'; img-src \'none\'; form-action \'none\'; base-uri \'none\'">',
             '<meta name="viewport" content="width=device-width,initial-scale=1"><title>采购报价核对</title>',
             '<style>body{font:16px system-ui;margin:2rem;color:#172c39;background:#f5f7fa}table{border-collapse:collapse;background:white;width:100%}th,td{border:1px solid #ccd5de;padding:.6rem;text-align:left;vertical-align:top}th{background:#e1eaf1}article{background:white;margin:1rem 0;padding:1rem;border-radius:.6rem}pre{white-space:pre-wrap;overflow-wrap:anywhere;font:13px monospace}small{color:#4c6272}h1{font-size:1.6rem}</style>',
             '<h1>电子元件采购报价核对</h1>',
             f'<p>固定评估时间：{e(audit["as_of"])} · 规则：{e(audit["rules_version"])} · 配置：{e(audit["config_version"])}</p>',
             '<p>按供应商配置顺序展示，不排名、不作采购建议。货值与含运可比金额分别列出；待确认表示条件不全。库存、到库与厂商交期均不代表客户收货承诺。</p>',
             '<table><tr><th>需求行 / 供应商</th><th>原始制造商 / 完整料号</th><th>订购 / 多购</th><th>货值</th><th>含运可比金额</th><th>有效性</th></tr>']
    for r in audit["comparison"]:
        parts.append(f'<tr><td>{e(r["demand_row"])} / {e(r["supplier"])}</td><td>{e(r["manufacturer"])} / {e(r["mpn"])}</td><td>{e(r.get("order_quantity"))} {e(r.get("order_unit"))} / {e(r.get("extra_quantity"))} {e(r.get("extra_unit"))}</td><td>{e(r["goods_amount"])} {e(r["currency"])}</td><td>{e(r["comparable_amount"])} {e(r["comparison_currency"])}</td><td>{e(r["validity"])}</td></tr>')
    parts.append('</table><h2>逐项待澄清</h2>')
    for issue in audit["clarifications"]:
        parts.append(f'<article><strong>{e(issue["id"])} · 需求行{e(issue["demand_row"])} · {e(issue["supplier"])}</strong><p>{e(issue["message"])}</p><details><summary>查看原始来源</summary><pre>{e(jdump(issue["evidence"]))}</pre></details></article>')
    parts.append('<h2>计算与交期证据</h2>')
    for r in audit["comparison"]:
        parts.append(f'<article><strong>需求行{e(r["demand_row"])} · {e(r["supplier"])}</strong><pre>{e(jdump({k:r.get(k) for k in ("match_kind", "mapping_chain", "constraints", "selected_tier", "calculation", "fx_calculation", "delivery")}))}</pre></article>')
    parts.append('<p><small>完整未转义原值、文件SHA-256、表名与单元格位于audit.json。CSV仅对导出文本作公式注入防护，内部匹配不使用转义值。此页离线、无脚本、无网络资源。</small></p></html>')
    (path / "review.html").write_text("\n".join(parts), encoding="utf-8")
    return path


def run(root_path, config_path, output_path):
    root_arg = Path(root_path).absolute()
    if any(p.is_symlink() for p in [root_arg, *root_arg.parents]):
        raise InputError("根目录及其父路径不能为符号链接")
    root = root_arg.resolve()
    if not root.is_dir():
        raise InputError("指定根目录不存在")
    data = read_bytes(root, config_path)
    config = strict_json(data)
    validate_config(config)
    config_source = {"file": config_path, "sha256": digest(data), "kind": "configuration", "pointer": ""}
    sources = [load_source(root, config.get("demand"), config_source, "/demand", "demand")]
    sources += [load_source(root, spec, config_source, f"/suppliers/{i}", "supplier") for i, spec in enumerate(config["suppliers"])]
    if not sources[0]["records"]:
        raise InputError("需求文件没有数据行")
    with localcontext() as ctx:
        ctx.prec = 256
        ctx.traps[Inexact] = True
        audit = Audit(config, config_source, sources).run()
    audit["implementation_sha256"] = digest(Path(__file__).read_bytes())
    return output(root, output_path, audit), audit


def main(argv=None):
    p = argparse.ArgumentParser(description="离线电子元件报价核对；先确认列映射，保留每字段证据；不排名/不采购")
    p.add_argument("--root", required=True, help="仅在此本地目录读取输入")
    p.add_argument("--config", required=True, help="root下相对JSON路径")
    p.add_argument("--output", required=True, help="root下全新输出目录相对路径；父目录须存在")
    args = p.parse_args(argv)
    try:
        path, audit = run(args.root, args.config, args.output)
    except (InputError, OSError, InvalidOperation, TypeError, AttributeError, OverflowError) as exc:
        print(f"未生成比价结果：{exc}", file=sys.stderr)
        return 2
    print(f"已生成：{path}\n可算货值候选{audit['summary']['calculated_candidates']}；完整可比候选{audit['summary']['comparable_candidates']}；待澄清{audit['summary']['clarification_count']}。无排名或采购建议。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
