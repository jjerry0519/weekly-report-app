from __future__ import annotations

import datetime as dt

import server


def mops_subject_search_text(record: dict[str, str], date_value: dt.date, keyword: str) -> str:
    """Query MOPS major-announcement subject full-text search."""
    market_kinds = ["C"]
    company_type = record.get("公司型態", "")
    if "上市" in company_type:
        market_kinds.insert(0, "L")
    elif "上櫃" in company_type:
        market_kinds.insert(0, "O")

    texts: list[str] = []
    for kind in dict.fromkeys(market_kinds):
        params = {
            "encodeURIComponent": "1",
            "step": "1",
            "firstin": "true",
            "id": "",
            "key": "",
            "TYPEK": "",
            "Stp": "4",
            "go": "false",
            "co_id": record.get("證券代號", ""),
            "COMPANY_ID": record.get("證券代號", ""),
            "r1": "1",
            "KIND": kind,
            "CODE": "",
            "keyWord": keyword,
            "Condition2": "2",
            "keyWord2": "",
            "year": f"{server.roc_year(date_value):03d}",
            "month1": str(date_value.month),
            "begin_day": "1",
            "end_day": "31",
            "Orderby": "1",
        }
        query_text = server.mops_query_text(("/mops/web/ajax_t51sb10", "/mops/web/t51sb10"), params)
        texts.append(query_text)
    return "\n".join(text for text in texts if text)




_original_parse_bond_short = server.parse_bond_short_from_announcement


def patched_parse_bond_short_from_announcement(text: str, record: dict[str, str], company_short: str = "") -> str:
    found = _original_parse_bond_short(text, record, company_short)
    if found and server.is_bond_product_name(found):
        return found

    compact = server.normalize_header(server.html_to_text(text))
    if not compact:
        return ""

    short = (
        company_short
        or server.public_company_short_name(record.get("證券代號", ""))
        or record.get("顯示名稱", "")
        or record.get("公司名稱", "")
    )
    short = server.normalize_stock_short_for_bond(short)
    ordinal = ""
    patterns = (
        r"第([一二三四五六七八九十百]+)次[^，。；;]{0,24}(?:轉換公司債|交換公司債|公司債)",
        r"國內([一二三四五六七八九十百]+)次[^，。；;]{0,24}(?:轉換公司債|交換公司債|公司債)",
        r"(?:轉換公司債|交換公司債|公司債)[^，。；;]{0,18}第([一二三四五六七八九十百]+)次",
    )
    for pattern in patterns:
        match = server.re.search(pattern, compact)
        if match:
            ordinal = server.chinese_ordinal_to_short(match.group(1))
            break
    if short and ordinal:
        return server.bond_name_with_ordinal(short, ordinal, record)
    return found


server.mops_subject_search_text = mops_subject_search_text
server.parse_bond_short_from_announcement = patched_parse_bond_short_from_announcement


if __name__ == "__main__":
    server.main()
