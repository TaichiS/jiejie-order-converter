#!/usr/bin/env python3
"""Fetch today's SHOPLINE orders and export a JJ official pre-conversion file."""
from __future__ import annotations

import argparse
import json
import os
import re
import sys
from collections import Counter
from datetime import datetime, time, timezone
from pathlib import Path
from typing import Any
from urllib.error import HTTPError, URLError
from urllib.parse import urlencode
from urllib.request import Request, urlopen
from zoneinfo import ZoneInfo

from openpyxl import Workbook


API_BASE_URL = "https://open.shopline.io/v1"
DEFAULT_USER_AGENT = "JiejieOrderConverter/0.1"
LOCAL_TZ = ZoneInfo("Asia/Taipei")
DEFAULT_OUTPUT_DIR = Path("測試資料")

JJ_OFFICIAL_INPUT_HEADER = [
    "訂單號碼",
    "訂單狀態",
    "付款狀態",
    "收件人",
    "完整地址",
    "收件人電話號碼",
    "發票號碼",
    "商品貨號",
    "商品名稱",
    "數量",
    "商品結帳價",
    "商品折扣優惠",
    "商品折扣金額",
    "點數折現分攤",
    "出貨備註",
    "送貨編號",
    "付款方式",
    "全家服務編號 / 7-11 店號",
    "加購品類型",
    "訂單備註",
    "發票開立日期",
    "運費",
]


def load_dotenv(path: Path) -> None:
    if not path.exists():
        return

    for raw_line in path.read_text(encoding="utf-8").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue

        key, value = line.split("=", 1)
        key = key.strip()
        value = value.strip().strip("'\"")
        if key and key not in os.environ:
            os.environ[key] = value


def utc_window_for_local_date(date_text: str | None) -> tuple[str, str, str]:
    if date_text:
        local_date = datetime.strptime(date_text, "%Y-%m-%d").date()
    else:
        local_date = datetime.now(LOCAL_TZ).date()

    start_local = datetime.combine(local_date, time.min, LOCAL_TZ)
    end_local = datetime.combine(local_date, time.max.replace(microsecond=0), LOCAL_TZ)

    start_utc = start_local.astimezone(timezone.utc).strftime("%Y-%m-%d %H:%M:%S")
    end_utc = end_local.astimezone(timezone.utc).strftime("%Y-%m-%d %H:%M:%S")
    return local_date.isoformat(), start_utc, end_utc


def get_nested(data: dict[str, Any], *keys: str) -> Any:
    current: Any = data
    for key in keys:
        if not isinstance(current, dict):
            return None
        current = current.get(key)
    return current


def money_dollars(value: Any) -> float:
    if not isinstance(value, dict):
        return 0.0
    dollars = value.get("dollars")
    if dollars is not None:
        try:
            return float(dollars)
        except (TypeError, ValueError):
            return 0.0
    cents = value.get("cents")
    try:
        return float(cents) if cents is not None else 0.0
    except (TypeError, ValueError):
        return 0.0


def money_label(value: Any) -> str:
    if not isinstance(value, dict):
        return ""
    label = value.get("label")
    if label:
        return str(label)
    dollars = money_dollars(value)
    currency = value.get("currency_iso") or ""
    return f"{currency} {dollars}" if dollars is not None else ""


def request_json(url: str, token: str, user_agent: str) -> dict[str, Any]:
    request = Request(
        url,
        headers={
            "Accept": "application/json",
            "Authorization": f"Bearer {token}",
            "User-Agent": user_agent,
        },
        method="GET",
    )

    try:
        with urlopen(request, timeout=30) as response:
            body = response.read().decode("utf-8")
    except HTTPError as exc:
        body = exc.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"SHOPLINE API 回應 HTTP {exc.code}: {body}") from exc
    except URLError as exc:
        raise RuntimeError(f"無法連線到 SHOPLINE API: {exc.reason}") from exc

    try:
        return json.loads(body)
    except json.JSONDecodeError as exc:
        raise RuntimeError(f"SHOPLINE API 回應不是 JSON: {body[:500]}") from exc


def fetch_orders(
    token: str,
    user_agent: str,
    created_after: str,
    created_before: str,
    per_page: int,
    max_pages: int,
) -> list[dict[str, Any]]:
    orders: list[dict[str, Any]] = []
    previous_id: str | None = None

    for page in range(1, max_pages + 1):
        params: dict[str, Any] = {
            "created_after": created_after,
            "created_before": created_before,
            "per_page": per_page,
            "sort_by": "asc",
        }
        if previous_id:
            params["previous_id"] = previous_id
        else:
            params["page"] = 1

        url = f"{API_BASE_URL}/orders?{urlencode(params)}"
        data = request_json(url, token, user_agent)
        items = data.get("items") or []
        if not isinstance(items, list):
            raise RuntimeError(f"SHOPLINE API 回應格式不符，items 不是 list: {type(items).__name__}")

        orders.extend(items)
        if len(items) < per_page:
            break

        last_id = items[-1].get("id")
        if not last_id:
            break
        previous_id = str(last_id)

    return orders


def summarize_order(order: dict[str, Any]) -> dict[str, Any]:
    payment_name = get_nested(order, "order_payment", "name_translations", "zh-hant")
    delivery_name = get_nested(order, "order_delivery", "name_translations", "zh-hant")

    return {
        "id": order.get("id"),
        "order_number": order.get("order_number"),
        "status": order.get("status"),
        "created_at": order.get("created_at"),
        "payment": payment_name or get_nested(order, "order_payment", "payment_type"),
        "payment_status": get_nested(order, "order_payment", "status"),
        "delivery": delivery_name or get_nested(order, "order_delivery", "platform"),
        "delivery_status": get_nested(order, "order_delivery", "delivery_status"),
        "total": money_label(order.get("total")),
        "items_count": len(order.get("subtotal_items") or []),
        "has_order_remarks": bool(order.get("order_remarks")),
    }


def local_datetime_text(value: Any) -> str:
    if not value:
        return ""
    text = str(value)
    try:
        dt = datetime.fromisoformat(text.replace("Z", "+00:00"))
    except ValueError:
        return text[:10]
    return dt.astimezone(LOCAL_TZ).strftime("%Y/%m/%d")


def translation_text(value: Any, default: str = "") -> str:
    if isinstance(value, dict):
        for key in ("zh-hant", "zh-tw", "zh-cn", "en"):
            if value.get(key):
                return str(value[key])
    return default


def clean_text(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip()


def is_addon_item(item: dict[str, Any]) -> bool:
    item_type = clean_text(item.get("item_type")).lower()
    title = translation_text(item.get("title_translations"), clean_text(item.get("title")))
    return item_type in {"addonproduct", "addon_product", "addon"} or "加購" in title


def full_address(order: dict[str, Any]) -> str:
    address = order.get("delivery_address") or {}
    parts = [
        address.get("country") or "台灣",
        address.get("postcode"),
        address.get("city") or address.get("state"),
        address.get("district"),
        address.get("address_1"),
        address.get("address_2"),
    ]
    return " ".join(clean_text(part) for part in parts if clean_text(part))


def invoice_number(order: dict[str, Any]) -> str:
    invoice = order.get("invoice") or {}
    if invoice.get("invoice_number"):
        return str(invoice["invoice_number"])
    invoices = order.get("invoices") or []
    if invoices and isinstance(invoices[0], dict):
        return clean_text(invoices[0].get("invoice_number"))
    return ""


def invoice_date(order: dict[str, Any]) -> str:
    invoice = order.get("invoice") or {}
    if invoice.get("invoice_date"):
        return local_datetime_text(invoice["invoice_date"])
    invoices = order.get("invoices") or []
    if invoices and isinstance(invoices[0], dict) and invoices[0].get("invoice_date"):
        return local_datetime_text(invoices[0]["invoice_date"])
    return local_datetime_text(order.get("created_at"))


def payment_name(order: dict[str, Any]) -> str:
    payment = order.get("order_payment") or {}
    return translation_text(payment.get("name_translations"), clean_text(payment.get("payment_type")))


def item_name(item: dict[str, Any]) -> str:
    return translation_text(item.get("title_translations"), clean_text(item.get("title")))


def item_unit_price(item: dict[str, Any]) -> float:
    quantity = item.get("quantity") or 1
    try:
        qty = float(quantity) or 1.0
    except (TypeError, ValueError):
        qty = 1.0

    total = money_dollars(item.get("total")) or money_dollars(item.get("discounted_total"))
    if total:
        return round(total / qty, 2)
    return money_dollars(item.get("item_price")) or money_dollars(item.get("price_sale")) or money_dollars(item.get("price"))


def item_discount(item: dict[str, Any]) -> float:
    total = money_dollars(item.get("total"))
    discounted_total = money_dollars(item.get("discounted_total"))
    if total and discounted_total and total > discounted_total:
        return round(total - discounted_total, 2)
    return 0.0


def order_note(order: dict[str, Any]) -> str:
    notes = []
    for key in ("order_remarks", "order_notes", "order_comments"):
        value = order.get(key)
        if isinstance(value, str) and value.strip():
            notes.append(value.strip())
    return "；".join(dict.fromkeys(notes))


def shipping_note(order: dict[str, Any]) -> str:
    values = [
        get_nested(order, "order_delivery", "remark"),
        get_nested(order, "delivery_address", "remarks"),
    ]
    return "；".join(dict.fromkeys(clean_text(v) for v in values if clean_text(v)))


def to_jjofficial_rows(orders: list[dict[str, Any]], include_cancelled: bool) -> list[list[Any]]:
    rows: list[list[Any]] = []

    for order in orders:
        if not include_cancelled and order.get("status") == "cancelled":
            continue

        delivery_data = order.get("delivery_data") or {}
        address = order.get("delivery_address") or {}
        payment = order.get("order_payment") or {}
        items = order.get("subtotal_items") or []
        freight = money_dollars(get_nested(order, "order_delivery", "total"))
        points = money_dollars(order.get("order_points_to_cash"))

        for index, item in enumerate(items):
            rows.append([
                order.get("order_number"),
                order.get("status"),
                payment.get("status"),
                address.get("recipient_name"),
                full_address(order),
                address.get("recipient_phone"),
                invoice_number(order),
                item.get("sku"),
                item_name(item),
                item.get("quantity") or 1,
                item_unit_price(item),
                "",
                item_discount(item),
                points if index == 0 else 0,
                shipping_note(order),
                delivery_data.get("tracking_number") or "",
                payment_name(order),
                delivery_data.get("location_code") or "",
                "主商品加購品" if is_addon_item(item) else "",
                order_note(order),
                invoice_date(order),
                freight if index == 0 else 0,
            ])

    return rows


def write_json(path: Path, orders: list[dict[str, Any]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(
        json.dumps({"items": orders}, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def write_jjofficial_xlsx(path: Path, rows: list[list[Any]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    wb = Workbook()
    ws = wb.active
    ws.title = "Sales"
    ws.append(JJ_OFFICIAL_INPUT_HEADER)
    for row in rows:
        ws.append(row)
    wb.save(path)


def safe_date_for_filename(date_text: str) -> str:
    return re.sub(r"\D", "", date_text)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="抓取今日 SHOPLINE 訂單資料")
    parser.add_argument("--date", help="指定台北日期 YYYY-MM-DD；未指定時使用今天")
    parser.add_argument("--per-page", type=int, default=50, help="每頁筆數，最大建議 50")
    parser.add_argument("--max-pages", type=int, default=10, help="最多抓取頁數，避免測試時無限抓取")
    parser.add_argument("--show", type=int, default=10, help="顯示前 N 筆摘要；設為 0 可隱藏明細")
    parser.add_argument("--output-dir", type=Path, default=DEFAULT_OUTPUT_DIR, help="輸出資料夾，預設為相對路徑 測試資料")
    parser.add_argument("--include-cancelled", action="store_true", help="整理 xlsx 時包含已取消訂單")
    parser.add_argument("--output-json", type=Path, help="自訂原始 JSON 路徑；未指定時寫入 output-dir")
    parser.add_argument("--output-xlsx", type=Path, help="自訂捷捷官網轉換前 xlsx 路徑；未指定時寫入 output-dir")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    load_dotenv(Path(".env"))

    token = os.environ.get("SHOPLINE_ACCESS_TOKEN")
    if not token:
        print("缺少 SHOPLINE_ACCESS_TOKEN，請先放到 .env 或環境變數。", file=sys.stderr)
        return 2

    user_agent = os.environ.get("SHOPLINE_USER_AGENT", DEFAULT_USER_AGENT)
    per_page = max(1, min(args.per_page, 50))
    local_date, created_after, created_before = utc_window_for_local_date(args.date)

    print(f"查詢台北日期：{local_date}")
    print(f"UTC 區間：{created_after} ~ {created_before}")
    print(f"User-Agent：{user_agent}")

    orders = fetch_orders(
        token=token,
        user_agent=user_agent,
        created_after=created_after,
        created_before=created_before,
        per_page=per_page,
        max_pages=args.max_pages,
    )

    print(f"取得訂單數：{len(orders)}")
    summaries = [summarize_order(order) for order in orders]
    status_counts = Counter(str(order.get("status") or "unknown") for order in orders)
    payment_counts = Counter(str(get_nested(order, "order_payment", "status") or "unknown") for order in orders)
    delivery_counts = Counter(str(get_nested(order, "order_delivery", "delivery_status") or "unknown") for order in orders)

    print(f"訂單狀態統計：{dict(status_counts)}")
    print(f"付款狀態統計：{dict(payment_counts)}")
    print(f"配送狀態統計：{dict(delivery_counts)}")

    if args.show > 0:
        print(f"前 {min(args.show, len(summaries))} 筆摘要：")
        print(json.dumps(summaries[: args.show], ensure_ascii=False, indent=2))

    filename_date = safe_date_for_filename(local_date)
    output_dir = args.output_dir
    json_path = args.output_json or output_dir / f"shopline_orders_{filename_date}_raw.json"
    xlsx_path = args.output_xlsx or output_dir / f"{filename_date}捷捷官網轉換前.xlsx"
    jj_rows = to_jjofficial_rows(orders, include_cancelled=args.include_cancelled)

    write_json(json_path, orders)
    write_jjofficial_xlsx(xlsx_path, jj_rows)

    skipped_cancelled = status_counts.get("cancelled", 0) if not args.include_cancelled else 0
    print(f"原始 JSON 已寫入：{json_path}")
    print(f"捷捷官網轉換前 xlsx 已寫入：{xlsx_path}")
    print(f"xlsx 明細列數：{len(jj_rows)}")
    if skipped_cancelled:
        print(f"xlsx 已略過取消訂單：{skipped_cancelled} 筆；如需包含，請加 --include-cancelled")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
