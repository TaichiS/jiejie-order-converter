"""
app/routes.py
Flask 路由：掃描、轉換、報表、下載、設定。
"""
from __future__ import annotations

import csv
import io
import json
import queue
import threading
import zipfile
from pathlib import Path

from flask import (Blueprint, Response, abort, current_app, jsonify, redirect,
                   render_template, request, send_file, url_for)

from app import db
from app.models import ConversionLog, ConversionError
from app.utils import SOURCE_LABELS, SOURCE_COLORS
import app.services as services

bp = Blueprint("main", __name__)

BASE_DIR    = Path(__file__).resolve().parent.parent
CONFIG_FILE = BASE_DIR / "config.json"

# SSE 進度佇列（簡易：單一使用者場景）
_progress_queue: queue.Queue = queue.Queue()


# ── 設定讀寫 ──────────────────────────────────────────────────────────────

def _load_config() -> dict:
    if CONFIG_FILE.exists():
        with open(CONFIG_FILE, encoding="utf-8") as f:
            return json.load(f)
    return {"default_folder": "", "default_operator": ""}


def _save_config(data: dict) -> None:
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


# ── 主頁 ──────────────────────────────────────────────────────────────────

@bp.route("/")
def index():
    config = _load_config()
    return render_template("index.html",
                           default_operator=config.get("default_operator", ""),
                           source_labels=SOURCE_LABELS)


# ── 資料夾瀏覽器 ────────────────────────────────────────────────────────

@bp.route("/browse-path")
def browse_path():
    path_str = request.args.get("path", "").strip()
    if not path_str:
        path_str = _load_config().get("default_folder", "") or str(Path.home())
    current = Path(path_str)
    if not current.exists() or not current.is_dir():
        return jsonify({"error": "路徑不存在"}), 400
    try:
        dirs = sorted(
            [{"name": d.name, "path": str(d)} for d in current.iterdir()
             if d.is_dir() and not d.name.startswith(".")],
            key=lambda x: x["name"].lower()
        )
    except PermissionError:
        dirs = []
    parent = current.parent
    return jsonify({
        "current": str(current),
        "dirs": dirs,
        "parent": str(parent) if parent != current else None,
    })


# ── 掃描資料夾 ────────────────────────────────────────────────────────────

@bp.route("/scan", methods=["POST"])
def scan():
    folder = request.json.get("folder", "").strip()
    if not folder:
        return jsonify({"error": "請輸入資料夾路徑"}), 400
    try:
        results = services.scan_folder(folder)
        return jsonify({"files": results, "folder": folder})
    except ValueError as e:
        return jsonify({"error": str(e)}), 400
    except Exception as e:
        return jsonify({"error": f"掃描失敗：{e}"}), 500


# ── 執行轉換（SSE）────────────────────────────────────────────────────────

@bp.route("/convert", methods=["POST"])
def convert():
    data           = request.json or {}
    folder          = data.get("folder", "").strip()
    operator        = data.get("operator", "").strip()
    forced_source   = data.get("source_type") or None
    selected_files  = data.get("selected_files") or None  # None = 全部

    if not folder:
        return jsonify({"error": "未提供資料夾路徑"}), 400

    app = current_app._get_current_object()

    def _run():
        with app.app_context():
            try:
                _progress_queue.put({"type": "progress", "message": "開始轉換..."})
                logs = services.run_conversion(folder, operator, forced_source,
                                               selected_files)
                all_success = sum(l.success_count for l in logs)
                all_fail    = sum(l.fail_count    for l in logs)
                all_manual  = sum(l.manual_count  for l in logs)
                combined_status = (
                    "success" if all(l.status == "success" for l in logs)
                    else "partial" if any(l.status == "success" for l in logs)
                    else logs[-1].status
                )
                _progress_queue.put({
                    "type":        "done",
                    "log_id":      logs[-1].id,
                    "success":     all_success,
                    "fail":        all_fail,
                    "manual":      all_manual,
                    "status":      combined_status,
                    "source_type": "+".join(l.source_type for l in logs),
                })
            except Exception as e:
                _progress_queue.put({"type": "error", "message": str(e)})

    threading.Thread(target=_run, daemon=True).start()
    return jsonify({"ok": True})


@bp.route("/convert/status")
def convert_status():
    def _stream():
        while True:
            try:
                msg = _progress_queue.get(timeout=30)
                yield f"data: {json.dumps(msg, ensure_ascii=False)}\n\n"
                if msg.get("type") in ("done", "error"):
                    break
            except queue.Empty:
                yield "data: {\"type\":\"heartbeat\"}\n\n"

    return Response(_stream(), mimetype="text/event-stream",
                    headers={"Cache-Control": "no-cache",
                             "X-Accel-Buffering": "no"})


# ── 報表 ──────────────────────────────────────────────────────────────────

@bp.route("/report")
def report():
    return render_template("report.html",
                           source_labels=SOURCE_LABELS,
                           source_colors=SOURCE_COLORS)


@bp.route("/report/data")
def report_data():
    source  = request.args.get("source")
    date_from = request.args.get("from")
    date_to   = request.args.get("to")

    query = ConversionLog.query.order_by(ConversionLog.created_at.desc())
    if source:
        query = query.filter_by(source_type=source)
    if date_from:
        query = query.filter(ConversionLog.created_at >= date_from)
    if date_to:
        query = query.filter(ConversionLog.created_at <= date_to + " 23:59:59")

    logs = query.limit(200).all()
    return jsonify([log.to_dict() for log in logs])


# ── 錯誤明細 ──────────────────────────────────────────────────────────────

@bp.route("/errors/<int:log_id>")
def error_detail(log_id: int):
    log    = ConversionLog.query.get_or_404(log_id)
    errors = ConversionError.query.filter_by(log_id=log_id).all()
    return render_template("error_detail.html", log=log, errors=errors,
                           resolved_ids=set(),
                           source_labels=SOURCE_LABELS)


# ── 下載 ──────────────────────────────────────────────────────────────────

@bp.route("/download")
def download():
    logs = ConversionLog.query.order_by(ConversionLog.created_at.desc()).limit(100).all()
    return render_template("download.html", logs=logs,
                           source_labels=SOURCE_LABELS,
                           source_colors=SOURCE_COLORS)


@bp.route("/download/<int:log_id>")
def download_zip(log_id: int):
    log = ConversionLog.query.get_or_404(log_id)
    files = [Path(p) for p in log.output_files_list if Path(p).exists()]
    if not files:
        abort(404)

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for f in files:
            zf.write(f, f.name)
    buf.seek(0)

    zip_name = f"轉換結果_{log.created_at.strftime('%Y%m%d_%H%M')}.zip"
    return send_file(buf, mimetype="application/zip",
                     as_attachment=True, download_name=zip_name)


@bp.route("/download/file")
def download_file():
    path = request.args.get("path", "")
    file_path = Path(path)
    if not file_path.exists() or not file_path.is_file():
        abort(404)
    # 安全性：確保路徑在 output/ 目錄下
    try:
        file_path.resolve().relative_to((BASE_DIR / "output").resolve())
    except ValueError:
        abort(403)
    return send_file(file_path, as_attachment=True, download_name=file_path.name)


# ── 別名新增 ──────────────────────────────────────────────────────────────

@bp.route("/alias/search")
def alias_search():
    """
    搜尋品號資料，供別名選擇用。
    Query: q=關鍵字, source_type=來源類型
    回傳: [{品號, 品名, 商品結帳價}]
    """
    import csv as _csv
    q           = request.args.get("q", "").strip().lower()
    source_type = request.args.get("source_type", "")
    if not q:
        return jsonify([])

    _ref_map = {
        "a1leage": services.BASE_DIR / "A1樂齡官網" / "品號資料.csv",
        "jjofficial": services.BASE_DIR / "捷捷寶寶粥官網" / "品號資料.csv",
        "yodee": services.BASE_DIR / "優迪通路" / "品號資料.csv",
    }
    ref_path = _ref_map.get(source_type, services.BASE_DIR / "reference" / "品號資料.csv")
    results  = []
    try:
        with open(ref_path, encoding="utf-8-sig", newline="") as f:
            for row in _csv.DictReader(f):
                name  = row.get("品名", "").strip()
                no    = row.get("品號", "").strip()
                price = row.get("商品結帳價", "").strip()
                if not name or not no or not price:
                    continue
                if q in name.lower() or q in no.lower():
                    results.append({"品號": no, "品名": name, "商品結帳價": price})
    except Exception:
        pass
    return jsonify(results[:8])


@bp.route("/alias", methods=["POST"])
def create_alias():
    """
    新增品號別名。
    Request JSON: {alias_name: str, target_no: str, source_type: str}

    a1leage：alias_name 為缺少的 SKU，複製 target_no 的完整品項資料並以 alias_name 為新品號。
    其他來源：alias_name 為缺少的品名，新增一行品號=target_no、品名=alias_name（空結帳價繼承）。
    """
    import csv as _csv
    data        = request.json or {}
    alias_name  = data.get("alias_name", "").strip()
    target_no   = data.get("target_no", "").strip()
    source_type = data.get("source_type", "").strip()

    if not alias_name or not target_no:
        return jsonify({"error": "alias_name 和 target_no 皆必填"}), 400

    _ref_map2 = {
        "a1leage": services.BASE_DIR / "A1樂齡官網" / "品號資料.csv",
        "jjofficial": services.BASE_DIR / "捷捷寶寶粥官網" / "品號資料.csv",
        "yodee": services.BASE_DIR / "優迪通路" / "品號資料.csv",
    }
    ref_path = _ref_map2.get(source_type, services.BASE_DIR / "reference" / "品號資料.csv")
    if not ref_path.exists():
        return jsonify({"error": "找不到品號資料.csv"}), 500

    with open(ref_path, encoding="utf-8-sig", newline="") as f:
        reader  = _csv.DictReader(f)
        headers = reader.fieldnames or []
        rows    = list(reader)

    if source_type == "a1leage":
        # a1leage 用 SKU 查找：複製目標品項的完整資料，品號改為缺少的 SKU
        if any(r.get("品號", "").strip() == alias_name for r in rows):
            return jsonify({"ok": True, "message": f"品號「{alias_name}」已存在"}), 200
        target_row = next((r for r in rows if r.get("品號", "").strip() == target_no), None)
        if not target_row:
            return jsonify({"error": f"找不到品號 {target_no}"}), 400
        new_row = dict(target_row)
        new_row["品號"] = alias_name
        if "來源" in headers:
            new_row["來源"] = "別名"
    else:
        # 其他來源：品名查找，新增繼承行
        if any(r.get("品名", "").strip() == alias_name for r in rows):
            return jsonify({"ok": True, "message": f"「{alias_name}」已存在"}), 200
        new_row = {h: "" for h in headers}
        new_row["品號"] = target_no
        new_row["品名"] = alias_name
        new_row["來源"] = "別名"

    with open(ref_path, "a", encoding="utf-8", newline="") as f:
        writer = _csv.DictWriter(f, fieldnames=headers)
        writer.writerow(new_row)

    return jsonify({"ok": True, "message": f"已新增別名「{alias_name}」→ {target_no}"}), 200


# ── 轉換 rollback ─────────────────────────────────────────────────────────

@bp.route("/convert/rollback/<int:log_id>", methods=["POST"])
def rollback_conversion(log_id: int):
    """刪除輸出檔、歸檔原始檔、DB 記錄，讓使用者可重新轉換。"""
    log = ConversionLog.query.get_or_404(log_id)

    # 刪除輸出檔
    for p in log.output_files_list:
        Path(p).unlink(missing_ok=True)

    # 刪除對應的歸檔原始檔（依 input_files 掃描 archive/）
    archive_root = BASE_DIR / "archive" / log.source_type
    for fname in log.input_files_list:
        for f in archive_root.rglob(f"*_{fname}"):
            f.unlink(missing_ok=True)

    # 刪除 DB 記錄（ConversionError 因 cascade 一起刪）
    db.session.delete(log)
    db.session.commit()
    return jsonify({"ok": True})


@bp.route("/errors/<int:log_id>/data")
def error_detail_data(log_id: int):
    """回傳某次轉換的錯誤列表（JSON）。"""
    errors = ConversionError.query.filter_by(log_id=log_id).all()
    return jsonify([e.to_dict() for e in errors])


# ── 說明 ──────────────────────────────────────────────────────────────────

@bp.route("/help")
def help_page():
    return render_template("help.html")


# ── 品號資料匯入 ──────────────────────────────────────────────────────────

@bp.route("/admin/import-products", methods=["GET", "POST"])
def import_products():
    """上傳品號資料統整.csv，全表覆蓋 unified_products。"""
    from app.models import UnifiedProduct

    result = None
    if request.method == "POST":
        file = request.files.get("csv_file")
        if not file or file.filename == "":
            result = {"success": 0, "errors": 1, "error_messages": ["未選擇檔案"]}
        else:
            try:
                # 清空現有資料（先不 commit，與後續匯入在同一筆交易中）
                UnifiedProduct.query.delete()

                stream = io.StringIO(file.stream.read().decode("utf-8-sig"))
                reader = csv.DictReader(stream)
                success = 0
                errors = 0
                error_messages: list[str] = []

                for row in reader:
                    try:
                        barcode = (row.get("條碼") or "").strip()
                        sku = (row.get("品號") or "").strip()
                        name = (row.get("品名") or "").strip()
                        channel = (row.get("通路") or "").strip()

                        if not sku or not name or not channel:
                            errors += 1
                            continue

                        def _float(val):
                            try:
                                return float(val.strip()) if val else None
                            except (ValueError, TypeError):
                                return None

                        def _int(val):
                            try:
                                return int(float(val.strip())) if val else None
                            except (ValueError, TypeError):
                                return None

                        up = UnifiedProduct(
                            barcode=barcode or None,
                            sku=sku,
                            name=name,
                            category=(row.get("類別") or "").strip() or None,
                            channel=channel,
                            quantity=_int(row.get("份數")) or 1,
                            pack_size=_int(row.get("包數")),
                            unit_price=_float(row.get("份數價格")),
                            pack_price=_float(row.get("包數價格")),
                            checkout_price=_float(row.get("商品結帳價")),
                        )
                        db.session.add(up)
                        success += 1
                    except Exception as e:
                        errors += 1
                        error_messages.append(f"第 {success + errors} 行錯誤：{e}")

                if errors:
                    db.session.rollback()
                    error_messages.insert(0, "資料未更新，請修正 CSV 後重新匯入")
                else:
                    # 補齊條碼：舊 product_barcodes 表仍有部分 CSV 缺少的條碼
                    try:
                        from app.models import ProductBarcode
                        barcode_map: dict[str, str] = {}
                        for pb in ProductBarcode.query.all():
                            if pb.sku not in barcode_map:
                                barcode_map[pb.sku] = pb.barcode
                        for sku, barcode in barcode_map.items():
                            UnifiedProduct.query.filter(
                                UnifiedProduct.sku == sku,
                                db.or_(
                                    UnifiedProduct.barcode.is_(None),
                                    UnifiedProduct.barcode == "",
                                ),
                            ).update({"barcode": barcode}, synchronize_session=False)
                    except Exception:
                        pass

                    db.session.commit()

                result = {
                    "success": success,
                    "errors": errors,
                    "error_messages": error_messages[:10],
                }
            except Exception as e:
                db.session.rollback()
                result = {"success": 0, "errors": 1, "error_messages": [f"匯入失敗：{e}"]}

    return render_template("import_products.html", result=result)


@bp.route("/admin/clear-data", methods=["POST"])
def clear_data():
    """清除所有轉換資料與輸出檔案（保留統一品號表）。"""
    try:
        ConversionError.query.delete()
        ConversionLog.query.delete()
        db.session.commit()

        # 刪除 output/ 與 archive/ 下的所有檔案
        for folder in (BASE_DIR / "output", BASE_DIR / "archive"):
            if folder.exists():
                for f in folder.rglob("*"):
                    if f.is_file():
                        f.unlink()

        return jsonify({"ok": True, "message": "已清除轉換資料與輸出檔案"})
    except Exception as e:
        db.session.rollback()
        return jsonify({"ok": False, "message": str(e)}), 500


@bp.route("/admin/download-product-template")
def download_product_template():
    """下載品號資料統整.csv 原檔。"""
    csv_path = BASE_DIR / "品號資料統整.csv"
    if not csv_path.exists():
        abort(404)
    return send_file(
        csv_path,
        mimetype="text/csv",
        as_attachment=True,
        download_name="品號資料統整.csv",
    )


@bp.route("/admin/product-search")
def product_search():
    """
    品號查詢 API。
    Query: q=關鍵字
    回傳: [{品號, 品名, 類別, 通路, 份數, 包數, 份數價格, 包數價格, 商品結帳價}]
    """
    from app.models import UnifiedProduct

    q = request.args.get("q", "").strip()
    if not q:
        return jsonify([])

    query = UnifiedProduct.query
    # 優先精確比對品號或條碼，再模糊比對品名
    try:
        results = (
            query.filter(
                db.or_(
                    UnifiedProduct.sku == q,
                    UnifiedProduct.barcode == q,
                    UnifiedProduct.name.contains(q),
                )
            )
            .order_by(
                db.case(
                    (UnifiedProduct.sku == q, 0),
                    (UnifiedProduct.barcode == q, 1),
                    else_=2,
                )
            )
            .limit(50)
            .all()
        )
    except Exception:
        results = []

    return jsonify(
        [
            {
                "條碼": r.barcode or "",
                "品號": r.sku,
                "品名": r.name,
                "類別": r.category or "",
                "通路": r.channel,
                "份數": r.quantity,
                "包數": r.pack_size or "",
                "份數價格": r.unit_price or "",
                "包數價格": r.pack_price or "",
                "商品結帳價": r.checkout_price or "",
            }
            for r in results
        ]
    )
