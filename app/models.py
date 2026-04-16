import json
from datetime import datetime
from app import db
from app.utils import SOURCE_LABELS


class ConversionLog(db.Model):
    __tablename__ = "conversion_logs"

    id            = db.Column(db.Integer, primary_key=True)
    created_at    = db.Column(db.DateTime, default=datetime.utcnow, nullable=False)
    operator      = db.Column(db.String(50))
    source_type   = db.Column(db.String(20), nullable=False)  # shopee|a1baby|leage
    input_files   = db.Column(db.Text)   # JSON 陣列（原始檔名）
    success_count = db.Column(db.Integer, default=0)
    fail_count    = db.Column(db.Integer, default=0)
    manual_count  = db.Column(db.Integer, default=0)
    output_files  = db.Column(db.Text)   # JSON 陣列（完整輸出路徑）
    status        = db.Column(db.String(20), nullable=False)  # completed|partial|failed

    errors = db.relationship("ConversionError", back_populates="log",
                             cascade="all, delete-orphan", lazy="dynamic")

    @property
    def input_files_list(self) -> list:
        return json.loads(self.input_files) if self.input_files else []

    @property
    def output_files_list(self) -> list:
        return json.loads(self.output_files) if self.output_files else []

    def to_dict(self) -> dict:
        return {
            "id":            self.id,
            "created_at":    self.created_at.strftime("%Y-%m-%d %H:%M"),
            "operator":      self.operator or "",
            "source_type":   self.source_type,
            "source_label":  SOURCE_LABELS.get(self.source_type, self.source_type),
            "input_files":   self.input_files_list,
            "success_count": self.success_count,
            "fail_count":    self.fail_count,
            "manual_count":  self.manual_count,
            "output_files":  self.output_files_list,
            "status":        self.status,
        }


class ConversionError(db.Model):
    __tablename__ = "conversion_errors"

    id              = db.Column(db.Integer, primary_key=True)
    log_id          = db.Column(db.Integer, db.ForeignKey("conversion_logs.id"), nullable=False)
    row_number      = db.Column(db.Integer)
    field_name      = db.Column(db.String(100))
    original_value  = db.Column(db.Text)
    reason          = db.Column(db.Text, nullable=False)
    candidates_json = db.Column(db.Text)   # JSON 陣列：[{品號,品名,商品結帳價,match_type}]
    source_file     = db.Column(db.String(255), default="")

    log = db.relationship("ConversionLog", back_populates="errors")

    @property
    def candidates(self) -> list:
        return json.loads(self.candidates_json) if self.candidates_json else []

    def to_dict(self) -> dict:
        return {
            "id":             self.id,
            "row_number":     self.row_number,
            "field_name":     self.field_name or "",
            "original_value": self.original_value or "",
            "reason":         self.reason,
            "candidates":     self.candidates,
            "source_file":    self.source_file or "",
            "source_type":    self.log.source_type if self.log else "",
            "source_label":   SOURCE_LABELS.get(self.log.source_type, self.log.source_type) if self.log else "",
        }


class Product(db.Model):
    __tablename__ = "products"

    sku        = db.Column(db.String(20),  primary_key=True)   # 品號
    name       = db.Column(db.String(100), nullable=False)      # 主品名
    category   = db.Column(db.String(50))                       # 類別
    quantity   = db.Column(db.Integer,  default=1)              # 份數
    pack_size  = db.Column(db.Integer)                          # 包數（nullable）
    unit_price = db.Column(db.Float)                            # 份數價格（nullable）
    erp_source = db.Column(db.String(50))                       # 來源（鼎新等）

    aliases    = db.relationship("ProductAlias",   back_populates="product",
                                 cascade="all, delete-orphan")
    barcodes   = db.relationship("ProductBarcode", back_populates="product",
                                 cascade="all, delete-orphan")
    prices     = db.relationship("ChannelPrice",   back_populates="product",
                                 cascade="all, delete-orphan")


class ProductAlias(db.Model):
    __tablename__ = "product_aliases"

    id      = db.Column(db.Integer, primary_key=True)
    sku     = db.Column(db.String(20), db.ForeignKey("products.sku"), nullable=False)
    alias   = db.Column(db.String(100), nullable=False)

    product = db.relationship("Product", back_populates="aliases")

    __table_args__ = (
        db.UniqueConstraint("sku", "alias", name="uq_sku_alias"),
    )


class ProductBarcode(db.Model):
    __tablename__ = "product_barcodes"

    id      = db.Column(db.Integer, primary_key=True)
    sku     = db.Column(db.String(20), db.ForeignKey("products.sku"), nullable=False)
    barcode = db.Column(db.String(50), nullable=False, unique=True)

    product = db.relationship("Product", back_populates="barcodes")


class ChannelPrice(db.Model):
    __tablename__ = "channel_prices"

    id      = db.Column(db.Integer, primary_key=True)
    sku     = db.Column(db.String(20), db.ForeignKey("products.sku"), nullable=False)
    channel = db.Column(db.String(20), nullable=False)
    price   = db.Column(db.Float,      nullable=False)

    product = db.relationship("Product", back_populates="prices")

    __table_args__ = (
        db.UniqueConstraint("sku", "channel", name="uq_sku_channel"),
    )


class UnifiedProduct(db.Model):
    """統一品號資料表：對應品號資料統整.csv"""
    __tablename__ = "unified_products"

    id             = db.Column(db.Integer, primary_key=True)
    barcode        = db.Column(db.String(50),  nullable=True,  index=True)
    sku            = db.Column(db.String(20),  nullable=False, index=True)
    name           = db.Column(db.String(100), nullable=False)
    category       = db.Column(db.String(50))
    channel        = db.Column(db.String(50),  nullable=False, index=True)
    quantity       = db.Column(db.Integer,     default=1)
    pack_size      = db.Column(db.Integer)
    unit_price     = db.Column(db.Float)       # 份數價格
    pack_price     = db.Column(db.Float)       # 包數價格
    checkout_price = db.Column(db.Float)       # 商品結帳價
    discount       = db.Column(db.Float)       # 加購折扣

    __table_args__ = (
        db.Index("ix_unified_products_sku_channel", "sku", "channel"),
    )
