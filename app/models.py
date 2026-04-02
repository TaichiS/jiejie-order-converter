import json
from datetime import datetime
from app import db


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
            "source_label":  {"shopee": "蝦皮", "a1baby": "A1婦幼展", "leage": "樂齡網"}.get(self.source_type, self.source_type),
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
        }
