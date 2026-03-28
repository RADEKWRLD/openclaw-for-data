"""Data models for engineering price data."""

from __future__ import annotations

import json
import re
from dataclasses import asdict, dataclass, field
from datetime import datetime
from pathlib import Path


@dataclass
class MonthlyPrice:
    month: str  # "2025-01"
    price_min: float
    price_max: float
    price_avg: float

    @classmethod
    def from_range(cls, month: str, price_min: float, price_max: float) -> MonthlyPrice:
        return cls(
            month=month,
            price_min=round(price_min, 2),
            price_max=round(price_max, 2),
            price_avg=round((price_min + price_max) / 2, 2),
        )

    @classmethod
    def from_single(cls, month: str, price: float) -> MonthlyPrice:
        return cls(month=month, price_min=price, price_max=price, price_avg=price)


@dataclass
class PriceItem:
    name: str  # "钢筋工"
    code: str  # "00030119"
    category: str  # "labor" or "material"
    unit: str  # "工日" or "t" or "m3"
    sheet: str = ""  # 来源 Sheet 名称，如 "建筑，安装"
    monthly: list[MonthlyPrice] = field(default_factory=list)

    def add_month(self, mp: MonthlyPrice) -> None:
        for existing in self.monthly:
            if existing.month == mp.month:
                existing.price_min = mp.price_min
                existing.price_max = mp.price_max
                existing.price_avg = mp.price_avg
                return
        self.monthly.append(mp)
        self.monthly.sort(key=lambda m: m.month)


@dataclass
class Meta:
    city: str
    period: str  # "2025-01~2025-12" or "2026-03"
    source: str = ""
    collected_at: str = field(default_factory=lambda: datetime.now().isoformat(timespec="seconds"))


@dataclass
class PriceData:
    meta: Meta
    prices: list[PriceItem] = field(default_factory=list)

    def find_item(self, name: str) -> PriceItem | None:
        for item in self.prices:
            if item.name == name:
                return item
        return None

    def to_dict(self) -> dict:
        return asdict(self)

    def to_json(self, indent: int = 2) -> str:
        return json.dumps(self.to_dict(), ensure_ascii=False, indent=indent)

    def save(self, output_dir: str | Path) -> Path:
        output_dir = Path(output_dir)
        output_dir.mkdir(parents=True, exist_ok=True)
        path = output_dir / "prices.json"
        path.write_text(self.to_json(), encoding="utf-8")
        return path

    @classmethod
    def load(cls, path: str | Path) -> PriceData:
        path = Path(path)
        data = json.loads(path.read_text(encoding="utf-8"))
        meta = Meta(**data["meta"])
        prices = []
        for item_data in data["prices"]:
            monthly = [MonthlyPrice(**m) for m in item_data.pop("monthly", [])]
            prices.append(PriceItem(**item_data, monthly=monthly))
        return cls(meta=meta, prices=prices)

    def merge(self, other: PriceData) -> None:
        """Merge another PriceData into this one (add months from other)."""
        for other_item in other.prices:
            existing = self.find_item(other_item.name)
            if existing is None:
                self.prices.append(other_item)
            else:
                for mp in other_item.monthly:
                    existing.add_month(mp)


def parse_price_range(val: str) -> tuple[float, float]:
    """Parse a price range string like '219.00-266.00' into (min, max).

    Also handles single values like '350.00' and dash-separated with spaces.
    """
    if not val or str(val).strip() in ("", "-", "NaN", "nan", "None"):
        raise ValueError(f"Empty or invalid price value: {val!r}")

    val = str(val).strip()
    # Match pattern like "219.00-266.00"
    match = re.match(r"^(\d+(?:\.\d+)?)\s*[-~]\s*(\d+(?:\.\d+)?)$", val)
    if match:
        return float(match.group(1)), float(match.group(2))

    # Single value
    try:
        v = float(val)
        return v, v
    except ValueError:
        raise ValueError(f"Cannot parse price value: {val!r}")


def get_data_dir(base_dir: str | Path, city: str) -> Path:
    return Path(base_dir) / "prices" / city
