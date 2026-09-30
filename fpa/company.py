"""
Company definition: loads config/company.toml and provides the financial-year
calendar (months as "YYYY-MM", April to March).
"""
import tomllib
from dataclasses import dataclass
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
CONFIG = ROOT / "config" / "company.toml"


@dataclass(frozen=True)
class Company:
    cfg: dict

    @property
    def name(self):
        return self.cfg["company"]["name"]

    @property
    def fy_start_month(self):
        return self.cfg["company"]["fy_start_month"]

    @property
    def budget_year(self):
        return self.cfg["company"]["budget_year"]

    def fy_months(self, fy_start_year=None):
        """The 12 months of a financial year, e.g. 2026 -> 2026-04 .. 2027-03."""
        y = self.budget_year if fy_start_year is None else fy_start_year
        return [add_months("%d-%02d" % (y, self.fy_start_month), i) for i in range(12)]

    def fy_label(self, fy_start_year=None):
        y = self.budget_year if fy_start_year is None else fy_start_year
        return "FY%d/%02d" % (y, (y + 1) % 100)

    def month_index(self, month):
        """1..12 within the financial year (April = 1)."""
        m = int(month[5:7])
        return (m - self.fy_start_month) % 12 + 1

    def season(self, series, month):
        return self.cfg["seasonality"][series][self.month_index(month) - 1]


def add_months(month, n):
    y, m = int(month[:4]), int(month[5:7])
    t = y * 12 + (m - 1) + n
    return "%d-%02d" % (t // 12, t % 12 + 1)


def month_range(first, last):
    out, m = [], first
    while m <= last:
        out.append(m)
        m = add_months(m, 1)
    return out


def load(path=None):
    with open(path or CONFIG, "rb") as f:
        return Company(tomllib.load(f))
