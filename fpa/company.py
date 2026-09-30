"""
Companies live in companies/<slug>/:

    company.toml      name, model, financial year, drivers or budget rules, cash rules
    accounts.csv      chart of accounts (fpa/accounts.py)
    commentary.toml   analyst notes by month and account (optional)
    forecast.toml     latest estimate and scenarios (optional)
    actuals/<month>/  monthly trial balance and KPI exports

model = "lettings" uses the driver-based template in fpa/drivers.py;
model = "rules" budgets each account by a simple rule (fpa/rules.py).
"""
import tomllib
from dataclasses import dataclass, field
from pathlib import Path

from .accounts import Chart

ROOT = Path(__file__).resolve().parent.parent
COMPANIES = ROOT / "companies"
DEFAULT = "kestrel-row"


@dataclass
class Company:
    cfg: dict
    slug: str = DEFAULT
    dir: Path = None
    chart: Chart = field(default=None, repr=False)

    @property
    def name(self):
        return self.cfg["company"]["name"]

    @property
    def model(self):
        return self.cfg["company"].get("model", "rules")

    @property
    def fy_start_month(self):
        return self.cfg["company"]["fy_start_month"]

    @property
    def budget_year(self):
        return self.cfg["company"]["budget_year"]

    @property
    def actuals_dir(self):
        return self.dir / "actuals"

    @property
    def commentary_path(self):
        return self.dir / "commentary.toml"

    @property
    def forecast_path(self):
        return self.dir / "forecast.toml"

    def fy_months(self, fy_start_year=None):
        """The 12 months of a financial year, e.g. 2026 -> 2026-04 .. 2027-03."""
        y = self.budget_year if fy_start_year is None else fy_start_year
        return [add_months("%d-%02d" % (y, self.fy_start_month), i) for i in range(12)]

    def fy_label(self, fy_start_year=None):
        y = self.budget_year if fy_start_year is None else fy_start_year
        if self.fy_start_month == 1:
            return "FY%d" % y
        return "FY%d/%02d" % (y, (y + 1) % 100)

    def month_index(self, month):
        """1..12 within the financial year."""
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


def available(root=None):
    d = Path(root or COMPANIES)
    return sorted(p.name for p in d.iterdir() if (p / "company.toml").exists()) if d.exists() else []


def load(slug=None, root=None):
    d = Path(root or COMPANIES) / (slug or DEFAULT)
    with open(d / "company.toml", "rb") as f:
        cfg = tomllib.load(f)
    return Company(cfg, d.name, d, Chart.load(d / "accounts.csv"))
