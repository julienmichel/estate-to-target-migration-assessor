"""
Live Azure retail pricing for the Estate-to-Target Migration Assessor.

Queries the public Azure Retail Prices API - no auth, no subscription needed - so
the BoQ carries real regional list prices rather than hardcoded guesses that rot.

Design notes:
  * Results are cached on disk (24h) because an estate of 200 databases would
    otherwise hammer the API with near-identical queries.
  * Any network failure degrades to an unpriced BoQ rather than aborting the run.
    A sizing without prices is still useful; a crash is not.
  * Prices are list / pay-as-you-go. Customer-specific discounting, MACC and
    Azure Hybrid Benefit are applied downstream as clearly-labelled estimates.
"""

from __future__ import annotations

import json
import time
import urllib.parse
import urllib.request
from pathlib import Path

CACHE_DIR = Path(__file__).resolve().parent.parent / ".cache"
CACHE_TTL_SECONDS = 24 * 60 * 60
API = "https://prices.azure.com/api/retail/prices"


def _cache_path(key: str) -> Path:
    safe = "".join(c if c.isalnum() else "_" for c in key)[:120]
    return CACHE_DIR / f"{safe}.json"


def _get_cached(key: str) -> list[dict] | None:
    path = _cache_path(key)
    if not path.exists():
        return None
    if time.time() - path.stat().st_mtime > CACHE_TTL_SECONDS:
        return None
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception:  # noqa: BLE001 - a corrupt cache entry is not fatal
        return None


def _put_cached(key: str, items: list[dict]) -> None:
    CACHE_DIR.mkdir(parents=True, exist_ok=True)
    try:
        _cache_path(key).write_text(json.dumps(items), encoding="utf-8")
    except Exception:  # noqa: BLE001 - caching is best-effort
        pass


def query(odata_filter: str, currency: str = "USD", timeout: int = 30) -> list[dict]:
    """Run one $filter against the retail price API, following pagination."""
    key = f"{currency}|{odata_filter}"
    cached = _get_cached(key)
    if cached is not None:
        return cached

    items: list[dict] = []
    url = f"{API}?currencyCode={currency}&$filter={urllib.parse.quote(odata_filter)}"
    pages = 0

    while url and pages < 10:
        req = urllib.request.Request(url, headers={"Accept": "application/json"})
        with urllib.request.urlopen(req, timeout=timeout) as resp:  # noqa: S310 - fixed public endpoint
            payload = json.loads(resp.read().decode("utf-8"))
        items.extend(payload.get("Items", []))
        url = payload.get("NextPageLink")
        pages += 1

    _put_cached(key, items)
    return items


class PriceBook:
    """Caches one region's SQL price list and answers sizing questions against it."""

    def __init__(self, region: str = "uaenorth", currency: str = "USD",
                 hours_per_month: int = 730):
        self.region = region
        self.currency = currency
        self.hours = hours_per_month
        self.available = False
        self.error = ""
        self._mi: list[dict] = []
        self._db: list[dict] = []
        self._storage: list[dict] = []

    def load(self) -> None:
        try:
            self._mi = query(
                f"serviceName eq 'SQL Managed Instance' and armRegionName eq '{self.region}' "
                f"and priceType eq 'Consumption'", self.currency)
            self._db = query(
                f"serviceName eq 'SQL Database' and armRegionName eq '{self.region}' "
                f"and priceType eq 'Consumption'", self.currency)
            self.available = bool(self._mi or self._db)
            if not self.available:
                self.error = f"No price data returned for region '{self.region}'. Check the region name."
        except Exception as exc:  # noqa: BLE001 - degrade, never abort
            self.available = False
            self.error = f"Azure Retail Prices API unreachable ({exc}). BoQ will be produced without prices."

    # -- lookup helpers ---------------------------------------------------

    @staticmethod
    def _vcores_in(item: dict) -> int | None:
        for field in ("armSkuName", "skuName", "meterName"):
            text = str(item.get(field, ""))
            for token in text.replace("_", " ").split():
                if token.isdigit():
                    n = int(token)
                    if 1 <= n <= 256:
                        return n
        return None

    def _best(self, pool: list[dict], vcores: int, tier: str,
              exclude: tuple[str, ...] = ()) -> dict | None:
        """Closest SKU at or above the required vCores, in the requested tier."""
        tier_token = "Business Critical" if tier == "BusinessCritical" else "General Purpose"
        candidates = []
        for item in pool:
            product = str(item.get("productName", ""))
            if tier_token.lower() not in product.lower():
                continue
            if any(x.lower() in product.lower() for x in exclude):
                continue
            if float(item.get("retailPrice", 0) or 0) <= 0:
                continue
            if "hour" not in str(item.get("unitOfMeasure", "")).lower():
                continue
            n = self._vcores_in(item)
            if n is None or n < vcores:
                continue
            candidates.append((n, float(item["retailPrice"]), item))

        if not candidates:
            return None
        candidates.sort(key=lambda t: (t[0], t[1]))
        return candidates[0][2]

    def price(self, target: str, vcores: int, tier: str, storage_gb: int) -> dict:
        """Estimated monthly list cost for one sized workload."""
        result = {
            "monthly": None,
            "monthly_ahb": None,
            "source": "unavailable",
            "detail": self.error or "Pricing not loaded",
        }
        if not self.available:
            return result

        if target == "SQL_MI":
            pool, label = self._mi, "SQL Managed Instance"
        elif target == "SQL_DB":
            pool, label = self._db, "SQL Database"
        else:
            result["detail"] = ("SQL Server on Azure VM pricing depends on the chosen VM family, OS and "
                                "licence model. Price this in the Azure pricing calculator.")
            return result

        # Exclude Hyperscale and serverless so we compare like for like.
        item = self._best(pool, vcores, tier, exclude=("Hyperscale", "Serverless"))
        if item is None:
            result["detail"] = f"No {label} {tier} SKU at or above {vcores} vCores in {self.region}."
            return result

        compute = float(item["retailPrice"]) * self.hours
        # Storage is billed per GB/month; a representative rate avoids a second
        # round-trip per workload and is clearly labelled as indicative.
        storage_rate = 0.115 if tier == "GeneralPurpose" else 0.25
        storage_cost = storage_gb * storage_rate
        total = compute + storage_cost

        result.update({
            "monthly": round(total, 2),
            # Azure Hybrid Benefit removes the SQL licence component from compute.
            "monthly_ahb": round(compute * 0.45 + storage_cost, 2),
            "source": "azure_retail_api",
            "detail": (
                f"{item.get('productName')} / {item.get('skuName')} at "
                f"{item['retailPrice']:.4f} {self.currency}/hr x {self.hours} hrs "
                f"= {compute:,.2f}, plus {storage_gb:,} GB storage at ~{storage_rate:.3f}/GB "
                f"= {storage_cost:,.2f}. Region {self.region}, list price."
            ),
        })
        return result


def apply_pricing(assessments, region: str, currency: str, hours: int) -> PriceBook:
    """Price every assessment in place and hand back the book for reporting."""
    book = PriceBook(region=region, currency=currency, hours_per_month=hours)
    book.load()
    for a in assessments:
        p = book.price(a.target, a.vcores, a.service_tier, a.storage_gb)
        a.est_monthly_usd = p["monthly"]
        a.est_monthly_ahb_usd = p["monthly_ahb"]
        a.price_source = p["detail"]
    return book
