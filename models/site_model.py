"""
Per-site model — reactive load curve, injection policy, savings estimator.
Trained from SAR data; retraining takes <5 seconds.
Saved/loaded as .pkl files in models/.
"""

import math
import pickle
import logging
from dataclasses import dataclass, field
from pathlib import Path
from datetime import datetime

from core.schema import PF_TARGET, SITE_CONFIG

log = logging.getLogger("hyesys.models.site_model")

MODELS_DIR = Path(__file__).parent


@dataclass
class SiteModel:
    site_id:     str
    solar:       bool  = True
    trained_at:  str   = ""
    n_records:   int   = 0

    # Hourly reactive load profile — avg kVAr per hour of day (index 0–23)
    hourly_avg_kvar:  list[float] = field(default_factory=lambda: [0.0] * 24)
    hourly_avg_kw:    list[float] = field(default_factory=lambda: [0.0] * 24)
    hourly_counts:    list[int]   = field(default_factory=lambda: [0]   * 24)

    # Overall statistics
    avg_kvar:    float = 0.0
    avg_kw:      float = 0.0
    max_kw:      float = 0.0
    avg_pf:      float = 0.0
    pf_std:      float = 0.0

    # Injection policy parameters (learned from SAR outcomes)
    inject_scale: float = 1.0   # scaling factor for injection magnitude
    hold_pf_band: float = 0.01  # don't inject if PF within this band of target

    def fit(self, meter_records: list[dict], sar_records: list[dict] | None = None):
        """
        Fit the site model from meter records and (optionally) SAR outcomes.
        meter_records: list of dicts with timestamp, kW, kVAr, PF fields.
        """
        if not meter_records:
            log.warning("[%s] No records to fit model.", self.site_id)
            return

        # ── Hourly load profile ────────────────────────────────
        hourly_kvar  = [[] for _ in range(24)]
        hourly_kw    = [[] for _ in range(24)]
        pf_vals      = []

        for rec in meter_records:
            try:
                ts   = rec.get("timestamp", "")
                hour = datetime.fromisoformat(ts).hour
                kvar = float(rec.get("kVAr", 0) or 0)
                kw   = float(rec.get("kW",   0) or 0)
                pf   = abs(float(rec.get("PF", 0) or 0))
                hourly_kvar[hour].append(kvar)
                hourly_kw[hour].append(kw)
                if pf > 0:
                    pf_vals.append(pf)
            except (ValueError, TypeError):
                continue

        self.hourly_avg_kvar = [
            round(sum(v) / len(v), 2) if v else 0.0 for v in hourly_kvar
        ]
        self.hourly_avg_kw = [
            round(sum(v) / len(v), 2) if v else 0.0 for v in hourly_kw
        ]
        self.hourly_counts = [len(v) for v in hourly_kvar]

        # ── Overall stats ──────────────────────────────────────
        all_kvar = [float(r.get("kVAr", 0) or 0) for r in meter_records]
        all_kw   = [float(r.get("kW",   0) or 0) for r in meter_records]

        self.avg_kvar  = round(sum(all_kvar) / len(all_kvar), 2) if all_kvar else 0.0
        self.avg_kw    = round(sum(all_kw)   / len(all_kw),   2) if all_kw   else 0.0
        self.max_kw    = round(max(all_kw, default=0.0),            2)
        self.avg_pf    = round(sum(pf_vals) / len(pf_vals),   4) if pf_vals  else 0.0
        self.n_records = len(meter_records)

        # PF standard deviation
        if pf_vals:
            mean = self.avg_pf
            variance = sum((x - mean) ** 2 for x in pf_vals) / len(pf_vals)
            self.pf_std = round(math.sqrt(variance), 4)

        # ── Learn injection policy from SAR outcomes ───────────
        if sar_records:
            self._fit_policy(sar_records)

        self.trained_at = datetime.utcnow().isoformat()
        log.info("[%s] Model fitted on %d records. avg_kVAr=%.1f avg_PF=%.3f",
                 self.site_id, self.n_records, self.avg_kvar, self.avg_pf)

    def _fit_policy(self, sar_records: list[dict]):
        """Adjust injection scale based on historical SAR outcomes."""
        positive = sum(1 for r in sar_records if r.get("outcome") == "POSITIVE")
        negative = sum(1 for r in sar_records if r.get("outcome") == "NEGATIVE")
        total    = len(sar_records)
        if total == 0:
            return

        positive_rate = positive / total
        if positive_rate > 0.7:
            self.inject_scale = min(self.inject_scale * 1.05, 1.2)
        elif positive_rate < 0.4:
            self.inject_scale = max(self.inject_scale * 0.95, 0.8)

    def predict_kvar_needed(self, hour: int, current_kw: float) -> float:
        """
        Predicts kVAr correction needed for this hour and load level.
        Uses the hourly load profile scaled to current kW.
        """
        profile_kw   = self.hourly_avg_kw[hour]
        profile_kvar = self.hourly_avg_kvar[hour]

        if profile_kw <= 0 or current_kw <= 0:
            return profile_kvar * self.inject_scale

        scale = current_kw / profile_kw
        return round(profile_kvar * scale * self.inject_scale, 2)

    def save(self, models_dir: Path = MODELS_DIR):
        path = models_dir / f"site_{self.site_id}.pkl"
        with open(path, "wb") as f:
            pickle.dump(self, f)
        log.info("Model saved: %s", path)
        return path

    @classmethod
    def load(cls, site_id: str, models_dir: Path = MODELS_DIR) -> "SiteModel | None":
        path = models_dir / f"site_{site_id}.pkl"
        if not path.exists():
            return None
        with open(path, "rb") as f:
            model = pickle.load(f)
        log.info("Model loaded: %s (trained %s, n=%d)", path, model.trained_at, model.n_records)
        return model

    def summary(self) -> str:
        lines = [
            f"Site:        {self.site_id}",
            f"Solar:       {self.solar}",
            f"Records:     {self.n_records}",
            f"Trained:     {self.trained_at}",
            f"Avg kW:      {self.avg_kw:.1f}",
            f"Max kW:      {self.max_kw:.1f}",
            f"Avg kVAr:    {self.avg_kvar:.1f}",
            f"Avg PF:      {self.avg_pf:.4f}",
            f"PF std:      {self.pf_std:.4f}",
            f"Inject scale:{self.inject_scale:.2f}",
        ]
        return "\n".join(lines)


def load_all_models(models_dir: Path = MODELS_DIR) -> dict[str, SiteModel]:
    """Load all .pkl site model files. Returns {site_id: SiteModel}."""
    models = {}
    for pkl in models_dir.glob("site_*.pkl"):
        site_id = pkl.stem.replace("site_", "", 1)
        model   = SiteModel.load(site_id, models_dir)
        if model:
            models[site_id] = model
    return models
