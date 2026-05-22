"""
Baoyuan HyESys Agent — Strategic oversight layer.

Uses the local `claude` CLI (Claude Code) as its reasoning brain to analyse
site data, develop strategy, and propose config changes to Agent 1 and Agent 2.
Proposals are sent to Jason via Telegram (@JOstocks_bot).
Jason approves in Claude Code; this script applies the approved changes.

Run analysis:     python sites/baoyuan/site_agent.py
Apply a proposal: python sites/baoyuan/site_agent.py --apply <proposal_id>
Reject a proposal: python sites/baoyuan/site_agent.py --reject <proposal_id>
"""

import argparse
import json
import logging
import sqlite3
import subprocess
import urllib.request
from datetime import datetime, timezone, timedelta
from pathlib import Path

# ─────────────────────────────────────────────
# PATHS & CONSTANTS
# ─────────────────────────────────────────────
SITE_DIR      = Path(__file__).parent
DB_PATH       = SITE_DIR / "data" / "baoyuan.db"
A1_CONFIG     = SITE_DIR / "agent1_config.json"
A2_CONFIG     = SITE_DIR / "agent2_config.json"
PROPOSALS_DIR = SITE_DIR / "proposals"
TG_CONFIG     = SITE_DIR / "telegram_config.json"

SITE_NAME     = "Baoyuan"
SITE_FULL     = "Baoyuan Industrial, Zhuji, Zhejiang, China"
CLAUDE_MODEL  = "claude-sonnet-4-6"
PROJECT_ROOT  = Path(__file__).parent.parent.parent

SGT = timezone(timedelta(hours=8))

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
log = logging.getLogger("baoyuan.site_agent")


# ─────────────────────────────────────────────
# TELEGRAM
# ─────────────────────────────────────────────
def _send_telegram(message: str) -> bool:
    cfg   = json.loads(TG_CONFIG.read_text())
    url   = f"https://api.telegram.org/bot{cfg['token']}/sendMessage"
    data  = json.dumps({
        "chat_id":    cfg["chat_id"],
        "text":       message,
        "parse_mode": "HTML",
    }).encode("utf-8")
    req = urllib.request.Request(
        url, data=data,
        headers={"Content-Type": "application/json"},
        method="POST",
    )
    try:
        with urllib.request.urlopen(req) as r:
            result = json.loads(r.read())
            return bool(result.get("ok"))
    except Exception as e:
        log.error("Telegram send failed: %s", e)
        return False


# ─────────────────────────────────────────────
# DATA GATHERING
# ─────────────────────────────────────────────
def _gather_site_data() -> dict:
    if not DB_PATH.exists():
        return {"error": "baoyuan.db not found — run agent1.py first"}

    conn = sqlite3.connect(str(DB_PATH))
    conn.row_factory = sqlite3.Row

    # Record counts by site + quality tag
    tag_rows = conn.execute("""
        SELECT site_id, quality_tag, COUNT(*) as n
        FROM meter_records GROUP BY site_id, quality_tag
    """).fetchall()
    record_counts: dict = {}
    for r in tag_rows:
        record_counts.setdefault(r["site_id"], {})[r["quality_tag"]] = r["n"]

    # Time range
    tr = conn.execute("""
        SELECT MIN(timestamp) as first, MAX(timestamp) as last, COUNT(*) as total
        FROM meter_records
    """).fetchone()

    # SAR outcome distribution
    sar_rows = conn.execute("""
        SELECT outcome, COUNT(*) as n FROM sar_log GROUP BY outcome
    """).fetchall()
    sar = {r["outcome"]: r["n"] for r in sar_rows}
    sar_total = sum(sar.values())

    # Phase current imbalance stats per site
    imbalance: dict = {}
    for site_id in ["BAOYUAN-CAPBANK1", "BAOYUAN-CAPBANK2"]:
        rows = conn.execute("""
            SELECT Ia, Ib, Ic FROM meter_records
            WHERE site_id=? AND quality_tag='CLEAN' AND Ia>0 AND Ib>0 AND Ic>0
        """, (site_id,)).fetchall()
        if rows:
            imbs = []
            for row in rows:
                ia, ib, ic = row["Ia"], row["Ib"], row["Ic"]
                i_max = max(ia, ib, ic)
                i_min = min(ia, ib, ic)
                if i_max > 0:
                    imbs.append((i_max - i_min) / i_max)
            if imbs:
                imbalance[site_id] = {
                    "mean_pct": round(sum(imbs) / len(imbs) * 100, 1),
                    "max_pct":  round(max(imbs) * 100, 1),
                    "samples":  len(imbs),
                }

    # Sample of recent SAR actions
    recent_sar = conn.execute("""
        SELECT site_id, timestamp, action, action_kVAr, outcome
        FROM sar_log ORDER BY timestamp DESC LIMIT 10
    """).fetchall()

    conn.close()

    return {
        "record_counts":    record_counts,
        "time_range":       {"first": tr["first"], "last": tr["last"], "total": tr["total"]},
        "sar_outcomes":     sar,
        "sar_total":        sar_total,
        "positive_pct":     round(sar.get("POSITIVE", 0) / max(sar_total, 1) * 100, 1),
        "negative_pct":     round(sar.get("NEGATIVE", 0) / max(sar_total, 1) * 100, 1),
        "imbalance_stats":  imbalance,
        "recent_sar":       [dict(r) for r in recent_sar],
    }


def _load_proposal_history() -> list[dict]:
    PROPOSALS_DIR.mkdir(exist_ok=True)
    history = []
    for f in sorted(PROPOSALS_DIR.glob("proposal_*.json")):
        try:
            history.append(json.loads(f.read_text()))
        except Exception:
            pass
    return history[-10:]


def _next_proposal_id() -> str:
    PROPOSALS_DIR.mkdir(exist_ok=True)
    return str(len(list(PROPOSALS_DIR.glob("proposal_*.json"))) + 1).zfill(3)


# ─────────────────────────────────────────────
# CLAUDE API — SYSTEM PROMPT (cached)
# ─────────────────────────────────────────────
_SYSTEM_PROMPT = """You are the Baoyuan HyESys Agent — the AI strategic brain responsible for improving electrical savings at Baoyuan Industrial (诸暨市葆元实业有限公司), Zhuji, Zhejiang, China.

## Site Overview
- HyESys H125 (125 kVA active digital compensator) installed at MSB1
- CapBank1 and CapBank2 are passive capacitor banks monitored via MQTT
- CRITICAL: CapBank MQTT meters are CURRENT-ONLY instruments — Ia, Ib, Ic are the ONLY meaningful fields. kW, kVAr, PF, and Voltage always report 0. This is normal hardware behaviour, NOT a fault.
- No live Main Grid meter — HyESys grid-level effect is only visible via periodic history reports from Baoyuan
- Site has heavy VFD loads causing very high harmonic distortion (THD 44–142% across phases)
- Baseline 3-phase current ~3,450 A, displacement PF 0.807–0.867, true apparent PF 0.508–0.739
- Observed I²R loss reduction: 5.2–7.5% per HyESys activation event (mean 6.7%)

## Your Sub-Agents
Agent 1 (agent1_config.json): Validates incoming CapBank MQTT data
- zero_current_threshold_A: current below this = SUSPECT (meter dropout)
- imbalance_threshold_pct: phase imbalance fraction above this = SUSPECT (e.g. 0.10 = 10%)

Agent 2 (agent2_config.json): Analyses records and logs decisions
- action_mode: "pf_pi_control" (PI controller on PF — INAPPROPRIATE for CapBank-only data) or "current_imbalance_monitor" (monitors Ia/Ib/Ic patterns — CORRECT for this site)
- pf_target: target PF for PI mode
- pi_control.k_p, k_i, i_max, deadband, dt_hours: PI tuning parameters
- current_imbalance_monitor.imbalance_alert_pct: threshold above which current imbalance is flagged as NEGATIVE outcome

## Experimental Phase Context
Baoyuan is in a data-collection phase. There is no real-time grid feedback. The workflow is:
1. Set HyESys to a mode → run for 2 weeks
2. Collect CapBank Ia/Ib/Ic live (the only real-time observable)
3. Receive Main Grid history from Baoyuan after the period
4. Correlate HyESys mode × CapBank patterns × grid consumption
5. Retrain strategy, repeat with different mode/settings

## Output Format
Return ONLY valid JSON — no markdown fences, no explanation outside the JSON:
{
  "analysis_summary": "2–3 sentences on key findings",
  "should_propose": true or false,
  "reason_no_proposal": "if false — why no change is needed now",
  "changes": [
    {
      "agent": "agent1" or "agent2",
      "config_file": "agent1_config.json" or "agent2_config.json",
      "parameter": "dot.notation.path to the parameter",
      "current_value": <current value>,
      "proposed_value": <proposed value>,
      "reason": "clear, specific reason why this change improves outcomes"
    }
  ],
  "expected_outcome": "what will improve and how it will be measured"
}"""


# ─────────────────────────────────────────────
# CALL CLAUDE API
# ─────────────────────────────────────────────
def _call_claude(data: dict, a1_cfg: dict, a2_cfg: dict, history: list[dict]) -> dict:
    user_msg = (
        f"Current date/time (SGT): {datetime.now(SGT).strftime('%Y-%m-%d %H:%M')}\n\n"
        f"## Site Data Summary\n{json.dumps(data, indent=2, default=str)}\n\n"
        f"## Current Agent 1 Config (agent1_config.json)\n{json.dumps(a1_cfg, indent=2)}\n\n"
        f"## Current Agent 2 Config (agent2_config.json)\n{json.dumps(a2_cfg, indent=2)}\n\n"
        f"## Proposal History (last {len(history)})\n"
        f"{json.dumps(history, indent=2, default=str) if history else 'No previous proposals.'}\n\n"
        f"Analyse the above and propose the next strategic change to improve electrical savings at Baoyuan."
    )

    full_prompt = _SYSTEM_PROMPT + "\n\n---\n\n" + user_msg

    log.info("Calling claude CLI (%s)...", CLAUDE_MODEL)
    result = subprocess.run(
        ["claude", "-p", full_prompt, "--model", CLAUDE_MODEL],
        capture_output=True,
        text=True,
        timeout=120,
        cwd=str(PROJECT_ROOT),
    )

    if result.returncode != 0:
        raise RuntimeError(f"claude CLI failed (exit {result.returncode}): {result.stderr[:500]}")

    raw = result.stdout.strip()
    log.info("Claude responded (%d chars)", len(raw))

    try:
        return json.loads(raw)
    except json.JSONDecodeError:
        import re
        m = re.search(r'\{.*\}', raw, re.DOTALL)
        if m:
            return json.loads(m.group())
        raise ValueError(f"Claude returned non-JSON output: {raw[:300]}")


# ─────────────────────────────────────────────
# TELEGRAM PROPOSAL FORMATTER
# ─────────────────────────────────────────────
def _format_proposal_message(proposal: dict) -> str:
    pid    = proposal["id"]
    basis  = proposal["data_basis"]
    result = proposal["claude_result"]
    changes = result.get("changes", [])

    lines = [
        f"🏭 <b>[{SITE_NAME}] HyESys Agent — Strategy Proposal #{pid}</b>",
        f"📍 {SITE_FULL}",
        "",
        f"📊 <b>Data basis:</b> {basis['total_records']} records",
        f"   SAR: ✅ POSITIVE {basis['positive_pct']}%  ❌ NEGATIVE {basis['negative_pct']}%",
        "",
        "🔍 <b>Analysis:</b>",
        result.get("analysis_summary", ""),
        "",
    ]

    if changes:
        plural = "s" if len(changes) > 1 else ""
        lines.append(f"💡 <b>Recommended change{plural} — {len(changes)} item(s):</b>")
        for i, c in enumerate(changes, 1):
            lines += [
                "",
                f"  📝 <b>Change {i} of {len(changes)}</b>",
                f"  File:      <code>{c['config_file']}</code>  ({c['agent'].upper()})",
                f"  Parameter: <code>{c['parameter']}</code>",
                f"  Current:   <code>{json.dumps(c['current_value'])}</code>",
                f"  Proposed:  <code>{json.dumps(c['proposed_value'])}</code>",
                f"  Why: {c['reason']}",
            ]
        lines += [
            "",
            "📈 <b>Expected outcome:</b>",
            result.get("expected_outcome", ""),
        ]
    else:
        lines.append("ℹ️ No config changes proposed at this time.")

    lines += [
        "",
        "─" * 28,
        f"📋 Proposal ID: <b>{pid}</b>",
        f'✅ To apply: tell HyESys Agent <b>"approve proposal {pid}"</b> in Claude Code.',
        f'❌ To reject: tell HyESys Agent <b>"reject proposal {pid}"</b> in Claude Code.',
    ]

    return "\n".join(lines)


# ─────────────────────────────────────────────
# RUN ANALYSIS → GENERATE PROPOSAL
# ─────────────────────────────────────────────
def run_analysis():
    log.info("Baoyuan HyESys Agent — starting analysis...")

    data     = _gather_site_data()
    a1_cfg   = json.loads(A1_CONFIG.read_text()) if A1_CONFIG.exists() else {}
    a2_cfg   = json.loads(A2_CONFIG.read_text()) if A2_CONFIG.exists() else {}
    history  = _load_proposal_history()

    if "error" in data:
        log.error(data["error"])
        return

    result = _call_claude(data, a1_cfg, a2_cfg, history)

    if not result.get("should_propose"):
        reason = result.get("reason_no_proposal", "No change needed.")
        log.info("No proposal: %s", reason)
        return

    pid = _next_proposal_id()
    proposal = {
        "id":         pid,
        "created_at": datetime.now(SGT).isoformat(),
        "status":     "pending",
        "site":       SITE_NAME,
        "data_basis": {
            "total_records": data.get("time_range", {}).get("total", 0),
            "positive_pct":  data.get("positive_pct", 0),
            "negative_pct":  data.get("negative_pct", 0),
        },
        "claude_result": result,
        "applied_at":  None,
    }

    PROPOSALS_DIR.mkdir(exist_ok=True)
    path = PROPOSALS_DIR / f"proposal_{pid}.json"
    path.write_text(json.dumps(proposal, indent=2, default=str))
    log.info("Proposal #%s saved to %s", pid, path)

    msg = _format_proposal_message(proposal)
    if _send_telegram(msg):
        log.info("Proposal #%s sent via Telegram.", pid)
    else:
        log.warning("Telegram send failed. Proposal content:")
        log.info(msg)


# ─────────────────────────────────────────────
# APPLY PROPOSAL
# ─────────────────────────────────────────────
def _set_nested(obj: dict, dotpath: str, value) -> None:
    keys = dotpath.split(".")
    for k in keys[:-1]:
        obj = obj.setdefault(k, {})
    obj[keys[-1]] = value


def apply_proposal(proposal_id: str):
    pid  = str(proposal_id).zfill(3)
    path = PROPOSALS_DIR / f"proposal_{pid}.json"

    if not path.exists():
        log.error("Proposal %s not found.", pid)
        return

    proposal = json.loads(path.read_text())

    if proposal["status"] == "applied":
        log.warning("Proposal %s already applied on %s.", pid, proposal["applied_at"])
        return

    changes = proposal["claude_result"].get("changes", [])
    a1_cfg  = json.loads(A1_CONFIG.read_text()) if A1_CONFIG.exists() else {}
    a2_cfg  = json.loads(A2_CONFIG.read_text()) if A2_CONFIG.exists() else {}

    change_lines = []
    for c in changes:
        cfg = a1_cfg if c["agent"] == "agent1" else a2_cfg
        _set_nested(cfg, c["parameter"], c["proposed_value"])
        change_lines.append(
            f"  • {c['config_file']} › <code>{c['parameter']}</code>: "
            f"<code>{json.dumps(c['current_value'])}</code> → "
            f"<code>{json.dumps(c['proposed_value'])}</code>"
        )
        log.info("Applied: %s › %s = %s", c["config_file"], c["parameter"], c["proposed_value"])

    A1_CONFIG.write_text(json.dumps(a1_cfg, indent=2))
    A2_CONFIG.write_text(json.dumps(a2_cfg, indent=2))

    proposal["status"]     = "applied"
    proposal["applied_at"] = datetime.now(SGT).isoformat()
    path.write_text(json.dumps(proposal, indent=2, default=str))

    confirm = (
        f"✅ <b>[{SITE_NAME}] Proposal #{pid} — Applied</b>\n\n"
        f"<b>Changes made to agent configs:</b>\n"
        + "\n".join(change_lines)
        + "\n\nAgent 1 and Agent 2 will use the new configuration on next run."
    )
    _send_telegram(confirm)
    log.info("Proposal #%s applied successfully.", pid)


def reject_proposal(proposal_id: str):
    pid  = str(proposal_id).zfill(3)
    path = PROPOSALS_DIR / f"proposal_{pid}.json"

    if not path.exists():
        log.error("Proposal %s not found.", pid)
        return

    proposal = json.loads(path.read_text())

    if proposal["status"] != "pending":
        log.warning("Proposal %s is already '%s'.", pid, proposal["status"])
        return

    proposal["status"]     = "rejected"
    proposal["applied_at"] = datetime.now(SGT).isoformat()
    path.write_text(json.dumps(proposal, indent=2, default=str))

    confirm = f"❌ <b>[{SITE_NAME}] Proposal #{pid} — Rejected</b>\nNo changes made. Next analysis will consider this rejection."
    _send_telegram(confirm)
    log.info("Proposal #%s rejected.", pid)


# ─────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────
def main():
    parser = argparse.ArgumentParser(description="Baoyuan HyESys Agent")
    group  = parser.add_mutually_exclusive_group()
    group.add_argument("--apply",  metavar="ID", help="Apply a pending proposal")
    group.add_argument("--reject", metavar="ID", help="Reject a pending proposal")
    args = parser.parse_args()

    if args.apply:
        apply_proposal(args.apply)
    elif args.reject:
        reject_proposal(args.reject)
    else:
        run_analysis()


if __name__ == "__main__":
    main()
