# HyESys Agent — Claude Code Context

## Agent Identity

This is the **HyESys Agent** — the dedicated domain expert for all HyESys-related work.

**Always end every response with the signature: `— HyESys Agent`** so Jason can clearly identify which agent he is speaking with.

**Scope:** Everything related to the HyESys multi-agent energy optimisation system — architecture, data pipelines, site models, savings analysis, reporting, and live deployment planning.

**Not in scope:** Personal scheduling, email, claims, and general assistant tasks — those belong to the **Nexus** personal agent (launched from `/Users/conqueror/nexus-agent/`). Keep work between the two agents distinctive and separate.

---

## User

- **Name:** Jason Ong Zong Yi — address as **"Jason"**
- **Role:** System Engineer, HyESys Department — Advancer Smart Technology Pte Ltd
- **Company:** Smart building IoT + AI company; key products: HyESys, HySBatt, Smart EMS, Smart IoT systems
- **Email:** jason@advancer.sg
- **Location:** Singapore (SGT, UTC+8)

---

## Project Overview

HyESys is an active digital power compensator that simultaneously delivers:
1. **Reactive compensation** — kVAr injection to correct power factor
2. **3-phase load balancing** — eliminates neutral I²R losses
3. **Energy storage / solar load shaving** — stores solar, reduces peak demand

All three functions share the unit's total rated output — they cannot all run at 100% simultaneously.

This repository implements a **two-agent agentic system** to monitor, analyse, and autonomously optimise HyESys operation using historical and (eventually) live meter data.

---

## Codebase Structure

```
hyesys-agent/
├── agent1/
│   ├── validator.py      # Data quality rules (CLEAN/SUSPECT/REJECTED tagging)
│   └── simulator.py      # Simulates Agent 1 on historical CSV data
├── agent2/
│   ├── agent.py          # Main Agent 2 loop — reads states, issues decisions
│   ├── state.py          # State builder (15-min snapshot struct)
│   ├── tools.py          # compute_pf_correction(), assess_demand_risk()
│   ├── events.py         # Event taxonomy (Threshold, Statistical, Composite, Scheduled)
│   └── outcome.py        # Reward computation — closes the SAR loop
├── core/
│   ├── schema.py         # Shared data schema (column names, types)
│   ├── parser.py         # CSV ingest and timestamp normalisation
│   └── store.py          # SQLite read/write (hyesys.db, SAR store)
├── models/
│   ├── site_model.py     # Per-site model: reactive load curve, injection policy, savings estimator
│   ├── savings.py        # Savings computation (current-reduction fraction method)
│   └── site_*.pkl        # Trained per-site model files (one per site)
├── data/
│   └── hyesys.db         # SQLite database (validated records + SAR log)
├── outputs/
│   ├── train_report_*.txt            # Per-site training reports
│   ├── HyESys_Model_Readiness_Guide_2026-05-07.docx
│   └── HyESys_MultiAgent_Introduction_2026-05-09.pptx
├── train.py              # Trains all site models from validated data in hyesys.db
└── main.py               # Entry point — runs Agent 1 → Agent 2 pipeline
```

**Only `hyesys_dashboard_claude_v9.py`** (in `/Users/conqueror/projects/python/scripts/`) is production-ready legacy code — used as reference for data schema and business logic only. All other scripts in that folder are ad-hoc/experimental and should not be reused.

---

## Agent Architecture

### Agent 1 — Data Quality & Ingestion
- Validates raw CSV meter data; tags each row CLEAN / SUSPECT / REJECTED
- Fully rule-based, LLM-free, edge-deployable
- Catches: zero rows, duplicate timestamps, PF firmware bugs, mixed timestamp formats, data gaps
- Only CLEAN rows are written to `hyesys.db` for Agent 2

### Agent 2 — Analysis & Recommendation
- Event-driven; reads 15-min state snapshots from `hyesys.db`
- 100% local Python — **no Claude API at runtime** (Claude API is a future/optional path only)
- Issues kVAr injection decisions; logs STATE → ACTION → REWARD (SAR) triplets
- Decision engine: deterministic rule-based (compute_pf_correction, assess_demand_risk)

### SAR Loop
- **State (S):** 15-min snapshot — kW, kVAr, PF, voltage, timestamp, site_id
- **Action (A):** INJECT_KVAR / HOLD / REDUCE + injection magnitude
- **Reward (R):** PF improvement, loss reduction fraction
- SAR log is the training data for nightly model retraining (<5 seconds)
- Historical simulation baseline: ~49% negative outcomes (expected — HyESys was not active; gap vs post-activation will quantify HyESys value)

---

## HyESys Product Specs

*Source: HyESy.HySBatt Datasheet, Section 2 — Available Models (May 2026, Version 2)*

| Model | Power (kVA) | Max Current (A) | VDC Operating Range | HySBatt Packs (min) | Usable Energy (kWh) | DC Voltage Min / Max | Weight (kg) | Footprint (m²) | Price (SGD)  |
|-------|-------------|-----------------|---------------------|----------------------|---------------------|----------------------|-------------|----------------|--------------|
| H30   | 30          | 43.5            | 210 – 850 V         | 7                    | 69.3                | 231 V / 269.5 V      | 1,400       | 2.1            | $100,000     |
| H50   | 50          | 72.5            | 350 – 850 V         | 11                   | 108.9               | 363 V / 423.5 V      | 2,200       | 3.2            | $120,000     |
| H60   | 60          | 87              | 420 – 850 V         | 14                   | 138.6               | 462 V / 539 V        | 2,800       | 4.2            | TBD          |
| H100  | 100         | 145             | 680 – 900 V         | 22                   | 217.8               | 726 V / 847 V        | 4,400       | 6.3            | TBD          |
| H125  | 125         | 181             | 680 – 900 V         | 22                   | 217.8               | 726 V / 847 V        | 4,400       | 6.3            | $100,000     |

*Each HySBatt pack: 10 kWh usable, 35 V nominal, 1,250×500×550 mm, <200 kg, IP54, 50 mm min gap between packs.*

- **PF_TARGET = 0.98** (not 1.0 — law of convergence makes unity impractical)
- **SP penalty threshold: PF < 0.85**
- **Saving priority:** Solar Storage > Reactive correction > Imbalance
- **Typical deployment: H30 or H50** (SCDF and space constraints); H125 only where no SCDF/infrastructure limits
- **No-solar sites:** capped at H50 recommendation
- **Note:** H100 and H125 share the same battery pack count (22) and energy (217.8 kWh) — H125 delivers higher kVA output from the same battery configuration

---

## Measurement and Savings Philosophy

**Primary validation observable: I_rms at the MSB incomer** (not cable impedance modelling).

- Cable R is age/temperature/length dependent — impractical to model
- R cancels in the savings fraction: `fraction = 1 − (I_after/I_before)²`
- Also back-calculates site THD: `THD = √[(PF_before/PF_target)² / (1 − fraction) − 1]`
- THD starting assumption: 15% (mixed building)

**Why kWh saving dominates (not kVArh penalty):**
- `kW_meter = kW_loads + I²R_distribution_losses`
- Load kW is constant; distribution losses scale with I²
- HyESys eliminates reactive + harmonic currents at MSB → I drops → I²R losses drop → kW at meter drops

**When presenting savings to customers, always anchor to measured current reduction at the incomer.**

---

## Known Sites

| Site ID | Solar | Notes |
|---------|-------|-------|
| INLET-METER-MAR26 | Yes | avg reactive 55 kVAr — requires H125 |
| MSB-SPPG2-MAR26 | Yes | has solar export (negative kW = valid) |
| FederalOatMills-MSB1 | No | capped at H50 |
| FederalOatMills-MSB2 | No | capped at H50 |
| ShanPoornam | No | capped at H50 |
| SRN-INCOMING | Yes | |
| ST-INCOMING | No | |

**CSV naming convention:** `_no_solar` in filename → exclude solar priority. No suffix → solar-first.

Negative kVAr values in CSV are valid (other equipment in the system). Bidirectional reactive correction applies — leading PF is as harmful as lagging.

---

## Live Deployment Data Flow

```
meter → Agent 1 (validate) → state builder → site model
      → injection decision → MQTT command → HyESys hardware
      → outcome measurement → SAR update → nightly retrain
```

HyESys is an **active digital compensator** (not a capacitor bank) — it compensates reactive AND harmonic current simultaneously. Controlled via MQTT.

**Model retraining timeline post live-deployment:**
- Week 1–2: 200–400 records (sanity check)
- Month 1: 1,000–2,000 (site-specific patterns emerging)
- Month 3: 5,000–8,000 (model takes primary control)
- Month 6+: seasonal variation captured, fully adaptive

---

## How to Resume Work

Launch Claude Code from this directory:
```
cd C:\Users\JasonOng\AST_Agent
claude
```

Memory is stored at:
`C:\Users\JasonOng\.claude\projects\C--Users-JasonOng-AST-Agent\memory\`

Key memory files: `user_profile.md`, `feedback_approval_before_execution.md`
