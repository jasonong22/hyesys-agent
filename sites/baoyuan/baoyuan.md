# Baoyuan Site — HyESys Agent Context

## Site Identity

- **Company:** Baoyuan Industrial (诸暨市葆元实业有限公司), Zhuji, Zhejiang, China
- **Site ID prefix:** `BAOYUAN-`
- **HyESys unit:** H125 (125 kVAr / 200 kWh)
- **Solar:** No
- **Currency:** CNY
- **Analysis period:** 20 March – 6 May 2026

---

## Electrical Architecture

```
Main Grid (10 kV)
│
├── MSB1 (not metered)
│   ├── CapBank1  ← BAOYUAN-CAPBANK1  (metered via MQTT)
│   ├── CapBank2  ← BAOYUAN-CAPBANK2  (metered via MQTT)
│   ├── Cabinet A (metered)
│   │   ├── Cabinet 1
│   │   ├── Cabinet 2 — HyESys H125  ← active digital compensator
│   │   ├── Cabinet 3
│   │   └── Cabinet 4
│   ├── Cabinet B (metered, feeders only)
│   └── Cabinet C (metered, feeders only)
│
└── MSB2 (not metered)
```

**Key constraint:** MSB1 and MSB2 have no meters. The HyESys effect is diluted at the Main Grid level because MSB2 load is not separately metered. This is documented as the main measurement limitation in the savings report.

**CapBank1 and CapBank2** are passive capacitor banks providing reactive compensation at the MSB1 bus (connected at the 10 kV / LV transformer secondary). They are separate from and operate in parallel with HyESys.

---

## MQTT Topics and Payload Format

### Topics
| Topic | Site ID | Device |
|---|---|---|
| `hyesys/data/dev/0086040215999997` | `BAOYUAN-CAPBANK1` | CapBank1 meter |
| `hyesys/data/dev/0086040215999996` | `BAOYUAN-CAPBANK2` | CapBank2 meter |

**Gateway SN:** `26022703840003` (same gateway publishes BMS data on `stsc/aems/message/26022703840003`)

### Payload Structure (nested JSON)
```json
{
  "method": "update",
  "reported": {
    "0_5_<device_id>": {
      "P":   "<kW total>",
      "Pa":  "<kW phase A>",   "Pb": "<kW phase B>",   "Pc": "<kW phase C>",
      "Q":   "<kVAr total>",
      "Qa":  "<kVAr phase A>", "Qb": "<kVAr phase B>", "Qc": "<kVAr phase C>",
      "S":   "<kVA total>",
      "Sa":  "<kVA phase A>",  "Sb": "<kVA phase B>",  "Sc": "<kVA phase C>",
      "PF":  "<power factor>",
      "PFa": "<PF phase A>",   "PFb": "<PF phase B>",  "PFc": "<PF phase C>",
      "Ua":  "<V phase A>",    "Ub":  "<V phase B>",   "Uc":  "<V phase C>",
      "Uab": "<V line AB>",    "Ubc": "<V line BC>",   "Uca": "<V line CA>",
      "Ia":  "<A phase A>",    "Ib":  "<A phase B>",   "Ic":  "<A phase C>",
      "Fr":  "<Hz frequency>",
      "EP":  "<kWh total energy>",
      "EPI": "<kWh import>",   "EPE": "<kWh export>",
      "state": "ONLINE/OFFLINE"
    }
  },
  "sn":       "<gateway_id>",
  "sendtime": <unix_timestamp_seconds>,
  "msgid":    <int>,
  "version":  1,
  "timestamp": <unix_timestamp_seconds>
}
```

**CapBank meters are current-only instruments.** Voltage (Ua/Ub/Uc), kW, kVAr, PF will always report 0 — this is expected hardware behaviour, not a fault. Only Ia/Ib/Ic are meaningful fields from these meters.

---

## Site Electrical Profile

*Source: Baoyuan HyESys Savings Report, 12 May 2026*

| Parameter | Value | Notes |
|---|---|---|
| Baseline kW (MSB1) | ~300 kW | Production load |
| Baseline I_rms 3-phase | ~3,450 A | Before HyESys activation |
| Displacement PF | 0.807–0.867 | Phase A–C |
| **True apparent PF** | **0.508–0.739** | Very low due to harmonics |
| THD — Phase A | ~126% | Extremely high — VFDs |
| THD — Phase B | ~44% | Moderate |
| THD — Phase C | ~142% | Critically high |
| I²R loss reduction | 5.2–7.5% per event | Mean 6.7% |
| Annual savings (conservative) | CNY 257,000 | Directly observed |
| Annual savings (upper bound) | CNY 390,000 | I²R fraction method |

**Note:** The large gap between displacement PF and true apparent PF is caused by very high harmonic content from VFDs (variable frequency drives). HyESys addresses both reactive and harmonic currents simultaneously — this is why the observed current reduction exceeds what pure reactive compensation predicts.

---

## Experimental Phase — Operational Model

Baoyuan is in an **experimental data-collection phase**. The end goal is an electrical savings solution derived from studying the relationship between HyESys operation and grid consumption. The path to that goal is iterative:

### System Components

| Component | Role | Data type | Availability |
|---|---|---|---|
| **HyESys H125** | Input — active digital compensator | Operational mode / settings | Controlled by AST |
| **CapBank 1** | Output — passive reactive bank | 3-phase current (Ia/Ib/Ic) only | Live via MQTT |
| **CapBank 2** | Output — passive reactive bank | 3-phase current (Ia/Ib/Ic) only | Live via MQTT |
| **Main Grid** | Output — total site consumption | kW, kWh, Voltage, Current, PF | Periodic (provided by Baoyuan) |

### Workflow

```
1. Set HyESys to an experimental mode
         ↓
2. Run blindly for a test period (e.g. 2 weeks)
   — CapBank 1 & 2 MQTT data collected live (partial signal only)
   — Main Grid effect NOT visible in real time (no meter)
         ↓
3. Receive Main Grid consumption history for the test period
   — kW, kWh, Voltage, Current, PF from Baoyuan's records
         ↓
4. Import history → correlate HyESys mode vs CapBank behaviour vs grid consumption
         ↓
5. Retrain model on the combined dataset
         ↓
6. Repeat with a different HyESys mode or setting
         ↓
7. Converge on optimal electrical savings configuration
```

### Key Constraints

- **Blind testing periods:** HyESys effect on the Main Grid cannot be observed in real time. Decisions during a test run cannot use grid feedback.
- **CapBank as proxy signal:** The only live observable is CapBank 1/2 current. Changes here give a partial, downstream indication of HyESys activity.
- **Periodic batch learning:** The model is retrained in batches after each Main Grid history import — not continuously.
- **Specifics TBD:** Experimental modes, test durations, model architecture, and savings quantification method will be defined in subsequent sessions.

---

## Known Issues

1. **Cotton contamination** — Cotton has been found blocking HyESys ventilation fans, causing overheating shutdowns. Needs compressed air cleaning.
2. **MSB1/MSB2 not metered** — HyESys effect is diluted at the Main Grid measurement point. Recommendation: install meters at MSB1 and MSB2 level.

---

## Agent 1 — Validation Context

- These are **current-only capacitor bank meters** — Ia/Ib/Ic are the only meaningful fields
- Voltage, kW, kVAr, PF always report 0 — this is normal hardware behaviour, not a fault
- SUSPECT: all currents = 0 (meter dropout or cap bank offline)
- SUSPECT: phase current imbalance > 10%
- All other records: CLEAN

## Agent 2 — Decision Context

Baoyuan Agent 2 operates differently from a standard real-time PI controller because there is no live Main Grid meter feedback.

**Current phase (experimental):**
- Monitors CapBank 1/2 current patterns as the only live signal
- Logs HyESys operational mode and timestamp for each test period
- Does NOT issue real-time injection commands (no live grid feedback to close the loop)

**After each Main Grid history import:**
- Ingests the periodic kW/kWh/Voltage/Current/PF history
- Correlates: HyESys mode × CapBank current changes × grid consumption delta
- Retrains model on the combined dataset

**End state (post-experimental):**
- Model trained on sufficient mode-vs-outcome data
- Agent 2 transitions to recommending or commanding optimal HyESys settings
- No solar priority override (solar = False)
- Recommended model: H125
