# Baoyuan Site — Reference Document

**Company:** Baoyuan Industrial (诸暨市葆元实业有限公司), Zhuji, Zhejiang, China
**HyESys unit:** H125 (125 kVA active digital compensator)
**Solar:** No  |  **Currency:** CNY  |  **Site ID prefix:** `BAOYUAN-`

---

## Electrical Architecture

```
Main Grid (10 kV)
│
├── MSB1 (not metered)
│   ├── CapBank1  ← BAOYUAN-CAPBANK1  (MQTT current-only meter)
│   ├── CapBank2  ← BAOYUAN-CAPBANK2  (MQTT current-only meter)
│   └── Cabinet A → Feeder 2 — HyESys H125  (active digital compensator)
│
└── MSB2 (not metered)
```

**Key constraint:** MSB1 and MSB2 have no grid meters. HyESys effect is diluted at the Main Grid measurement point because MSB2 load is not separately observable.

---

## Site Electrical Profile

*Source: Baoyuan HyESys Savings Report, 12 May 2026*

| Parameter | Value | Notes |
|---|---|---|
| Baseline kW | ~300 kW | MSB1 production load |
| Baseline I_rms | ~3,450 A | 3-phase, before HyESys |
| Displacement PF | 0.807–0.867 | Phase A–C |
| True apparent PF | 0.508–0.739 | Very low — heavy harmonic content |
| THD Phase A / B / C | ~126% / 44% / 142% | Caused by VFDs |
| I²R loss reduction | 5.2–7.5% per event (mean 6.7%) | Directly observed |
| Annual savings (conservative) | CNY 257,000 | |
| Annual savings (upper bound) | CNY 390,000 | I²R fraction method |

The gap between displacement PF and true apparent PF is caused by VFDs (variable frequency drives). HyESys compensates both reactive and harmonic currents — which is why observed current reduction exceeds pure reactive compensation predictions.

---

## MQTT — Data Ingestion (Agent 1)

### Topics

| Topic | Site ID | Device | Data |
|---|---|---|---|
| `hyesys/data/dev/0086040215999997` | `BAOYUAN-CAPBANK1` | CapBank 1 meter | Current only |
| `hyesys/data/dev/0086040215999996` | `BAOYUAN-CAPBANK2` | CapBank 2 meter | Current only |
| `stsc/aems/message/26022703840003` | `BAOYUAN-HYESYS` | HyESys H125 | Full electrical + BMS |

**CapBank meters are current-only instruments.** Voltage, kW, kVAr, PF always report 0 — this is normal hardware behaviour, not a fault. Only Ia/Ib/Ic are meaningful.

### CapBank Payload Format
```json
{
  "reported": {
    "0_5_<device_id>": {
      "Ia": "<A>", "Ib": "<A>", "Ic": "<A>",
      "Ua": 0, "Ub": 0, "Uc": 0,
      "P": 0, "Q": 0, "PF": 0
    }
  },
  "sendtime": "<unix_seconds>"
}
```

### HyESys H125 Payload Format (`stsc/aems/message/26022703840003`)
All data nested under `raw["data"]`. Timestamp is `reportTimeTs` in milliseconds.

**`type: "pcs_v3"`** — electrical readings:
```json
{
  "type": "pcs_v3",
  "reportTimeTs": "<ms>",
  "data": {
    "voltageA": "<V>", "voltageB": "<V>", "voltageC": "<V>",
    "currentA": "<A>", "currentB": "<A>", "currentC": "<A>",
    "activePowerTotal": "<kW>",   "reactivePowerTotal": "<kVar>",
    "apparentPowerTotal": "<kVA>", "powerFactorTotal": "<PF>",
    "activePowerA": "<kW>",  "activePowerB": "<kW>",  "activePowerC": "<kW>",
    "reactivePowerA": "<kVar>", "reactivePowerB": "<kVar>", "reactivePowerC": "<kVar>",
    "hz": "<Hz>", "temperature": "<°C>",
    "inputPower": "<kW>", "inputVoltage": "<V>", "inputCurrent": "<A>"
  }
}
```

**`type: "bms"`** — battery management:
```json
{
  "type": "bms",
  "reportTimeTs": "<ms>",
  "data": {
    "singleVoltageAvg": "<V>",
    "singleVoltageMax": "<V>", "singleVoltageMin": "<V>",
    "soc": "<0–100>", "soh": "<0–100>",
    "voltage": "<V>", "current": "<A>",
    "tempMain": "<°C>"
  }
}
```

---

## MQTT — Command Output (Agent 2)

**Command topic:** `stsc/aems/cabinet/26022703840003/multi/operate/tx`
*(Confirmed format: Jason 2026-05-26)*

**kVAr setpoint:**
```json
{
  "cabinetId": "26022703840003", "index": 1,
  "key": "set_reactive_power", "remote": true,
  "params": { "reactivePowerA": 0, "reactivePowerB": 0, "reactivePowerC": 0 }
}
```

**kW setpoint:**
```json
{
  "cabinetId": "26022703840003", "index": 1,
  "key": "set_active_power", "remote": true,
  "params": { "activePowerA": 0, "activePowerB": 0, "activePowerC": 0 }
}
```

**Per-phase hard limit: `|kW_per_phase| + |kVAr_per_phase| ≤ 40`**
Both commands are always sent as separate messages. kW is always sent (even when 0) to prevent stale setpoints.

---

## Experimental Phase

Baoyuan has no live Main Grid meter — real-time closed-loop optimisation is not possible. The approach is iterative blind testing:

```
1. Set HyESys to a mode → run for ~2 weeks
2. Collect CapBank Ia/Ib/Ic live (only real-time observable)
3. Receive Main Grid history CSV from Baoyuan after the period
4. Import → correlate HyESys mode × CapBank response × grid consumption
5. Retrain model → repeat with refined settings
```

### Current experiment: `sweep_001` (kvar_sweep_experiment)
- Sweeps kVAr injection: 0 → 120 kVAr in +1.0 kVAr steps, 1 hour per step (~5 days total)
- Study inputs: HyESys kVAr (1), Study outputs: CapBank1 current, CapBank2 current, Main Grid (3)
- BMS safety: pause + charge at −10 kW/phase if `singleVoltageAvg < 3.20 V`; resume when `≥ 3.45 V`
- PCS temp safety: hard cap if temp ≥ 83°C

### Main Grid history import
Script: `sites/baoyuan/import_maingrid.py` — ingests periodic kW/kWh/PF CSV from Baoyuan into `maingrid_history` table for post-period correlation analysis.

---

## Known Issues

1. **Cotton contamination** — Cotton blocks HyESys ventilation fans, causing overheating shutdowns. Clean with compressed air periodically.
2. **MSB1/MSB2 not metered** — HyESys effect is diluted at Main Grid level. Recommended fix: install meters at MSB1 and MSB2.
