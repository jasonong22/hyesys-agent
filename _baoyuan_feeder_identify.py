#!/usr/bin/env python3
"""
Baoyuan — Feeder identification test.
  1. Baseline: collect feeder Q + H125 pcs_v3 for 30s
  2. Send 10 kVAr/phase to all 3 H125 units (kW = 0)
  3. Wait 90s for pcs_v3 to reflect
  4. Compare feeder Q delta → identify feeder
  5. Send all-zero to all 3 units; confirm reflected
"""

import json, sys, time, threading
from datetime import datetime

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

import paho.mqtt.client as mqtt

BROKER   = "loragw.advastech.com"
PORT     = 1883
USERNAME = "TYKJadmin"
PASSWORD = "TYKJ2018."
SGT_OFFSET = 8 * 3600

TEST_KVAR = 10.0   # kVAr per phase per unit  (10 × 3 phases × 3 units = 90 kVAr total expected delta)

SITES = [
    {"label": "FY001", "cab": "20261014BY001",
     "cmd_topic": "stsc/aems/cabinet/20261014BY001/multi/operate/tx"},
    {"label": "FY002", "cab": "20261014BY002",
     "cmd_topic": "stsc/aems/cabinet/20261014BY002/multi/operate/tx"},
    {"label": "H125",  "cab": "26022703840003",
     "cmd_topic": "stsc/aems/cabinet/26022703840003/multi2/operate/tx"},
]

CABINET_TOPIC = "hyesys/data/dev/26022703840003"
FEEDER_KEYS   = [
    "0_1_12204023110190",
    "0_2_1","0_2_2","0_2_3","0_2_4",
    "0_3_1","0_3_2","0_3_3","0_3_4",
    "0_4_1","0_4_2","0_4_3","0_4_4",
]
FEEDER_LABELS = {
    "0_1_12204023110190":"CABB-INCOMER",
    "0_2_1":"CABB-F1","0_2_2":"CABB-F2","0_2_3":"CABB-F3","0_2_4":"CABB-F4",
    "0_3_1":"CABC-F1","0_3_2":"CABC-F2","0_3_3":"CABC-F3","0_3_4":"CABC-F4",
    "0_4_1":"CABA-F1","0_4_2":"CABA-F2","0_4_3":"CABA-F3","0_4_4":"CABA-F4",
}

lock         = threading.Lock()
feeder_snap  = {}   # key -> {"P","Q","Ia","Ib","Ic"}
pcs_snap     = {}   # cab -> {"kVAr_a","kVAr_b","kVAr_c", "ts_ms"}

def _f(v):
    try: return float(v)
    except: return 0.0

def ts_sgt(unix_s):
    dt = datetime.utcfromtimestamp(unix_s + SGT_OFFSET)
    return dt.strftime("%H:%M:%S SGT")

def on_connect(client, userdata, flags, rc):
    if rc == 0:
        client.subscribe(CABINET_TOPIC, qos=0)
        for s in SITES:
            client.subscribe(f"stsc/aems/message/{s['cab']}", qos=0)
        print(f"  Connected — subscribed to {1 + len(SITES)} topics.")
    else:
        print(f"  Connect failed rc={rc}")

def on_message(client, userdata, msg):
    try:
        data = json.loads(msg.payload.decode("utf-8", errors="replace"))
    except Exception:
        return
    with lock:
        if msg.topic == CABINET_TOPIC:
            ts  = data.get("sendtime")
            rep = data.get("reported", {})
            for key in FEEDER_KEYS:
                if key in rep:
                    m = rep[key]
                    feeder_snap[key] = {
                        "P":  _f(m.get("P")), "Q": _f(m.get("Q")),
                        "Ia": _f(m.get("Ia")),"Ib":_f(m.get("Ib")),"Ic":_f(m.get("Ic")),
                        "ts": float(ts) if ts else None,
                    }
        else:
            for s in SITES:
                if msg.topic == f"stsc/aems/message/{s['cab']}" and data.get("type") == "pcs_v3":
                    d = data.get("data", {})
                    pcs_snap[s["cab"]] = {
                        "kVAr_a": _f(d.get("reactivePowerA")),
                        "kVAr_b": _f(d.get("reactivePowerB")),
                        "kVAr_c": _f(d.get("reactivePowerC")),
                        "ts_ms":  data.get("reportTimeTs"),
                    }

def snapshot_feeders():
    with lock:
        return {k: dict(v) for k, v in feeder_snap.items()}

def snapshot_pcs():
    with lock:
        return {k: dict(v) for k, v in pcs_snap.items()}

def send_kvar(client, site, kvar):
    payload = json.dumps({
        "cabinetId": site["cab"], "index": 1, "remote": True,
        "key": "set_reactive_power",
        "params": {"reactivePowerA": kvar, "reactivePowerB": kvar, "reactivePowerC": kvar},
    })
    info = client.publish(site["cmd_topic"], payload, qos=1)
    info.wait_for_publish(timeout=10)
    return info.is_published()

def send_kw(client, site, kw):
    payload = json.dumps({
        "cabinetId": site["cab"], "index": 1, "remote": True,
        "key": "set_active_power",
        "params": {"activePowerA": kw, "activePowerB": kw, "activePowerC": kw},
    })
    info = client.publish(site["cmd_topic"], payload, qos=1)
    info.wait_for_publish(timeout=10)
    return info.is_published()

def wait_collect(seconds, label):
    for remaining in range(seconds, 0, -5):
        time.sleep(5)
        with lock:
            nf = len(feeder_snap)
            nh = len(pcs_snap)
        print(f"  [{label}] {remaining-5:3d}s left — feeders:{nf}/13  pcs:{nh}/3",
              end="\r", flush=True)
    print()

def print_feeder_table(snap, delta=None):
    hdr = f"  {'Key':<22} {'Label':<15} {'P (kW)':>9} {'Q (kVAr)':>10}"
    if delta:
        hdr += f" {'dQ':>8}"
    hdr += f"  {'Ia':>6} {'Ib':>6} {'Ic':>6}"
    print(hdr)
    print("  " + "-" * (len(hdr) - 2))
    for key in FEEDER_KEYS:
        label = FEEDER_LABELS.get(key, key)
        d = snap.get(key)
        if d:
            row = (f"  {key:<22} {label:<15} {d['P']:>9.2f} {d['Q']:>10.2f}")
            if delta is not None:
                dq = delta.get(key, 0.0)
                marker = "  <-- MATCH" if abs(dq) > 10 else ""
                row += f" {dq:>+8.2f}{marker}"
            row += f"  {d['Ia']:>6.1f} {d['Ib']:>6.1f} {d['Ic']:>6.1f}"
            print(row)
        else:
            print(f"  {key:<22} {label:<15}  -- no data --")

def main():
    client = mqtt.Client(mqtt.CallbackAPIVersion.VERSION1,
                         client_id=f"feeder-id-{int(time.time())}")
    client.username_pw_set(USERNAME, PASSWORD)
    client.on_connect = on_connect
    client.on_message = on_message

    print(f"\nConnecting to {BROKER}:{PORT} ...")
    client.connect(BROKER, PORT, keepalive=60)
    client.loop_start()

    # ── STEP 1: Baseline ──────────────────────────────────────────────────
    print("\n[Step 1] Collecting baseline (30s) ...")
    wait_collect(30, "baseline")
    baseline = snapshot_feeders()
    baseline_pcs = snapshot_pcs()

    print(f"\n  Baseline feeder Q:")
    for key in FEEDER_KEYS:
        d = baseline.get(key)
        q = f"{d['Q']:.2f} kVAr" if d else "no data"
        print(f"    {FEEDER_LABELS.get(key,key):<15}: {q}")

    print(f"\n  Baseline H125 pcs_v3:")
    for s in SITES:
        d = baseline_pcs.get(s["cab"])
        if d:
            tot = d["kVAr_a"] + d["kVAr_b"] + d["kVAr_c"]
            print(f"    {s['label']}: {d['kVAr_a']:.2f} / {d['kVAr_b']:.2f} / {d['kVAr_c']:.2f} kVAr/phase  (total {tot:.2f})")
        else:
            print(f"    {s['label']}: -- no data --")

    # ── STEP 2: Send test kVAr ───────────────────────────────────────────
    print(f"\n[Step 2] Sending {TEST_KVAR} kVAr/phase to all 3 units ...")
    for s in SITES:
        ok_kw   = send_kw(client, s, 0.0)
        ok_kvar = send_kvar(client, s, TEST_KVAR)
        status  = "OK" if (ok_kw and ok_kvar) else "PUBACK failed"
        print(f"  {s['label']} ({s['cab']})  cmd_topic: .../{s['cmd_topic'].split('/')[-3]}/...  -> {status}")

    # ── STEP 3: Wait for pcs_v3 to reflect ───────────────────────────────
    print(f"\n[Step 3] Waiting 90s for units to reflect command ...")
    wait_collect(90, "waiting")

    after_pcs = snapshot_pcs()
    print(f"\n  Post-command H125 pcs_v3:")
    responded = []
    for s in SITES:
        d = after_pcs.get(s["cab"])
        if d:
            tot = d["kVAr_a"] + d["kVAr_b"] + d["kVAr_c"]
            expected = TEST_KVAR * 3
            delta_pcs = tot - (baseline_pcs.get(s["cab"], {}).get("kVAr_a", 0) +
                               baseline_pcs.get(s["cab"], {}).get("kVAr_b", 0) +
                               baseline_pcs.get(s["cab"], {}).get("kVAr_c", 0))
            flag = "  RESPONDED" if abs(delta_pcs) > 5 else "  no change"
            if abs(delta_pcs) > 5:
                responded.append(s["label"])
            print(f"    {s['label']}: {d['kVAr_a']:.2f}/{d['kVAr_b']:.2f}/{d['kVAr_c']:.2f} kVAr  "
                  f"total={tot:.2f}  dQ={delta_pcs:+.2f}{flag}")
        else:
            print(f"    {s['label']}: -- no data --")

    # ── STEP 4: Feeder delta ─────────────────────────────────────────────
    print(f"\n[Step 4] Feeder Q delta (after - baseline):")
    after_feeders = snapshot_feeders()
    delta = {}
    for key in FEEDER_KEYS:
        b = baseline.get(key, {}).get("Q", 0.0)
        a = after_feeders.get(key, {}).get("Q", 0.0)
        delta[key] = a - b
    print_feeder_table(after_feeders, delta)

    # Find best match
    best_key = max(FEEDER_KEYS, key=lambda k: delta.get(k, 0.0))
    best_dq  = delta.get(best_key, 0.0)
    print()
    if abs(best_dq) > 10:
        print(f"  ** IDENTIFIED: {FEEDER_LABELS[best_key]} ({best_key})  dQ = {best_dq:+.2f} kVAr **")
    else:
        print(f"  ** No clear match (max dQ = {best_dq:+.2f} kVAr) — units may not have responded **")
        print(f"     Units confirmed responded: {responded or 'none'}")

    # ── STEP 5: Send all-zero ────────────────────────────────────────────
    print(f"\n[Step 5] Sending all-zero to all 3 units ...")
    for s in SITES:
        ok_kw   = send_kw(client, s, 0.0)
        ok_kvar = send_kvar(client, s, 0.0)
        status  = "OK" if (ok_kw and ok_kvar) else "PUBACK failed"
        print(f"  {s['label']} -> {status}")

    print(f"\n[Step 6] Waiting 90s to confirm zeros ...")
    wait_collect(90, "confirming zeros")

    final_pcs = snapshot_pcs()
    print(f"\n  Final pcs_v3 confirmation:")
    for s in SITES:
        d = final_pcs.get(s["cab"])
        if d:
            tot_kvar = d["kVAr_a"] + d["kVAr_b"] + d["kVAr_c"]
            ok = "ZEROED" if abs(tot_kvar) < 3.0 else "still non-zero"
            print(f"    {s['label']}: kVAr = {d['kVAr_a']:.2f}/{d['kVAr_b']:.2f}/{d['kVAr_c']:.2f}  total={tot_kvar:.2f}  [{ok}]")
        else:
            print(f"    {s['label']}: -- no data --")

    client.loop_stop()
    client.disconnect()
    print("\nDone.")


if __name__ == "__main__":
    main()
