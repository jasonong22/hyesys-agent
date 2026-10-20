#!/usr/bin/env python3
"""E-STOP: send all-zero to all 3 H125 units and confirm pcs_v3 reflects."""
import json, sys, time, threading
from datetime import datetime

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

import paho.mqtt.client as mqtt

BROKER, PORT, USERNAME, PASSWORD = "loragw.advastech.com", 1883, "TYKJadmin", "TYKJ2018."
SGT_OFFSET = 8 * 3600

SITES = [
    {"label": "FY001", "cab": "20261014BY001", "cmd": "stsc/aems/cabinet/20261014BY001/multi/operate/tx"},
    {"label": "FY002", "cab": "20261014BY002", "cmd": "stsc/aems/cabinet/20261014BY002/multi/operate/tx"},
    {"label": "H125",  "cab": "26022703840003", "cmd": "stsc/aems/cabinet/26022703840003/multi2/operate/tx"},
]

lock     = threading.Lock()
pcs_snap = {}  # cab -> {kVAr_a/b/c, kW_a/b/c, ts_ms}
connected = threading.Event()

def _f(v):
    try: return float(v)
    except: return 0.0

def ts_sgt(ms):
    return datetime.utcfromtimestamp(ms/1000 + SGT_OFFSET).strftime("%H:%M:%S SGT")

def on_connect(client, userdata, flags, rc):
    if rc == 0:
        for s in SITES:
            client.subscribe(f"stsc/aems/message/{s['cab']}", qos=0)
        connected.set()

def on_message(client, userdata, msg):
    try:
        data = json.loads(msg.payload.decode("utf-8", errors="replace"))
    except: return
    if data.get("type") != "pcs_v3": return
    with lock:
        for s in SITES:
            if msg.topic == f"stsc/aems/message/{s['cab']}":
                d = data.get("data", {})
                pcs_snap[s["cab"]] = {
                    "kW_a":   _f(d.get("activePowerA")),
                    "kW_b":   _f(d.get("activePowerB")),
                    "kW_c":   _f(d.get("activePowerC")),
                    "kVAr_a": _f(d.get("reactivePowerA")),
                    "kVAr_b": _f(d.get("reactivePowerB")),
                    "kVAr_c": _f(d.get("reactivePowerC")),
                    "ts_ms":  _f(data.get("reportTimeTs")),
                }

client = mqtt.Client(mqtt.CallbackAPIVersion.VERSION1, client_id=f"estop-{int(time.time())}")
client.username_pw_set(USERNAME, PASSWORD)
client.on_connect = on_connect
client.on_message = on_message
client.connect(BROKER, PORT, keepalive=60)
client.loop_start()
connected.wait(timeout=10)

# ── Send all-zero immediately ─────────────────────────────────────────────────
print("\n*** E-STOP — sending all-zero to all 3 units ***\n")
for s in SITES:
    for key in ("set_active_power", "set_reactive_power"):
        field = "activePower" if "active" in key else "reactivePower"
        payload = json.dumps({
            "cabinetId": s["cab"], "index": 1, "remote": True, "key": key,
            "params": {f"{field}A": 0.0, f"{field}B": 0.0, f"{field}C": 0.0},
        })
        info = client.publish(s["cmd"], payload, qos=1)
        info.wait_for_publish(timeout=10)
    print(f"  {s['label']} ({s['cab']}) -> zero sent")

send_ts = time.time()
print("\nWaiting for pcs_v3 to confirm zeros (up to 90s)...")

deadline = time.time() + 90
confirmed = set()
while time.time() < deadline:
    time.sleep(2)
    with lock:
        snap = {k: dict(v) for k, v in pcs_snap.items()}
    for s in SITES:
        if s["cab"] in confirmed: continue
        d = snap.get(s["cab"])
        if d and d["ts_ms"]/1000 > send_ts:
            confirmed.add(s["cab"])
    if len(confirmed) == 3:
        break
    pending = [s["label"] for s in SITES if s["cab"] not in confirmed]
    print(f"  waiting for: {pending}", end="\r", flush=True)

print()
print("\n  Final pcs_v3 readings:")
print(f"  {'Unit':<8} {'kW-A':>7} {'kW-B':>7} {'kW-C':>7}  {'kVAr-A':>8} {'kVAr-B':>8} {'kVAr-C':>8}  Status    Timestamp")
print("  " + "-" * 80)
with lock:
    snap = {k: dict(v) for k, v in pcs_snap.items()}
for s in SITES:
    d = snap.get(s["cab"])
    if d:
        tot_kvar = d["kVAr_a"] + d["kVAr_b"] + d["kVAr_c"]
        tot_kw   = d["kW_a"]   + d["kW_b"]   + d["kW_c"]
        zeroed   = "ZEROED" if abs(tot_kvar) < 5.0 and abs(tot_kw) < 3.0 else "CHECK"
        fresh    = "fresh" if d["ts_ms"]/1000 > send_ts else "stale"
        ts       = ts_sgt(d["ts_ms"])
        print(f"  {s['label']:<8} {d['kW_a']:>7.2f} {d['kW_b']:>7.2f} {d['kW_c']:>7.2f}  "
              f"{d['kVAr_a']:>8.2f} {d['kVAr_b']:>8.2f} {d['kVAr_c']:>8.2f}  {zeroed}/{fresh}  {ts}")
    else:
        print(f"  {s['label']:<8}  -- no data --")

client.loop_stop()
client.disconnect()
print("\nDone.")
