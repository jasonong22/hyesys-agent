#!/usr/bin/env python3
"""Quick BMS SOC reader — prints SOC for all 3 H125 units."""
import json, sys, time, threading
from datetime import datetime

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

import paho.mqtt.client as mqtt

BROKER, PORT, USERNAME, PASSWORD = "loragw.advastech.com", 1883, "TYKJadmin", "TYKJ2018."
SGT_OFFSET = 8 * 3600
SITES = [
    {"label": "FY001", "cab": "20261014BY001"},
    {"label": "FY002", "cab": "20261014BY002"},
    {"label": "H125",  "cab": "26022703840003"},
]
lock = threading.Lock()
bms  = {}  # cab -> data

def _f(v):
    try: return float(v)
    except: return None

def on_connect(client, userdata, flags, rc):
    if rc == 0:
        for s in SITES:
            client.subscribe(f"stsc/aems/message/{s['cab']}", qos=0)
        print("  Connected.")

def on_message(client, userdata, msg):
    try:
        data = json.loads(msg.payload.decode("utf-8", errors="replace"))
    except: return
    if data.get("type") != "bms": return
    with lock:
        for s in SITES:
            if msg.topic == f"stsc/aems/message/{s['cab']}":
                d = data.get("data", {})
                bms[s["cab"]] = {
                    "soc": _f(d.get("soc")), "soh": _f(d.get("soh")),
                    "volt": _f(d.get("voltage")), "cur": _f(d.get("current")),
                    "cell_avg": _f(d.get("singleVoltageAvg")),
                    "cell_max": _f(d.get("singleVoltageMax")),
                    "cell_min": _f(d.get("singleVoltageMin")),
                    "temp": _f(d.get("tempMain")),
                    "ts_ms": data.get("reportTimeTs"),
                }

client = mqtt.Client(mqtt.CallbackAPIVersion.VERSION1, client_id=f"soc-{int(time.time())}")
client.username_pw_set(USERNAME, PASSWORD)
client.on_connect = on_connect
client.on_message = on_message
client.connect(BROKER, PORT, keepalive=60)
client.loop_start()

print("Waiting for BMS data (up to 90s)...")
for r in range(90, 0, -5):
    time.sleep(5)
    with lock: n = len(bms)
    print(f"  {r-5}s left — {n}/3 BMS received", end="\r", flush=True)
    if n == 3: break

client.loop_stop(); client.disconnect()
print("\n")
print(f"  {'Unit':<8} {'SOC%':>6} {'SOH%':>6} {'CellAvg':>9} {'CellMax':>9} {'CellMin':>9} {'Temp':>6}  Timestamp")
print("  " + "-" * 72)
for s in SITES:
    d = bms.get(s["cab"])
    if d:
        ts = datetime.utcfromtimestamp(d["ts_ms"]/1000 + SGT_OFFSET).strftime("%H:%M:%S SGT") if d["ts_ms"] else "--"
        flag = "  <-- ~37%" if d["soc"] and abs(d["soc"] - 37) < 8 else ""
        print(f"  {s['label']:<8} {d['soc']:>6.1f} {d['soh']:>6.1f} {d['cell_avg']:>9.3f} "
              f"{d['cell_max']:>9.3f} {d['cell_min']:>9.3f} {d['temp']:>6.1f}  {ts}{flag}")
    else:
        print(f"  {s['label']:<8}  -- no BMS data --")
