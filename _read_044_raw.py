#!/usr/bin/env python3
"""Print the raw 0_4_4 meter fields exactly as they arrive."""
import json, sys, time
from datetime import datetime

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

import paho.mqtt.client as mqtt

BROKER, PORT, USERNAME, PASSWORD = "loragw.advastech.com", 1883, "TYKJadmin", "TYKJ2018."
SGT_OFFSET = 8 * 3600
got = threading.Event() if False else None  # inline below

import threading
done = threading.Event()
result = {}

def on_connect(client, userdata, flags, rc):
    if rc == 0:
        client.subscribe("hyesys/data/dev/26022703840003", qos=0)
        print("  Connected, waiting for 0_4_4 batch...")

def on_message(client, userdata, msg):
    try:
        data = json.loads(msg.payload.decode("utf-8", errors="replace"))
        rep  = data.get("reported", {})
        if "0_4_4" in rep:
            result["raw"]      = rep["0_4_4"]
            result["sendtime"] = data.get("sendtime")
            result["all_keys_in_batch"] = list(rep.keys())
            done.set()
    except Exception:
        pass

client = mqtt.Client(mqtt.CallbackAPIVersion.VERSION1, client_id=f"raw044-{int(time.time())}")
client.username_pw_set(USERNAME, PASSWORD)
client.on_connect = on_connect
client.on_message = on_message
client.connect(BROKER, PORT, keepalive=60)
client.loop_start()
done.wait(timeout=120)
client.loop_stop(); client.disconnect()

if not result:
    print("No data received.")
    sys.exit(1)

sgt = datetime.utcfromtimestamp(float(result["sendtime"]) + SGT_OFFSET).strftime("%Y-%m-%d %H:%M:%S SGT")
print(f"\n  Timestamp : {sgt}")
print(f"\n  0_4_4 raw fields:")
print(f"  {'-'*40}")
for k, v in result["raw"].items():
    try:    fv = f"{float(v):.4f}"
    except: fv = str(v)
    print(f"  {k:<12} = {fv}")

print(f"\n  Other meters present in same batch: {result['all_keys_in_batch']}")

# Map to Excel columns
print(f"\n  Excel upstream meter column mapping:")
print(f"  {'Excel Col':<12} {'Col Letter':<10} {'Field':<12} {'Value'}")
print(f"  {'-'*50}")
raw = result["raw"]
def fv(k):
    try: return f"{float(raw.get(k,0)):.3f}"
    except: return "N/A"

mappings = [
    ("kW_A",   "Q (17)", "P",  fv("P")),
    ("kW_B",   "R (18)", "—",  "no per-phase kW"),
    ("kW_C",   "S (19)", "—",  "no per-phase kW"),
    ("kVAr_A", "T (20)", "Q",  fv("Q")),
    ("kVAr_B", "U (21)", "—",  "no per-phase kVAr"),
    ("kVAr_C", "V (22)", "—",  "no per-phase kVAr"),
]
for col, letter, field, val in mappings:
    print(f"  {col:<12} {letter:<10} {field:<12} {val}")
