#!/usr/bin/env python3
"""
Live side-by-side monitor: 0_4_4 feeder vs H125 pcs_v3.
Prints each message as it arrives for 120s.
"""

import json, sys, time
from datetime import datetime

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

import paho.mqtt.client as mqtt

BROKER     = "loragw.advastech.com"
PORT       = 1883
USERNAME   = "TYKJadmin"
PASSWORD   = "TYKJ2018."
LISTEN_S   = 120
SGT_OFFSET = 8 * 3600

CABINET_TOPIC = "hyesys/data/dev/26022703840003"
H125_TOPIC    = "stsc/aems/message/26022703840003"

def ts_sgt(unix_s):
    return datetime.utcfromtimestamp(unix_s + SGT_OFFSET).strftime("%H:%M:%S SGT")

def _f(v):
    try: return float(v)
    except: return 0.0

def on_connect(client, userdata, flags, rc):
    if rc == 0:
        client.subscribe(CABINET_TOPIC, qos=0)
        client.subscribe(H125_TOPIC,    qos=0)
        print("  Connected.\n")
        print(f"  {'Time':>12}  {'Source':<22}  {'P/kW':>8}  {'Q/kVAr':>8}  {'Ia':>7}  {'Ib':>7}  {'Ic':>7}")
        print("  " + "-" * 80)

def on_message(client, userdata, msg):
    try:
        data = json.loads(msg.payload.decode("utf-8", errors="replace"))
    except Exception:
        return

    now = ts_sgt(time.time())

    if msg.topic == CABINET_TOPIC:
        ts  = data.get("sendtime")
        rep = data.get("reported", {})

        # Print ALL feeders present in this batch
        for key, m in rep.items():
            p  = _f(m.get("P"))
            q  = _f(m.get("Q"))
            ia = _f(m.get("Ia"))
            ib = _f(m.get("Ib"))
            ic = _f(m.get("Ic"))
            marker = "  <== 0_4_4" if key == "0_4_4" else ""
            print(f"  {now:>12}  {key:<22}  {p:>8.2f}  {q:>8.2f}  {ia:>7.2f}  {ib:>7.2f}  {ic:>7.2f}{marker}")

    elif msg.topic == H125_TOPIC and data.get("type") == "pcs_v3":
        d    = data.get("data", {})
        ts_s = _f(data.get("reportTimeTs", 0)) / 1000
        kva  = _f(d.get("reactivePowerA"))
        kvb  = _f(d.get("reactivePowerB"))
        kvc  = _f(d.get("reactivePowerC"))
        kw   = _f(d.get("activePowerTotal"))
        kvt  = kva + kvb + kvc
        print(f"  {now:>12}  {'H125 pcs_v3':<22}  {kw:>8.2f}  {kvt:>8.2f}  "
              f"  [kVAr/ph: {kva:.2f}/{kvb:.2f}/{kvc:.2f}]")

def main():
    client = mqtt.Client(mqtt.CallbackAPIVersion.VERSION1,
                         client_id=f"check044-{int(time.time())}")
    client.username_pw_set(USERNAME, PASSWORD)
    client.on_connect = on_connect
    client.on_message = on_message

    print(f"\nConnecting to {BROKER}:{PORT} ...")
    client.connect(BROKER, PORT, keepalive=60)
    client.loop_start()
    time.sleep(LISTEN_S)
    client.loop_stop()
    client.disconnect()
    print("\nDone.")

if __name__ == "__main__":
    main()
