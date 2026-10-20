#!/usr/bin/env python3
"""
Baoyuan FAT — 3-site H125 data topic sampler.
Listens for LISTEN_S seconds, prints latest timestamp + message types per cabinet.
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
LISTEN_S = 90

SITES = [
    {"label": "FY001", "cabinet_id": "20261014BY001",   "power": 125},
    {"label": "FY002", "cabinet_id": "20261014BY002",   "power": 125},
    {"label": "H125",  "cabinet_id": "26022703840003",  "power": 125},
]

SGT_OFFSET = 8 * 3600

def ts_sgt(ms):
    sgt = datetime.utcfromtimestamp(ms / 1000 + SGT_OFFSET)
    return sgt.strftime("%Y-%m-%d %H:%M:%S SGT")

# state[cabinet_id] = { type: {ts_ms, data} }
state = {s["cabinet_id"]: {} for s in SITES}
lock  = threading.Lock()

data_topics = {s["cabinet_id"]: f"stsc/aems/message/{s['cabinet_id']}" for s in SITES}


def on_connect(client, userdata, flags, rc):
    if rc == 0:
        for cab, topic in data_topics.items():
            client.subscribe(topic, qos=0)
        print(f"  Connected — subscribed to {len(data_topics)} data topics.")
    else:
        print(f"  Connect failed rc={rc}")


def on_message(client, userdata, msg):
    try:
        data  = json.loads(msg.payload.decode("utf-8", errors="replace"))
        mtype = data.get("type", "unknown")
        ts_ms = data.get("reportTimeTs")
        if not ts_ms:
            return
        # Match topic back to cabinet_id
        for cab, topic in data_topics.items():
            if msg.topic == topic:
                with lock:
                    state[cab][mtype] = {"ts_ms": int(ts_ms), "data": data.get("data", {})}
                break
    except Exception:
        pass


def main():
    client = mqtt.Client(mqtt.CallbackAPIVersion.VERSION1,
                         client_id=f"fat-check-{int(time.time())}")
    client.username_pw_set(USERNAME, PASSWORD)
    client.on_connect = on_connect
    client.on_message = on_message

    print(f"\nConnecting to {BROKER}:{PORT} ...")
    client.connect(BROKER, PORT, keepalive=60)
    client.loop_start()

    for remaining in range(LISTEN_S, 0, -10):
        time.sleep(10)
        with lock:
            total = sum(len(v) for v in state.values())
        print(f"  {remaining-10:3d}s remaining — {total} message type(s) seen across all sites",
              end="\r", flush=True)

    client.loop_stop()
    client.disconnect()

    print("\n")
    print("=" * 68)
    print(f"  Baoyuan FAT — H125 Data Topic Summary  ({datetime.utcfromtimestamp(time.time() + SGT_OFFSET).strftime('%H:%M:%S SGT')})")
    print("=" * 68)

    with lock:
        snap = {cab: dict(types) for cab, types in state.items()}

    for site in SITES:
        cab   = site["cabinet_id"]
        label = site["label"]
        types = snap.get(cab, {})
        print(f"\n  [{label}]  cabinet: {cab}  |  data topic: stsc/aems/message/{cab}")
        if not types:
            print("    - NO DATA received")
            continue
        # Sort by ts_ms descending to show latest first
        for mtype, info in sorted(types.items(), key=lambda x: -x[1]["ts_ms"]):
            print(f"    - type={mtype:<20}  latest: {ts_sgt(info['ts_ms'])}")

    print("\n" + "=" * 68)

    # Command topic summary
    print("\n  Command topics (for sending kW / kVar / restart):")
    for site in SITES:
        cab   = site["cabinet_id"]
        label = site["label"]
        # H125 26022703840003 uses multi2; others assumed multi
        suffix = "multi2" if cab == "26022703840003" else "multi"
        print(f"    [{label}]  stsc/aems/cabinet/{cab}/{suffix}/operate/tx")
    print()


if __name__ == "__main__":
    main()
