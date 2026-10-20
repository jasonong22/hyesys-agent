#!/usr/bin/env python3
"""
One-shot Baoyuan MQTT topic sampler.
Listens to all 4 topics for LISTEN_S seconds, then prints latest timestamp per meter.
"""

import json, time, threading, sys
from datetime import datetime, timezone

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

import paho.mqtt.client as mqtt

BROKER   = "loragw.advastech.com"
PORT     = 1883
USERNAME = "TYKJadmin"
PASSWORD = "TYKJ2018."
LISTEN_S = 120    # seconds to collect; cabinet meters arrive ~every 90 s

TOPICS = [
    "hyesys/data/dev/0086040215999997",   # CapBank 1
    "hyesys/data/dev/0086040215999996",   # CapBank 2
    "hyesys/data/dev/26022703840003",     # 13 cabinet meters (batched)
    "stsc/aems/message/26022703840003",   # H125 PCS + BMS
]

# latest = { display_label: unix_float_seconds }
latest = {}
lock   = threading.Lock()

CABINET_METER_IDS = [
    "0_1_12204023110190",   # BAOYUAN-CABB-INCOMER
    "0_2_1",                # BAOYUAN-CABB-F1
    "0_2_2",                # BAOYUAN-CABB-F2
    "0_2_3",                # BAOYUAN-CABB-F3
    "0_2_4",                # BAOYUAN-CABB-F4
    "0_3_1",                # BAOYUAN-CABC-F1
    "0_3_2",                # BAOYUAN-CABC-F2
    "0_3_3",                # BAOYUAN-CABC-F3
    "0_3_4",                # BAOYUAN-CABC-F4
    "0_4_1",                # BAOYUAN-CABA-F1
    "0_4_2",                # BAOYUAN-CABA-F2
    "0_4_3",                # BAOYUAN-CABA-F3
    "0_4_4",                # BAOYUAN-CABA-F4 (H125 feeder meter)
]

SITE_LABELS = {
    "hyesys/data/dev/0086040215999997": "BAOYUAN-CAPBANK1",
    "hyesys/data/dev/0086040215999996": "BAOYUAN-CAPBANK2",
    "0_1_12204023110190": "BAOYUAN-CABB-INCOMER",
    "0_2_1": "BAOYUAN-CABB-F1",
    "0_2_2": "BAOYUAN-CABB-F2",
    "0_2_3": "BAOYUAN-CABB-F3",
    "0_2_4": "BAOYUAN-CABB-F4",
    "0_3_1": "BAOYUAN-CABC-F1",
    "0_3_2": "BAOYUAN-CABC-F2",
    "0_3_3": "BAOYUAN-CABC-F3",
    "0_3_4": "BAOYUAN-CABC-F4",
    "0_4_1": "BAOYUAN-CABA-F1",
    "0_4_2": "BAOYUAN-CABA-F2",
    "0_4_3": "BAOYUAN-CABA-F3",
    "0_4_4": "BAOYUAN-CABA-F4 [H125 feeder]",
    "stsc/aems/message/26022703840003:pcs_v3": "BAOYUAN-HYESYS (pcs_v3)",
    "stsc/aems/message/26022703840003:bms":    "BAOYUAN-HYESYS (bms)",
}


def on_connect(client, userdata, flags, rc):
    code = rc if isinstance(rc, int) else rc.value
    if code == 0:
        print(f"  Connected. Subscribing to {len(TOPICS)} topics ...")
        for t in TOPICS:
            client.subscribe(t, qos=0)
    else:
        print(f"  Connect failed rc={code}")


def on_message(client, userdata, msg):
    topic = msg.topic
    try:
        data = json.loads(msg.payload.decode("utf-8", errors="replace"))
    except Exception:
        return

    with lock:
        # CapBank meters
        if topic in ("hyesys/data/dev/0086040215999997",
                     "hyesys/data/dev/0086040215999996"):
            ts = data.get("sendtime")
            if ts:
                latest[topic] = float(ts)

        # Cabinet batch meter
        elif topic == "hyesys/data/dev/26022703840003":
            ts = data.get("sendtime")
            if ts:
                ts_f = float(ts)
                reported = data.get("reported", {})
                # Update each sub-meter that's present in this batch
                for mid in CABINET_METER_IDS:
                    if mid in reported:
                        latest[mid] = ts_f
                # If no reported keys matched, store at topic level anyway
                if not any(mid in reported for mid in CABINET_METER_IDS):
                    latest["hyesys/data/dev/26022703840003:raw"] = ts_f

        # H125 PCS + BMS
        elif topic == "stsc/aems/message/26022703840003":
            mtype = data.get("type")
            ts_ms = data.get("reportTimeTs")
            if ts_ms:
                key = f"stsc/aems/message/26022703840003:{mtype}"
                latest[key] = float(ts_ms) / 1000.0


SGT_OFFSET = 8 * 3600

def ts_to_sgt(unix_s):
    if unix_s is None:
        return "--"
    sgt = datetime.utcfromtimestamp(unix_s + SGT_OFFSET)
    return sgt.strftime("%Y-%m-%d %H:%M:%S SGT")


def main():
    client = mqtt.Client(mqtt.CallbackAPIVersion.VERSION1,
                         client_id=f"baoyuan-check-{int(time.time())}")
    client.username_pw_set(USERNAME, PASSWORD)
    client.on_connect = on_connect
    client.on_message = on_message

    print(f"\nConnecting to {BROKER}:{PORT} ...")
    client.connect(BROKER, PORT, keepalive=60)
    client.loop_start()

    print(f"Listening for {LISTEN_S} seconds ...\n")
    for remaining in range(LISTEN_S, 0, -10):
        time.sleep(10)
        with lock:
            n = len(latest)
        print(f"  {remaining - 10}s remaining — {n} meter(s) seen so far", end="\r", flush=True)

    client.loop_stop()
    client.disconnect()

    # ── Print results ─────────────────────────────────────────────────────────
    print("\n")
    print("=" * 72)
    print(f"  Baoyuan — Latest MQTT Timestamps  (sampled {LISTEN_S}s at {datetime.now().strftime('%H:%M:%S')} SGT)")
    print("=" * 72)

    # Define display order
    order = [
        ("BAOYUAN-CAPBANK1",             "hyesys/data/dev/0086040215999997"),
        ("BAOYUAN-CAPBANK2",             "hyesys/data/dev/0086040215999996"),
        ("BAOYUAN-CABB-INCOMER",         "0_1_12204023110190"),
        ("BAOYUAN-CABB-F1",              "0_2_1"),
        ("BAOYUAN-CABB-F2",              "0_2_2"),
        ("BAOYUAN-CABB-F3",              "0_2_3"),
        ("BAOYUAN-CABB-F4",              "0_2_4"),
        ("BAOYUAN-CABC-F1",              "0_3_1"),
        ("BAOYUAN-CABC-F2",              "0_3_2"),
        ("BAOYUAN-CABC-F3",              "0_3_3"),
        ("BAOYUAN-CABC-F4",              "0_3_4"),
        ("BAOYUAN-CABA-F1",              "0_4_1"),
        ("BAOYUAN-CABA-F2",              "0_4_2"),
        ("BAOYUAN-CABA-F3",              "0_4_3"),
        ("BAOYUAN-CABA-F4 [H125 feeder]","0_4_4"),
        ("BAOYUAN-HYESYS pcs_v3",        "stsc/aems/message/26022703840003:pcs_v3"),
        ("BAOYUAN-HYESYS bms",           "stsc/aems/message/26022703840003:bms"),
    ]

    print(f"\n  {'Meter':<30}  {'Latest Timestamp (UTC+8 SGT)':>29}  Status")
    print("  " + "-" * 68)

    with lock:
        snap = dict(latest)

    now_s = time.time()
    for label, key in order:
        ts_s = snap.get(key)
        if ts_s:
            age = now_s - ts_s
            ts_str = ts_to_sgt(ts_s)
            if age < 300:
                status = "[LIVE]"
            elif age < 3600:
                status = "[STALE]"
            else:
                status = "[OLD]"
        else:
            ts_str = "-- no data received --"
            status = "[SILENT]"
        print(f"  {label:<30}  {ts_str:>29}  {status}")

    print()
    if snap:
        oldest = min(snap.values())
        newest = max(snap.values())
        print(f"  Oldest data : {ts_to_sgt(oldest)}")
        print(f"  Newest data : {ts_to_sgt(newest)}")
    print("=" * 72)


if __name__ == "__main__":
    main()
