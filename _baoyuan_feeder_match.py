#!/usr/bin/env python3
"""
Reads all cabinet sub-meters from hyesys/data/dev/26022703840003
and pcs_v3 from all 3 H125 units simultaneously.
Prints a feeder table and identifies which feeder = sum of 3 H125 outputs.
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
SGT_OFFSET = 8 * 3600

# Cabinet sub-meter topic
CABINET_TOPIC = "hyesys/data/dev/26022703840003"

# All known sub-meter keys in order
FEEDER_KEYS = [
    "0_1_12204023110190",
    "0_2_1", "0_2_2", "0_2_3", "0_2_4",
    "0_3_1", "0_3_2", "0_3_3", "0_3_4",
    "0_4_1", "0_4_2", "0_4_3", "0_4_4",
]
FEEDER_LABELS = {
    "0_1_12204023110190": "CABB-INCOMER",
    "0_2_1": "CABB-F1", "0_2_2": "CABB-F2",
    "0_2_3": "CABB-F3", "0_2_4": "CABB-F4",
    "0_3_1": "CABC-F1", "0_3_2": "CABC-F2",
    "0_3_3": "CABC-F3", "0_3_4": "CABC-F4",
    "0_4_1": "CABA-F1", "0_4_2": "CABA-F2",
    "0_4_3": "CABA-F3", "0_4_4": "CABA-F4",
}

# 3 H125 data topics
H125_SITES = [
    {"label": "FY001", "cab": "20261014BY001"},
    {"label": "FY002", "cab": "20261014BY002"},
    {"label": "H125",  "cab": "26022703840003"},
]
H125_TOPICS = {s["cab"]: f"stsc/aems/message/{s['cab']}" for s in H125_SITES}

lock           = threading.Lock()
feeder_latest  = {}   # key -> {"P": kW, "Q": kVAr, "Ia":A, "Ib":A, "Ic":A, "ts": unix_s}
h125_pcs_latest = {}  # cab -> {"kW_a","kW_b","kW_c","kVAr_a","kVAr_b","kVAr_c","ts_ms"}


def on_connect(client, userdata, flags, rc):
    if rc == 0:
        client.subscribe(CABINET_TOPIC, qos=0)
        for topic in H125_TOPICS.values():
            client.subscribe(topic, qos=0)
        print(f"  Connected — subscribed to {1 + len(H125_TOPICS)} topics.")
    else:
        print(f"  Connect failed rc={rc}")


def on_message(client, userdata, msg):
    try:
        data = json.loads(msg.payload.decode("utf-8", errors="replace"))
    except Exception:
        return

    with lock:
        # Cabinet sub-meters
        if msg.topic == CABINET_TOPIC:
            ts = data.get("sendtime")
            reported = data.get("reported", {})
            for key in FEEDER_KEYS:
                if key in reported:
                    m = reported[key]
                    def _f(v):
                        try: return float(v)
                        except: return 0.0
                    feeder_latest[key] = {
                        "P":  _f(m.get("P")),
                        "Q":  _f(m.get("Q")),
                        "Ia": _f(m.get("Ia")),
                        "Ib": _f(m.get("Ib")),
                        "Ic": _f(m.get("Ic")),
                        "ts": float(ts) if ts else None,
                    }

        # H125 pcs_v3
        for cab, topic in H125_TOPICS.items():
            if msg.topic == topic and data.get("type") == "pcs_v3":
                d = data.get("data", {})
                def _f(v):
                    try: return float(v)
                    except: return 0.0
                h125_pcs_latest[cab] = {
                    "kW_a":   _f(d.get("activePowerA")),
                    "kW_b":   _f(d.get("activePowerB")),
                    "kW_c":   _f(d.get("activePowerC")),
                    "kVAr_a": _f(d.get("reactivePowerA")),
                    "kVAr_b": _f(d.get("reactivePowerB")),
                    "kVAr_c": _f(d.get("reactivePowerC")),
                    "ts_ms":  data.get("reportTimeTs"),
                }


def ts_sgt(unix_s):
    if not unix_s:
        return "--"
    sgt = datetime.utcfromtimestamp(unix_s + SGT_OFFSET)
    return sgt.strftime("%H:%M:%S SGT")


def main():
    client = mqtt.Client(mqtt.CallbackAPIVersion.VERSION1,
                         client_id=f"feeder-match-{int(time.time())}")
    client.username_pw_set(USERNAME, PASSWORD)
    client.on_connect = on_connect
    client.on_message = on_message

    print(f"\nConnecting to {BROKER}:{PORT} ...")
    client.connect(BROKER, PORT, keepalive=60)
    client.loop_start()

    for remaining in range(LISTEN_S, 0, -10):
        time.sleep(10)
        with lock:
            nf = len(feeder_latest)
            nh = len(h125_pcs_latest)
        print(f"  {remaining-10:3d}s left — {nf}/13 feeders | {nh}/3 H125 pcs_v3", end="\r", flush=True)

    client.loop_stop()
    client.disconnect()

    print("\n")

    with lock:
        fl  = {k: dict(v) for k, v in feeder_latest.items()}
        h125 = {k: dict(v) for k, v in h125_pcs_latest.items()}

    # ── H125 totals ────────────────────────────────────────────────────────
    print("=" * 72)
    print("  H125 Units — pcs_v3 Readings")
    print("=" * 72)
    print(f"  {'Unit':<8} {'kW-A':>7} {'kW-B':>7} {'kW-C':>7} {'kW-Total':>10}  "
          f"{'kVAr-A':>7} {'kVAr-B':>7} {'kVAr-C':>7} {'kVAr-Total':>11}  Timestamp")
    print("  " + "-" * 90)

    sum_kW   = 0.0
    sum_kVAr = 0.0

    for site in H125_SITES:
        cab   = site["cab"]
        label = site["label"]
        d = h125.get(cab)
        if d:
            kw_tot   = d["kW_a"]   + d["kW_b"]   + d["kW_c"]
            kvar_tot = d["kVAr_a"] + d["kVAr_b"] + d["kVAr_c"]
            sum_kW   += kw_tot
            sum_kVAr += kvar_tot
            ts_str = ts_sgt(d["ts_ms"] / 1000) if d["ts_ms"] else "--"
            print(f"  {label:<8} {d['kW_a']:>7.2f} {d['kW_b']:>7.2f} {d['kW_c']:>7.2f} {kw_tot:>10.2f}  "
                  f"{d['kVAr_a']:>7.2f} {d['kVAr_b']:>7.2f} {d['kVAr_c']:>7.2f} {kvar_tot:>11.2f}  {ts_str}")
        else:
            print(f"  {label:<8} -- no pcs_v3 data received --")

    print("  " + "-" * 90)
    print(f"  {'SUM':<8} {'':>7} {'':>7} {'':>7} {sum_kW:>10.2f}  "
          f"{'':>7} {'':>7} {'':>7} {sum_kVAr:>11.2f}")

    # ── Feeder table ────────────────────────────────────────────────────────
    print()
    print("=" * 72)
    print("  Cabinet Sub-meters — hyesys/data/dev/26022703840003")
    print("=" * 72)
    print(f"  {'Key':<22} {'Label':<15} {'P (kW)':>9} {'Q (kVAr)':>10} "
          f"{'Ia (A)':>8} {'Ib (A)':>8} {'Ic (A)':>8}  Timestamp  Match?")
    print("  " + "-" * 100)

    MATCH_TOL = 5.0   # kW tolerance for declaring a match

    for key in FEEDER_KEYS:
        label = FEEDER_LABELS.get(key, key)
        d = fl.get(key)
        if d:
            match = ""
            if h125_pcs_latest:
                if abs(d["P"] - sum_kW) < MATCH_TOL:
                    match = "  <-- MATCH (kW)"
                elif abs(abs(d["P"]) - abs(sum_kW)) < MATCH_TOL:
                    match = "  <-- MATCH (|kW|)"
            print(f"  {key:<22} {label:<15} {d['P']:>9.2f} {d['Q']:>10.2f} "
                  f"{d['Ia']:>8.2f} {d['Ib']:>8.2f} {d['Ic']:>8.2f}  "
                  f"{ts_sgt(d['ts'])}{match}")
        else:
            print(f"  {key:<22} {label:<15}  -- no data --")

    print()
    print(f"  3x H125 sum:  kW = {sum_kW:.2f}  |  kVAr = {sum_kVAr:.2f}  (tolerance +/-{MATCH_TOL} kW)")
    print("=" * 72)
    print()


if __name__ == "__main__":
    main()
