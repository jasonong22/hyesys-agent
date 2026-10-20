#!/usr/bin/env python3
"""
Sign convention test — +kVAr then -kVAr.
After each command reads pcs_v3 from all 3 H125 units and Q from 0_4_4.
"""

import json, sys, time, threading
from datetime import datetime

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

import paho.mqtt.client as mqtt

BROKER     = "loragw.advastech.com"
PORT       = 1883
USERNAME   = "TYKJadmin"
PASSWORD   = "TYKJ2018."
SGT_OFFSET = 8 * 3600

SITES = [
    {"label": "FY001", "cab": "20261014BY001",
     "cmd_topic": "stsc/aems/cabinet/20261014BY001/multi/operate/tx"},
    {"label": "FY002", "cab": "20261014BY002",
     "cmd_topic": "stsc/aems/cabinet/20261014BY002/multi/operate/tx"},
    {"label": "H125",  "cab": "26022703840003",
     "cmd_topic": "stsc/aems/cabinet/26022703840003/multi2/operate/tx"},
]

CABINET_TOPIC = "hyesys/data/dev/26022703840003"
TARGET_FEEDER = "0_4_4"

lock     = threading.Lock()
pcs_data = {}   # cab -> latest pcs_v3 dict
f044     = {}   # latest 0_4_4 feeder reading

def _f(v):
    try: return float(v)
    except: return 0.0

def ts_sgt(unix_s):
    return datetime.utcfromtimestamp(unix_s + SGT_OFFSET).strftime("%H:%M:%S SGT")

def on_connect(client, userdata, flags, rc):
    if rc == 0:
        client.subscribe(CABINET_TOPIC, qos=0)
        for s in SITES:
            client.subscribe(f"stsc/aems/message/{s['cab']}", qos=0)
        print("  Connected.")

def on_message(client, userdata, msg):
    try:
        data = json.loads(msg.payload.decode("utf-8", errors="replace"))
    except Exception:
        return
    with lock:
        if msg.topic == CABINET_TOPIC:
            rep = data.get("reported", {})
            if TARGET_FEEDER in rep:
                m = rep[TARGET_FEEDER]
                f044.update({
                    "P":  _f(m.get("P")),  "Q": _f(m.get("Q")),
                    "Ia": _f(m.get("Ia")), "Ib": _f(m.get("Ib")), "Ic": _f(m.get("Ic")),
                    "ts": _f(data.get("sendtime")),
                })
        else:
            for s in SITES:
                if msg.topic == f"stsc/aems/message/{s['cab']}" and data.get("type") == "pcs_v3":
                    d = data.get("data", {})
                    pcs_data[s["cab"]] = {
                        "kVAr_a": _f(d.get("reactivePowerA")),
                        "kVAr_b": _f(d.get("reactivePowerB")),
                        "kVAr_c": _f(d.get("reactivePowerC")),
                        "kW_tot": _f(d.get("activePowerTotal")),
                        "ts_ms":  _f(data.get("reportTimeTs")),
                    }

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

def wait_fresh(label, seconds):
    for r in range(seconds, 0, -5):
        time.sleep(5)
        with lock:
            np = len(pcs_data)
            nf = 1 if f044 else 0
        print(f"  [{label}] {r-5:3d}s left — pcs:{np}/3  0_4_4:{'yes' if nf else 'no '}",
              end="\r", flush=True)
    print()

def print_results(label, kvar_cmd):
    print(f"\n  {'='*62}")
    print(f"  Command: {kvar_cmd:+.1f} kVAr/phase  [{label}]")
    print(f"  {'='*62}")

    with lock:
        pcs_snap = {k: dict(v) for k, v in pcs_data.items()}
        f_snap   = dict(f044)

    # PCS table
    print(f"\n  PCS pcs_v3 readings:")
    print(f"  {'Unit':<8} {'kVAr-A':>9} {'kVAr-B':>9} {'kVAr-C':>9} {'kVAr-Total':>12}  Timestamp")
    print(f"  {'-'*58}")
    pcs_total = 0.0
    for s in SITES:
        d = pcs_snap.get(s["cab"])
        if d:
            tot = d["kVAr_a"] + d["kVAr_b"] + d["kVAr_c"]
            pcs_total += tot
            ts = ts_sgt(d["ts_ms"] / 1000) if d["ts_ms"] else "--"
            print(f"  {s['label']:<8} {d['kVAr_a']:>+9.2f} {d['kVAr_b']:>+9.2f} "
                  f"{d['kVAr_c']:>+9.2f} {tot:>+12.2f}  {ts}")
        else:
            print(f"  {s['label']:<8}  -- no data --")
    print(f"  {'-'*58}")
    print(f"  {'SUM':<8} {'':>9} {'':>9} {'':>9} {pcs_total:>+12.2f}")

    # 0_4_4 table
    print(f"\n  0_4_4 (CABA-F4) feeder meter:")
    print(f"  {'Field':<12} {'Value':>12}")
    print(f"  {'-'*26}")
    if f_snap:
        print(f"  {'P (kW)':<12} {f_snap['P']:>+12.2f}")
        print(f"  {'Q (kVAr)':<12} {f_snap['Q']:>+12.2f}")
        print(f"  {'Ia (A)':<12} {f_snap['Ia']:>12.2f}")
        print(f"  {'Ib (A)':<12} {f_snap['Ib']:>12.2f}")
        print(f"  {'Ic (A)':<12} {f_snap['Ic']:>12.2f}")
        ts_f = ts_sgt(f_snap["ts"]) if f_snap.get("ts") else "--"
        print(f"  {'Timestamp':<12} {ts_f:>12}")
    else:
        print(f"  -- no data --")

    # Tally check
    if f_snap and pcs_snap:
        q044 = f_snap["Q"]
        sign_pcs  = "+" if pcs_total >= 0 else "-"
        sign_044  = "+" if q044 >= 0 else "-"
        ratio = q044 / pcs_total if abs(pcs_total) > 1 else 0
        print(f"\n  Tally check:")
        print(f"    PCS total kVAr : {pcs_total:>+8.2f}  (sign: {sign_pcs})")
        print(f"    0_4_4 Q        : {q044:>+8.2f}  (sign: {sign_044})")
        print(f"    Ratio 044/PCS  : {ratio:>8.3f}  (expect ~-1.0 if opposite sign)")
        if abs(abs(ratio) - 1.0) < 0.3:
            same = "SAME" if ratio > 0 else "OPPOSITE"
            print(f"    --> Signs are {same}, magnitudes tally OK")
        else:
            print(f"    --> Magnitudes do NOT tally (check meter polarity/connection)")


def main():
    client = mqtt.Client(mqtt.CallbackAPIVersion.VERSION1,
                         client_id=f"sign-test-{int(time.time())}")
    client.username_pw_set(USERNAME, PASSWORD)
    client.on_connect = on_connect
    client.on_message = on_message

    print(f"\nConnecting to {BROKER}:{PORT} ...")
    client.connect(BROKER, PORT, keepalive=60)
    client.loop_start()

    print("\n[Warmup — 20s to collect initial data]")
    wait_fresh("warmup", 20)

    for kvar_cmd in [+10.0, -10.0]:
        label = "POSITIVE" if kvar_cmd > 0 else "NEGATIVE"
        print(f"\n{'='*62}")
        print(f"  Sending kW=0, kVAr={kvar_cmd:+.1f}/phase to all 3 units ...")
        for s in SITES:
            ok_kw   = send_kw(client, s, 0.0)
            ok_kvar = send_kvar(client, s, kvar_cmd)
            status  = "OK" if (ok_kw and ok_kvar) else "PUBACK fail"
            print(f"    {s['label']} -> {status}")

        print(f"\n  Waiting 90s for pcs_v3 + 0_4_4 to reflect ...")
        wait_fresh(label, 90)
        print_results(label, kvar_cmd)

    # Zero all
    print(f"\n\n{'='*62}")
    print("  Sending all-zero to all 3 units ...")
    for s in SITES:
        send_kw(client, s, 0.0)
        send_kvar(client, s, 0.0)
        print(f"    {s['label']} -> zeroed")

    print("\n  Waiting 90s to confirm zeros ...")
    wait_fresh("zero-confirm", 90)

    with lock:
        pcs_snap = {k: dict(v) for k, v in pcs_data.items()}
        f_snap   = dict(f044)

    print("\n  Final confirmation:")
    for s in SITES:
        d = pcs_snap.get(s["cab"])
        if d:
            tot = d["kVAr_a"] + d["kVAr_b"] + d["kVAr_c"]
            ok  = "ZEROED" if abs(tot) < 5.0 else "NOT zeroed"
            print(f"    {s['label']}: kVAr total = {tot:+.2f}  [{ok}]")
    if f_snap:
        print(f"    0_4_4 Q: {f_snap['Q']:+.2f} kVAr  |  Ia={f_snap['Ia']:.1f} A")

    client.loop_stop()
    client.disconnect()
    print("\nDone.")

if __name__ == "__main__":
    main()
