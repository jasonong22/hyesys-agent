"""
mqtt_probe.py — Capture raw MQTT messages, including command RX topic.
"""

import json
import time
import paho.mqtt.client as mqtt

BROKER_HOST = "loragw.advastech.com"
BROKER_PORT = 1883
USERNAME    = "TYKJadmin"
PASSWORD    = "TYKJ2018."
TIMEOUT_SEC = 60

TOPICS = [
    ("stsc/aems/message/26022703840003",              0),  # HyESys data (all types)
    ("stsc/aems/cabinet/26022703840003/multi/operate/rx", 0),  # Command echo / response
    ("stsc/aems/cabinet/26022703840003/#",            0),  # All cabinet sub-topics
]

msg_count = 0
start_time = time.time()


def on_connect(client, userdata, flags, rc, properties):
    if rc == 0:
        print(f"Connected to {BROKER_HOST}:{BROKER_PORT}")
        client.subscribe(TOPICS)
        print(f"Subscribed to {len(TOPICS)} topics — waiting {TIMEOUT_SEC}s...\n")
    else:
        print(f"Connection refused: {rc}")


def on_message(client, userdata, msg):
    global msg_count
    msg_count += 1
    raw_str = msg.payload.decode("utf-8", errors="replace")
    print(f"{'='*70}")
    print(f"MSG #{msg_count}  topic={msg.topic}")
    print(f"{'='*70}")
    try:
        parsed = json.loads(raw_str)
        print(json.dumps(parsed, indent=2, ensure_ascii=False))
    except json.JSONDecodeError:
        print(f"[RAW] {raw_str[:500]}")
    print()


client = mqtt.Client(mqtt.CallbackAPIVersion.VERSION2)
client.username_pw_set(USERNAME, PASSWORD)
client.on_connect = on_connect
client.on_message = on_message
client.connect(BROKER_HOST, BROKER_PORT, 60)
client.loop_start()

try:
    while time.time() - start_time < TIMEOUT_SEC:
        time.sleep(1)
except KeyboardInterrupt:
    pass

client.loop_stop()
client.disconnect()
print(f"\nDone. Messages captured: {msg_count}")
