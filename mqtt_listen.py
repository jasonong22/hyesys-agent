"""
Live MQTT listener — subscribes to a topic and prints every message with timestamp.
Run in a separate terminal window: python mqtt_listen.py
Press Ctrl+C to stop.
"""

import json
import time
import paho.mqtt.client as mqtt

BROKER_HOST = "loragw.advastech.com"
BROKER_PORT = 1883
MQTT_USERNAME = "TYKJadmin"
MQTT_PASSWORD = "TYKJ2018."

TOPIC = "stsc/aems/message/e45f01FFFEe9380f"

msg_count = 0


def on_connect(client, userdata, flags, reason_code, properties):
    if reason_code == 0:
        print(f"\n  Broker  : {BROKER_HOST}:{BROKER_PORT}")
        print(f"  Topic   : {TOPIC}")
        print(f"  Status  : Connected — waiting for messages...\n")
        print("=" * 72)
        client.subscribe(TOPIC, qos=0)
    else:
        print(f"Connection failed — reason: {reason_code}")


def on_disconnect(client, userdata, flags, reason_code, properties):
    if reason_code != 0:
        print(f"\n[WARN] Disconnected unexpectedly ({reason_code}) — reconnecting...")


def on_message(client, userdata, msg):
    global msg_count
    msg_count += 1
    ts = time.strftime("%Y-%m-%d %H:%M:%S")
    raw = msg.payload.decode("utf-8", errors="replace")

    print(f"[{ts}]  #{msg_count}  {msg.topic}")

    # Pretty-print if JSON, otherwise raw
    try:
        data = json.loads(raw)
        for k, v in data.items():
            print(f"  {k:<20} {v}")
    except (json.JSONDecodeError, AttributeError):
        print(f"  {raw}")

    print("-" * 72)


def main():
    print("\n HyESys MQTT Live Listener")
    print(" ─────────────────────────")

    client = mqtt.Client(mqtt.CallbackAPIVersion.VERSION2)
    client.username_pw_set(MQTT_USERNAME, MQTT_PASSWORD)
    client.on_connect    = on_connect
    client.on_disconnect = on_disconnect
    client.on_message    = on_message

    client.connect(BROKER_HOST, BROKER_PORT, keepalive=60)

    try:
        client.loop_forever()
    except KeyboardInterrupt:
        print(f"\n Stopped. Total messages received: {msg_count}")
    finally:
        client.disconnect()


if __name__ == "__main__":
    main()
