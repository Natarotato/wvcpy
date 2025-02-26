import os
import tinytuya
from datetime import datetime
import json

CONFIG_FILE = "config.json"
DEFAULT_CONFIG = {
    "API_KEY": "9kvfdejycye8958kh4ve",
    "API_SECRET": "759d51d9861c4ab68f0232836203ad2d",
    "REGION": "eu",
    "INVERTERS": [
        {"device_id": "bf53440f72897ffb99vda7", "ip": "172.17.4.184", "local_key": "_i^q|6=G2ghV69^a", "sheet": "Inverter 1"},
        {"device_id": "bf8d502c9fc3759ec6vgum", "ip": "172.17.6.176", "local_key": "|{#4oCF!URiNK3Ek", "sheet": "Inverter 2"},
        {"device_id": "bf53440f72897ffb99vda7", "ip": "172.17.4.184", "local_key": "_i^q|6=G2ghV69^a", "sheet": "Inverter 3"},  # Mirroring Inverter 1 as placeholder
        {"device_id": "bf8d502c9fc3759ec6vgum", "ip": "172.17.6.176", "local_key": "|{#4oCF!URiNK3Ek", "sheet": "Inverter 4"}   # Mirroring Inverter 2 as placeholder
    ],
    "RECORDING_WINDOW": {"start": "06:00", "stop": "20:00"},
    "FETCH_INTERVAL": 300,
    "SAVE_DIR": "data"
}

def load_config():
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, 'r') as f:
            config = json.load(f)
            existing_sheets = {inv["sheet"] for inv in config["INVERTERS"]}
            for default_inv in DEFAULT_CONFIG["INVERTERS"]:
                if default_inv["sheet"] not in existing_sheets:
                    config["INVERTERS"].append(default_inv)
            start_time = datetime.strptime(config["RECORDING_WINDOW"]["start"], "%H:%M").time()
            stop_time = datetime.strptime(config["RECORDING_WINDOW"]["stop"], "%H:%M").time()
            config["RECORDING_WINDOW"] = (start_time, stop_time)
            print(f"Loaded CONFIG with inverters: {[inv['sheet'] for inv in config['INVERTERS']]}")
            return config
    default = DEFAULT_CONFIG.copy()
    start_time = datetime.strptime(default["RECORDING_WINDOW"]["start"], "%H:%M").time()
    stop_time = datetime.strptime(default["RECORDING_WINDOW"]["stop"], "%H:%M").time()
    default["RECORDING_WINDOW"] = (start_time, stop_time)
    print(f"Loaded DEFAULT_CONFIG with inverters: {[inv['sheet'] for inv in default['INVERTERS']]}")
    return default

CONFIG = load_config()
CLOUD = tinytuya.Cloud(
    apiRegion=CONFIG["REGION"],
    apiKey=CONFIG["API_KEY"],
    apiSecret=CONFIG["API_SECRET"],
    apiDeviceID=CONFIG["INVERTERS"][0]["device_id"]  # Use Inverter 1's device_id as default, but will need adjustment for multiple devices later
)