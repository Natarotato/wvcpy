import tinytuya
import base64
import struct
from datetime import datetime
from typing import Dict, Optional, Tuple
import random
from config import CONFIG, CLOUD

def decode_tuya_value(encoded_value: str) -> Optional[Tuple[int, ...]]:
    try:
        decoded_bytes = base64.b64decode(encoded_value)
        return struct.unpack(f">{len(decoded_bytes)//2}H", decoded_bytes)
    except Exception as e:
        print(f"⚠️ Base64 decoding error: {e}")
        return None

def fetch_inverter_data(device_id: str) -> Optional[Dict]:
    try:
        from __main__ import app
        if app.simulate_mode and not device_id:
            sim_data = {
                "timestamp": datetime.now().strftime("%H:%M:%S"),
                "important_dps": {
                    "reverse_energy_total (kWh)": random.uniform(0, 10),
                    "temp_current (°C)": random.uniform(20, 30),
                    "ac_power (W)": random.uniform(50, 200)
                },
                "extracted": {
                    "phase_a": {
                        "ac_voltage": random.uniform(220, 240),
                        "frequency": random.uniform(49.9, 50.1),
                        "ac_current (A)": random.uniform(0.2, 1.0)
                    },
                    "pv1_dc_data": {
                        "dc_voltage": random.uniform(250, 350),
                        "dc_current": random.uniform(0.1, 0.5),
                        "dc_power": random.uniform(50, 150)
                    }
                }
            }
            print(f"🎭 Simulated data for device {device_id}")
            return sim_data
    except (ImportError, AttributeError):
        pass

    try:
        CLOUD.apiDeviceID = device_id
        status = CLOUD.getstatus(device_id)
        if not status or "result" not in status:
            print(f"❌ Failed to get status for device {device_id}")
            return None

        result = status["result"]
        data = result if isinstance(result, list) else [result]

        important_dps = {
            item["code"]: item["value"]
            for item in data
            if item.get("code", "") in ["reverse_energy_total", "temp_current", "ac_power"]
        }
        
        if "reverse_energy_total" in important_dps:
            important_dps["reverse_energy_total"] = important_dps["reverse_energy_total"] / 100
        if "ac_power" in important_dps:
            important_dps["ac_power"] = important_dps["ac_power"] / 10

        phase_a = next((decode_tuya_value(item["value"]) for item in data 
                        if item.get("code") == "phase_a"), None)
        pv1_dc_data = next((decode_tuya_value(item["value"]) for item in data 
                            if item.get("code") == "pv1_dc_data"), None)

        phase_a_data = {"ac_voltage": None, "frequency": None}
        if phase_a and len(phase_a) >= 2:
            phase_a_data = {"ac_voltage": phase_a[0] / 10, "frequency": phase_a[-1] / 10}

        pv1_dc_data_extracted = {"dc_voltage": None, "dc_current": None, "dc_power": None}
        if pv1_dc_data and len(pv1_dc_data) >= 3:
            pv1_dc_data_extracted = {
                "dc_voltage": pv1_dc_data[0] / 10,
                "dc_current": pv1_dc_data[1] / 10,
                "dc_power": pv1_dc_data[2] / 10
            }

        ac_current = (important_dps.get("ac_power") / phase_a_data["ac_voltage"] 
                      if phase_a_data["ac_voltage"] and "ac_power" in important_dps else None)

        return {
            "timestamp": datetime.now().strftime("%H:%M:%S"),
            "important_dps": {
                "reverse_energy_total (kWh)": important_dps.get("reverse_energy_total", "N/A"),
                "temp_current (°C)": important_dps.get("temp_current", "N/A"),
                "ac_power (W)": important_dps.get("ac_power", "N/A")
            },
            "extracted": {
                "phase_a": {**phase_a_data, "ac_current (A)": ac_current},
                "pv1_dc_data": pv1_dc_data_extracted
            }
        }
    except Exception as e:
        print(f"❌ Unexpected error for device {device_id}: {e}")
        return None