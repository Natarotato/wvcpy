import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from inverter_monitoring.config import CONFIG, CONFIG_FILE, CLOUD
import json
from datetime import datetime
import tinytuya

def open_settings(self):
    settings_win = tk.Toplevel(self.root)
    settings_win.title("Settings")
    settings_win.geometry("400x700")

    api_frame = ttk.LabelFrame(settings_win, text="API Settings", padding="5")
    api_frame.pack(fill="x", padx=5, pady=5)
    ttk.Label(api_frame, text="API Key:").grid(row=0, column=0, sticky="w")
    api_key_entry = ttk.Entry(api_frame, width=30)
    api_key_entry.insert(0, CONFIG["API_KEY"])
    api_key_entry.grid(row=0, column=1)
    ttk.Label(api_frame, text="API Secret:").grid(row=1, column=0, sticky="w")
    api_secret_entry = ttk.Entry(api_frame, width=30)
    api_secret_entry.insert(0, CONFIG["API_SECRET"])
    api_secret_entry.grid(row=1, column=1)
    ttk.Label(api_frame, text="Region:").grid(row=2, column=0, sticky="w")
    region_entry = ttk.Entry(api_frame, width=30)
    region_entry.insert(0, CONFIG["REGION"])
    region_entry.grid(row=2, column=1)

    recording_frame = ttk.LabelFrame(settings_win, text="Recording Settings", padding="5")
    recording_frame.pack(fill="x", padx=5, pady=5)
    ttk.Label(recording_frame, text="Start Time (HH:MM):").grid(row=0, column=0, sticky="w")
    start_time_entry = ttk.Entry(recording_frame, width=10)
    start_time_entry.insert(0, CONFIG["RECORDING_WINDOW"][0].strftime("%H:%M"))
    start_time_entry.grid(row=0, column=1, sticky="w")
    ttk.Label(recording_frame, text="Stop Time (HH:MM):").grid(row=1, column=0, sticky="w")
    stop_time_entry = ttk.Entry(recording_frame, width=10)
    stop_time_entry.insert(0, CONFIG["RECORDING_WINDOW"][1].strftime("%H:%M"))
    stop_time_entry.grid(row=1, column=1, sticky="w")
    ttk.Label(recording_frame, text="Interval (seconds):").grid(row=2, column=0, sticky="w")
    interval_entry = ttk.Entry(recording_frame, width=10)
    interval_entry.insert(0, str(CONFIG["FETCH_INTERVAL"]))
    interval_entry.grid(row=2, column=1, sticky="w")

    inverter_frame = ttk.LabelFrame(settings_win, text="Inverters", padding="5")
    inverter_frame.pack(fill="x", padx=5, pady=5)
    inverter_entries = []
    for i, inv in enumerate(CONFIG["INVERTERS"]):
        ttk.Label(inverter_frame, text=f"Inverter {i+1}").grid(row=i*5, column=0, columnspan=2)
        entries = {}
        for j, (key, label) in enumerate([("device_id", "Device ID:"), ("ip", "IP:"), ("local_key", "Local Key:"), ("sheet", "Sheet Name:")]):
            ttk.Label(inverter_frame, text=label).grid(row=i*5+j+1, column=0, sticky="w")
            entry = ttk.Entry(inverter_frame, width=30)
            entry.insert(0, inv[key])
            entry.grid(row=i*5+j+1, column=1)
            entries[key] = entry
        inverter_entries.append(entries)

    save_frame = ttk.LabelFrame(settings_win, text="Data Save Location", padding="5")
    save_frame.pack(fill="x", padx=5, pady=5)
    save_dir_entry = ttk.Entry(save_frame, width=30)
    save_dir_entry.insert(0, CONFIG["SAVE_DIR"])
    save_dir_entry.grid(row=0, column=0)
    ttk.Button(save_frame, text="Browse", command=lambda: save_dir_entry.insert(0, filedialog.askdirectory())).grid(row=0, column=1)

    ttk.Button(settings_win, text="Save", command=lambda: save_settings(self, 
        api_key_entry.get(), api_secret_entry.get(), region_entry.get(), 
        start_time_entry.get(), stop_time_entry.get(), interval_entry.get(),
        inverter_entries, save_dir_entry.get())).pack(pady=10)

def save_settings(self, api_key, api_secret, region, start_time, stop_time, interval, inverter_entries, save_dir):
    global CONFIG, CLOUD
    try:
        start = datetime.strptime(start_time, "%H:%M").time()
        stop = datetime.strptime(stop_time, "%H:%M").time()
        interval_int = int(interval)
        if interval_int <= 0:
            raise ValueError("Interval must be positive")
    except ValueError as e:
        messagebox.showerror("Invalid Input", f"Error: {e}. Use HH:MM for times and a positive integer for interval.")
        return

    CONFIG["API_KEY"] = api_key
    CONFIG["API_SECRET"] = api_secret
    CONFIG["REGION"] = region
    CONFIG["RECORDING_WINDOW"] = (start, stop)
    CONFIG["FETCH_INTERVAL"] = interval_int
    CONFIG["INVERTERS"] = [
        {key: entries[key].get() for key in ["device_id", "ip", "local_key", "sheet"]}
        for entries in inverter_entries
    ]
    CONFIG["SAVE_DIR"] = save_dir

    config_to_save = CONFIG.copy()
    config_to_save["RECORDING_WINDOW"] = {
        "start": start.strftime("%H:%M"),
        "stop": stop.strftime("%H:%M")
    }
    with open(CONFIG_FILE, 'w') as f:
        json.dump(config_to_save, f, indent=4)

    CLOUD = tinytuya.Cloud(
        apiRegion=CONFIG["REGION"],
        apiKey=CONFIG["API_KEY"],
        apiSecret=CONFIG["API_SECRET"],
        apiDeviceID=CONFIG["INVERTERS"][0]["device_id"]
    )
    self.log.insert(tk.END, f"[{datetime.now()}] Configuration updated\n")
    if self.running:
        self.refresh_data()