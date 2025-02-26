import tkinter as tk
from tkinter import ttk
import threading
from tkinter import filedialog
from tkinter import messagebox
from matplotlib.ticker import MaxNLocator
import tinytuya
from .tabs import setup_tab, handle_range_selection, prompt_specific_hour, enable_zoom, on_press, on_release
from .graphs import update_power_graph, update_voltage_graph, update_current_graph, update_energy_graph, update_all_graphs, resize_graphs
from datetime import datetime, timedelta
import time
from inverter_monitoring.data import fetch_inverter_data
from inverter_monitoring.file_ops import load_historical_data, write_to_excel, export_historical_data
from inverter_monitoring.config import CONFIG, CONFIG_FILE, CLOUD
import json
import pandas as pd
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib.pyplot as plt

class InverterGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Inverter Monitoring Dashboard")
        self.running = False
        self.last_update = "Never"
        self.resize_timer = None
        self.simulate_mode = False

        # Dynamically bind graph methods to self
        self.update_power_graph = update_power_graph.__get__(self, InverterGUI)
        self.update_voltage_graph = update_voltage_graph.__get__(self, InverterGUI)
        self.update_current_graph = update_current_graph.__get__(self, InverterGUI)
        self.update_energy_graph = update_energy_graph.__get__(self, InverterGUI)
        self.update_all_graphs = update_all_graphs.__get__(self, InverterGUI)

        # Dynamically bind tab-related methods to self
        self.handle_range_selection = handle_range_selection.__get__(self, InverterGUI)
        self.prompt_specific_hour = prompt_specific_hour.__get__(self, InverterGUI)
        self.enable_zoom = enable_zoom.__get__(self, InverterGUI)
        self.on_press = on_press.__get__(self, InverterGUI)
        self.on_release = on_release.__get__(self, InverterGUI)

        self.main_frame = ttk.Frame(self.root, padding="10")
        self.main_frame.grid(row=0, column=0, sticky="nsew")

        self.notebook = ttk.Notebook(self.main_frame)
        self.notebook.grid(row=0, column=0, sticky="nsew")

        self.tabs = {}
        self.values = {}
        self.graphs = {}
        self.status_lights = {}

        print(f"Creating tabs for inverters: {[inv['sheet'] for inv in CONFIG['INVERTERS']]}")
        for i, inverter in enumerate(CONFIG["INVERTERS"]):
            tab = ttk.Frame(self.notebook)
            self.notebook.add(tab, text=inverter["sheet"])
            tab_id = inverter["device_id"] if inverter["device_id"] else f"unconfigured_{i}"
            self.tabs[tab_id] = tab
            self.values[tab_id] = {}
            self.status_lights[tab_id] = None
            setup_tab(self, tab, inverter["device_id"], inverter["sheet"], tab_id)

            # Initialize graphs with default range (ensure new options are available)
            self.handle_range_selection(tab_id, "All")

        self.control_frame = ttk.Frame(self.main_frame)
        self.control_frame.grid(row=1, column=0, pady=5, sticky="ew")
        self.start_button = tk.Button(self.control_frame, text="Start", command=self.start_monitoring, bg="gray", fg="white")
        self.start_button.grid(row=0, column=0, padx=5)
        self.stop_button = tk.Button(self.control_frame, text="Stop", command=self.stop_monitoring, bg="gray", fg="white")
        self.stop_button.grid(row=0, column=1, padx=5)
        ttk.Button(self.control_frame, text="Refresh", command=self.refresh_data).grid(row=0, column=2, padx=5)
        ttk.Button(self.control_frame, text="Settings", command=self.open_settings).grid(row=0, column=3, padx=5)
        self.last_update_label = ttk.Label(self.control_frame, text=f"Last Update: {self.last_update}")
        self.last_update_label.grid(row=0, column=4, padx=5)
        ttk.Button(self.control_frame, text="Simulate", command=self.toggle_simulate).grid(row=0, column=5, padx=5)
        ttk.Button(self.control_frame, text="Export Data", command=self.export_historical_data).grid(row=0, column=6, padx=5)

        self.log = tk.Text(self.main_frame, height=5, width=80, font=("Arial", 10))
        self.log.grid(row=2, column=0, pady=10, sticky="ew")

        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        self.main_frame.columnconfigure(0, weight=1)
        self.main_frame.rowconfigure(0, weight=1)

        self.root.bind("<Configure>", self.on_resize)

    # ... (rest of the methods remain the same)

    def setup_tab(self, tab, device_id, sheet_name, tab_id):
        data_frame = ttk.LabelFrame(tab, text="Current Values", padding="5")
        data_frame.grid(row=0, column=0, padx=5, pady=5, sticky="nsew")

        labels = ["AC Power (W)", "AC Voltage (V)", "Frequency (Hz)", "DC Power (W)", 
                  "DC Voltage (V)", "DC Current (A)", "Temperature (°C)", "Reverse Energy (kWh)"]
        for i, label in enumerate(labels):
            ttk.Label(data_frame, text=label, font=("Arial", 10)).grid(row=i, column=0, sticky="w", padx=5)
            value_label = ttk.Label(data_frame, text="N/A", font=("Arial", 10))
            value_label.grid(row=i, column=1, sticky="w", padx=5)
            self.values[tab_id][label] = value_label
        
        status_frame = ttk.Frame(data_frame)
        status_frame.grid(row=0, column=2, rowspan=len(labels), padx=5)
        ttk.Label(status_frame, text="Status", font=("Arial", 10)).grid(row=0, column=0)
        status_canvas = tk.Canvas(status_frame, width=20, height=20)
        status_canvas.grid(row=1, column=0)
        status_canvas.create_oval(2, 2, 18, 18, fill="grey" if device_id else "lightgrey", tags="status")
        self.status_lights[tab_id] = status_canvas
        
        if not device_id:
            self.notebook.tab(tab, state="disabled")

        graphs_frame = ttk.LabelFrame(tab, text="Trends", padding="5")
        graphs_frame.grid(row=0, column=1, padx=5, pady=5, sticky="nsew")
        tab.columnconfigure(1, weight=1)
        tab.rowconfigure(0, weight=1)
        graphs_frame.columnconfigure((0, 1), weight=1)
        graphs_frame.rowconfigure((0, 1, 2), weight=1)

        range_frame = ttk.Frame(graphs_frame)
        range_frame.grid(row=0, column=0, columnspan=2, pady=5)
        range_var = tk.StringVar(value="All")
        ttk.OptionMenu(range_frame, range_var, "All", "Last Hour", "Last Day", "All", 
                       command=lambda val: self.update_all_graphs(tab_id)).pack()

        initial_width, initial_height = 8, 5

        power_frame = ttk.Frame(graphs_frame)
        power_frame.grid(row=1, column=0, padx=5, pady=5, sticky="nsew")
        power_select = tk.StringVar(value="Both")
        ttk.OptionMenu(power_frame, power_select, "Both", "AC", "DC", "Both", 
                       command=lambda val: self.update_power_graph(tab_id, val)).pack(fill="x")
        power_fig, power_ax = plt.subplots(figsize=(initial_width, initial_height))
        power_fig.subplots_adjust(left=0.1, right=0.95, bottom=0.3, top=0.9)
        power_canvas = FigureCanvasTkAgg(power_fig, master=power_frame)
        power_widget = power_canvas.get_tk_widget()
        power_widget.pack(fill="both", expand=True)

        voltage_frame = ttk.Frame(graphs_frame)
        voltage_frame.grid(row=1, column=1, padx=5, pady=5, sticky="nsew")
        voltage_select = tk.StringVar(value="Both")
        ttk.OptionMenu(voltage_frame, voltage_select, "Both", "AC", "DC", "Both", 
                       command=lambda val: self.update_voltage_graph(tab_id, val)).pack(fill="x")
        voltage_fig, voltage_ax = plt.subplots(figsize=(initial_width, initial_height))
        voltage_fig.subplots_adjust(left=0.1, right=0.95, bottom=0.3, top=0.9)
        voltage_canvas = FigureCanvasTkAgg(voltage_fig, master=voltage_frame)
        voltage_widget = voltage_canvas.get_tk_widget()
        voltage_widget.pack(fill="both", expand=True)

        current_frame = ttk.Frame(graphs_frame)
        current_frame.grid(row=2, column=0, padx=5, pady=5, sticky="nsew")
        current_select = tk.StringVar(value="Both")
        ttk.OptionMenu(current_frame, current_select, "Both", "AC", "DC", "Both", 
                       command=lambda val: self.update_current_graph(tab_id, val)).pack(fill="x")
        current_fig, current_ax = plt.subplots(figsize=(initial_width, initial_height))
        current_fig.subplots_adjust(left=0.1, right=0.95, bottom=0.3, top=0.9)
        current_canvas = FigureCanvasTkAgg(current_fig, master=current_frame)
        current_widget = current_canvas.get_tk_widget()
        current_widget.pack(fill="both", expand=True)

        energy_frame = ttk.Frame(graphs_frame)
        energy_frame.grid(row=2, column=1, padx=5, pady=5, sticky="nsew")
        ttk.Label(energy_frame, text="Reverse Energy").pack(fill="x")
        energy_fig, energy_ax = plt.subplots(figsize=(initial_width, initial_height))
        energy_fig.subplots_adjust(left=0.1, right=0.95, bottom=0.3, top=0.9)
        energy_canvas = FigureCanvasTkAgg(energy_fig, master=energy_frame)
        energy_widget = energy_canvas.get_tk_widget()
        energy_widget.pack(fill="both", expand=True)

        historical_data = load_historical_data(sheet_name)
        self.graphs[tab_id] = {
            "power_fig": power_fig, "power_ax": power_ax, "power_canvas": power_canvas, "power_select": power_select,
            "voltage_fig": voltage_fig, "voltage_ax": voltage_ax, "voltage_canvas": voltage_canvas, "voltage_select": voltage_select,
            "current_fig": current_fig, "current_ax": current_ax, "current_canvas": current_canvas, "current_select": current_select,
            "energy_fig": energy_fig, "energy_ax": energy_ax, "energy_canvas": energy_canvas,
            "range_var": range_var,
            "historical_data": historical_data
        }
        self.update_all_graphs(tab_id)

    def toggle_simulate(self):
        self.simulate_mode = not self.simulate_mode
        self.log.insert(tk.END, f"[{datetime.now()}] Simulation {'enabled' if self.simulate_mode else 'disabled'}\n")
        for i, inverter in enumerate(CONFIG["INVERTERS"]):
            tab_id = inverter["device_id"] if inverter["device_id"] else f"unconfigured_{i}"
            tab = self.tabs[tab_id]
            if not inverter["device_id"]:
                state = "normal" if self.simulate_mode else "disabled"
                self.notebook.tab(tab, state=state)
        if self.simulate_mode:
            self.refresh_data()

    def export_historical_data(self):
        export_historical_data(self.graphs, self.log)

    def on_resize(self, event):
        if self.resize_timer is not None:
            self.root.after_cancel(self.resize_timer)
        self.resize_timer = self.root.after(200, self.resize_graphs)

    def resize_graphs(self):
        for tab_id, tab in self.tabs.items():
            graphs_frame = tab.winfo_children()[1]
            width = graphs_frame.winfo_width() / 100
            height = graphs_frame.winfo_height() / 200
            graphs = self.graphs[tab_id]
            for fig_key in ["power_fig", "voltage_fig", "current_fig", "energy_fig"]:
                fig = graphs[fig_key]
                new_width = max(4, width / 2 - 0.5)
                new_height = max(3, height - 0.5)
                fig.set_size_inches(new_width, new_height)
                fig.subplots_adjust(left=0.1, right=0.95, bottom=0.28, top=0.9)
                graphs[f"{fig_key.replace('_fig', '')}_canvas"].draw()

    def update_data(self):
        while self.running:
            current_time = datetime.now().time()
            if CONFIG["RECORDING_WINDOW"][0] <= current_time <= CONFIG["RECORDING_WINDOW"][1]:
                for i, inverter in enumerate(CONFIG["INVERTERS"]):
                    tab_id = inverter["device_id"] if inverter["device_id"] else f"unconfigured_{i}"
                    data = fetch_inverter_data(inverter["device_id"])
                    if data:
                        self.update_display(tab_id, data, inverter["sheet"])
                        write_to_excel(data, inverter["sheet"])
                        self.log.insert(tk.END, f"[{datetime.now()}] Data updated for {inverter['sheet']}\n")
                        self.status_lights[tab_id].itemconfig("status", fill="green")
                    else:
                        self.log.insert(tk.END, f"[{datetime.now()}] Failed to fetch data for {inverter['sheet']}\n")
                        self.status_lights[tab_id].itemconfig("status", fill="orange")
                    self.log.see(tk.END)
                    self.last_update = datetime.now().strftime("%H:%M:%S")
                    self.last_update_label.config(text=f"Last Update: {self.last_update}")
            time.sleep(CONFIG["FETCH_INTERVAL"])

    def refresh_data(self):
        for i, inverter in enumerate(CONFIG["INVERTERS"]):
            tab_id = inverter["device_id"] if inverter["device_id"] else f"unconfigured_{i}"
            data = fetch_inverter_data(inverter["device_id"])
            if data:
                self.update_display(tab_id, data, inverter["sheet"])
                self.status_lights[tab_id].itemconfig("status", fill="green")
            else:
                self.status_lights[tab_id].itemconfig("status", fill="orange")
            self.last_update = datetime.now().strftime("%H:%M:%S")
            self.last_update_label.config(text=f"Last Update: {self.last_update}")

    def update_display(self, tab_id, data, sheet_name):
        def format_value(val):
            return f"{val:.2f}" if isinstance(val, (int, float)) else "N/A"
        self.values[tab_id]["AC Power (W)"].config(text=format_value(data["important_dps"]["ac_power (W)"]))
        self.values[tab_id]["AC Voltage (V)"].config(text=format_value(data["extracted"]["phase_a"]["ac_voltage"]))
        self.values[tab_id]["Frequency (Hz)"].config(text=format_value(data["extracted"]["phase_a"]["frequency"]))
        self.values[tab_id]["DC Power (W)"].config(text=format_value(data["extracted"]["pv1_dc_data"]["dc_power"]))
        self.values[tab_id]["DC Voltage (V)"].config(text=format_value(data["extracted"]["pv1_dc_data"]["dc_voltage"]))
        self.values[tab_id]["DC Current (A)"].config(text=format_value(data["extracted"]["pv1_dc_data"]["dc_current"]))
        temp = data["important_dps"]["temp_current (°C)"]
        self.values[tab_id]["Temperature (°C)"].config(
            text=format_value(temp),
            foreground="red" if isinstance(temp, (int, float)) and temp > 50 else "black"
        )
        self.values[tab_id]["Reverse Energy (kWh)"].config(text=format_value(data["important_dps"]["reverse_energy_total (kWh)"]))

        historical_data = self.graphs[tab_id]["historical_data"]
        new_row = pd.DataFrame([{
            "Timestamp": datetime.now(),
            "Reverse Energy (kWh)": data["important_dps"]["reverse_energy_total (kWh)"],
            "Temp (°C)": data["important_dps"]["temp_current (°C)"],
            "AC Power (W)": data["important_dps"]["ac_power (W)"],
            "AC Voltage (V)": data["extracted"]["phase_a"]["ac_voltage"],
            "Frequency (Hz)": data["extracted"]["phase_a"]["frequency"],
            "AC Current (A)": data["extracted"]["phase_a"]["ac_current (A)"],
            "DC Voltage (V)": data["extracted"]["pv1_dc_data"]["dc_voltage"],
            "DC Current (A)": data["extracted"]["pv1_dc_data"]["dc_current"],
            "DC Power (W)": data["extracted"]["pv1_dc_data"]["dc_power"]
        }])
        self.graphs[tab_id]["historical_data"] = pd.concat([historical_data, new_row], ignore_index=True)
        self.update_all_graphs(tab_id)

    def update_all_graphs(self, tab_id):
        self.update_power_graph(tab_id, self.graphs[tab_id]["power_select"].get())
        self.update_voltage_graph(tab_id, self.graphs[tab_id]["voltage_select"].get())
        self.update_current_graph(tab_id, self.graphs[tab_id]["current_select"].get())
        self.update_energy_graph(tab_id)

    def update_power_graph(self, tab_id, option):
        graph_data = self.graphs[tab_id]
        historical_data = graph_data["historical_data"]
        graph_data["power_ax"].clear()
        range_var = graph_data["range_var"].get()
        if not historical_data.empty and pd.api.types.is_datetime64_any_dtype(historical_data["Timestamp"]):
            if range_var == "Last Hour":
                historical_data = historical_data[historical_data["Timestamp"] > datetime.now() - timedelta(hours=1)]
            elif range_var == "Last Day":
                historical_data = historical_data[historical_data["Timestamp"] > datetime.now() - timedelta(days=1)]
            time_only = historical_data["Timestamp"].dt.strftime("%H:%M:%S").fillna("N/A")
            if option == "AC" and not historical_data["AC Power (W)"].isna().all():
                graph_data["power_ax"].plot(time_only, historical_data["AC Power (W)"], 'b-', label="AC Power (W)")
            elif option == "DC" and not historical_data["DC Power (W)"].isna().all():
                graph_data["power_ax"].plot(time_only, historical_data["DC Power (W)"], 'g-', label="DC Power (W)")
            elif option == "Both":
                if not historical_data["AC Power (W)"].isna().all():
                    graph_data["power_ax"].plot(time_only, historical_data["AC Power (W)"], 'b-', label="AC Power (W)")
                if not historical_data["DC Power (W)"].isna().all():
                    graph_data["power_ax"].plot(time_only, historical_data["DC Power (W)"], 'g-', label="DC Power (W)")
        graph_data["power_ax"].set_title("Power Trends")
        graph_data["power_ax"].legend(loc="upper left")
        graph_data["power_ax"].grid(True)
        graph_data["power_ax"].tick_params(axis='x', rotation=45)
        graph_data["power_ax"].xaxis.set_major_locator(MaxNLocator(nbins=6))
        graph_data["power_ax"].set_xlabel("Time (HH:MM:SS)")
        graph_data["power_fig"].tight_layout()
        graph_data["power_canvas"].draw()

    def update_voltage_graph(self, tab_id, option):
        graph_data = self.graphs[tab_id]
        historical_data = graph_data["historical_data"]
        graph_data["voltage_ax"].clear()
        range_var = graph_data["range_var"].get()
        if not historical_data.empty and pd.api.types.is_datetime64_any_dtype(historical_data["Timestamp"]):
            if range_var == "Last Hour":
                historical_data = historical_data[historical_data["Timestamp"] > datetime.now() - timedelta(hours=1)]
            elif range_var == "Last Day":
                historical_data = historical_data[historical_data["Timestamp"] > datetime.now() - timedelta(days=1)]
            time_only = historical_data["Timestamp"].dt.strftime("%H:%M:%S").fillna("N/A")
            if option == "AC" and not historical_data["AC Voltage (V)"].isna().all():
                graph_data["voltage_ax"].plot(time_only, historical_data["AC Voltage (V)"], 'b-', label="AC Voltage (V)")
            elif option == "DC" and not historical_data["DC Voltage (V)"].isna().all():
                graph_data["voltage_ax"].plot(time_only, historical_data["DC Voltage (V)"], 'g-', label="DC Voltage (V)")
            elif option == "Both":
                if not historical_data["AC Voltage (V)"].isna().all():
                    graph_data["voltage_ax"].plot(time_only, historical_data["AC Voltage (V)"], 'b-', label="AC Voltage (V)")
                if not historical_data["DC Voltage (V)"].isna().all():
                    graph_data["voltage_ax"].plot(time_only, historical_data["DC Voltage (V)"], 'g-', label="DC Voltage (V)")
        graph_data["voltage_ax"].set_title("Voltage Trends")
        graph_data["voltage_ax"].legend(loc="upper left")
        graph_data["voltage_ax"].grid(True)
        graph_data["voltage_ax"].tick_params(axis='x', rotation=45)
        graph_data["voltage_ax"].xaxis.set_major_locator(MaxNLocator(nbins=6))
        graph_data["voltage_ax"].set_xlabel("Time (HH:MM:SS)")
        graph_data["voltage_fig"].tight_layout()
        graph_data["voltage_canvas"].draw()

    def update_current_graph(self, tab_id, option):
        graph_data = self.graphs[tab_id]
        historical_data = graph_data["historical_data"]
        graph_data["current_ax"].clear()
        range_var = graph_data["range_var"].get()
        if not historical_data.empty and pd.api.types.is_datetime64_any_dtype(historical_data["Timestamp"]):
            if range_var == "Last Hour":
                historical_data = historical_data[historical_data["Timestamp"] > datetime.now() - timedelta(hours=1)]
            elif range_var == "Last Day":
                historical_data = historical_data[historical_data["Timestamp"] > datetime.now() - timedelta(days=1)]
            time_only = historical_data["Timestamp"].dt.strftime("%H:%M:%S").fillna("N/A")
            if option == "AC" and not historical_data["AC Current (A)"].isna().all():
                graph_data["current_ax"].plot(time_only, historical_data["AC Current (A)"], 'b-', label="AC Current (A)")
            elif option == "DC" and not historical_data["DC Current (A)"].isna().all():
                graph_data["current_ax"].plot(time_only, historical_data["DC Current (A)"], 'g-', label="DC Current (A)")
            elif option == "Both":
                if not historical_data["AC Current (A)"].isna().all():
                    graph_data["current_ax"].plot(time_only, historical_data["AC Current (A)"], 'b-', label="AC Current (A)")
                if not historical_data["DC Current (A)"].isna().all():
                    graph_data["current_ax"].plot(time_only, historical_data["DC Current (A)"], 'g-', label="DC Current (A)")
        graph_data["current_ax"].set_title("Current Trends")
        graph_data["current_ax"].legend(loc="upper left")
        graph_data["current_ax"].grid(True)
        graph_data["current_ax"].tick_params(axis='x', rotation=45)
        graph_data["current_ax"].xaxis.set_major_locator(MaxNLocator(nbins=6))
        graph_data["current_ax"].set_xlabel("Time (HH:MM:SS)")
        graph_data["current_fig"].tight_layout()
        graph_data["current_canvas"].draw()

    def update_energy_graph(self, tab_id):
        graph_data = self.graphs[tab_id]
        historical_data = graph_data["historical_data"]
        graph_data["energy_ax"].clear()
        range_var = graph_data["range_var"].get()
        if not historical_data.empty and pd.api.types.is_datetime64_any_dtype(historical_data["Timestamp"]):
            if range_var == "Last Hour":
                historical_data = historical_data[historical_data["Timestamp"] > datetime.now() - timedelta(hours=1)]
            elif range_var == "Last Day":
                historical_data = historical_data[historical_data["Timestamp"] > datetime.now() - timedelta(days=1)]
            time_only = historical_data["Timestamp"].dt.strftime("%H:%M:%S").fillna("N/A")
            if not historical_data["Reverse Energy (kWh)"].isna().all():
                graph_data["energy_ax"].plot(time_only, historical_data["Reverse Energy (kWh)"], 'm-', label="Energy (kWh)")
        graph_data["energy_ax"].set_title("Energy Trends")
        graph_data["energy_ax"].legend(loc="upper left")
        graph_data["energy_ax"].grid(True)
        graph_data["energy_ax"].tick_params(axis='x', rotation=45)
        graph_data["energy_ax"].xaxis.set_major_locator(MaxNLocator(nbins=6))
        graph_data["energy_ax"].set_xlabel("Time (HH:MM:SS)")
        graph_data["energy_fig"].tight_layout()
        graph_data["energy_canvas"].draw()

    def start_monitoring(self):
        if not self.running:
            self.running = True
            self.thread = threading.Thread(target=self.update_data, daemon=True)
            self.thread.start()
            self.log.insert(tk.END, f"[{datetime.now()}] Monitoring started\n")
            self.start_button.config(bg="green")
            self.stop_button.config(bg="gray")

    def stop_monitoring(self):
        if self.running:
            self.running = False
            self.log.insert(tk.END, f"[{datetime.now()}] Monitoring stopped\n")
            self.start_button.config(bg="gray")
            self.stop_button.config(bg="red")

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

        ttk.Button(settings_win, text="Save", command=lambda: self.save_settings(
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