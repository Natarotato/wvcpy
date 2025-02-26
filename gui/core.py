import tkinter as tk
from tkinter import ttk, messagebox
from functools import partial

import pandas as pd
from .tabs import InverterTab
from inverter_monitoring.file_ops import save_data, load_historical_data
from inverter_monitoring.config import CONFIG
import time
from datetime import datetime

class InverterMonitoringGUI:
    def __init__(self, root, inverters):
        self.root = root
        self.inverters = inverters
        self.tabs = {}
        self.graphs = {}
        self.values = {inverter: {} for inverter in inverters}
        self.status_lights = {}
        
        # Create notebook for tabs
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(expand=True, fill="both")
        
        # Create tabs for each inverter
        self.create_tabs()
        
        # Create menu
        self.create_menu()
        
        # Start data update loop
        self.update_data()

    def create_tabs(self):
        for i, inverter in enumerate(self.inverters):
            sheet_name = f"{CONFIG['sheet_prefix']}{i + 1}"
            device_id = inverter if i < 2 else None  # Real IDs for Inverters 1 and 2, None for 3 and 4 (placeholder)
            tab = ttk.Frame(self.notebook)
            self.notebook.add(tab, text=inverter)
            self.tabs[inverter] = tab
            InverterTab.setup_tab(self, tab, device_id, sheet_name, inverter)

    def create_menu(self):
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        
        view_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="View", menu=view_menu)
        view_menu.add_command(label="Last Hour", command=partial(self.handle_range_selection, None, "Last Hour"))
        view_menu.add_command(label="Last Day", command=partial(self.handle_range_selection, None, "Last Day"))
        view_menu.add_command(label="Last 7 Days", command=partial(self.handle_range_selection, None, "Last 7 Days"))
        view_menu.add_command(label="All", command=partial(self.handle_range_selection, None, "All"))
        view_menu.add_command(label="Specific Hour", command=partial(self.handle_range_selection, None, "Specific Hour"))
        view_menu.add_command(label="Zoom", command=partial(self.enable_zoom, None))

    def update_data(self):
        for inverter in self.inverters:
            sheet_name = f"{CONFIG['sheet_prefix']}{self.inverters.index(inverter) + 1}"
            try:
                # Placeholder for real data collection (replace with actual API or hardware calls later)
                current_data = self.collect_real_data(inverter, sheet_name)
                if current_data is not None:
                    self.update_values(inverter, current_data)
                    save_data(current_data, sheet_name)
                    self.update_graphs(inverter)
            except Exception as e:
                print(f"❌ Failed to get status for device {inverter}: {e}")
        self.root.after(5000, self.update_data)  # Update every 5 seconds

    def collect_real_data(self, inverter, sheet_name):
        # Placeholder for real data collection (replace with actual API or hardware calls later)
        # For now, return None or an empty DataFrame for Inverters 3 and 4
        if inverter in ['Inverter 1', 'Inverter 2']:
            historical_data = load_historical_data(sheet_name)
            if not historical_data.empty:
                # Simulate a simple update (replace with real data later)
                current_time = datetime.now()
                sample = pd.DataFrame({
                    "Timestamp": [current_time],
                    "AC Power (W)": [0],  # Placeholder values
                    "AC Voltage (V)": [0],
                    "Frequency (Hz)": [0],
                    "DC Power (W)": [0],
                    "DC Voltage (V)": [0],
                    "DC Current (A)": [0],
                    "Temperature (°C)": [0],
                    "Reverse Energy (kWh)": [0]
                })
                return sample
        return None  # No data for Inverters 3 and 4 for now

    def update_values(self, inverter, data):
        for column in data.columns:
            if column != "Timestamp":
                value = data[column].iloc[0] if not pd.isna(data[column].iloc[0]) else "N/A"
                self.values[inverter][column].config(text=str(value))
        # Update status light (green for real inverters, grey for unconfigured)
        status = "green" if inverter in ['Inverter 1', 'Inverter 2'] else "lightgrey"
        self.status_lights[inverter].itemconfig("status", fill=status)

    def update_graphs(self, inverter):
        from .graphs import update_all_graphs  # Relative import within gui package
        update_all_graphs(self, inverter)

    def handle_range_selection(self, tab_id, value):
        print(f"Handling range selection: tab_id={tab_id}, value={value}")  # Debug print
        if value == "Specific Hour":
            self.prompt_specific_hour(tab_id)
        elif value == "Zoom":
            self.enable_zoom(tab_id)
        else:
            for inverter in self.inverters:
                self.graphs[inverter]["range_var"].set(value)
                self.update_graphs(inverter)

    def prompt_specific_hour(self, tab_id):
        def on_submit():
            hour_str = hour_entry.get()
            try:
                hour = datetime.strptime(hour_str, "%H:%M").time()
                current_date = datetime.now().date()
                start = datetime.combine(current_date, hour)
                end = start.replace(hour=start.hour + 1, minute=0, second=0, microsecond=0)
                for inverter in self.inverters:
                    self.graphs[inverter]["range_var"].set(f"Custom-{start.isoformat()}-{end.isoformat()}")
                    self.update_graphs(inverter)
                dialog.destroy()
            except ValueError:
                messagebox.showerror("Invalid Input", "Please enter time in HH:MM format (e.g., 14:30)")

        dialog = tk.Toplevel(self.root)
        dialog.title("Select Specific Hour")
        dialog.geometry("300x150")
        ttk.Label(dialog, text="Enter Time (HH:MM):").pack(pady=10)
        hour_entry = ttk.Entry(dialog, width=10)
        hour_entry.pack(pady=5)
        ttk.Button(dialog, text="Submit", command=on_submit).pack(pady=10)

    def enable_zoom(self, tab_id):
        # Handle tab_id=None (called from menu) by applying zoom to all tabs
        if tab_id is None:
            for inverter in self.inverters:
                graph_data = self.graphs[inverter]
                for fig_key in ["power_fig", "voltage_fig", "current_fig", "energy_fig"]:
                    fig = graph_data[fig_key]
                    fig.canvas.mpl_connect('button_press_event', lambda event: self.on_press(event, inverter, fig_key))
                    fig.canvas.mpl_connect('button_release_event', lambda event: self.on_release(event, inverter, fig_key))
                self.update_graphs(inverter)
        else:
            graph_data = self.graphs[tab_id]
            for fig_key in ["power_fig", "voltage_fig", "current_fig", "energy_fig"]:
                fig = graph_data[fig_key]
                fig.canvas.mpl_connect('button_press_event', lambda event: self.on_press(event, tab_id, fig_key))
                fig.canvas.mpl_connect('button_release_event', lambda event: self.on_release(event, tab_id, fig_key))
            self.update_graphs(tab_id)

    def on_press(self, event, tab_id, fig_key):
        self.graphs[tab_id][f"{fig_key}_zoom_start"] = (event.xdata, event.ydata) if event.xdata and event.ydata else None

    def on_release(self, event, tab_id, fig_key):
        if hasattr(self.graphs[tab_id], f"{fig_key}_zoom_start") and self.graphs[tab_id][f"{fig_key}_zoom_start"]:
            zoom_start = self.graphs[tab_id][f"{fig_key}_zoom_start"]
            zoom_end = (event.xdata, event.ydata) if event.xdata and event.ydata else None
            if zoom_end:
                ax = self.graphs[tab_id][f"{fig_key.replace('_fig', '_ax')}"]
                ax.set_xlim(min(zoom_start[0], zoom_end[0]), max(zoom_start[0], zoom_end[0]))
                ax.set_ylim(min(zoom_start[1], zoom_end[1]), max(zoom_start[1], zoom_end[1]))
                self.graphs[tab_id][f"{fig_key}_canvas"].draw()
            del self.graphs[tab_id][f"{fig_key}_zoom_start"]