import sys
from tkinter import ttk
import tkinter as tk
from tkinter import messagebox
from .graphs import update_all_graphs
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from inverter_monitoring.file_ops import load_historical_data  # Absolute import
from inverter_monitoring.config import CONFIG
from datetime import datetime

print(f"sys.path in tabs.py: {sys.path}")  # Debug print to check module search path

class InverterTab:
    @staticmethod
    def setup_tab(self, tab, device_id, sheet_name, tab_id):
        # Current Values Frame
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

        # Trends Frame (only graph selection dropdowns)
        graphs_frame = ttk.LabelFrame(tab, text="Trends", padding="0")  # No padding
        graphs_frame.grid(row=0, column=1, padx=0, pady=0, sticky="nsew")  # No padding
        tab.columnconfigure(1, weight=1)
        tab.rowconfigure(0, weight=1)
        graphs_frame.columnconfigure((0, 1), weight=1)
        graphs_frame.rowconfigure((0, 1), weight=1)  # Adjust for only graph frames

        initial_width, initial_height = 8, 5

        # Power Graph
        power_frame = ttk.Frame(graphs_frame, padding="0")  # No frame padding
        power_frame.grid(row=0, column=0, padx=0, pady=0, sticky="nsew")  # No padding
        power_select = tk.StringVar(value="Both")
        ttk.OptionMenu(power_frame, power_select, "Both", "AC", "DC", "Both", 
                       command=lambda val: self.update_power_graph(tab_id, val)).pack(fill="x", padx=0, pady=0)  # No padding
        # Ensure a completely clean figure with a single subplot, with balanced margins
        power_fig = plt.Figure(figsize=(initial_width, initial_height), facecolor='white', dpi=100)  # Explicit DPI
        if hasattr(power_fig, 'axes'):
            for ax in power_fig.axes:
                ax.remove()  # Remove any default axes
        power_fig.clf()  # Explicitly clear any default subplots
        power_ax = power_fig.add_subplot(111)  # Single subplot
        power_fig.subplots_adjust(left=0.12, right=0.88, bottom=0.15, top=0.88)  # Keep current margins
        # Removed tight_layout to eliminate potential warnings, relying on subplots_adjust for layout
        # power_fig.tight_layout(pad=2.0)  # Commented out to prevent warnings
        # Reset and recreate canvas with resize event handling
        if hasattr(self, 'graphs') and tab_id in self.graphs and self.graphs[tab_id].get("power_canvas"):
            self.graphs[tab_id]["power_canvas"].get_tk_widget().destroy()  # Destroy old canvas
        power_canvas = FigureCanvasTkAgg(power_fig, master=power_frame)
        power_canvas.mpl_connect('resize_event', lambda event: power_canvas.draw())  # Handle resize events
        power_canvas.draw()  # Explicitly draw the figure after setup
        power_widget = power_canvas.get_tk_widget()
        power_widget.pack(fill="both", expand=True, padx=0, pady=0)  # Tight packing, no padding

        # Voltage Graph
        voltage_frame = ttk.Frame(graphs_frame, padding="0")
        voltage_frame.grid(row=0, column=1, padx=0, pady=0, sticky="nsew")
        voltage_select = tk.StringVar(value="Both")
        ttk.OptionMenu(voltage_frame, voltage_select, "Both", "AC", "DC", "Both", 
                       command=lambda val: self.update_voltage_graph(tab_id, val)).pack(fill="x", padx=0, pady=0)
        voltage_fig = plt.Figure(figsize=(initial_width, initial_height), facecolor='white', dpi=100)
        if hasattr(voltage_fig, 'axes'):
            for ax in voltage_fig.axes:
                ax.remove()
        voltage_fig.clf()
        voltage_ax = voltage_fig.add_subplot(111)  # Single subplot
        voltage_fig.subplots_adjust(left=0.12, right=0.88, bottom=0.15, top=0.88)  # Keep current margins
        # Removed tight_layout to eliminate potential warnings, relying on subplots_adjust for layout
        # voltage_fig.tight_layout(pad=2.0)  # Commented out to prevent warnings
        if hasattr(self, 'graphs') and tab_id in self.graphs and self.graphs[tab_id].get("voltage_canvas"):
            self.graphs[tab_id]["voltage_canvas"].get_tk_widget().destroy()
        voltage_canvas = FigureCanvasTkAgg(voltage_fig, master=voltage_frame)
        voltage_canvas.mpl_connect('resize_event', lambda event: voltage_canvas.draw())
        voltage_canvas.draw()  # Explicitly draw the figure
        voltage_widget = voltage_canvas.get_tk_widget()
        voltage_widget.pack(fill="both", expand=True, padx=0, pady=0)

        # Current Graph
        current_frame = ttk.Frame(graphs_frame, padding="0")
        current_frame.grid(row=1, column=0, padx=0, pady=0, sticky="nsew")
        current_select = tk.StringVar(value="Both")
        ttk.OptionMenu(current_frame, current_select, "Both", "AC", "DC", "Both", 
                       command=lambda val: self.update_current_graph(tab_id, val)).pack(fill="x", padx=0, pady=0)
        current_fig = plt.Figure(figsize=(initial_width, initial_height), facecolor='white', dpi=100)
        if hasattr(current_fig, 'axes'):
            for ax in current_fig.axes:
                ax.remove()
        current_fig.clf()
        current_ax = current_fig.add_subplot(111)  # Single subplot
        current_fig.subplots_adjust(left=0.12, right=0.88, bottom=0.15, top=0.88)  # Keep current margins
        # Removed tight_layout to eliminate potential warnings, relying on subplots_adjust for layout
        # current_fig.tight_layout(pad=2.0)  # Commented out to prevent warnings
        if hasattr(self, 'graphs') and tab_id in self.graphs and self.graphs[tab_id].get("current_canvas"):
            self.graphs[tab_id]["current_canvas"].get_tk_widget().destroy()
        current_canvas = FigureCanvasTkAgg(current_fig, master=current_frame)
        current_canvas.mpl_connect('resize_event', lambda event: current_canvas.draw())
        current_canvas.draw()  # Explicitly draw the figure
        current_widget = current_canvas.get_tk_widget()
        current_widget.pack(fill="both", expand=True, padx=0, pady=0)

        # Energy Graph
        energy_frame = ttk.Frame(graphs_frame, padding="0")
        energy_frame.grid(row=1, column=1, padx=0, pady=0, sticky="nsew")
        ttk.Label(energy_frame, text="Reverse Energy").pack(fill="x", padx=0, pady=0)
        energy_fig = plt.Figure(figsize=(initial_width, initial_height), facecolor='white', dpi=100)
        if hasattr(energy_fig, 'axes'):
            for ax in energy_fig.axes:
                ax.remove()
        energy_fig.clf()
        energy_ax = energy_fig.add_subplot(111)  # Single subplot
        energy_fig.subplots_adjust(left=0.12, right=0.88, bottom=0.15, top=0.88)  # Keep current margins
        # Removed tight_layout to eliminate potential warnings, relying on subplots_adjust for layout
        # energy_fig.tight_layout(pad=2.0)  # Commented out to prevent warnings
        if hasattr(self, 'graphs') and tab_id in self.graphs and self.graphs[tab_id].get("energy_canvas"):
            self.graphs[tab_id]["energy_canvas"].get_tk_widget().destroy()
        energy_canvas = FigureCanvasTkAgg(energy_fig, master=energy_frame)
        energy_canvas.mpl_connect('resize_event', lambda event: energy_canvas.draw())
        energy_canvas.draw()  # Explicitly draw the figure
        energy_widget = energy_canvas.get_tk_widget()
        energy_widget.pack(fill="both", expand=True, padx=0, pady=0)

        historical_data = load_historical_data(sheet_name)
        self.graphs[tab_id] = {
            "power_fig": power_fig, "power_ax": power_ax, "power_canvas": power_canvas, "power_select": power_select,
            "voltage_fig": voltage_fig, "voltage_ax": voltage_ax, "voltage_canvas": voltage_canvas, "voltage_select": voltage_select,
            "current_fig": current_fig, "current_ax": current_ax, "current_canvas": current_canvas, "current_select": current_select,
            "energy_fig": energy_fig, "energy_ax": energy_ax, "energy_canvas": energy_canvas,
            "range_var": tk.StringVar(value="All"),  # Keep range_var for graph updates, controlled by menu
            "historical_data": historical_data
        }

        # Initialize graphs with default range (handled by menu)
        self.handle_range_selection(tab_id, "All")