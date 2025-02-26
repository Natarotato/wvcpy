import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.ticker import MaxNLocator
import pandas as pd
from datetime import datetime, timedelta
import tkinter as tk

def update_power_graph(self, tab_id, option):
    graph_data = self.graphs[tab_id]
    historical_data = graph_data["historical_data"].copy()
    # Clear the figure completely and re-add the single subplot with explicit checks and debugging
    print(f"Updating power graph for tab {tab_id}, initial axes: {len(graph_data['power_fig'].axes)}")  # Debug: check initial axes
    if hasattr(graph_data["power_fig"], 'axes'):
        for ax in graph_data["power_fig"].axes:
            ax.remove()  # Remove all existing axes
    graph_data["power_fig"].clear()
    graph_data["power_ax"] = graph_data["power_fig"].add_subplot(111)
    print(f"Power graph updated for tab {tab_id}, axes after update: {len(graph_data['power_fig'].axes)}")  # Debug: check axes after update
    graph_data["power_fig"].set_facecolor('white')  # Ensure clean background
    
    range_var = graph_data["range_var"].get()
    if not historical_data.empty and pd.api.types.is_datetime64_any_dtype(historical_data["Timestamp"]):
        historical_data["Timestamp"] = pd.to_datetime(historical_data["Timestamp"], errors='coerce')
        
        # Handle range filtering
        if range_var == "Last Hour":
            historical_data = historical_data[historical_data["Timestamp"] > datetime.now() - timedelta(hours=1)]
        elif range_var == "Last Day":
            historical_data = historical_data[historical_data["Timestamp"] > datetime.now() - timedelta(days=1)]
        elif range_var == "Last 7 Days":
            historical_data = historical_data[historical_data["Timestamp"] > datetime.now() - timedelta(days=7)]
        elif range_var.startswith("Custom-"):
            try:
                start_str, end_str = range_var.split("-")[1], range_var.split("-")[2]
                start = datetime.fromisoformat(start_str)
                end = datetime.fromisoformat(end_str)
                historical_data = historical_data[
                    (historical_data["Timestamp"] >= start) & (historical_data["Timestamp"] <= end)
                ]
            except (ValueError, IndexError):
                historical_data = historical_data  # Fall back to all data if parsing fails
        elif range_var == "All":
            pass  # No filtering for "All"

        time_only = historical_data["Timestamp"].dt.strftime("%H:%M:%S").fillna("N/A")
        
        # Plot only the selected option with dynamic y-limits to prevent overlap
        if option == "AC" and not historical_data["AC Power (W)"].isna().all():
            graph_data["power_ax"].plot(time_only, historical_data["AC Power (W)"], 'b-', label="AC Power (W)")
        elif option == "DC" and not historical_data["DC Power (W)"].isna().all():
            graph_data["power_ax"].plot(time_only, historical_data["DC Power (W)"], 'g-', label="DC Power (W)")
        elif option == "Both":
            if not historical_data["AC Power (W)"].isna().all() and not historical_data["DC Power (W)"].isna().all():
                # Calculate y-limits to prevent overlap
                ac_data = historical_data["AC Power (W)"].dropna()
                dc_data = historical_data["DC Power (W)"].dropna()
                if not ac_data.empty and not dc_data.empty:
                    y_min = min(ac_data.min(), dc_data.min()) - 10  # Add buffer
                    y_max = max(ac_data.max(), dc_data.max()) + 10
                    graph_data["power_ax"].set_ylim(y_min, y_max)
                graph_data["power_ax"].plot(time_only, historical_data["AC Power (W)"], 'b-', label="AC Power (W)")
                graph_data["power_ax"].plot(time_only, historical_data["DC Power (W)"], 'g-', label="DC Power (W)")
            elif not historical_data["AC Power (W)"].isna().all():
                graph_data["power_ax"].plot(time_only, historical_data["AC Power (W)"], 'b-', label="AC Power (W)")
            elif not historical_data["DC Power (W)"].isna().all():
                graph_data["power_ax"].plot(time_only, historical_data["DC Power (W)"], 'g-', label="DC Power (W)")
    
    graph_data["power_ax"].set_title("Power Trends")
    graph_data["power_ax"].legend(loc="upper left")
    graph_data["power_ax"].grid(True)
    graph_data["power_ax"].tick_params(axis='x', rotation=45)
    graph_data["power_ax"].xaxis.set_major_locator(MaxNLocator(nbins=6))
    graph_data["power_ax"].set_xlabel("Time (HH:MM:SS)")
    graph_data["power_fig"].subplots_adjust(left=0.12, right=0.88, bottom=0.15, top=0.88)  # Keep current margins
    # Removed tight_layout to eliminate warnings, relying on subplots_adjust for layout
    graph_data["power_canvas"].draw()

def update_voltage_graph(self, tab_id, option):
    graph_data = self.graphs[tab_id]
    historical_data = graph_data["historical_data"].copy()
    print(f"Updating voltage graph for tab {tab_id}, initial axes: {len(graph_data['voltage_fig'].axes)}")
    if hasattr(graph_data["voltage_fig"], 'axes'):
        for ax in graph_data["voltage_fig"].axes:
            ax.remove()
    graph_data["voltage_fig"].clear()
    graph_data["voltage_ax"] = graph_data["voltage_fig"].add_subplot(111)
    print(f"Voltage graph updated for tab {tab_id}, axes after update: {len(graph_data['voltage_fig'].axes)}")
    graph_data["voltage_fig"].set_facecolor('white')
    
    range_var = graph_data["range_var"].get()
    if not historical_data.empty and pd.api.types.is_datetime64_any_dtype(historical_data["Timestamp"]):
        historical_data["Timestamp"] = pd.to_datetime(historical_data["Timestamp"], errors='coerce')
        
        if range_var == "Last Hour":
            historical_data = historical_data[historical_data["Timestamp"] > datetime.now() - timedelta(hours=1)]
        elif range_var == "Last Day":
            historical_data = historical_data[historical_data["Timestamp"] > datetime.now() - timedelta(days=1)]
        elif range_var == "Last 7 Days":
            historical_data = historical_data[historical_data["Timestamp"] > datetime.now() - timedelta(days=7)]
        elif range_var.startswith("Custom-"):
            try:
                start_str, end_str = range_var.split("-")[1], range_var.split("-")[2]
                start = datetime.fromisoformat(start_str)
                end = datetime.fromisoformat(end_str)
                historical_data = historical_data[
                    (historical_data["Timestamp"] >= start) & (historical_data["Timestamp"] <= end)
                ]
            except (ValueError, IndexError):
                historical_data = historical_data  # Fall back to all data if parsing fails
        elif range_var == "All":
            pass

        time_only = historical_data["Timestamp"].dt.strftime("%H:%M:%S").fillna("N/A")
        
        if option == "AC" and not historical_data["AC Voltage (V)"].isna().all():
            graph_data["voltage_ax"].plot(time_only, historical_data["AC Voltage (V)"], 'b-', label="AC Voltage (V)")
        elif option == "DC" and not historical_data["DC Voltage (V)"].isna().all():
            graph_data["voltage_ax"].plot(time_only, historical_data["DC Voltage (V)"], 'g-', label="DC Voltage (V)")
        elif option == "Both":
            if not historical_data["AC Voltage (V)"].isna().all() and not historical_data["DC Voltage (V)"].isna().all():
                ac_data = historical_data["AC Voltage (V)"].dropna()
                dc_data = historical_data["DC Voltage (V)"].dropna()
                if not ac_data.empty and not dc_data.empty:
                    y_min = min(ac_data.min(), dc_data.min()) - 5  # Add buffer
                    y_max = max(ac_data.max(), dc_data.max()) + 5
                    graph_data["voltage_ax"].set_ylim(y_min, y_max)
                graph_data["voltage_ax"].plot(time_only, historical_data["AC Voltage (V)"], 'b-', label="AC Voltage (V)")
                graph_data["voltage_ax"].plot(time_only, historical_data["DC Voltage (V)"], 'g-', label="DC Voltage (V)")
            elif not historical_data["AC Voltage (V)"].isna().all():
                graph_data["voltage_ax"].plot(time_only, historical_data["AC Voltage (V)"], 'b-', label="AC Voltage (V)")
            elif not historical_data["DC Voltage (V)"].isna().all():
                graph_data["voltage_ax"].plot(time_only, historical_data["DC Voltage (V)"], 'g-', label="DC Voltage (V)")
    
    graph_data["voltage_ax"].set_title("Voltage Trends")
    graph_data["voltage_ax"].legend(loc="upper left")
    graph_data["voltage_ax"].grid(True)
    graph_data["voltage_ax"].tick_params(axis='x', rotation=45)
    graph_data["voltage_ax"].xaxis.set_major_locator(MaxNLocator(nbins=6))
    graph_data["voltage_ax"].set_xlabel("Time (HH:MM:SS)")
    graph_data["voltage_fig"].subplots_adjust(left=0.12, right=0.88, bottom=0.15, top=0.88)  # Keep current margins
    # Removed tight_layout to eliminate warnings, relying on subplots_adjust for layout
    graph_data["voltage_canvas"].draw()

def update_current_graph(self, tab_id, option):
    graph_data = self.graphs[tab_id]
    historical_data = graph_data["historical_data"].copy()
    print(f"Updating current graph for tab {tab_id}, initial axes: {len(graph_data['current_fig'].axes)}")
    if hasattr(graph_data["current_fig"], 'axes'):
        for ax in graph_data["current_fig"].axes:
            ax.remove()
    graph_data["current_fig"].clear()
    graph_data["current_ax"] = graph_data["current_fig"].add_subplot(111)
    print(f"Current graph updated for tab {tab_id}, axes after update: {len(graph_data['current_fig'].axes)}")
    graph_data["current_fig"].set_facecolor('white')
    
    range_var = graph_data["range_var"].get()
    if not historical_data.empty and pd.api.types.is_datetime64_any_dtype(historical_data["Timestamp"]):
        historical_data["Timestamp"] = pd.to_datetime(historical_data["Timestamp"], errors='coerce')
        
        if range_var == "Last Hour":
            historical_data = historical_data[historical_data["Timestamp"] > datetime.now() - timedelta(hours=1)]
        elif range_var == "Last Day":
            historical_data = historical_data[historical_data["Timestamp"] > datetime.now() - timedelta(days=1)]
        elif range_var == "Last 7 Days":
            historical_data = historical_data[historical_data["Timestamp"] > datetime.now() - timedelta(days=7)]
        elif range_var.startswith("Custom-"):
            try:
                start_str, end_str = range_var.split("-")[1], range_var.split("-")[2]
                start = datetime.fromisoformat(start_str)
                end = datetime.fromisoformat(end_str)
                historical_data = historical_data[
                    (historical_data["Timestamp"] >= start) & (historical_data["Timestamp"] <= end)
                ]
            except (ValueError, IndexError):
                historical_data = historical_data  # Fall back to all data if parsing fails
        elif range_var == "All":
            pass

        time_only = historical_data["Timestamp"].dt.strftime("%H:%M:%S").fillna("N/A")
        
        if option == "AC" and not historical_data["AC Current (A)"].isna().all():
            graph_data["current_ax"].plot(time_only, historical_data["AC Current (A)"], 'b-', label="AC Current (A)")
        elif option == "DC" and not historical_data["DC Current (A)"].isna().all():
            graph_data["current_ax"].plot(time_only, historical_data["DC Current (A)"], 'g-', label="DC Current (A)")
        elif option == "Both":
            if not historical_data["AC Current (A)"].isna().all() and not historical_data["DC Current (A)"].isna().all():
                ac_data = historical_data["AC Current (A)"].dropna()
                dc_data = historical_data["DC Current (A)"].dropna()
                if not ac_data.empty and not dc_data.empty:
                    y_min = min(ac_data.min(), dc_data.min()) - 0.5  # Add buffer
                    y_max = max(ac_data.max(), dc_data.max()) + 0.5
                    graph_data["current_ax"].set_ylim(y_min, y_max)
                graph_data["current_ax"].plot(time_only, historical_data["AC Current (A)"], 'b-', label="AC Current (A)")
                graph_data["current_ax"].plot(time_only, historical_data["DC Current (A)"], 'g-', label="DC Current (A)")
            elif not historical_data["AC Current (A)"].isna().all():
                graph_data["current_ax"].plot(time_only, historical_data["AC Current (A)"], 'b-', label="AC Current (A)")
            elif not historical_data["DC Current (A)"].isna().all():
                graph_data["current_ax"].plot(time_only, historical_data["DC Current (A)"], 'g-', label="DC Current (A)")
    
    graph_data["current_ax"].set_title("Current Trends")
    graph_data["current_ax"].legend(loc="upper left")
    graph_data["current_ax"].grid(True)
    graph_data["current_ax"].tick_params(axis='x', rotation=45)
    graph_data["current_ax"].xaxis.set_major_locator(MaxNLocator(nbins=6))
    graph_data["current_ax"].set_xlabel("Time (HH:MM:SS)")
    graph_data["current_fig"].subplots_adjust(left=0.12, right=0.88, bottom=0.15, top=0.88)  # Keep current margins
    # Removed tight_layout to eliminate warnings, relying on subplots_adjust for layout
    graph_data["current_canvas"].draw()

def update_energy_graph(self, tab_id):
    graph_data = self.graphs[tab_id]
    historical_data = graph_data["historical_data"].copy()
    print(f"Updating energy graph for tab {tab_id}, initial axes: {len(graph_data['energy_fig'].axes)}")
    if hasattr(graph_data["energy_fig"], 'axes'):
        for ax in graph_data["energy_fig"].axes:
            ax.remove()
    graph_data["energy_fig"].clear()
    graph_data["energy_ax"] = graph_data["energy_fig"].add_subplot(111)
    print(f"Energy graph updated for tab {tab_id}, axes after update: {len(graph_data['energy_fig'].axes)}")
    graph_data["energy_fig"].set_facecolor('white')
    
    range_var = graph_data["range_var"].get()
    if not historical_data.empty and pd.api.types.is_datetime64_any_dtype(historical_data["Timestamp"]):
        historical_data["Timestamp"] = pd.to_datetime(historical_data["Timestamp"], errors='coerce')
        
        if range_var == "Last Hour":
            historical_data = historical_data[historical_data["Timestamp"] > datetime.now() - timedelta(hours=1)]
        elif range_var == "Last Day":
            historical_data = historical_data[historical_data["Timestamp"] > datetime.now() - timedelta(days=1)]
        elif range_var == "Last 7 Days":
            historical_data = historical_data[historical_data["Timestamp"] > datetime.now() - timedelta(days=7)]
        elif range_var.startswith("Custom-"):
            try:
                start_str, end_str = range_var.split("-")[1], range_var.split("-")[2]
                start = datetime.fromisoformat(start_str)
                end = datetime.fromisoformat(end_str)
                historical_data = historical_data[
                    (historical_data["Timestamp"] >= start) & (historical_data["Timestamp"] <= end)
                ]
            except (ValueError, IndexError):
                historical_data = historical_data  # Fall back to all data if parsing fails
        elif range_var == "All":
            pass

        time_only = historical_data["Timestamp"].dt.strftime("%H:%M:%S").fillna("N/A")
        if not historical_data["Reverse Energy (kWh)"].isna().all():
            graph_data["energy_ax"].plot(time_only, historical_data["Reverse Energy (kWh)"], 'm-', label="Energy (kWh)")
    
    graph_data["energy_ax"].set_title("Energy Trends")
    graph_data["energy_ax"].legend(loc="upper left")
    graph_data["energy_ax"].grid(True)
    graph_data["energy_ax"].tick_params(axis='x', rotation=45)
    graph_data["energy_ax"].xaxis.set_major_locator(MaxNLocator(nbins=6))
    graph_data["energy_ax"].set_xlabel("Time (HH:MM:SS)")
    graph_data["energy_fig"].subplots_adjust(left=0.12, right=0.88, bottom=0.15, top=0.88)  # Keep current margins
    # Removed tight_layout to eliminate warnings, relying on subplots_adjust for layout
    graph_data["energy_canvas"].draw()

def update_all_graphs(self, tab_id):
    self.update_power_graph(tab_id, self.graphs[tab_id]["power_select"].get())
    self.update_voltage_graph(tab_id, self.graphs[tab_id]["voltage_select"].get())
    self.update_current_graph(tab_id, self.graphs[tab_id]["current_select"].get())
    self.update_energy_graph(tab_id)

def resize_graphs(self):
    for tab_id, tab in self.tabs.items():
        graphs_frame = tab.winfo_children()[1]
        # Use actual canvas dimensions for more precise sizing, assuming 100 DPI
        width = graphs_frame.winfo_width() / 100  # Convert pixels to inches
        height = graphs_frame.winfo_height() / 200  # Convert pixels to inches
        graphs = self.graphs[tab_id]
        for fig_key in ["power_fig", "voltage_fig", "current_fig", "energy_fig"]:
            fig = graphs[fig_key]
            # Calculate new sizes to match canvas exactly, no buffers
            new_width = max(1, width / 2)  # Ensure minimum width of 1 inch
            new_height = max(1, height - 0.1)  # Ensure minimum height of 1 inch, minimal buffer
            fig.set_size_inches(new_width, new_height)
            fig.subplots_adjust(left=0.12, right=0.88, bottom=0.15, top=0.88)  # Keep current margins
            # Removed tight_layout to eliminate warnings, relying on subplots_adjust for layout
            graphs[f"{fig_key.replace('_fig', '')}_canvas"].draw()  # Redraw to apply new size and margins