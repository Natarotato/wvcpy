import sys
import os
import tkinter as tk

# Add project root to sys.path to ensure inverter_monitoring is found
project_root = os.path.dirname(os.path.abspath(__file__))
if project_root not in sys.path:
    sys.path.insert(0, project_root)  # Insert at the beginning for priority

from gui.core import InverterMonitoringGUI
from inverter_monitoring.config import CONFIG
from inverter_monitoring.file_ops import save_data, load_historical_data

def main():
    print(f"sys.path in main.py: {sys.path}")  # Debug print to check module search path
    
    # Load configuration
    inverters = CONFIG['inverters']
    
    # Initialize GUI
    root = tk.Tk()
    root.title("Inverter Monitoring Dashboard")
    app = InverterMonitoringGUI(root, inverters)
    
    # Start the main event loop
    root.mainloop()

if __name__ == "__main__":
    main()