import psutil
import time
import threading
import tkinter as tk
from tkinter import ttk, messagebox
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib.animation as animation
from matplotlib.figure import Figure
from collections import deque
import configparser
import platform
import logging
import os
import sys
from threading import RLock
import ctypes
import webbrowser


logging.basicConfig(
    level=logging.INFO,  # Change back from DEBUG to INFO for release
    format='%(asctime)s - %(levelname)s - %(message)s',
    filename='speedmeter.log',
    filemode='a'
)

CONFIG_FILE = "config.ini"

# Define color themes
THEMES = {
    "dark": {
        "bg": "#1E1E1E",
        "fg": "white",
        "highlight_bg": "#323232",
        "plot_bg": "#121212",
        "grid_color": "#333333",
        "button_bg": "#2196F3",
        "button_fg": "white",
        "button_active_bg": "#0D47A1",
        "button_active_fg": "white",
        "close_bg": "#FF5252",
        "close_active_bg": "#FF0000",
        "menu_bg": "#212121",
        "menu_fg": "white",
        "menu_active_bg": "#424242",
        "dl_color": "#4CAF50",
        "ul_color": "#FFC107",
        "cpu_color": "#FF5722",
        "ram_color": "#2196F3",
        "status_color": "#888888"
    },
    "light": {
        "bg": "#F5F5F5",
        "fg": "black",
        "highlight_bg": "#E0E0E0",
        "plot_bg": "#FFFFFF",
        "grid_color": "#CCCCCC",
        "button_bg": "#2196F3",
        "button_fg": "white",
        "button_active_bg": "#0D47A1",
        "button_active_fg": "white",
        "close_bg": "#FF5252",
        "close_active_bg": "#FF0000",
        "menu_bg": "#FFFFFF",
        "menu_fg": "black",
        "menu_active_bg": "#E0E0E0",
        "dl_color": "#4CAF50",
        "ul_color": "#FFC107",
        "cpu_color": "#FF5722",
        "ram_color": "#2196F3",
        "status_color": "#888888"
    },
    "system": {}
}

def load_config():
    config = configparser.ConfigParser()
    
    if not os.path.exists(CONFIG_FILE):
        config["Settings"] = {
            "theme": "dark",
            "speed_unit": "None",
            "update_interval": "0.5",
            "show_system_stats": "True"
        }
        save_config(config)
    else:
        config.read(CONFIG_FILE)
        if not config.has_section("Settings"):
            config.add_section("Settings")
        for key, default in {"theme": "dark", "speed_unit": "None", 
                             "update_interval": "0.5", "show_system_stats": "True"}.items():
            if not config.has_option("Settings", key):
                config.set("Settings", key, default)
    
    return config

def save_config(config):
    try:
        with open(CONFIG_FILE, "w") as configfile:
            config.write(configfile)
    except Exception as e:
        logging.error(f"Error saving config: {e}")

def format_speed(speed_bps, force_unit=None):
    if force_unit == "None":
        force_unit = None
        
    speed_Bps = speed_bps / 8.0
        
    if force_unit:
        if force_unit == "kbps":
            return f"{speed_Bps/1000:.1f}", "KB/s"
        elif force_unit == "Mbps":
            return f"{speed_Bps/1000000:.1f}", "MB/s"
        elif force_unit == "Gbps":
            return f"{speed_Bps/1000000000:.2f}", "GB/s"
        else:
            return f"{speed_Bps:.0f}", "B/s"
    else:
        if speed_Bps < 1000:
            return f"{speed_Bps:.0f}", "B/s"
        elif speed_Bps < 1000000:
            return f"{speed_Bps/1000:.1f}", "KB/s"
        elif speed_Bps < 1000000000:
            return f"{speed_Bps/1000000:.1f}", "MB/s"
        else:
            return f"{speed_Bps/1000000000:.2f}", "GB/s"

class EnhancedNetworkMonitor:
    def __init__(self):
        self.download_speed = 0.0
        self.upload_speed = 0.0
        self.running = True
        self.lock = threading.Lock()
        
        self.system_stats_lock = RLock()
        self.cpu_usage = 0.0
        self.cpu_per_core = []
        self.ram_usage = 0.0
        self.ram_used = 0
        self.ram_total = 0
        self.top_processes = []
        
        self.core_count = psutil.cpu_count(logical=True) or 1
        self.cpu_per_core = [0.0] * self.core_count
        
        self.last_received = psutil.net_io_counters().bytes_recv
        self.last_sent = psutil.net_io_counters().bytes_sent
        self.last_time = time.time()
        
        self.system_stats_thread = threading.Thread(target=self.update_system_stats, daemon=True)
        self.system_stats_thread.start()
        
        # Add an interface filter to allow user to select which network interface to monitor
        self.selected_interface = None  # Will use combined stats by default
        # Initialize active_method attribute
        self.active_method = "Monitoring all interfaces"
    
    def update_speeds(self):
        last_received = self.last_received
        last_sent = self.last_sent
        last_time = self.last_time
        
        while self.running:
            try:
                time.sleep(0.5)  # This is fine for normal monitoring
                
                current_time = time.time()
                try:
                    # Get per-interface counters to detect actual active interfaces
                    net_io_per_nic = psutil.net_io_counters(pernic=True)
                    
                    # If a specific interface is selected, use its stats
                    if self.selected_interface and self.selected_interface in net_io_per_nic:
                        network_stats = net_io_per_nic[self.selected_interface]
                        self.active_method = f"Monitoring {self.selected_interface}"
                    else:
                        # Get combined stats from all interfaces
                        network_stats = psutil.net_io_counters()
                        self.active_method = "Monitoring all interfaces"
                    
                    # Log active interfaces for debugging
                    logging.debug(f"Active interfaces: {list(net_io_per_nic.keys())}")
                    
                    # Get combined stats from all interfaces
                    current_received = network_stats.bytes_recv
                    current_sent = network_stats.bytes_sent
                    
                    time_delta = current_time - last_time
                    
                    # Log raw values for debugging
                    logging.debug(f"Time delta: {time_delta:.2f}s, Bytes received: {current_received-last_received}, Bytes sent: {current_sent-last_sent}")
                    
                    if current_received < last_received or current_sent < last_sent:
                        logging.warning("Network counter reset detected")
                        last_received = current_received
                        last_sent = current_sent
                        last_time = current_time
                        continue

                    with self.lock:
                        dl_bytes = current_received - last_received
                        ul_bytes = current_sent - last_sent
                        
                        # Convert bytes to bits per second
                        self.download_speed = (dl_bytes * 8) / time_delta
                        self.upload_speed = (ul_bytes * 8) / time_delta
                        
                        # Add sanity check for abnormally high values
                        if self.download_speed > 1e12:  # More than ~1 TB/s is likely an error
                            logging.warning(f"Abnormally high download speed detected: {self.download_speed} bps")
                            self.download_speed = 0
                            
                        if self.upload_speed > 1e12:
                            logging.warning(f"Abnormally high upload speed detected: {self.upload_speed} bps")
                            self.upload_speed = 0
                        
                        # Log calculated speeds for debugging
                        dl_text, dl_unit = format_speed(self.download_speed)
                        ul_text, ul_unit = format_speed(self.upload_speed)
                        logging.debug(f"Download: {dl_text} {dl_unit}, Upload: {ul_text} {ul_unit}")

                    last_received = current_received
                    last_sent = current_sent
                    last_time = current_time
                    
                except Exception as e:
                    logging.error(f"Error getting network stats: {e}")
                    time.sleep(0.5)

            except Exception as e:
                logging.error(f"Error updating speeds: {e}")
                time.sleep(0.5)
    
    def update_system_stats(self):
        windows_counters = False
        
        if platform.system() == "Windows":
            try:
                import win32pdh
                import win32pdhutil
                import win32api
                import win32con
                import win32process
                windows_counters = True
                
                self.hq = win32pdh.OpenQuery()
                self.processor_path = win32pdh.MakeCounterPath((None, "Processor", "_Total", None, 0, "% Processor Time"))
                self.processor_counter = win32pdh.AddCounter(self.hq, self.processor_path)
                win32pdh.CollectQueryData(self.hq)
                
                logging.info("Windows performance counters activated with Task Manager precision")
            except Exception as e:
                windows_counters = False
                logging.warning(f"Windows performance counters unavailable: {e}")
        
        task_manager_interval = 0.5

        while self.running:
            try:
                if platform.system() == "Windows" and windows_counters:
                    try:
                        win32pdh.CollectQueryData(self.hq)
                        time.sleep(task_manager_interval)
                        win32pdh.CollectQueryData(self.hq)
                        _, processor_raw = win32pdh.GetFormattedCounterValue(self.processor_counter, win32pdh.PDH_FMT_DOUBLE)
                        cpu_percent = processor_raw
                        
                    except Exception as e:
                        try:
                            import ctypes
                            from ctypes.wintypes import FILETIME
                            
                            class FILETIME(ctypes.Structure):
                                _fields_ = [("dwLowDateTime", ctypes.c_ulong),
                                          ("dwHighDateTime", ctypes.c_ulong)]
                            
                            kernel32 = ctypes.windll.kernel32
                            idle_time = FILETIME()
                            kernel_time = FILETIME()
                            user_time = FILETIME()
                            
                            kernel32.GetSystemTimes(ctypes.byref(idle_time), 
                                                 ctypes.byref(kernel_time),
                                                 ctypes.byref(user_time))
                            
                            idle_initial = (idle_time.dwHighDateTime << 32) + idle_time.dwLowDateTime
                            kernel_initial = (kernel_time.dwHighDateTime << 32) + kernel_time.dwLowDateTime
                            user_initial = (user_time.dwHighDateTime << 32) + user_time.dwLowDateTime
                            
                            time.sleep(task_manager_interval)
                            
                            kernel32.GetSystemTimes(ctypes.byref(idle_time), 
                                                 ctypes.byref(kernel_time),
                                                 ctypes.byref(user_time))
                                                 
                            idle_end = (idle_time.dwHighDateTime << 32) + idle_time.dwLowDateTime
                            kernel_end = (kernel_time.dwHighDateTime << 32) + kernel_time.dwLowDateTime
                            user_end = (user_time.dwHighDateTime << 32) + user_time.dwLowDateTime
                            
                            idle_delta = idle_end - idle_initial
                            kernel_delta = kernel_end - kernel_initial
                            user_delta = user_end - user_initial
                            
                            system_delta = kernel_delta + user_delta
                            
                            if system_delta > 0:
                                cpu_percent = 100.0 * (1.0 - idle_delta / float(system_delta))
                            else:
                                cpu_percent = 0.0
                                
                        except Exception as inner_e:
                            logging.error(f"All Windows-specific methods failed, using psutil: {inner_e}")
                            cpu_percent = psutil.cpu_percent(interval=0.1)
                            
                else:
                    cpu_percent = psutil.cpu_percent(interval=0.1)
                    
                cpu_percent = max(0.0, min(100.0, cpu_percent))
                
                try:
                    cpu_per_core = psutil.cpu_percent(interval=0, percpu=True)
                except Exception as core_e:
                    logging.error(f"Error getting per-core CPU: {core_e}")
                    cpu_per_core = [0.0] * self.core_count
                
                try:
                    mem = psutil.virtual_memory()
                    ram_percent = mem.percent
                    ram_used = mem.used
                    ram_total = mem.total
                except Exception as e:
                    logging.error(f"Error getting memory info: {e}")
                    ram_percent = 0.0
                    ram_used = 0
                    ram_total = 1
                
                try:
                    processes = []
                    for proc in psutil.process_iter(['pid', 'name']):
                        try:
                            proc.cpu_percent()
                        except:
                            pass
                    
                    time.sleep(0.02)
                    
                    for proc in psutil.process_iter(['pid', 'name']):
                        try:
                            cpu_usage = proc.cpu_percent()
                            if cpu_usage > 0.5:
                                name = proc.name() if hasattr(proc, 'name') and callable(proc.name) else proc.info.get('name', 'Unknown')
                                processes.append((cpu_usage, name))
                        except:
                            pass
                            
                    processes.sort(reverse=True)
                    top_processes = processes[:3]
                except Exception as e:
                    logging.error(f"Error getting process info: {e}")
                    top_processes = []

                with self.system_stats_lock:
                    self.cpu_usage = cpu_percent
                    self.cpu_per_core = cpu_per_core
                    self.ram_usage = ram_percent
                    self.ram_used = ram_used
                    self.ram_total = ram_total
                    self.top_processes = top_processes
                
                time.sleep(0.02)
                
            except Exception as e:
                logging.error(f"Error updating system stats: {e}")
                time.sleep(0.1)
    
    def get_system_stats(self):
        with self.system_stats_lock:
            return {
                "cpu_percent": self.cpu_usage,
                "cpu_per_core": self.cpu_per_core,
                "ram_percent": self.ram_usage,
                "ram_used": self.ram_used,
                "ram_total": self.ram_total,
                "top_processes": self.top_processes
            }
    
    def get_speeds(self):
        with self.lock:
            return self.download_speed, self.upload_speed
    
    def get_monitoring_method(self):
        """Returns the current network monitoring method"""
        return getattr(self, 'active_method', "Monitoring all interfaces")
    
    def stop(self):
        self.running = False
    
    # Add method to allow selecting a specific interface
    def set_interface(self, interface_name=None):
        """Set a specific network interface to monitor or None for all interfaces"""
        self.selected_interface = interface_name
        
    def get_available_interfaces(self):
        """Return a list of available network interfaces"""
        try:
            interfaces = psutil.net_io_counters(pernic=True)
            # Filter out loopback interfaces
            return [nic for nic in interfaces.keys() if not nic.startswith('lo')]
        except Exception as e:
            logging.error(f"Error getting interfaces: {e}")
            return []

class NetworkSpeedApp:
    def __init__(self, root):
        self.root = root
        self.root.title("")
        self.root.iconify()
        self.root.withdraw()
        self.root.attributes('-alpha', 0)
        self.root.attributes("-toolwindow", True)
        self.root.wm_state('iconic')

        if platform.system() == "Windows":
            try:
                import win32gui
                import win32con
                hwnd = win32gui.GetParent(root.winfo_id())
                style = win32gui.GetWindowLong(hwnd, win32con.GWL_EXSTYLE)
                style = style & ~win32con.WS_EX_APPWINDOW | win32con.WS_EX_TOOLWINDOW
                win32gui.SetWindowLong(hwnd, win32con.GWL_EXSTYLE, style)
            except Exception as e:
                logging.warning(f"Could not hide root window: {e}")
        
        config = load_config()
        self.current_theme = config.get("Settings", "theme", fallback="dark")
        self.speed_unit = config.get("Settings", "speed_unit", fallback="None")
        self.show_system_stats = config.getboolean("Settings", "show_system_stats", fallback=True)
        
        if self.current_theme == "system":
            self.detect_system_theme()
        
        self.window = tk.Toplevel(root)
        self.window.title("")
        self.window.attributes("-topmost", True)
        self.window.overrideredirect(True)
        
        screen_width = root.winfo_screenwidth()
        screen_height = root.winfo_screenheight()
        
        taskbar_height = 40
        
        window_width = 180
        window_height = 85
        
        x_position = screen_width - window_width - 5
        y_position = screen_height - window_height - taskbar_height - 5
        
        self.window.geometry(f"{window_width}x{window_height}+{x_position}+{y_position}")
        
        self.main_frame = tk.Frame(self.window)
        self.main_frame.pack(fill=tk.BOTH, expand=True, padx=2, pady=2)
        
        self.border_frame = tk.Frame(self.main_frame, bd=1)
        self.border_frame.pack(fill=tk.BOTH, expand=True, padx=1, pady=1)
        
        self.content_frame = tk.Frame(self.border_frame)
        self.content_frame.pack(fill=tk.BOTH, expand=True, padx=1, pady=1)
        
        # Create main button frame - set background color first
        self.button_frame = tk.Frame(self.content_frame, bg=THEMES[self.current_theme]["bg"])
        self.button_frame.pack(side=tk.TOP, fill=tk.X, padx=0, pady=0)
        
        # Create left side frame with background color
        self.left_frame = tk.Frame(self.button_frame, bg=THEMES[self.current_theme]["bg"])
        self.left_frame.pack(side=tk.LEFT)
        
        # Create help button with explicit background
        self.help_button = tk.Button(self.left_frame, text="?", command=self.show_menu,
                                   font=("Arial", 8, "bold"), relief=tk.FLAT, 
                                   padx=0, pady=0, bd=0, width=2,
                                   highlightthickness=0, borderwidth=0,
                                   bg=THEMES[self.current_theme]["bg"], 
                                   fg=THEMES[self.current_theme]["button_fg"],
                                   activebackground=THEMES[self.current_theme]["button_active_bg"],
                                   activeforeground=THEMES[self.current_theme]["button_active_fg"])
        
        # Create help area frame with background
        self.help_area_frame = tk.Frame(self.left_frame, width=24, height=20, 
                                      cursor="hand2", bg=THEMES[self.current_theme]["bg"])
        self.help_area_frame.pack(side=tk.LEFT, padx=0, pady=0)
        
        # Create center status label with background
        self.status_label = tk.Label(self.button_frame, text="", font=("Arial", 7),
                                  anchor="center", bg=THEMES[self.current_theme]["bg"], 
                                  fg=THEMES[self.current_theme]["status_color"])
        self.status_label.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=0, pady=0)
        
        # Create right side frame with background
        self.right_frame = tk.Frame(self.button_frame, bg=THEMES[self.current_theme]["bg"])
        self.right_frame.pack(side=tk.RIGHT)
        
        # Create close button with explicit background
        self.close_button = tk.Button(self.right_frame, text="×", command=self.close_app,
                                    font=("Arial", 8, "bold"), relief=tk.FLAT, 
                                    padx=0, pady=0, bd=0, width=2, 
                                    highlightthickness=0, borderwidth=0,
                                    bg=THEMES[self.current_theme]["bg"], 
                                    fg=THEMES[self.current_theme]["button_fg"],
                                    activebackground=THEMES[self.current_theme]["close_active_bg"],
                                    activeforeground=THEMES[self.current_theme]["button_active_fg"])
        
        # Create close area frame with background
        self.close_area_frame = tk.Frame(self.right_frame, width=24, height=20, 
                                       cursor="hand2", bg=THEMES[self.current_theme]["bg"])
        self.close_area_frame.pack(side=tk.RIGHT, padx=0, pady=0)
        
        # No need to apply styles again since we already set them in the creation

        self.monitor = EnhancedNetworkMonitor()
        
        self.data_frame = tk.Frame(self.content_frame)
        self.data_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=0, pady=0)
        
        self.data_frame.columnconfigure(0, weight=1)
        self.data_frame.columnconfigure(1, weight=0)
        self.data_frame.rowconfigure(0, weight=1)
        self.data_frame.rowconfigure(1, weight=1)
        self.data_frame.rowconfigure(2, weight=0)
        
        self.data_points = 40
        self.download_data = deque([0] * self.data_points, maxlen=self.data_points)
        self.upload_data = deque([0] * self.data_points, maxlen=self.data_points)
        self.cpu_data = deque([0] * self.data_points, maxlen=self.data_points)
        self.ram_data = deque([0] * self.data_points, maxlen=self.data_points)
        
        self.fig = Figure(figsize=(0.95, 0.35), dpi=100)
        self.fig.subplots_adjust(left=0.02, right=0.98, bottom=0.02, top=0.98, hspace=0.1)
        
        self.ax1 = self.fig.add_subplot(2, 1, 1)
        self.ax2 = self.fig.add_subplot(2, 1, 2)
        
        self.canvas = FigureCanvasTkAgg(self.fig, master=self.data_frame)
        self.canvas.draw()
        self.canvas.get_tk_widget().grid(row=0, column=0, rowspan=2, sticky="nsew", padx=(2, 0))
        
        self.dl_label = tk.Label(self.data_frame, text="↓ 0 B/s", font=("Consolas", 9),
                             anchor="w", padx=2)
        self.dl_label.grid(row=0, column=1, sticky="sw", padx=(0, 2), pady=(0, 2))
        
        self.ul_label = tk.Label(self.data_frame, text="↑ 0 B/s", font=("Consolas", 9),
                             anchor="w", padx=2)
        self.ul_label.grid(row=1, column=1, sticky="nw", padx=(0, 2), pady=(2, 0))
        
        self.stats_frame = tk.Frame(self.data_frame)
        self.stats_frame.grid(row=2, column=0, columnspan=2, sticky="ew", padx=2, pady=(2, 0))
        
        if not self.show_system_stats:
            self.stats_frame.grid_remove()
            
        self.stats_frame.columnconfigure(0, weight=1)
        self.stats_frame.columnconfigure(1, weight=1)
        
        self.cpu_frame = tk.Frame(self.stats_frame)
        self.cpu_frame.grid(row=0, column=0, sticky="ew", padx=(0, 1))
        
        self.cpu_label = tk.Label(self.cpu_frame, text="CPU: 0%", font=("Consolas", 8),
                               anchor="w", padx=1)
        self.cpu_label.pack(side=tk.TOP, fill=tk.X)
        
        self.cpu_canvas = tk.Canvas(self.cpu_frame, height=5, highlightthickness=0)
        self.cpu_canvas.pack(side=tk.BOTTOM, fill=tk.X, expand=True, pady=(2, 0))
        
        self.ram_frame = tk.Frame(self.stats_frame)
        self.ram_frame.grid(row=0, column=1, sticky="ew", padx=(1, 0))
        
        self.ram_label = tk.Label(self.ram_frame, text="RAM: 0%", font=("Consolas", 8), 
                               anchor="w", padx=1)
        self.ram_label.pack(side=tk.TOP, fill=tk.X)
        
        self.ram_canvas = tk.Canvas(self.ram_frame, height=5, highlightthickness=0)
        self.ram_canvas.pack(side=tk.BOTTOM, fill=tk.X, expand=True, pady=(2, 0))
        
        self.cpu_tooltip = ToolTip(self.cpu_label, "Loading process data...")
        self.ram_tooltip = ToolTip(self.ram_label, "System memory usage")
        
        self.style = ttk.Style()
        
        self.ani = None
        
        self.apply_theme(self.current_theme)
        
        self.monitor_thread = threading.Thread(target=self.monitor.update_speeds, daemon=True)
        self.monitor_thread.start()
        
        self.update_status()
        
        self.window.bind("<Button-1>", self.start_move)
        self.window.bind("<ButtonRelease-1>", self.stop_move)
        self.window.bind("<B1-Motion>", self.on_motion)
        
        self.setup_button_effects()
        
        if platform.system() == "Windows":
            try:
                import win32gui
                import win32con
                hwnd = win32gui.GetParent(root.winfo_id())
                style = win32gui.GetWindowLong(hwnd, win32con.GWL_EXSTYLE)
                style = style & ~win32con.WS_EX_APPWINDOW | win32con.WS_EX_TOOLWINDOW
                win32gui.SetWindowLong(hwnd, win32con.GWL_EXSTYLE, style)
            except:
                pass

        self.fix_menu_style()

    def setup_button_effects(self):
        """Add hover effects to buttons"""
        # Help button hover effects
        self.help_button.bind("<Enter>", self.on_help_hover_enter)
        self.help_button.bind("<Leave>", self.on_help_hover_leave)
        
        # Close button hover effects  
        self.close_button.bind("<Enter>", self.on_close_hover_enter)
        self.close_button.bind("<Leave>", self.on_close_hover_leave)
        
        # Bind events to show buttons on hover
        self.close_area_frame.bind("<Enter>", self.show_close_button)
        self.help_area_frame.bind("<Enter>", self.show_help_button)
        
        # Bind event to hide buttons when mouse leaves
        self.button_frame.bind("<Leave>", self.check_hide_buttons)
        self.close_button.bind("<Leave>", self.check_hide_buttons)
        self.help_button.bind("<Leave>", self.check_hide_buttons)

    def show_close_button(self, event=None):
        """Show close button when mouse enters the close area"""
        self.close_area_frame.pack_forget()
        # Set bg color before packing to avoid flash, and ensure no borders
        theme = THEMES[self.current_theme]
        self.close_button.config(
            bg=theme["close_bg"], 
            fg=theme["button_fg"],
            relief=tk.FLAT,
            borderwidth=0,
            highlightthickness=0,
            padx=0,
            pady=0
        )
        self.close_button.pack(in_=self.right_frame, side=tk.RIGHT, padx=0, pady=0, ipadx=0, ipady=0)

    def show_help_button(self, event=None):
        """Show help button when mouse enters the help area"""
        self.help_area_frame.pack_forget()
        # Set bg color before packing to avoid flash, and ensure no borders
        theme = THEMES[self.current_theme]
        self.help_button.config(
            bg=theme["button_bg"], 
            fg=theme["button_fg"],
            relief=tk.FLAT,
            borderwidth=0,
            highlightthickness=0,
            padx=0,
            pady=0
        )
        self.help_button.pack(in_=self.left_frame, side=tk.LEFT, padx=0, pady=0, ipadx=0, ipady=0)

    def check_hide_buttons(self, event=None):
        """Hide buttons when mouse leaves their areas"""
        x, y = self.window.winfo_pointerxy()
        
        # Get coordinates for left frame
        left_frame_x = self.left_frame.winfo_rootx()
        left_frame_y = self.left_frame.winfo_rooty()
        left_frame_width = self.left_frame.winfo_width()
        left_frame_height = self.left_frame.winfo_height()
        
        # Get coordinates for right frame
        right_frame_x = self.right_frame.winfo_rootx()
        right_frame_y = self.right_frame.winfo_rooty()
        right_frame_width = self.right_frame.winfo_width()
        right_frame_height = self.right_frame.winfo_height()
        
        # Check if pointer is in help button area
        in_help_area = (
            left_frame_x <= x <= left_frame_x + left_frame_width and 
            left_frame_y <= y <= left_frame_y + left_frame_height
        )
        
        # Check if pointer is in close button area
        in_close_area = (
            right_frame_x <= x <= right_frame_x + right_frame_width and 
            right_frame_y <= y <= right_frame_y + right_frame_height
        )
        
        # Hide help button if pointer not in help area
        if not in_help_area:
            self.help_button.pack_forget()
            self.help_area_frame.pack(in_=self.left_frame, side=tk.LEFT, padx=0, pady=0)  # Remove padding
        
        # Hide close button if pointer not in close area
        if not in_close_area:
            self.close_button.pack_forget()
            self.close_area_frame.pack(in_=self.right_frame, side=tk.RIGHT, padx=0, pady=0)  # Remove padding

    def on_help_hover_enter(self, event):
        # Don't change background color immediately to prevent flash
        self.window.after(10, lambda: self.help_button.config(bg="#0D47A1"))
        
    def on_help_hover_leave(self, event):
        # Set to theme color to prevent flashing
        theme = THEMES[self.current_theme]
        self.window.after(10, lambda: self.help_button.config(bg=theme["button_bg"]))
        
    def on_close_hover_enter(self, event):
        # Don't change background color immediately to prevent flash
        self.window.after(10, lambda: self.close_button.config(bg="#FF0000"))
        
    def on_close_hover_leave(self, event):
        # Set to theme color to prevent flashing
        theme = THEMES[self.current_theme]
        self.window.after(10, lambda: self.close_button.config(bg=theme["close_bg"]))
    
    def update_status(self):
        try:
            method = self.monitor.get_monitoring_method()
            # Use original text format without extra padding
            self.status_label.config(text=method)
            
        except Exception as e:
            logging.error(f"Error updating status: {e}")
        
        self.window.after(2000, self.update_status)
    
    def detect_system_theme(self):
        system = platform.system()
        try:
            if system == "Windows":
                try:
                    import winreg
                    key = winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                                         r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize")
                    value, _ = winreg.QueryValueEx(key, "AppsUseLightTheme")
                    if value == 0:
                        THEMES["system"] = THEMES["dark"].copy()
                    else:
                        THEMES["system"] = THEMES["light"].copy()
                    
                    # Apply the updated system theme if we're using it
                    if self.current_theme == "system":
                        self.update_theme_colors()
                except Exception as e:
                    logging.warning(f"Could not detect system theme: {e}")
                    THEMES["system"] = THEMES["dark"].copy()
            else:
                THEMES["system"] = THEMES["dark"].copy()
        except Exception as e:
            logging.error(f"Error in system theme detection: {e}")
            THEMES["system"] = THEMES["dark"].copy()

    def apply_theme(self, theme_name):
        try:
            if theme_name not in THEMES:
                logging.warning(f"Theme '{theme_name}' not found, using dark theme instead")
                theme_name = "dark"
                
            self.current_theme = theme_name
            self.update_theme_colors()
            
        except Exception as e:
            logging.error(f"Error applying theme: {e}")
            self.current_theme = "dark"
            try:
                self.update_theme_colors()
            except:
                logging.critical("Could not apply any theme")

    def update_theme_colors(self):
        if self.current_theme not in THEMES:
            self.current_theme = "dark"
            
        theme = THEMES[self.current_theme]
        
        required_keys = [
            "bg", "fg", "highlight_bg", "plot_bg", "grid_color", 
            "button_bg", "button_fg", "button_active_bg", "button_active_fg",
            "close_bg", "close_active_bg", "menu_bg", "menu_fg", "menu_active_bg",
            "dl_color", "ul_color", "cpu_color", "ram_color", "status_color"
        ]
        
        for key in required_keys:
            if key not in theme:
                if key in THEMES["dark"]:
                    theme[key] = THEMES["dark"][key]
                else:
                    if key == "bg":
                        theme[key] = "#1E1E1E"
                    elif key == "fg":
                        theme[key] = "white"
                    else:
                        theme[key] = "#2196F3"
        
        self.window.configure(bg=theme["bg"])
        self.main_frame.configure(bg=theme["bg"])
        self.border_frame.configure(bg=theme["highlight_bg"])
        self.content_frame.configure(bg=theme["bg"])
        self.button_frame.configure(bg=theme["bg"])
        self.data_frame.configure(bg=theme["bg"])
        self.close_area_frame.configure(bg=theme["bg"])
        self.help_area_frame.configure(bg=theme["bg"])
        
        self.help_button.config(bg=theme["button_bg"], fg=theme["button_fg"],
                              activebackground=theme["button_active_bg"], 
                              activeforeground=theme["button_active_fg"])
        self.close_button.config(bg=theme["close_bg"], fg=theme["button_fg"],
                               activebackground=theme["close_active_bg"], 
                               activeforeground=theme["button_active_fg"])
        self.status_label.config(bg=theme["bg"], fg=theme["status_color"])
        
        self.dl_label.config(bg=theme["bg"], fg=theme["dl_color"])
        self.ul_label.config(bg=theme["bg"], fg=theme["ul_color"])
        
        self.stats_frame.configure(bg=theme["bg"])
        self.cpu_frame.configure(bg=theme["bg"])
        self.ram_frame.configure(bg=theme["bg"])
        self.cpu_label.configure(bg=theme["bg"], fg=theme["cpu_color"])
        self.ram_label.configure(bg=theme["bg"], fg=theme["ram_color"])
        self.cpu_canvas.configure(bg=theme["plot_bg"])
        self.ram_canvas.configure(bg=theme["plot_bg"])
        
        self.fig.patch.set_facecolor(theme["bg"])
        for ax in [self.ax1, self.ax2]:
            ax.set_facecolor(theme["plot_bg"])
            ax.grid(True, color=theme["grid_color"], alpha=0.5)
            ax.spines['top'].set_visible(False)
            ax.spines['right'].set_visible(False)
            ax.spines['bottom'].set_visible(False)
            ax.spines['left'].set_visible(False)
        
        self.canvas.draw()
    
    def show_menu(self):
        # Apply style directly for root menu
        root = self.window._root() if hasattr(self.window, '_root') else self.window
        root.option_add('*Menu.borderWidth', '0')
        root.option_add('*Menu.activeBorderWidth', '0')

        # Create menu with no border and additional styling - removing unsupported options
        menu = tk.Menu(self.window, tearoff=0, 
                      bg=THEMES[self.current_theme]["menu_bg"], 
                      fg=THEMES[self.current_theme]["menu_fg"], 
                      activebackground=THEMES[self.current_theme]["menu_active_bg"], 
                      activeforeground=THEMES[self.current_theme]["fg"],
                      bd=0, relief=tk.FLAT, borderwidth=0)
        
        # Create theme submenu with consistent styling - removing unsupported options
        theme_menu = tk.Menu(menu, tearoff=0, 
                           bg=THEMES[self.current_theme]["menu_bg"], 
                           fg=THEMES[self.current_theme]["menu_fg"], 
                           activebackground=THEMES[self.current_theme]["menu_active_bg"], 
                           activeforeground=THEMES[self.current_theme]["fg"],
                           bd=0, relief=tk.FLAT, borderwidth=0)
        theme_menu.add_command(label="Dark Theme", command=lambda: self.change_theme("dark"))
        theme_menu.add_command(label="Light Theme", command=lambda: self.change_theme("light"))
        theme_menu.add_command(label="System Default", command=lambda: self.change_theme("system"))
        
        # Create unit submenu with consistent styling - removing unsupported options
        unit_menu = tk.Menu(menu, tearoff=0, 
                          bg=THEMES[self.current_theme]["menu_bg"], 
                          fg=THEMES[self.current_theme]["menu_fg"], 
                          activebackground=THEMES[self.current_theme]["menu_active_bg"], 
                          activeforeground=THEMES[self.current_theme]["fg"],
                          bd=0, relief=tk.FLAT, borderwidth=0)
        unit_menu.add_command(label="Auto", command=lambda: self.set_speed_unit("None"))
        unit_menu.add_command(label="KB/s", command=lambda: self.set_speed_unit("kbps"))
        unit_menu.add_command(label="MB/s", command=lambda: self.set_speed_unit("Mbps"))
        unit_menu.add_command(label="GB/s", command=lambda: self.set_speed_unit("Gbps"))
        
        # Add interface selection submenu
        interface_menu = tk.Menu(menu, tearoff=0, 
                             bg=THEMES[self.current_theme]["menu_bg"], 
                             fg=THEMES[self.current_theme]["menu_fg"], 
                             activebackground=THEMES[self.current_theme]["menu_active_bg"], 
                             activeforeground=THEMES[self.current_theme]["fg"],
                             bd=0, relief=tk.FLAT, borderwidth=0)
                             
        # Add "All Interfaces" option
        interface_menu.add_command(label="All Interfaces", 
                               command=lambda: self.select_interface(None))
        
        # Add available interfaces
        for interface in self.monitor.get_available_interfaces():
            interface_menu.add_command(label=interface, 
                                   command=lambda intf=interface: self.select_interface(intf))
        
        menu.add_cascade(label="Network Interface", menu=interface_menu)
        menu.add_cascade(label="Theme", menu=theme_menu)
        menu.add_cascade(label="Speed Unit", menu=unit_menu)
        menu.add_separator()
        
        if self.show_system_stats:
            menu.add_command(label="Hide CPU/RAM", command=self.toggle_system_stats)
        else:
            menu.add_command(label="Show CPU/RAM", command=self.toggle_system_stats)
        
        menu.add_command(label="Reset Application", command=self.reset_app)
        
        menu.add_separator()
        menu.add_command(label="About", command=self.show_about)
        
        x = self.help_button.winfo_rootx()
        y = self.help_button.winfo_rooty() + self.help_button.winfo_height()
        
        # Custom implementation for smoother menu showing - direct call to tk
        menu._tclCommands = []
        menu.tk_popup(x, y, 0)
        
        return "break"  # Prevent default handling
    
    def close_app(self):
        self.monitor.stop()
        if self.monitor_thread.is_alive():
            self.monitor_thread.join(0.5)
        self.window.destroy()
        self.root.destroy()
    
    def change_theme(self, theme_name):
        self.apply_theme(theme_name)
        config = load_config()
        config.set("Settings", "theme", theme_name)
        save_config(config)
    
    def set_speed_unit(self, unit):
        self.speed_unit = unit
        config = load_config()
        config.set("Settings", "speed_unit", str(unit))
        save_config(config)
    
    def toggle_system_stats(self):
        self.show_system_stats = not self.show_system_stats
        
        config = load_config()
        config.set("Settings", "show_system_stats", str(self.show_system_stats))
        save_config(config)
        
        if self.show_system_stats:
            self.stats_frame.grid(row=2, column=0, columnspan=2, sticky="ew", padx=2, pady=(2, 0))
        else:
            self.stats_frame.grid_remove()
            
        self.window.update_idletasks()
    
    def reset_app(self):
        theme = THEMES[self.current_theme]
        if not messagebox.askyesno("Reset Application", 
                                "Are you sure you want to reset the application?",
                                parent=self.window):
            return
            
        self.monitor.stop()
        if self.monitor_thread.is_alive():
            self.monitor_thread.join(1.0)
        
        try:
            if os.path.exists(CONFIG_FILE):
                os.remove(CONFIG_FILE)
        except Exception as e:
            logging.error(f"Error removing config file: {e}")
        
        self.window.destroy()
        self.root.destroy()
        
        import subprocess
        python = sys.executable
        script_path = os.path.abspath(__file__)
        subprocess.Popen([python, script_path])
        sys.exit(0)
    
    def show_about(self):
        about_theme = THEMES["dark"]
        
        about_root = tk.Toplevel(self.window)
        about_root.withdraw()
        about_root.overrideredirect(True)
        about_root.configure(bg=about_theme["bg"])
        
        width, height = 320, 238

        screen_width = self.window.winfo_screenwidth()
        screen_height = self.window.winfo_screenheight()
        x = (screen_width - width) // 2
        y = (screen_height - height) // 2
        
        about_root.geometry(f"{width}x{height}+{x}+{y}")
        border_frame = tk.Frame(about_root, bg=about_theme["highlight_bg"], bd=1)
        border_frame.pack(fill=tk.BOTH, expand=True, padx=1, pady=1)
        
        main_frame = tk.Frame(border_frame, bg=about_theme["bg"])
        main_frame.pack(fill=tk.BOTH, expand=True, padx=1, pady=1)

        title_bar = tk.Frame(main_frame, bg=about_theme["bg"], height=25)
        title_bar.pack(fill=tk.X, side=tk.TOP)
        title_bar.pack_propagate(False)

        window_title = tk.Label(title_bar, text="About Bit Meter", 
                             bg=about_theme["bg"], fg=about_theme["fg"],
                             font=("Arial", 9))
        window_title.pack(side=tk.LEFT, padx=10)
        
        close_btn = tk.Label(title_bar, text="×", bg=about_theme["bg"], 
                           fg=about_theme["fg"], font=("Arial", 12, "bold"),
                           cursor="hand2")
        close_btn.pack(side=tk.RIGHT, padx=10)
        close_btn.bind("<Button-1>", lambda e: about_root.destroy())
        close_btn.bind("<Enter>", lambda e: close_btn.config(fg=about_theme["close_bg"]))
        close_btn.bind("<Leave>", lambda e: close_btn.config(fg=about_theme["fg"]))
        
        separator = tk.Frame(main_frame, height=1, bg=about_theme["highlight_bg"])
        separator.pack(fill=tk.X, padx=0, pady=0)

        content_frame = tk.Frame(main_frame, bg=about_theme["bg"], padx=20, pady=10)
        content_frame.pack(fill=tk.BOTH, expand=True, padx=0, pady=0)
        
        title_label = tk.Label(content_frame, text="Bit Meter", 
                            font=("Arial", 14, "bold"),
                            bg=about_theme["bg"], fg=about_theme["fg"])
        title_label.pack(pady=(5, 5))
        
        version_label = tk.Label(content_frame, text="Version 1.0",
                              font=("Arial", 9),
                              bg=about_theme["bg"], fg=about_theme["fg"])
        version_label.pack(pady=(0, 10))
        
        desc_label = tk.Label(content_frame, 
                            text="A lightweight tool to monitor network speeds,\nCPU and RAM usage in real-time.",
                            justify=tk.CENTER,
                            bg=about_theme["bg"], fg=about_theme["fg"],
                            font=("Arial", 9))
        desc_label.pack(pady=5)
        
        # Add GitHub link
        github_link = tk.Label(content_frame,
                           text="Check my GitHub project",
                           font=("Arial", 8, "underline"), 
                           justify=tk.CENTER,
                           bg=about_theme["bg"], 
                           fg=about_theme["button_bg"],  # Use button color for link
                           cursor="hand2")  # Change cursor to hand when hovering
        github_link.pack(pady=2)
        
        # Function to open the GitHub link
        def open_github(event):
            webbrowser.open_new("https://github.com/Bijoy121/Bit-Meter")
        
        # Bind the click event
        github_link.bind("<Button-1>", open_github)
        
        credits_label = tk.Label(content_frame,
                              text="Made with love BD\nby Bijoy Basak",
                              font=("Arial", 8), justify=tk.CENTER,
                              bg=about_theme["bg"], fg=about_theme["status_color"])
        credits_label.pack(pady=8)
        
        button_frame = tk.Frame(content_frame, bg=about_theme["bg"])
        button_frame.pack(pady=(5, 5))
        
        ok_button = tk.Button(button_frame, text="OK", 
                            bg=about_theme["button_bg"], 
                            fg=about_theme["button_fg"],
                            activebackground=about_theme["button_active_bg"],
                            activeforeground=about_theme["button_active_fg"],
                            relief=tk.FLAT, bd=0, padx=20, pady=2,
                            command=about_root.destroy)
        ok_button.pack()
        
        def start_move(event):
            about_root.x = event.x
            about_root.y = event.y
        
        def on_motion(event):
            deltax = event.x - about_root.x
            deltay = event.y - about_root.y
            x = about_root.winfo_x() + deltax
            y = about_root.winfo_y() + deltay
            about_root.geometry(f"+{x}+{y}")
        
        title_bar.bind("<ButtonPress-1>", start_move)
        title_bar.bind("<B1-Motion>", on_motion)
        window_title.bind("<ButtonPress-1>", start_move)
        window_title.bind("<B1-Motion>", on_motion)
        

        about_root.bind("<Escape>", lambda e: about_root.destroy())              
        about_root.attributes("-topmost", True)
        about_root.update_idletasks() 

        about_root.deiconify()

        self._about_window_ref = about_root

    def start_move(self, event):
        if event.widget not in [self.close_button, self.help_button]:
            self.x = event.x
            self.y = event.y
    
    def stop_move(self, event):
        self.x = None
        self.y = None
    
    def on_motion(self, event):
        if hasattr(self, 'x') and self.x is not None:
            deltax = event.x - self.x
            deltay = event.y - self.y
            x = self.window.winfo_x() + deltax
            y = self.window.winfo_y() + deltay
            self.window.geometry(f"+{x}+{y}")
    
    def update_plot(self, i):
        try:
            theme = THEMES[self.current_theme]
            
            dl_speed, ul_speed = self.monitor.get_speeds()
            
            system_stats = self.monitor.get_system_stats()
            cpu_percent = system_stats["cpu_percent"]
            ram_percent = system_stats["ram_percent"]
            ram_used = system_stats["ram_used"]
            ram_total = system_stats["ram_total"]
            cpu_per_core = system_stats.get("cpu_per_core", [])
            top_processes = system_stats.get("top_processes", [])
            
            if hasattr(self, 'cpu_tooltip'):
                top_process_text = "Top CPU processes:\n"
                if top_processes:
                    for i, proc in enumerate(top_processes):
                        if len(proc) >= 2 and proc[0] > 0:
                            top_process_text += f"• {proc[1]}: {int(proc[0])}%\n"
                else:
                    top_process_text += "No processes with significant CPU usage"
                self.cpu_tooltip.update_text(top_process_text)

            if hasattr(self, 'ram_tooltip'):
                ram_gb_used = ram_used / (1024**3)
                ram_gb_total = ram_total / (1024**3)
                ram_tooltip_text = f"Memory usage: {int(ram_percent)}%\n"
                ram_tooltip_text += f"Used: {ram_gb_used:.1f} GB\n"
                ram_tooltip_text += f"Total: {ram_gb_total:.1f} GB"
                self.ram_tooltip.update_text(ram_tooltip_text)
            
            self.download_data.append(dl_speed)
            self.upload_data.append(ul_speed)
            self.cpu_data.append(cpu_percent)
            self.ram_data.append(ram_percent)
            
            dl_text, dl_unit = format_speed(dl_speed, self.speed_unit)
            ul_text, ul_unit = format_speed(ul_speed, self.speed_unit)
            
            self.dl_label.config(text=f"D(↓) {dl_text} {dl_unit}")
            self.ul_label.config(text=f"U(↑) {ul_text} {ul_unit}")
            
            self.cpu_label.config(text=f"CPU: {int(cpu_percent)}%")
            self.ram_label.config(text=f"RAM: {int(ram_percent)}%")
            
            width = self.cpu_canvas.winfo_width()
            height = self.cpu_canvas.winfo_height()
            if width > 1:
                self.cpu_canvas.delete("all")
                bar_width = max(1, int(width * (cpu_percent / 100.0)))
                self.cpu_canvas.create_rectangle(0, 0, bar_width, height, 
                                             fill=theme["cpu_color"], outline="")
                
            width = self.ram_canvas.winfo_width()
            height = self.ram_canvas.winfo_height()
            if width > 1:
                self.ram_canvas.delete("all")
                bar_width = max(1, int(width * (ram_percent / 100.0)))
                self.ram_canvas.create_rectangle(0, 0, bar_width, height, 
                                             fill=theme["ram_color"], outline="")
                
            # Store current background color before clearing
            bg_color = theme["plot_bg"]
            
            # Clear with specific background color
            self.ax1.clear()
            self.ax2.clear()
            
            # Calculate smooth max values to prevent frequent rescaling
            # Only rescale when really needed (values exceed current scale by 20% or drop below 50%)
            if not hasattr(self, 'current_max_dl'):
                self.current_max_dl = max(self.download_data) if max(self.download_data) > 0 else 1
                self.current_max_ul = max(self.upload_data) if max(self.upload_data) > 0 else 1
            else:
                max_dl = max(self.download_data) if max(self.download_data) > 0 else 1
                max_ul = max(self.upload_data) if max(self.upload_data) > 0 else 1
                
                # Only increase scale if new max exceeds current by 20%
                if max_dl > self.current_max_dl * 1.2:
                    self.current_max_dl = max_dl
                # Only decrease scale if new max is less than 50% of current
                elif max_dl < self.current_max_dl * 0.5 and max_dl > 0:
                    self.current_max_dl = max(max_dl * 2, 1)  # Smoother downscaling
                    
                # Same logic for upload
                if max_ul > self.current_max_ul * 1.2:
                    self.current_max_ul = max_ul
                elif max_ul < self.current_max_ul * 0.5 and max_ul > 0:
                    self.current_max_ul = max(max_ul * 2, 1)  # Smoother downscaling
            
            # Set y-limits with smoothed values and some padding
            self.ax1.set_ylim(0, self.current_max_dl * 1.2)
            self.ax2.set_ylim(0, self.current_max_ul * 1.2)
            
            # Fill plot background explicitly before drawing data
            self.ax1.patch.set_facecolor(bg_color)
            self.ax2.patch.set_facecolor(bg_color)
            
            # Draw plots
            self.ax1.fill_between(range(self.data_points), list(self.download_data), 
                                color=theme["dl_color"], alpha=0.3)
            self.ax1.plot(range(self.data_points), list(self.download_data), 
                        color=theme["dl_color"], linewidth=1.0)
            
            self.ax2.fill_between(range(self.data_points), list(self.upload_data), 
                                color=theme["ul_color"], alpha=0.3)
            self.ax2.plot(range(self.data_points), list(self.upload_data), 
                        color=theme["ul_color"], linewidth=1.0)
            
            # Apply consistent styling for both axes
            for ax in [self.ax1, self.ax2]:
                ax.set_facecolor(bg_color)
                ax.grid(True, color=theme["grid_color"], alpha=0.5)
                ax.set_xlim(0, self.data_points-1)
                ax.set_xticks([])
                ax.set_yticks([])
                ax.spines['top'].set_visible(False)
                ax.spines['right'].set_visible(False)
                ax.spines['bottom'].set_visible(False)
                ax.spines['left'].set_visible(False)
                
            # Draw without using blit to ensure full redraw when needed
            self.canvas.draw()
        
        except Exception as e:
            logging.error(f"Error updating plot: {e}")

    def fix_menu_style(self):
        """Configure Tkinter menu styles to remove all borders"""
        try:
            self.window.option_add('*Menu.borderWidth', '0')
            self.window.option_add('*Menu.activeBorderWidth', '0')
            self.window.option_add('*Menu.relief', 'flat')
            self.window.option_add('*Menu.activeRelief', 'flat')
        except Exception as e:
            logging.warning(f"Could not set menu styles: {e}")

    def select_interface(self, interface_name):
        """Select a specific network interface to monitor"""
        self.monitor.set_interface(interface_name)

class ToolTip:
    def __init__(self, widget, text=""):
        self.widget = widget
        self.text = text
        self.tip_window = None
        
        self.widget.bind("<Enter>", self.show_tip)
        self.widget.bind("<Leave>", self.hide_tip)
        
    def show_tip(self, event=None):
        if self.tip_window:
            return
            
        x = self.widget.winfo_rootx() + 20
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 5
        
        self.tip_window = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)
        tw.wm_geometry(f"+{x}+{y}")
        tw.attributes("-topmost", True)
        
        frame = tk.Frame(tw, borderwidth=1, relief="solid", background="#FFFFEA")
        frame.pack(fill="both", expand=True)
        
        self.label = tk.Label(frame, text=self.text, justify=tk.LEFT,
                      background="#FFFFEA", foreground="#000000",
                      font=("Arial", "8", "normal"), padx=5, pady=3,
                      wraplength=250)
        self.label.pack(padx=1, pady=1)
        
    def hide_tip(self, event=None):
        if self.tip_window:
            self.tip_window.destroy()
            self.tip_window = None
            
    def update_text(self, text):
        self.text = text
        if hasattr(self, 'label') and self.label and self.tip_window:
            self.label.config(text=self.text)

_animations = []

def main():
    # Add proper error handling for no network connection
    try:
        # Check if a network connection exists
        psutil.net_io_counters()
    except Exception as e:
        logging.error(f"Network error: {e}")
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("Network Error", 
                            "Unable to access network interfaces. Please check your connection and try again.")
        sys.exit(1)

    root = tk.Tk()
    root.withdraw()
    root.attributes('-alpha', 0)
    root.attributes("-topmost", False)
    root.attributes("-toolwindow", True)
    root.wm_state('iconic')
    
    # Create app first before manipulating windows
    app = NetworkSpeedApp(root)
    
    # Let Tk process events and create windows before accessing handles
    root.update_idletasks()
    
    if platform.system() == "Windows":
        try:
            root.wm_attributes("-toolwindow", 1)
            
            import win32gui
            import win32con
            
            # Let's ensure the window is created by forcing another update
            root.update_idletasks()
            
            # Check if the root window ID is valid
            if root.winfo_exists() and root.winfo_id():
                hwnd = win32gui.GetParent(root.winfo_id())
                
                # Verify the handle is valid before using it
                if hwnd and win32gui.IsWindow(hwnd):
                    style = win32gui.GetWindowLong(hwnd, win32con.GWL_EXSTYLE)
                    style = style & ~win32con.WS_EX_APPWINDOW | win32con.WS_EX_TOOLWINDOW
                    win32gui.SetWindowLong(hwnd, win32con.GWL_EXSTYLE, style)
                    win32gui.SetWindowPos(hwnd, 0, 0, 0, 0, 0,
                                        win32con.SWP_NOMOVE | win32con.SWP_NOSIZE | 
                                        win32con.SWP_NOZORDER | win32con.SWP_FRAMECHANGED)
        except Exception as e:
            logging.warning(f"Could not hide root window: {e}")
            # Non-critical error, application will still function
    
    # Use blitting and a higher interval for smoother animations
    anim = animation.FuncAnimation(
        app.fig, 
        app.update_plot, 
        interval=200,  # Increase interval to reduce flashing (was 100ms)
        cache_frame_data=False,
        blit=False  # Setting to False can sometimes help with flashing issues
    )
    
    app.ani = anim
    root._anim_ref = anim
    _animations.append(anim)
    anim._fig = app.fig
    
    # Apply matplotlib backend configurations
    from matplotlib import rcParams
    rcParams['figure.autolayout'] = True  # Use tight layout to avoid resizing
    
    root.mainloop()

if __name__ == "__main__":
    main()