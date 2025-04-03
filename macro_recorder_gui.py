import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from ttkthemes import ThemedTk
import json
import os
import keyboard
import mouse
import time
from pynput.mouse import Controller as MouseController
from pynput.keyboard import Controller as KeyboardController
from datetime import datetime
import threading
import pystray
from PIL import Image, ImageDraw
import win32gui
import win32con

class MacroRecorderGUI:
    def __init__(self):
        self.root = ThemedTk(theme="arc")
        self.root.title("Macro Recorder")
        self.root.geometry("480x640")
        
        # Store the window state
        self.is_minimized = False
        
        # Initialize recorder variables
        self.recording = False
        self.events = []
        self.mouse = MouseController()
        self.keyboard = KeyboardController()
        self.start_time = None
        self.macro_dir = "macros"
        
        # System tray icon
        self.icon = None
        self.setup_system_tray()
        
        # Default hotkeys
        self.default_hotkeys = {
            "start_recording": "f7",
            "stop_recording": "esc",
            "play_macro": "f8"
        }
        
        # Default recording settings
        self.default_recording_settings = {
            "record_keyboard": True,
            "record_mouse": True,
            "record_mouse_movement": True,
            "record_mouse_clicks": True,
            "record_mouse_scroll": True,
            "minimum_mouse_movement": 5,  # Minimum pixels between recorded mouse movements
            "minimize_to_tray": True,  # New setting for system tray behavior
            "loop_playback": False  # New setting for loop playback
        }
        
        # Load or create settings
        self.load_settings()
        
        # Create macros directory if it doesn't exist
        if not os.path.exists(self.macro_dir):
            os.makedirs(self.macro_dir)
        
        self.create_gui()
        self.load_macro_list()
        
        # Bind window events
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.root.bind("<Configure>", self.on_window_configure)
        
    def on_window_configure(self, event=None):
        # Check if this is actually a minimize event
        if event and event.widget == self.root:
            current_state = win32gui.GetWindowState(self.root.winfo_id())
            is_minimized = current_state == win32con.SW_SHOWMINIMIZED
            
            if is_minimized and not self.is_minimized:
                self.is_minimized = True
                if self.settings["recording"].get("minimize_to_tray", True):
                    self.hide_window()
            elif not is_minimized:
                self.is_minimized = False
    
    def create_tray_icon(self):
        # Create a simple square icon
        icon_size = 64
        image = Image.new('RGB', (icon_size, icon_size), color='white')
        draw = ImageDraw.Draw(image)
        
        # Draw a red circle when recording, otherwise blue
        color = 'red' if self.recording else 'blue'
        margin = 4
        draw.ellipse([margin, margin, icon_size - margin, icon_size - margin], fill=color)
        
        return image
    
    def setup_system_tray(self):
        menu = (
            pystray.MenuItem("Show", self.show_window),
            pystray.MenuItem("Start Recording", self.start_recording),
            pystray.MenuItem("Stop Recording", self.stop_recording),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem("Exit", self.quit_app)
        )
        
        self.icon = pystray.Icon("macro_recorder", self.create_tray_icon(), "Macro Recorder", menu)
        
    def update_tray_icon(self):
        if self.icon:
            self.icon.icon = self.create_tray_icon()
    
    def show_window(self, icon=None, item=None):
        self.root.deiconify()
        self.root.lift()
        self.root.focus_force()
    
    def hide_window(self):
        self.root.withdraw()
    
    def on_closing(self):
        if messagebox.askokcancel("Quit", "Do you want to quit Macro Recorder?"):
            self.quit_app()
    
    def quit_app(self, icon=None):
        self.icon.stop()
        self.root.quit()
    
    def load_settings(self):
        try:
            with open("settings.json", "r") as f:
                self.settings = json.load(f)
                # Add any missing recording settings
                if "recording" not in self.settings:
                    self.settings["recording"] = self.default_recording_settings
                for key, value in self.default_recording_settings.items():
                    if key not in self.settings["recording"]:
                        self.settings["recording"][key] = value
        except FileNotFoundError:
            self.settings = {
                "hotkeys": self.default_hotkeys.copy(),
                "playback_speed": 1.0,
                "repeat_count": 1,
                "repeat_delay": 0.0,
                "recording": self.default_recording_settings.copy()
            }
            self.save_settings()
    
    def save_settings(self):
        with open("settings.json", "w") as f:
            json.dump(self.settings, f, indent=4)
    
    def create_gui(self):
        # Create notebook for tabs
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(expand=True, fill='both', padx=5, pady=5)
        
        # Main tab
        self.main_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.main_frame, text='Macros')
        
        # Status frame
        status_frame = ttk.LabelFrame(self.main_frame, text="Status")
        status_frame.pack(fill='x', padx=5, pady=5)
        
        self.status_label = ttk.Label(status_frame, text="Ready")
        self.status_label.pack(padx=5, pady=5)
        
        # Control frame
        control_frame = ttk.LabelFrame(self.main_frame, text="Controls")
        control_frame.pack(fill='x', padx=5, pady=5)
        
        ttk.Button(control_frame, text="Start Recording", command=self.start_recording).pack(side='left', padx=5, pady=5)
        ttk.Button(control_frame, text="Stop Recording", command=self.stop_recording).pack(side='left', padx=5, pady=5)
        ttk.Button(control_frame, text="Play Selected", command=self.play_selected_macro).pack(side='left', padx=5, pady=5)
        ttk.Button(control_frame, text="Minimize to Tray", command=self.hide_window).pack(side='right', padx=5, pady=5)
        
        # Macro list frame
        list_frame = ttk.LabelFrame(self.main_frame, text="Saved Macros")
        list_frame.pack(fill='both', expand=True, padx=5, pady=5)
        
        # Add scrollbar to macro list
        scrollbar = ttk.Scrollbar(list_frame)
        scrollbar.pack(side='right', fill='y')
        
        self.macro_list = tk.Listbox(list_frame, yscrollcommand=scrollbar.set)
        self.macro_list.pack(fill='both', expand=True, padx=5, pady=5)
        scrollbar.config(command=self.macro_list.yview)
        
        # Macro list buttons
        list_button_frame = ttk.Frame(list_frame)
        list_button_frame.pack(fill='x', padx=5, pady=5)
        
        ttk.Button(list_button_frame, text="Rename", command=self.rename_macro).pack(side='left', padx=5)
        ttk.Button(list_button_frame, text="Delete", command=self.delete_macro).pack(side='left', padx=5)
        ttk.Button(list_button_frame, text="Refresh", command=self.load_macro_list).pack(side='left', padx=5)
        
        # Settings tab
        self.settings_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.settings_frame, text='Settings')
        
        # Initialize recording_vars dictionary
        self.recording_vars = {}
        
        # Create a canvas with scrollbar for settings
        settings_canvas = tk.Canvas(self.settings_frame)
        scrollbar = ttk.Scrollbar(self.settings_frame, orient="vertical", command=settings_canvas.yview)
        scrollable_frame = ttk.Frame(settings_canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: settings_canvas.configure(scrollregion=settings_canvas.bbox("all"))
        )
        
        settings_canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        settings_canvas.configure(yscrollcommand=scrollbar.set)
        
        settings_canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Recording settings frame
        recording_frame = ttk.LabelFrame(scrollable_frame, text="Recording Settings")
        recording_frame.grid(row=0, column=0, padx=5, pady=5, sticky="ew")
        
        # Input device settings
        device_frame = ttk.Frame(recording_frame)
        device_frame.pack(fill='x', padx=5, pady=5)
        
        # Create two columns for checkboxes
        left_column = ttk.Frame(device_frame)
        left_column.pack(side='left', fill='x', expand=True, padx=5)
        right_column = ttk.Frame(device_frame)
        right_column.pack(side='left', fill='x', expand=True, padx=5)
        
        # Left column checkboxes
        self.recording_vars["record_keyboard"] = tk.BooleanVar(value=self.settings["recording"]["record_keyboard"])
        ttk.Checkbutton(left_column, text="Record Keyboard", 
                       variable=self.recording_vars["record_keyboard"]).pack(anchor='w')
        
        self.recording_vars["record_mouse"] = tk.BooleanVar(value=self.settings["recording"]["record_mouse"])
        ttk.Checkbutton(left_column, text="Record Mouse", 
                       variable=self.recording_vars["record_mouse"]).pack(anchor='w')
        
        # Right column checkboxes
        self.recording_vars["record_mouse_movement"] = tk.BooleanVar(value=self.settings["recording"]["record_mouse_movement"])
        ttk.Checkbutton(right_column, text="Record Mouse Movement", 
                       variable=self.recording_vars["record_mouse_movement"]).pack(anchor='w')
        
        self.recording_vars["record_mouse_clicks"] = tk.BooleanVar(value=self.settings["recording"]["record_mouse_clicks"])
        ttk.Checkbutton(right_column, text="Record Mouse Clicks", 
                       variable=self.recording_vars["record_mouse_clicks"]).pack(anchor='w')
        
        self.recording_vars["record_mouse_scroll"] = tk.BooleanVar(value=self.settings["recording"]["record_mouse_scroll"])
        ttk.Checkbutton(right_column, text="Record Mouse Scroll", 
                       variable=self.recording_vars["record_mouse_scroll"]).pack(anchor='w')
        
        # Mouse movement threshold
        threshold_frame = ttk.Frame(recording_frame)
        threshold_frame.pack(fill='x', padx=5, pady=5)
        
        ttk.Label(threshold_frame, text="Min. Movement Threshold:").pack(side='left')
        self.recording_vars["minimum_mouse_movement"] = tk.StringVar(value=str(self.settings["recording"]["minimum_mouse_movement"]))
        ttk.Entry(threshold_frame, textvariable=self.recording_vars["minimum_mouse_movement"], width=5).pack(side='left', padx=5)
        
        # System tray settings
        self.recording_vars["minimize_to_tray"] = tk.BooleanVar(value=self.settings["recording"].get("minimize_to_tray", True))
        ttk.Checkbutton(recording_frame, text="Minimize to System Tray", 
                       variable=self.recording_vars["minimize_to_tray"]).pack(anchor='w', padx=5, pady=5)
        
        # Hotkeys frame
        hotkeys_frame = ttk.LabelFrame(scrollable_frame, text="Hotkeys")
        hotkeys_frame.grid(row=1, column=0, padx=5, pady=5, sticky="ew")
        
        # Hotkey settings in a grid
        self.hotkey_vars = {}
        for i, (key, default) in enumerate(self.settings["hotkeys"].items()):
            ttk.Label(hotkeys_frame, text=f"{key.replace('_', ' ').title()}:").grid(row=i, column=0, padx=5, pady=2)
            self.hotkey_vars[key] = tk.StringVar(value=default)
            entry = ttk.Entry(hotkeys_frame, textvariable=self.hotkey_vars[key], width=10)
            entry.grid(row=i, column=1, padx=5, pady=2)
            ttk.Button(hotkeys_frame, text="Set", command=lambda k=key: self.set_hotkey(k)).grid(row=i, column=2, padx=5, pady=2)
        
        # Playback settings frame
        playback_frame = ttk.LabelFrame(scrollable_frame, text="Playback Settings")
        playback_frame.grid(row=2, column=0, padx=5, pady=5, sticky="ew")
        
        # Playback settings in a grid
        settings_grid = ttk.Frame(playback_frame)
        settings_grid.pack(fill='x', padx=5, pady=5)
        
        # Row 0: Playback Speed
        ttk.Label(settings_grid, text="Playback Speed:").grid(row=0, column=0, padx=5, pady=2)
        self.playback_speed_var = tk.StringVar(value=str(self.settings.get("playback_speed", 1.0)))
        ttk.Entry(settings_grid, textvariable=self.playback_speed_var, width=5).grid(row=0, column=1, padx=5, pady=2)
        
        # Row 1: Repeat Count
        ttk.Label(settings_grid, text="Repeat Count:").grid(row=1, column=0, padx=5, pady=2)
        self.repeat_count_var = tk.StringVar(value=str(self.settings.get("repeat_count", 1)))
        ttk.Entry(settings_grid, textvariable=self.repeat_count_var, width=5).grid(row=1, column=1, padx=5, pady=2)
        
        # Row 2: Repeat Delay
        ttk.Label(settings_grid, text="Repeat Delay (s):").grid(row=2, column=0, padx=5, pady=2)
        self.repeat_delay_var = tk.StringVar(value=str(self.settings.get("repeat_delay", 0.0)))
        ttk.Entry(settings_grid, textvariable=self.repeat_delay_var, width=5).grid(row=2, column=1, padx=5, pady=2)
        
        # Loop playback checkbox
        self.loop_playback_var = tk.BooleanVar(value=self.settings["recording"].get("loop_playback", False))
        ttk.Checkbutton(playback_frame, text="Loop Playback", 
                       variable=self.loop_playback_var).pack(anchor='w', padx=5, pady=5)
        
        # Save settings button
        ttk.Button(scrollable_frame, text="Save Settings", command=self.save_current_settings).grid(row=3, column=0, pady=10)
        
        # Set up global hotkeys
        self.setup_global_hotkeys()
    
    def setup_global_hotkeys(self):
        for key, hotkey in self.settings["hotkeys"].items():
            if key == "start_recording":
                keyboard.add_hotkey(hotkey, self.start_recording)
            elif key == "stop_recording":
                keyboard.add_hotkey(hotkey, self.stop_recording)
            elif key == "play_macro":
                keyboard.add_hotkey(hotkey, self.play_selected_macro)
    
    def set_hotkey(self, key):
        self.status_label.config(text=f"Press new hotkey for {key}...")
        self.root.update()
        
        new_hotkey = keyboard.read_event(suppress=True).name
        self.hotkey_vars[key].set(new_hotkey)
        self.status_label.config(text="Ready")
    
    def save_current_settings(self):
        # Update settings from GUI variables
        self.settings["playback_speed"] = float(self.playback_speed_var.get())
        self.settings["repeat_count"] = int(self.repeat_count_var.get())
        self.settings["repeat_delay"] = float(self.repeat_delay_var.get())
        
        # Update recording settings
        for key, var in self.recording_vars.items():
            if isinstance(var, tk.BooleanVar):
                self.settings["recording"][key] = var.get()
            else:
                self.settings["recording"][key] = int(var.get())
        
        # Update loop playback setting
        self.settings["recording"]["loop_playback"] = self.loop_playback_var.get()
        
        self.save_settings()
        
        # Refresh hotkeys
        keyboard.unhook_all()
        self.setup_global_hotkeys()
        
        messagebox.showinfo("Success", "Settings saved successfully!")
    
    def load_macro_list(self):
        self.macro_list.delete(0, tk.END)
        if os.path.exists(self.macro_dir):
            for file in os.listdir(self.macro_dir):
                if file.endswith('.json'):
                    self.macro_list.insert(tk.END, file[:-5])
    
    def rename_macro(self):
        selection = self.macro_list.curselection()
        if not selection:
            messagebox.showwarning("Warning", "Please select a macro to rename")
            return
        
        old_name = self.macro_list.get(selection[0])
        new_name = tk.simpledialog.askstring("Rename Macro", "Enter new name:", initialvalue=old_name)
        
        if new_name:
            old_path = os.path.join(self.macro_dir, f"{old_name}.json")
            new_path = os.path.join(self.macro_dir, f"{new_name}.json")
            os.rename(old_path, new_path)
            self.load_macro_list()
    
    def delete_macro(self):
        selection = self.macro_list.curselection()
        if not selection:
            messagebox.showwarning("Warning", "Please select a macro to delete")
            return
        
        if messagebox.askyesno("Confirm Delete", "Are you sure you want to delete this macro?"):
            macro_name = self.macro_list.get(selection[0])
            os.remove(os.path.join(self.macro_dir, f"{macro_name}.json"))
            self.load_macro_list()
    
    def start_recording(self, icon=None):
        if not self.settings["recording"]["record_keyboard"] and not self.settings["recording"]["record_mouse"]:
            messagebox.showwarning("Warning", "Please enable at least one input device in settings")
            return
            
        self.recording = True
        self.events = []
        self.start_time = time.time()
        self.status_label.config(text="Recording...")
        self.last_mouse_pos = None
        
        # Update tray icon to show recording state
        self.update_tray_icon()
        
        # Start recording in a separate thread
        self.record_thread = threading.Thread(target=self._record)
        self.record_thread.daemon = True
        self.record_thread.start()
    
    def _record(self):
        if self.settings["recording"]["record_keyboard"]:
            keyboard.hook(self.on_keyboard_event)
        if self.settings["recording"]["record_mouse"]:
            mouse.hook(self.on_mouse_event)
    
    def stop_recording(self, icon=None):
        if self.recording:
            self.recording = False
            keyboard.unhook_all()
            mouse.unhook_all()
            self.status_label.config(text="Recording stopped")
            
            # Update tray icon to show stopped state
            self.update_tray_icon()
            
            # Ask for macro name and save
            self.show_window()  # Show window to get user input
            name = tk.simpledialog.askstring("Save Macro", "Enter macro name:")
            if name:
                self.save_macro(name)
                self.load_macro_list()
            
            # Restore hotkeys
            self.setup_global_hotkeys()
    
    def on_keyboard_event(self, event):
        if event.event_type == 'down':
            if event.name == self.settings["hotkeys"]["stop_recording"] and self.recording:
                self.stop_recording()
                return
            
            current_time = time.time() - self.start_time
            self.events.append({
                'type': 'keyboard',
                'event': 'press',
                'key': event.name,
                'time': current_time
            })
    
    def on_mouse_event(self, event):
        if not self.recording:
            return
        
        current_time = time.time() - self.start_time
        
        # Handle mouse clicks
        if hasattr(event, 'button') and self.settings["recording"]["record_mouse_clicks"]:
            self.events.append({
                'type': 'mouse',
                'event': event.event_type,
                'button': event.button,
                'position': (event.x, event.y),
                'time': current_time
            })
        
        # Handle mouse movement
        elif hasattr(event, 'x') and self.settings["recording"]["record_mouse_movement"]:
            # Check if movement exceeds minimum threshold
            if self.last_mouse_pos is None:
                self.last_mouse_pos = (event.x, event.y)
            else:
                dx = event.x - self.last_mouse_pos[0]
                dy = event.y - self.last_mouse_pos[1]
                distance = (dx * dx + dy * dy) ** 0.5
                
                if distance >= self.settings["recording"]["minimum_mouse_movement"]:
                    self.events.append({
                        'type': 'mouse',
                        'event': 'move',
                        'position': (event.x, event.y),
                        'time': current_time
                    })
                    self.last_mouse_pos = (event.x, event.y)
        
        # Handle mouse scroll
        elif hasattr(event, 'wheel') and self.settings["recording"]["record_mouse_scroll"]:
            self.events.append({
                'type': 'mouse',
                'event': 'scroll',
                'delta': event.wheel,
                'time': current_time
            })
    
    def save_macro(self, name):
        filename = os.path.join(self.macro_dir, f"{name}.json")
        with open(filename, 'w') as f:
            json.dump(self.events, f)
        self.status_label.config(text=f"Macro saved as: {name}")
    
    def play_selected_macro(self):
        selection = self.macro_list.curselection()
        if not selection:
            messagebox.showwarning("Warning", "Please select a macro to play")
            return
        
        macro_name = self.macro_list.get(selection[0])
        filename = os.path.join(self.macro_dir, f"{macro_name}.json")
        
        try:
            with open(filename, 'r') as f:
                events = json.load(f)
            
            # Start playback in a separate thread
            playback_thread = threading.Thread(target=self._play_macro, args=(events,))
            playback_thread.daemon = True
            playback_thread.start()
            
        except Exception as e:
            messagebox.showerror("Error", f"Error playing macro: {str(e)}")
    
    def _play_macro(self, events):
        self.status_label.config(text="Playing macro...")
        repeat_count = self.settings["repeat_count"]
        repeat_delay = self.settings["repeat_delay"]
        playback_speed = self.settings["playback_speed"]
        loop_playback = self.settings["recording"]["loop_playback"]
        
        while True:  # Loop indefinitely if loop_playback is True
            for _ in range(repeat_count):
                last_time = 0
                for event in events:
                    if not self.recording:  # Don't play events while recording
                        time_to_wait = (event['time'] - last_time) / playback_speed
                        if time_to_wait > 0:
                            time.sleep(time_to_wait)
                        
                        if event['type'] == 'keyboard':
                            if event['event'] == 'press':
                                keyboard.press(event['key'])
                                keyboard.release(event['key'])
                        
                        elif event['type'] == 'mouse':
                            if event['event'] == 'move':
                                self.mouse.position = event['position']
                            elif event['event'] == 'click':
                                mouse.move(event['position'][0], event['position'][1])
                                mouse.click(event['button'])
                            elif event['event'] == 'double click':
                                mouse.move(event['position'][0], event['position'][1])
                                mouse.double_click(event['button'])
                            elif event['event'] == 'scroll':
                                mouse.wheel(event['delta'])
                        
                        last_time = event['time']
                
                if _ < repeat_count - 1:  # Don't delay after the last repeat
                    time.sleep(repeat_delay)
            
            if not loop_playback:
                break  # Exit the loop if loop_playback is False
        
        self.status_label.config(text="Ready")

def main():
    app = MacroRecorderGUI()
    
    # Start system tray icon in a separate thread
    icon_thread = threading.Thread(target=app.icon.run)
    icon_thread.daemon = True
    icon_thread.start()
    
    app.root.mainloop()

if __name__ == "__main__":
    main() 