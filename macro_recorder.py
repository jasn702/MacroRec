import keyboard
import mouse
import json
import time
from pynput.mouse import Controller as MouseController
from pynput.keyboard import Controller as KeyboardController
from datetime import datetime
import os

class MacroRecorder:
    def __init__(self):
        self.recording = False
        self.events = []
        self.mouse = MouseController()
        self.keyboard = KeyboardController()
        self.start_time = None
        self.macro_dir = "macros"
        
        # Create macros directory if it doesn't exist
        if not os.path.exists(self.macro_dir):
            os.makedirs(self.macro_dir)

    def start_recording(self):
        print("Recording started... Press 'Esc' to stop recording.")
        self.recording = True
        self.events = []
        self.start_time = time.time()
        
        # Start listening to events
        keyboard.hook(self.on_keyboard_event)
        mouse.hook(self.on_mouse_event)

    def stop_recording(self):
        self.recording = False
        keyboard.unhook_all()
        mouse.unhook_all()
        print("Recording stopped.")
        return self.events

    def on_keyboard_event(self, event):
        if event.event_type == 'down':
            if event.name == 'esc' and self.recording:
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
        if hasattr(event, 'button'):
            self.events.append({
                'type': 'mouse',
                'event': event.event_type,
                'button': event.button,
                'position': (event.x, event.y),
                'time': current_time
            })
        elif hasattr(event, 'x'):
            self.events.append({
                'type': 'mouse',
                'event': 'move',
                'position': (event.x, event.y),
                'time': current_time
            })

    def save_macro(self, name=None):
        if not name:
            name = datetime.now().strftime("%Y%m%d_%H%M%S")
        
        filename = os.path.join(self.macro_dir, f"{name}.json")
        with open(filename, 'w') as f:
            json.dump(self.events, f)
        print(f"Macro saved as: {filename}")
        return filename

    def load_macro(self, filename):
        with open(filename, 'r') as f:
            self.events = json.load(f)
        return self.events

    def play_macro(self, events=None):
        if events is None:
            events = self.events
        
        if not events:
            print("No events to play!")
            return

        print("Playing macro in 3 seconds...")
        time.sleep(3)
        
        last_time = 0
        for event in events:
            # Wait for the appropriate time
            time_to_wait = event['time'] - last_time
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
            
            last_time = event['time']

def main():
    recorder = MacroRecorder()
    print("Macro Recorder")
    print("Press 'F7' to start recording")
    print("Press 'F8' to play last recorded macro")
    print("Press 'F9' to save last recorded macro")
    print("Press 'Ctrl+Q' to quit")

    def on_f7():
        recorder.start_recording()

    def on_f8():
        if recorder.events:
            recorder.play_macro()
        else:
            print("No macro recorded yet!")

    def on_f9():
        if recorder.events:
            recorder.save_macro()
        else:
            print("No macro recorded yet!")

    keyboard.add_hotkey('f7', on_f7)
    keyboard.add_hotkey('f8', on_f8)
    keyboard.add_hotkey('f9', on_f9)

    keyboard.wait('ctrl+q')
    print("Program terminated.")

if __name__ == "__main__":
    main() 