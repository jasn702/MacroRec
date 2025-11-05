# Macro Recorder

A powerful and user-friendly macro recording and playback application for Windows. Record and replay keyboard and mouse actions with customizable settings.

![Macro Recorder Screenshot](screenshot.png)

## Features

- Record keyboard and mouse actions
- Customizable recording settings
- System tray integration
- Hotkey support
- Adjustable playback speed
- Repeat and loop playback options
- Save and load macros
- Modern and intuitive GUI

## Installation

### Prerequisites
- Windows 10 or later
- Python 3.8 or later (for development)

### Download
1. Go to the [Releases](https://github.com/jasn702/MacroRec/releases) page
2. Download the latest `Macro Recorder.exe`
3. Run the executable - no installation required

### Development Setup
1. Clone the repository:
   ```bash
   git clone https://github.com/jasn702/MacroRec.git
   cd macro-recorder
   ```

2. Create a virtual environment:
   ```bash
   python -m venv venv
   venv\Scripts\activate
   ```

3. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

4. Run the application:
   ```bash
   python macro_recorder_gui.py
   ```

## Usage

### Recording Macros
1. Click "Start Recording" or press F7
2. Perform your actions (keyboard and mouse)
3. Click "Stop Recording" or press Esc
4. Save your macro with a descriptive name

### Playing Macros
1. Select a macro from the list
2. Click "Play Selected" or press F8
3. Adjust playback settings in the Settings tab:
   - Playback speed
   - Repeat count
   - Repeat delay
   - Loop playback

### Settings
- **Recording Settings**
  - Record keyboard/mouse
  - Mouse movement threshold
  - System tray behavior
- **Hotkeys**
  - Start recording: F7
  - Stop recording: Esc
  - Play macro: F8
- **Playback Settings**
  - Playback speed
  - Repeat count
  - Repeat delay
  - Loop playback

## Building the Executable

To build the standalone executable:

```bash
pyinstaller macro_recorder_gui.spec
```

The executable will be created in the `dist` directory.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## Acknowledgments

- [PyInstaller](https://www.pyinstaller.org/) for creating standalone executables
- [pynput](https://github.com/moses-palmer/pynput) for input device control

- [ttkthemes](https://github.com/RedFantom/ttkthemes) for modern GUI themes 

