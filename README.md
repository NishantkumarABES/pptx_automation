# PPTX Automation

## Overview

`pptx_automation` is a Python script designed to automate the control of PowerPoint presentations using hand gestures, audio activation, and speech synthesis. This project leverages various libraries such as `mediapipe`, `pyaudio`, `gtts`, and `win32com.client` to create an interactive and dynamic presentation experience.

## Features

- **Hand Gesture Recognition**: Hand gestures control slide navigation.
- **Audio Activation**: Activate the system using audio input (e.g., finger snap).
- **Speech Synthesis**: Convert text to speech for audio feedback.
- **PowerPoint Integration**: Directly interact with PowerPoint presentations.

## Requirements

- Python 3.6+
- `python-pptx`
- `cv2` (OpenCV)
- `numpy`
- `pyaudio`
- `gtts`
- `mediapipe`
- `pygame`
- `PyQt5`
- `pywin32`

## Installation

1. Clone the repository:

```bash
git clone https://github.com/yourusername/pptx_automation.git
cd pptx_automation
```

2. Install the required packages:

```bash
pip install -r requirements.txt
```

3. Install additional dependencies:

```bash
pip install opencv-python numpy pyaudio gtts mediapipe pygame pyqt5 pywin32
```



### Example

Here’s a basic example of how the script works:

1. **Activation Function**:
    - Listens for a specific audio signal (e.g., finger snap) to activate the system.
    
    ```python
    def activation_function():
        p = pyaudio.PyAudio()
        stream = p.open(format=pyaudio.paInt16, channels=1, rate=RATE, input=True, frames_per_buffer=CHUNK)
        print("Listening for Activation...")
        active = False
        while True:
            data = stream.read(CHUNK)
            audio_data = np.frombuffer(data, dtype=np.int16)
            audio_data_abs = np.abs(audio_data)
            avg_energy = np.mean(audio_data_abs)
            if avg_energy > THRESHOLD:
                print("Finger snap detected!")
                speak("System Activated!")
                active = True
                break

        stream.stop_stream()
        stream.close()
        p.terminate()
    ```

2. **Presentation Control**:
    - Opens and controls a PowerPoint presentation using the `win32com.client` library.
    
    ```python
    def pptx_controller(file_path):
        presentation = pptx_app.Presentations.Open(FileName=file_path, ReadOnly=1)
        presentation.SlideShowSettings.run()
        def moveRight():
            time.sleep(1)
            presentation.SlideShowWindow.View.Next()
        def moveLeft():
            time.sleep(1)
            presentation.SlideShowWindow.View.Previous()
    ```

3. **Speech Synthesis**:
    - Converts text to speech using the `gtts` library and plays the audio using `pygame`.
    
    ```python
    def speak(text):
        print("\\n" + text)
        sound = gTTS(text=text, lang='hi')
        audio_data = io.BytesIO()
        sound.write_to_fp(audio_data)
        audio_data.seek(0)
        pygame.mixer.init()
        pygame.mixer.music.load(audio_data)
        pygame.mixer.music.play()
        while pygame.mixer.music.get_busy():
            continue
    ```

## Contributing

Contributions are welcome! Please open an issue or submit a pull request for any changes.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Contact

For any inquiries or issues, please get in touch with Nishant Kumar at nishant543099@gmail.com.

