# docx2mpy2
Docx to MP3 Converter

# Word to MP3 Converter

This is a simple desktop application built with Python and Tkinter that allows you to convert Microsoft Word documents (`.docx`) into MP3 audio files using text-to-speech (TTS) technology.

The app lets you choose different voices (if available on your system), adjust speech rate and volume, and then save the audio version of your document.

---

## 🧰 features

- Load `.docx` files and extract text
- Convert text to speech using `pyttsx3`
- Save the audio directly as an `.mp3` file
- Adjustable voice, speed, and volume settings
- Progress bar and status updates
- Cross-platform (Windows, macOS, Linux)

---

## 🖥️ requirements

- Python 3.7+
- A compatible TTS engine installed on your system

### dependencies

You can install the required Python packages using:

```bash
pip install python-docx pyttsx3 tqdm
```

> Note: On some systems, `pyttsx3` might require additional voices or TTS support. On Windows, it uses the built-in SAPI5. On macOS, it uses NSSpeechSynthesizer. On Linux, it may depend on `espeak` or similar.

---

## 🚀 how to use

1. Run the script:

```bash
python your_script_name.py
```

2. The GUI will open.

3. Click **"Buscar"** to select a `.docx` file.

4. Choose your preferred **voice** from the dropdown (or leave as "default").

5. Adjust the **speech rate** and **volume** sliders.

6. Click **"Convertir a MP3"** to begin the conversion.

7. Wait for the process to complete. The MP3 file will be saved in the same folder as your document.

---

## 📁 project structure

```
word_to_mp3_converter/
│
├── main.py               # Main application script
└── README.md             # Project documentation
```

---

## 🧩 limitations and notes

- This tool only works with `.docx` files (not `.doc`, `.pdf`, etc.)
- Long documents might take longer to process.
- It does not support combining multiple voices or adding effects.
- The generated audio file is saved as a single `.mp3` file.

---

## 📸 screenshots

![GUI Screenshot](#) *(Optional: Add image here)*

---

## 🤖 acknowledgments

- [python-docx](https://python-docx.readthedocs.io/)
- [pyttsx3](https://pyttsx3.readthedocs.io/)
- Tkinter GUI module (standard in Python)

---

## 📃 license

This project is provided under the [MIT License](https://opensource.org/licenses/MIT). Feel free to use and modify it.

---

## ✍️ author

Developed by Sergio A. Hernandez — feel free to fork, share or improve.
