import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from docx import Document
import pyttsx3
from tqdm import tqdm

class WordToMP3Converter:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Conversor de Word a MP3")
        self.root.geometry("500x300")
        
        self.input_file = tk.StringVar()
        self.output_file = tk.StringVar()
        self.voice_selected = tk.StringVar(value="default")
        self.speech_rate = tk.IntVar(value=150)
        self.volume_level = tk.DoubleVar(value=0.9)
        
        self.setup_ui()
    
    def setup_ui(self):
        main_frame = tk.Frame(self.root, padx=20, pady=20)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        tk.Label(main_frame, text="Archivo Word (.docx):").grid(row=0, column=0, sticky="w")
        tk.Entry(main_frame, textvariable=self.input_file, width=40).grid(row=0, column=1)
        tk.Button(main_frame, text="Buscar", command=self.select_file).grid(row=0, column=2)
        
        tk.Label(main_frame, text="Voz:").grid(row=1, column=0, sticky="w")
        tk.OptionMenu(main_frame, self.voice_selected, *self.get_available_voices()).grid(row=1, column=1, sticky="w")
        
        tk.Label(main_frame, text="Velocidad (palabras/min):").grid(row=2, column=0, sticky="w")
        tk.Scale(main_frame, from_=50, to=300, variable=self.speech_rate, orient=tk.HORIZONTAL).grid(row=2, column=1, sticky="ew")
        
        tk.Label(main_frame, text="Volumen:").grid(row=3, column=0, sticky="w")
        tk.Scale(main_frame, from_=0.1, to=1.0, resolution=0.1, variable=self.volume_level, orient=tk.HORIZONTAL).grid(row=3, column=1, sticky="ew")
        
        tk.Button(main_frame, text="Convertir a MP3", command=self.convert_to_mp3, bg="#4CAF50", fg="white").grid(row=4, column=1, pady=10)
        
        self.progress = ttk.Progressbar(main_frame, orient="horizontal", length=300, mode="determinate")
        self.progress.grid(row=5, column=0, columnspan=3, pady=10)
        
        self.status_label = tk.Label(main_frame, text="", fg="blue")
        self.status_label.grid(row=6, column=0, columnspan=3)
    
    def select_file(self):
        file_path = filedialog.askopenfilename(
            title="Seleccionar archivo Word",
            filetypes=[("Documentos Word", "*.docx"), ("Todos los archivos", "*.*")]
        )
        if file_path:
            self.input_file.set(file_path)
            output_path = os.path.splitext(file_path)[0] + ".mp3"
            self.output_file.set(output_path)
    
    def get_available_voices(self):
        engine = pyttsx3.init()
        voices = engine.getProperty('voices')
        return ["default"] + [voice.name for voice in voices]
    
    def update_progress(self, value):
        self.progress['value'] = value
        self.root.update_idletasks()
    
    def convert_to_mp3(self):
        input_file = self.input_file.get()
        if not input_file:
            messagebox.showerror("Error", "¡Selecciona un archivo Word primero!")
            return
        
        try:
            self.status_label.config(text="Procesando...", fg="blue")
            self.update_progress(0)
            
            doc = Document(input_file)
            paragraphs = [p.text for p in doc.paragraphs if p.text.strip()]
            
            engine = pyttsx3.init()
            
            if self.voice_selected.get() != "default":
                voices = engine.getProperty('voices')
                for voice in voices:
                    if voice.name == self.voice_selected.get():
                        engine.setProperty('voice', voice.id)
                        break
            
            engine.setProperty('rate', self.speech_rate.get())
            engine.setProperty('volume', self.volume_level.get())
            
            # Guardar directamente a MP3 (sin combinar segmentos)
            output_path = self.output_file.get()
            engine.save_to_file('\n'.join(paragraphs), output_path)
            engine.runAndWait()
            
            self.status_label.config(text="¡Conversión completada!", fg="green")
            messagebox.showinfo("Éxito", f"Archivo MP3 generado:\n{output_path}")
            self.update_progress(100)
            
        except Exception as e:
            messagebox.showerror("Error", f"Error durante la conversión:\n{str(e)}")
            self.status_label.config(text="Error", fg="red")
    
    def run(self):
        self.root.mainloop()

if __name__ == "__main__":
    app = WordToMP3Converter()
    app.run()