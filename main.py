import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pdf2image import convert_from_path
import pytesseract
import threading
import queue
import os
import webbrowser
import subprocess
from langdetect import detect
from PyPDF2 import PdfMerger
import io

# ================= CONFIGURACIÓN =================
TESSERACT_PATH = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
POPPLER_PATH = r"C:\poppler\Library\bin"

pytesseract.pytesseract.tesseract_cmd = TESSERACT_PATH

DPI = 200                       # OCR más rápido
ESCALA = 2                      # Reducir imagen
CONFIG_RAPIDO = "--oem 1 --psm 3"

IDIOMAS_MAPA = {
    "es": "spa",
    "en": "eng",
    "fr": "fra",
    "de": "deu",
    "it": "ita",
    "pt": "por"
}
# =================================================

stop_process = False
q = queue.Queue()
pdf_salida = ""

# ---------- GUI ----------
def cargar_pdf():
    pdf = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
    if pdf:
        iniciar_ocr(pdf)

def iniciar_ocr(pdf):
    global stop_process, pdf_salida
    stop_process = False
    progress["value"] = 0
    estado.set("Preparando OCR...")
    pdf_salida = pdf.replace(".pdf", "_OCR_EDITABLE.pdf")

    threading.Thread(
        target=ocr_worker,
        args=(pdf,),
        daemon=True
    ).start()

def detener():
    global stop_process
    stop_process = True

# ---------- OCR ----------
def detectar_idioma(img):
    try:
        texto = pytesseract.image_to_string(
            img,
            lang="eng",
            config=CONFIG_RAPIDO
        )
        code = detect(texto)
        return IDIOMAS_MAPA.get(code, "eng")
    except:
        return "eng"

def ocr_worker(pdf):
    try:
        images = convert_from_path(
            pdf,
            poppler_path=POPPLER_PATH,
            dpi=DPI
        )

        total = len(images)
        q.put(("max", total))

        # Detectar idioma SOLO una vez
        img_test = images[0].resize(
            (images[0].width // ESCALA, images[0].height // ESCALA)
        )
        idioma = detectar_idioma(img_test)

        merger = PdfMerger()

        for i, img in enumerate(images, start=1):
            if stop_process:
                q.put(("stop", None))
                merger.close()
                return

            q.put(("page", i, total))

            img = img.resize(
                (img.width // ESCALA, img.height // ESCALA)
            )

            pdf_bytes = pytesseract.image_to_pdf_or_hocr(
                img,
                lang=idioma,
                config=CONFIG_RAPIDO,
                extension="pdf"
            )

            merger.append(io.BytesIO(pdf_bytes))
            q.put(("progress", i))

        with open(pdf_salida, "wb") as f:
            merger.write(f)

        merger.close()
        q.put(("done", pdf_salida))

    except Exception as e:
        q.put(("error", str(e)))

# ---------- ACTUALIZAR GUI ----------
def actualizar():
    try:
        while True:
            data = q.get_nowait()
            tipo = data[0]

            if tipo == "max":
                progress["maximum"] = data[1]

            elif tipo == "progress":
                progress["value"] = data[1]

            elif tipo == "page":
                estado.set(f"Página {data[1]} de {data[2]}")

            elif tipo == "done":
                estado.set("Completado")
                messagebox.showinfo("Listo", "PDF OCR editable creado")
                abrir_en_chrome(data[1])

            elif tipo == "stop":
                estado.set("Proceso detenido")
                messagebox.showwarning("Detenido", "OCR detenido")

            elif tipo == "error":
                estado.set("Error")
                messagebox.showerror("Error", data[1])

    except queue.Empty:
        pass

    root.after(50, actualizar)

# ---------- UTIL ----------
def abrir_en_chrome(path):
    try:
        webbrowser.get("chrome").open_new(path)
    except:
        chrome = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
        if os.path.exists(chrome):
            subprocess.Popen([chrome, path])

# ---------- INTERFAZ ----------
root = tk.Tk()
root.title("OCR PDF Rápido → Editable")
root.geometry("480x260")

frame = tk.Frame(root)
frame.pack(pady=15)

tk.Button(frame, text="Cargar PDF", width=24, command=cargar_pdf).pack(pady=5)
tk.Button(frame, text="Detener", width=24, command=detener).pack(pady=5)

progress = ttk.Progressbar(root, length=400, mode="determinate")
progress.pack(pady=10)

estado = tk.StringVar(value="Esperando PDF...")
tk.Label(root, textvariable=estado, font=("Arial", 10)).pack(pady=5)

root.after(50, actualizar)
root.mainloop()
