import os
import fitz  # PyMuPDF
import cv2
from tkinter import filedialog

# Seleccionar el archivo PDF
pdf_file = filedialog.askopenfilename(title="Selecciona el archivo PDF", filetypes=[("Archivos PDF", "*.pdf")])

# Convertir el PDF a imágenes
doc = fitz.open(pdf_file)
image_files = []
for page_index in range(len(doc)):
    page = doc[page_index]
    image = page.get_pixmap()
    image_file = f"page{page_index+1}.png"
    image.save(image_file)
    image_files.append(image_file)

# Combinar las imágenes en un video
fourcc = cv2.VideoWriter_fourcc(*"mp4v")
video_fps = 1
height, width = cv2.imread(image_files[0]).shape[:2]
video_writer = cv2.VideoWriter("hhh.mp4", fourcc, video_fps, (width, height))

for image_file in image_files:
    frame = cv2.imread(image_file)
    video_writer.write(frame)
    os.remove(image_file)

video_writer.release()

# Seleccionar la carpeta de descarga
download_folder = filedialog.askdirectory(title="Selecciona la carpeta de descarga")
video_file = os.path.join(download_folder, "hhh.mp4")
os.rename("hhh.mp4", video_file)
