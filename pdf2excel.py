import os
import re
import tkinter as tk
from tkinter import filedialog

import pandas as pd
import pdfplumber


def pdf_a_excel(pdf_path, excel_path):
    rows = []

    # Abrir el PDF y extraer filas con datos
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                for line in text.split("\n"):
                    if re.match(r"^\d+\s+\d{12}", line):
                        rows.append(line)

    # Procesar filas y separarlas en columnas
    data = []
    for row in rows:
        parts = row.split()

        servicio_id = parts[0]
        guia = parts[1]

        # Extraer nombre del destinatario
        destinatario_parts = []
        i = 2
        while i < len(parts) and not re.match(r"^[A-Z]{3}-[A-Z]{3}$", parts[i]):
            destinatario_parts.append(parts[i])
            i += 1
        destinatario = " ".join(destinatario_parts)

        destino = parts[i] if i < len(parts) else ""
        i += 1

        tipo_servicio = parts[i] if i < len(parts) else ""
        i += 1

        fecha_prod = parts[i] if i < len(parts) else ""
        i += 1

        resto = parts[i:]
        while len(resto) < 9:
            resto.append(None)

        total = resto[8]

        data.append(
            {
                "Guía": guia,
                "Destinatario": destinatario,
                "Fecha Prod": fecha_prod,
                "Total": total,
            }
        )

    # Crear DataFrame y exportar a Excel
    df = pd.DataFrame(data)
    df.to_excel(excel_path, index=False)


if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()

    # Seleccionar archivo PDF
    pdf_path = filedialog.askopenfilename(
        title="Selecciona el archivo PDF", filetypes=[("Archivos PDF", "*.pdf")]
    )
    if not pdf_path:
        exit()

    # Seleccionar carpeta de destino
    carpeta_destino = filedialog.askdirectory(
        title="Selecciona la carpeta para guardar el Excel"
    )
    if not carpeta_destino:
        exit()

    # Buscar último número usado en los archivos factura_envia_X.xlsx
    contador = 1
    while True:
        nombre_archivo = f"factura_envia_{contador}.xlsx"
        ruta_excel = os.path.join(carpeta_destino, nombre_archivo)
        if not os.path.exists(ruta_excel):
            break
        contador += 1

    # Guardar con el siguiente número
    pdf_a_excel(pdf_path, ruta_excel)
    print(f"✅ Archivo guardado como: {ruta_excel}")
