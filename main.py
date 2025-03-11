import os
import tkinter as tk
from tkinter import filedialog
from ttkbootstrap import Style
from ttkbootstrap.dialogs import Messagebox
from ttkbootstrap.widgets import Frame, Button, Treeview, Progressbar
from lector import extract_barcodes_from_tiff
from openpyxl import Workbook
from PIL import Image
import logging
from PIL import Image, ImageTk  # Asegúrate de importar tanto Image como ImageTk



class TiffToPdfApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Decodificador DIA - 0 ")

        # Cargar el icono con Pillow
        imagen = Image.open("barcode.png")  # Reemplaza con la ruta correcta de tu imagen
        icono = ImageTk.PhotoImage(imagen)  # Usamos ImageTk para convertir la imagen

        # Establecer el icono de la ventana
        self.root.iconphoto(True, icono)

        # Configuración adicional
        self.style = Style(theme="darkly")
        self.is_running = False
        self.folder_path = None

        self.create_widgets()


    def create_widgets(self):
        # Frame principal
        self.main_frame = Frame(self.root, padding=10)
        self.main_frame.pack(fill="both", expand=True)

        # Etiqueta y botón para seleccionar carpeta
        self.label = tk.Label(self.main_frame, text="Selecciona una carpeta con archivos TIFF:")
        self.label.grid(row=0, column=0, columnspan=2, pady=10, sticky="w")

        self.browse_button = Button(
            self.main_frame, text="Seleccionar carpeta", command=self.browse_folder
        )
        self.browse_button.grid(row=1, column=0, padx=5, pady=5)

        self.cancel_button = Button(
            self.main_frame, text="Cancelar", command=self.cancel_task, state=tk.DISABLED
        )
        self.cancel_button.grid(row=1, column=1, padx=5, pady=5)

        self.clear_button = Button(
            self.main_frame, text="Limpiar tabla", command=self.clear_table
        )
        self.clear_button.grid(row=1, column=2, padx=5, pady=5)

        # Tabla de resultados
        self.barcode_table = Treeview(
            self.main_frame,
            columns=("Folio", "Página", "Código de Barras"),
            show="headings",
            height=15,
            bootstyle="info",
        )
        self.barcode_table.heading("Folio", text="Folio")
        self.barcode_table.heading("Página", text="Página")
        self.barcode_table.heading("Código de Barras", text="Código de Barras")
        self.barcode_table.grid(row=2, column=0, columnspan=3, padx=10, pady=10)

        # Barra de progreso
        self.progress_bar = Progressbar(
            self.main_frame, orient="horizontal", length=400, mode="determinate"
        )
        self.progress_bar.grid(row=3, column=0, columnspan=3, pady=10)

        # Botón para guardar el Excel
        self.save_button = Button(
            self.main_frame, text="Exportar Excel", command=self.save_to_file, state=tk.DISABLED
        )
        self.save_button.grid(row=4, column=1, pady=10)

        # Etiqueta de estado
        self.status_label = tk.Label(
            self.main_frame, text="", fg="blue", font=("Helvetica", 12, "bold")
        )
        self.status_label.grid(row=5, column=0, columnspan=3)

        # Firma
        self.signature_frame = Frame(self.root, padding=10)
        self.signature_frame.pack(side="bottom", fill="x")

        self.signature_label = tk.Label(
            self.signature_frame, text="BY RITO", font=("Helvetica", 10, "italic"), fg="gray"
        )
        self.signature_label.pack()

    def browse_folder(self):
        self.folder_path = filedialog.askdirectory()
        if self.folder_path:
            self.process_folder(self.folder_path)

    def process_folder(self, folder_path):
        self.is_running = True
        self.cancel_button.config(state=tk.NORMAL)
        self.status_label.config(
            text="Procesando archivos... Esto puede tardar unos minutos.", fg="orange")
        self.root.update_idletasks()

        all_barcodes = {}
        try:
            filenames = [f for f in os.listdir(folder_path) if f.lower().endswith(('.tif', '.tiff'))]

            if not filenames:
                self.show_message(
                    "Información", "No se encontraron archivos TIFF en la carpeta seleccionada.")
                self.status_label.config(
                    text="No se encontraron archivos TIFF.", fg="red")
                self.is_running = False
                return

            # Calcular el total de páginas
            total_pages = 0
            for filename in filenames:
                file_path = os.path.join(folder_path, filename)
                with Image.open(file_path) as img:
                    total_pages += img.n_frames  # `n_frames` es el número de páginas en el TIFF

            self.progress_bar["maximum"] = total_pages
            processed_pages = 0

            for filename in filenames:
                if not self.is_running:
                    break

                file_path = os.path.join(folder_path, filename)
                barcodes = extract_barcodes_from_tiff(file_path)
                all_barcodes[filename] = barcodes

                # Procesar cada página
                with Image.open(file_path) as img:
                    for page_idx in range(img.n_frames):
                        if not self.is_running:
                            break

                        img.seek(page_idx)  # Mover a la página específica
                        self.convert_tiff_to_pdf(file_path, barcodes, page_idx)

                        # Actualizar barra de progreso y estado
                        processed_pages += 1
                        self.progress_bar["value"] = processed_pages
                        self.update_status(processed_pages, total_pages)

            self.populate_table(all_barcodes)

            if self.is_running:
                self.show_message(
                    "Éxito", "El proceso ha finalizado y los PDFs se han generado correctamente.")
                self.status_label.config(
                    text="Proceso completado con éxito.", fg="green")
        except Exception as e:
            logging.error("Error al procesar los archivos: %s", e)
            self.show_message("Error", f"Ocurrió un error: {e}")
            self.status_label.config(text="Error en el procesamiento.", fg="red")
        finally:
            self.is_running = False
            self.cancel_button.config(state=tk.DISABLED)

    def update_status(self, processed_pages, total_pages):
        status_text = f"Procesando páginas... ({processed_pages}/{total_pages})"
        self.status_label.config(text=status_text)
        self.root.update_idletasks()

    def convert_tiff_to_pdf(self, tiff_path, barcodes, page_idx):
        try:
            img = Image.open(tiff_path)
            base_filename = os.path.splitext(os.path.basename(tiff_path))[0]

            img.seek(page_idx)
            page_img = img.copy()

            barcode_data = barcodes.get(page_idx, "sin_codigo")  # Obtener datos del código para la página actual
            pdf_filename = f"{base_filename}_page_{page_idx + 1}_{barcode_data}.pdf"
            pdf_path = os.path.join(os.path.dirname(tiff_path), pdf_filename)

            page_img.save(pdf_path, "PDF", resolution=100.0)
        except Exception as e:
            logging.error("Error al convertir TIFF a PDF en la página %s: %s", page_idx + 1, e)

    def populate_table(self, barcodes):
        for row in self.barcode_table.get_children():
            self.barcode_table.delete(row)

        for filename, barcode_data in barcodes.items():
            for page, barcode in barcode_data:
                folio = os.path.splitext(os.path.basename(filename))[0]
                self.barcode_table.insert("", "end", values=(folio, page, barcode))

        self.save_button.config(state=tk.NORMAL)

    def save_to_file(self):
        if self.folder_path and self.barcode_table.get_children():
            folder_name = os.path.basename(self.folder_path)
            filename = os.path.join(self.folder_path, f"{folder_name}_folios.xlsx")

            try:
                wb = Workbook()
                ws = wb.active
                ws.title = "Códigos de Barras"
                ws.append(["Folio", "Página", "Código de Barras"])

                for child in self.barcode_table.get_children():
                    item = self.barcode_table.item(child)
                    folio, page, barcode = item["values"]
                    ws.append([folio, page, barcode])

                wb.save(filename)
                self.show_message("Información", f"El archivo Excel se guardó en:\n{filename}.")
            except Exception as e:
                self.show_message("Error", f"No se pudo guardar el archivo Excel: {e}")
        else:
            self.show_message("Advertencia", "No hay datos disponibles para guardar.")

    def clear_table(self):
        for row in self.barcode_table.get_children():
            self.barcode_table.delete(row)
        self.save_button.config(state=tk.DISABLED)
        self.status_label.config(text="")

    def show_message(self, title, message):
        Messagebox.show_info(title=title, message=message)

    def cancel_task(self):
        self.is_running = False
        self.cancel_button.config(state=tk.DISABLED)


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO)
    root = tk.Tk()
    app = TiffToPdfApp(root)
    root.mainloop()
