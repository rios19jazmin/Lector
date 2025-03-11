from PIL import Image
from pyzbar.pyzbar import decode

def extract_barcodes_from_tiff(tiff_file_path):
    page_barcodes = []

    try:
        img = Image.open(tiff_file_path)
        num_pages = img.n_frames

        for page_idx in range(num_pages):
            img.seek(page_idx)
            barcodes = decode(img)

            if barcodes:
                for barcode in barcodes:
                    barcode_data = '0' + barcode.data.decode('utf-8')
                    page_barcodes.append((page_idx + 1, barcode_data))
            else:
                page_barcodes.append((page_idx + 1, "No contiene código de barras"))

    except FileNotFoundError:
        raise FileNotFoundError(f"No se encontró el archivo {tiff_file_path}")
    except Exception as e:
        raise Exception(f"Error al procesar el archivo TIFF: {e}")

    return page_barcodes

