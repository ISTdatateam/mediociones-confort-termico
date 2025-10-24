"""External resource helpers for Word reports."""
from io import BytesIO
import re
from typing import Optional

import qrcode
import requests
from PIL import ImageOps


def generate_qr_code(url: str, border: int = 10, box_size: int = 2) -> BytesIO:
    qr = qrcode.QRCode(
        version=1,
        error_correction=qrcode.constants.ERROR_CORRECT_L,
        box_size=box_size,
        border=2,
    )
    qr.add_data(url)
    qr.make(fit=True)
    img = qr.make_image(fill_color="black", back_color="white").convert("RGB")
    padded_img = ImageOps.expand(img, border=border, fill="white")
    bio = BytesIO()
    padded_img.save(bio, format="PNG")
    bio.seek(0)
    return bio


def descargar_imagen_gdrive(url_foto: str) -> Optional[BytesIO]:
    match = re.search(r"id=(.*)", url_foto)
    if not match:
        return None
    file_id = match.group(1)
    download_url = f"https://drive.google.com/uc?export=download&id={file_id}"
    try:
        response = requests.get(download_url)
    except Exception:
        return None
    if response.status_code == 200:
        return BytesIO(response.content)
    return None


__all__ = ["generate_qr_code", "descargar_imagen_gdrive"]
