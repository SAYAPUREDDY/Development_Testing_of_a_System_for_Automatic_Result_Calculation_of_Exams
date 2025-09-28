import zipfile
from pathlib import Path
from PIL import Image, ImageDraw
from typing import List, Dict, Tuple, Any, Optional
import os

class PdfUtils:
    def save_annotated_pdf(pdf_folder: Path, out_name: str = "annotated_all.pdf") -> Optional[str]:
        imgs: List[Image.Image] = []
        page_folders = sorted([p for p in pdf_folder.glob("image_*") if p.is_dir()],
                            key=lambda p: int(p.name.split("_")[1]))
        for pf in page_folders:
            det = pf / "detected.jpg"
            if det.exists():
                imgs.append(Image.open(det).convert("RGB"))
        if not imgs:
            return None
        out_path = pdf_folder / out_name
        imgs[0].save(out_path, save_all=True, append_images=imgs[1:])
        return str(out_path)
    def make_zip_from_folder(folder: Path, zip_path: Path) -> str:
        with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
            for root, _, files in os.walk(folder):
                for f in files:
                    file_path = Path(root) / f
                    arcname = file_path.relative_to(folder)
                    zf.write(file_path, arcname)
        return str(zip_path)