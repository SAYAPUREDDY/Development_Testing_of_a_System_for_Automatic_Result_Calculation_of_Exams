from pathlib import Path
from uuid import uuid4
import cv2
from ultralytics import YOLO

# Load YOLO model once
model = YOLO("C:/Users/91965/SSST/automatic_exam_grading_system/weights/best.pt")

UPLOAD_DIR = Path("uploads")
UPLOAD_DIR.mkdir(parents=True, exist_ok=True)

def process_image_with_yolo(file_path: str, output_dir: Path = None, save_original: bool = True) -> dict:
    try:
        image_path = Path(file_path)
        if not image_path.exists():
            raise FileNotFoundError(f"File not found: {file_path}")

        img = cv2.imread(str(image_path))
        if img is None:
            raise ValueError("Invalid image format or unreadable file.")

        if output_dir is None:
            unique_id = uuid4().hex
            output_dir = Path("outputs") / unique_id
        output_dir.mkdir(parents=True, exist_ok=True)

        cropped_folder = output_dir / "crops"
        cropped_folder.mkdir(parents=True, exist_ok=True)

        detected_path = output_dir / "detected.jpg"

        results = model.predict(img)
        cropped_image_paths = []

        if results and results[0].boxes is not None:
            for i, (box, cls_id) in enumerate(zip(results[0].boxes.xyxy, results[0].boxes.cls)):
                x1, y1, x2, y2 = map(int, box[:4])
                crop_img = img[y1:y2, x1:x2]

                # Save RAW crop (for Excel embedding)
                class_name = model.names[int(cls_id.item())]
                raw_crop_path = cropped_folder / f"{i}_{class_name}_raw.jpg"
                cv2.imwrite(str(raw_crop_path), crop_img)

                # Preprocess for OCR
                gray = cv2.cvtColor(crop_img, cv2.COLOR_BGR2GRAY)
                gray = cv2.GaussianBlur(gray, (3, 3), 0)
                th = cv2.adaptiveThreshold(
                    gray, 255, cv2.ADAPTIVE_THRESH_MEAN_C, cv2.THRESH_BINARY, 15, 11
                )
                h, w = th.shape[:2]
                if max(h, w) < 200:
                    th = cv2.resize(th, (w * 2, h * 2), interpolation=cv2.INTER_CUBIC)
                th = cv2.cvtColor(th, cv2.COLOR_GRAY2BGR)

                # Save PREPROCESSED crop
                crop_path = cropped_folder / f"{i}_{class_name}.jpg"
                cv2.imwrite(str(crop_path), th)

                cropped_image_paths.append({
                    "label": class_name,
                    "path": str(crop_path),        # preprocessed (OCR)
                    "raw_path": str(raw_crop_path), # raw (Excel embedding)
                    "bbox": [x1, y1, x2, y2]
                })

                # Draw bounding boxes
                cv2.rectangle(img, (x1, y1), (x2, y2), (0, 255, 0), 2)

            cv2.imwrite(str(detected_path), img)
        else:
            detected_path = None

        return {
            "original": str(image_path) if save_original else None,
            "detected_image": str(detected_path) if detected_path else None,
            "cropped_folder": str(cropped_folder),
            "cropped_images": cropped_image_paths
        }

    except Exception as e:
        raise RuntimeError(f"Image processing failed: {str(e)}")



