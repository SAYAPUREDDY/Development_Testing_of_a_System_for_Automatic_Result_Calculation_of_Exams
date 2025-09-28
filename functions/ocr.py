from pathlib import Path
from PIL import Image
from .yolo_detection import process_image_with_yolo

from doctr.models import recognition_predictor
from doctr.io import DocumentFile


model = recognition_predictor(pretrained=True)

# # Initialize the OCR predictor once
# from transformers import TrOCRProcessor, VisionEncoderDecoderModel

# # Load model and processor
# processor = TrOCRProcessor.from_pretrained('microsoft/trocr-base-handwritten')
# ocr_model = VisionEncoderDecoderModel.from_pretrained('microsoft/trocr-base-handwritten')

def process_and_ocr_image(image_path: str, output_dir: Path = None) -> dict:
    yolo_result = process_image_with_yolo(image_path, output_dir=output_dir)

    # Prepare extracted text results
    extracted_texts = []

    for crop in yolo_result.get("cropped_images", []):
        label = crop["label"]
        path = Path(crop["path"])

        raw_path = Path(crop["raw_path"])
        # print(raw_path)

        image = DocumentFile.from_images(path)
        # pixel_values = processor(images=image, return_tensors="pt").pixel_values
        # generated_ids = ocr_model.generate(pixel_values)
        # result = processor.batch_decode(generated_ids, skip_special_tokens=True)
        result=model(image)

        extracted_texts.append({
            "label": label,
            "text": result,
            "bbox": crop.get("bbox"),
            "image_path": str(path),
            "raw_path":raw_path
        })
    
    return {
        "original": yolo_result.get("original"),
        "detected_image": yolo_result.get("detected_image"),
        "cropped_folder": yolo_result.get("cropped_folder"),
        "results": extracted_texts
    }


