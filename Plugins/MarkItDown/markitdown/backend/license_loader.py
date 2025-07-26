import os
import aspose.words as aw
import aspose.cells as ac
import aspose.slides as slides

def _load_license(LicenseClass, product_name: str):
    license_path = os.getenv("ASPOSE_LICENSE_PATH")
    if not license_path:
        print(f"[Aspose] No license path set for {product_name}. Running in evaluation mode.")
        return
    try:
        license = LicenseClass()
        license.set_license(license_path)
        print(f"[Aspose] {product_name} license applied successfully.")
    except Exception as e:
        print(f"[Aspose] Failed to apply {product_name} license: {e}. Running in evaluation mode.")

def apply_words_license():
    _load_license(aw.License, "Words")

def apply_cells_license():
    _load_license(ac.License, "Cells")

def apply_slides_license():
    _load_license(slides.License, "Slides")
