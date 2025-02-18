import openpyxl
import os
import json
from openpyxl.drawing.image import Image
from PIL import Image as PILImage

# Define paths
EXCEL_FILE = "C:/Users/Syed/Desktop/HCHW/hchw-backend/uploads/student_data.xlsx"
UPLOADS_FOLDER = "C:/Users/Syed/Desktop/HCHW/hchw-backend/uploads/photos/"

# Ensure the uploads folder exists
if not os.path.exists(UPLOADS_FOLDER):
    os.makedirs(UPLOADS_FOLDER)

# Open the workbook
wb = openpyxl.load_workbook(EXCEL_FILE, data_only=True)
ws = wb.active  # Assuming the first sheet contains student data

# Identify column index for Admission Number (assuming it's in Column C)
admission_col_index = None
for col_idx, cell in enumerate(ws[1], start=1):  # Check header row (row 1)
    if cell.value and "Admin_No" in str(cell.value):
        admission_col_index = col_idx
        break

if admission_col_index is None:
    raise ValueError("❌ Could not find 'Admin_No' column in the Excel file.")

# Dictionary to store admission number to photo mapping
photo_mapping = {}

# Extract images and match them to admission numbers
image_count = 0  # Keep track of the number of images
for image in ws._images:
    # Image position in Excel
    row_index = image.anchor._from.row + 1  # Convert 0-based to 1-based index

    # Get Admission Number from the same row
    admission_no = ws.cell(row=row_index, column=admission_col_index).value

    if admission_no:  # Ensure there's a valid admission number
        # Save the image
        img_filename = f"photo_{admission_no}.png"
        img_path = os.path.join(UPLOADS_FOLDER, img_filename)

        # Convert openpyxl Image to PIL Image for saving
        pil_img = PILImage.open(image.ref)  # Open the image
        pil_img.save(img_path, "PNG")  # Save the image as PNG

        # Add mapping
        photo_mapping[str(admission_no)] = img_filename
        image_count += 1
        print(f"✅ Saved {img_filename} for Admission No: {admission_no}")

# Save mapping as JSON
json_path = os.path.join(UPLOADS_FOLDER, "photo_mapping.json")
with open(json_path, "w") as json_file:
    json.dump(photo_mapping, json_file, indent=4)

print(f"\n✅ Extraction complete! {image_count} images saved.")
print(f"📂 Images saved in: {UPLOADS_FOLDER}")
print(f"📝 Mapping saved as: {json_path}")
