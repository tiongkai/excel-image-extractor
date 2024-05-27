import openpyxl
from PIL import Image
from io import BytesIO

# Load the Excel file
wb = openpyxl.load_workbook('file.xlsx')

# Select the first sheet
sheet = wb['Worksheet1']
p=0
# Iterate through all images in the sheet
for image in sheet._images:
    # Get the image data
    img_data = image._data()
    p+=1
    # Convert the data to a PIL Image object
    img = Image.open(BytesIO(img_data))
    print(p)
    # Save the image as a .jpg file
    img.save('output_folder/output_{}.jpg'.format(p), 'JPEG')