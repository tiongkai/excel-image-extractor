import openpyxl
from openpyxl_image_loader import SheetImageLoader
import os

# Load the Excel file
wb = openpyxl.load_workbook('M7-230524_clean.xlsx')

# Select the first sheet
sheet = wb['Worksheet1']
image_loader = SheetImageLoader(sheet)
image_cell = 'B'
camera_cell = 'A'
POI_enrol = 'C'
POI_name = 'D'
time_cell = 'G'
date_cell = 'H'

POI_output = 'Extractions/POI/'
Wild_output = 'Extractions/Wild/'

for row in range(sheet.max_row-2):
    index = str(row+2)
    if os.path.exists(POI_output+sheet[POI_name+index].value+'.png'):
        print()
    else:
        image = image_loader.get(POI_enrol+index)
        image.save(POI_output+sheet[POI_name+index].value+'.png')
        os.makedirs('Extractions/Wild/'+sheet[POI_name+index].value)

    image_wild = image_loader.get(image_cell+index)
    date = sheet[date_cell+index].value.replace('-','')
    time = sheet[time_cell+index].value.replace(':','')[:6]
    filename = f"{row}_{sheet[camera_cell+index].value}_{date}{time}_NA"
    image_wild.save('Extractions/Wild/'+sheet[POI_name+index].value+'/'+filename+'.png')