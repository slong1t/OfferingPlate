from openpyxl import Workbook
from openpyxl.drawing.image import Image

wb = Workbook()

# grab the active worksheet
ws = wb.active

rdir=r"C:\Users\Sam\Documents\Workspace\OfferingPlate"
import glob
import os
flist=glob.glob(os.path.join(rdir,"*.png"))
col = 1
row = 1
for f in flist:
    img = Image(f)
    print (img.image.width,img.image.height)
    size = 2*128,3*128 
    img.image.thumbnail(size)
    #my_image = img.image.resize((int(img.image.width * 246/256), int(img.image.height * 1/256)))
    cell = 'A' + str(row)
    ws.add_image(img,cell)
    row += 6

# Save the file
wb.save("sample.xlsx")