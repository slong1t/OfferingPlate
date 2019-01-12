import xlsxwriter
# Create an new Excel file and add a worksheet.
rdir=r"C:\Users\Sam\Documents\Scans\20181118"
workbook = xlsxwriter.Workbook(rdir + '\images.xlsx')
worksheet = workbook.add_worksheet()

# Widen the first column to make the text clearer.
worksheet.set_column('A:A', 50)
worksheet.set_default_row(150)
import glob
import os
flist=glob.glob(os.path.join(rdir,"*.jpg"))
col = 1
row = 2
for f in flist:
    #print (img.image.width,img.image.height)
    image_cell = 'A' + str(row)
    worksheet.insert_image(image_cell, f, {'x_scale': 0.6, 'y_scale': 0.6})
    name_cell = 'B' + str(row)
    worksheet.write(name_cell, f)
    row += 1

workbook.close()