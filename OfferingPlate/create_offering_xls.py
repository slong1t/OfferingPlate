from PIL import Image
import xlsxwriter
import io
from datetime import date
from os import path

def calculate_scale(file_path, bound_size):
    # check the image size without loading it into memory
    im = Image.open(file_path)
    original_width, original_height = im.size
    if  (original_height > original_width):
        temp_width = original_width
        temp_height = original_height
        original_width = temp_height
        original_height = temp_width
         
    # calculate the resize factor, keeping original aspect and staying within boundary
    bound_width, bound_height = bound_size
    #print "bound " + str(bound_width) + " " + str(bound_height)
    ratios = (float(bound_width) / original_width, float(bound_height) / original_height)
    #print (file_path, original_width, original_height, ratios )
    return min(ratios)


def get_resized_image_data(file_path, bound_width_height):
    # get the image and resize it
    im = Image.open(file_path)
    original_width, original_height = im.size
    if  (original_height > original_width):
        size = original_height,original_width
        im = im.rotate(90, expand=1).resize(size)
        #im.show()

    im.thumbnail(bound_width_height, Image.ANTIALIAS)  # ANTIALIAS is important if shrinking

    # stuff the image data into a bytestream that excel can read
    im_bytes = io.BytesIO()
    im.save(im_bytes, format='PNG')
    return im_bytes

def write_title_row(row):
    worksheet.set_column('A:A', 55)
    worksheet.set_row(0, 20)  
    worksheet.freeze_panes(row,row)
    worksheet.write('A'+str(row), 'Image', cell_format)
    worksheet.write('B'+str(row), 'Cash', cell_format)
    worksheet.set_column('B:B', 17)
    worksheet.write('C'+str(row), 'Check', cell_format)
    worksheet.set_column('C:C', 17)
    worksheet.write('D'+str(row), '4110 Tithes and Offerings', cell_format)
    worksheet.set_column('D:D', 17)
    worksheet.write('E'+str(row), '4112 Building Fund', cell_format)
    worksheet.set_column('E:E', 17)
    worksheet.write('F'+str(row), '4108 Youth Camp', cell_format)
    worksheet.set_column('F:F', 17)
    worksheet.write('G'+str(row), '4103 Maintenance Fund', cell_format)
    worksheet.set_column('G:G', 17)
    worksheet.write('H'+str(row), '??? Fund', cell_format)
    worksheet.set_column('H:H', 17)
    return row+1

def write_scanned_image_rows(row):
    import glob
    bound_width_height = (128*3, 128*2)
    flist=glob.glob(path.join(scandir,"*.jpg"))
    for f in flist:
        image_cell = 'A' + str(row)
        image_data = get_resized_image_data(f, bound_width_height)
        im = Image.open(image_data)
        im.seek(0)  # reset the "file" for excel to read it.    
        worksheet.insert_image(image_cell, f, {'image_data': image_data})
        worksheet.write('B'+str(row), '', num_format)
        worksheet.write('C'+str(row), '', num_format)
        worksheet.write('D'+str(row), '', num_format)
        formula = '=B' + str(row) + ' + C' + str(row) + ' - SUM(E' + str(row) + ':L' + str(row) + ')'
        worksheet.write_formula('D'+str(row), formula,num_format)
        worksheet.write('E'+str(row), '', num_format)
        worksheet.write('F'+str(row), '', num_format)
        worksheet.write('G'+str(row), '', num_format)
        worksheet.write('H'+str(row), '', num_format)
        worksheet.write('I'+str(row), '', num_format)
        worksheet.set_row(row-1,137)
        row += 1

    center_format = workbook.add_format()
    center_format.set_align('center')
    center_format.set_align('vcenter')
    center_format.set_font_size(22)
    worksheet.write('A'+str(row), 'Undesignated Cash', center_format)
    formula = '=B' + str(row) + ' - SUM(E' + str(row) + ':L' + str(row) + ')'
    worksheet.write_formula('D'+str(row), formula,num_format)

    worksheet.write('B'+str(row), '', num_format)
    worksheet.write('C'+str(row), '', num_format)
    worksheet.write('E'+str(row), '', num_format)
    worksheet.write('F'+str(row), '', num_format)
    worksheet.write('G'+str(row), '', num_format)
    worksheet.write('H'+str(row), '', num_format)
    worksheet.write('I'+str(row), '', num_format)
 
    worksheet.set_row(row-1,137)
    row += 1
    
    return row

def write_summary_row(row):
    worksheet.set_row(row-1, 20)  
    worksheet.write('A'+str(row), 'Total', cell_format)
    cell_format.set_num_format(8)
    worksheet.write_formula('B'+str(row), '=SUM(B1:B'+str(row-1)+')',cell_format)
    worksheet.write('C'+str(row), '', cell_format)
    worksheet.write_formula('C'+str(row), '=SUM(C1:C'+str(row-1)+')',cell_format)
    worksheet.write('D'+str(row), '', cell_format)
    worksheet.write_formula('D'+str(row), '=SUM(D1:D'+str(row-1)+')',cell_format)
    worksheet.write('E'+str(row), '', cell_format)
    worksheet.write_formula('E'+str(row), '=SUM(E1:E'+str(row-1)+')',cell_format)
    worksheet.write('F'+str(row), '', cell_format)
    worksheet.write_formula('F'+str(row), '=SUM(F1:F'+str(row-1)+')',cell_format)
    worksheet.write('G'+str(row), '', cell_format)
    worksheet.write_formula('G'+str(row), '=SUM(G1:G'+str(row-1)+')',cell_format)
    worksheet.write('H'+str(row), '', cell_format)
    worksheet.write_formula('H'+str(row), '=SUM(H1:H'+str(row-1)+')',cell_format)
    return row+1

def write_deposit_slip(row):
    row += 1
    worksheet.write('A'+str(row), 'Total Cash', cell_format)
    worksheet.write_formula('B'+str(row), '=B'+str(row-2),cell_format)
    row += 1
    worksheet.write('A'+str(row), 'Total Checks', cell_format)
    worksheet.write_formula('B'+str(row), '=C'+str(row-3),cell_format)
    row += 1
    worksheet.write('A'+str(row), 'Total Deposit', cell_format)
    worksheet.write_formula('B'+str(row), '=B'+str(row-2) + '+B'+str(row-1),cell_format)
    row += 1
    center_format = workbook.add_format()
    center_format.set_align('center')
    center_format.set_align('vcenter')
    center_format.set_font_size(22)
    worksheet.write('A'+str(row), 'Paste Deposit Slip Here', center_format)
    worksheet.set_row(row-1,137)
    row += 1
    
    return row+1
    
    

##################### MAIN ###############################

# add a dialog if directory doesn't exists to pick one
today = date.today()
userpath = path.expanduser('~\OneDrive\Documents\Scans\\2019 Receipts')
#userpath = path.expanduser('~\OneDrive\Scans')
scandir = userpath + "\\" + str(today.year) + str(today.month).zfill(2) + str(today.day).zfill(2)
if path.exists(scandir) == False:
    exit (scandir + " directory doesn't exist !!!!")

xlsFile =  scandir + '\offering' + str(today.year) + str(today.month).zfill(2) + str(today.day).zfill(2) +'.xlsx'
if path.isfile(xlsFile):
    xlsFile = scandir + '\offering' + str(today.year) + str(today.month).zfill(2) + str(today.day).zfill(2) + '_NEW.xlsx'
workbook = xlsxwriter.Workbook(xlsFile)
worksheet = workbook.add_worksheet()

cell_format = workbook.add_format()
cell_format.set_pattern(1)  # This is optional when using a solid fill.
cell_format.set_bg_color('blue')
cell_format.set_font_color('white')
cell_format.set_font_size(16)

num_format = workbook.add_format()
num_format.set_font_size(16)
num_format.set_num_format(8)

row = 1
row = write_title_row(row)
row = write_scanned_image_rows(row)
row = write_summary_row(row)
row = write_deposit_slip(row)

workbook.close()