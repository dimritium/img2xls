from PIL import Image
from PIL import ImageOps

# import PIL
import xlsxwriter

def getImageDetails( img ):    
    im = Image.open( img )
    im = ImageOps.mirror(im)
    im = im.transpose(Image.ROTATE_90)
    im.thumbnail((320, 400), Image.BILINEAR)
    im.save('hckrmn.thumbnail.jpeg', "JPEG")
    pix = im.load()
    return pix, im.size

def RGB(r, g, b):
    return '#{:02x}{:02x}{:02x}'.format(r, g, b)

def populateCell(wb, ws, rgb_v, xls_row, col, col_v):
    cell_format = wb.add_format()
    cell_format.set_bg_color(rgb_v)
    ws.write(xls_row, col, col_v, cell_format)    

def populateXls( pixels, size ):
    rows, cols = size
    wb = xlsxwriter.Workbook('Img2Xls.xlsx')
    ws = wb.add_worksheet()
    for row in range(rows):
        xls_row = row*3
        for col in range(cols):
            lis_rgb = pixels[row, col]

            for i in range (3):
                if i == 0:
                    rgb_v = RGB(lis_rgb[0], 0, 0)
                elif i == 1:
                    rgb_v = RGB(0, lis_rgb[1], 0)
                else:
                    rgb_v = RGB(0, 0, lis_rgb[2])
                
                populateCell(wb, ws, rgb_v, xls_row + i, col, lis_rgb[i])
    wb.close()

if __name__ == "__main__":
    img = 'img/hkrmn.jpg'
    pixels, size = getImageDetails(img)
    populateXls( pixels, size )
    # print (pixels[0, 0], size)