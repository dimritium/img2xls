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

def populateXls( pixels, size ):
    rows, cols = size
    wb = xlsxwriter.Workbook('Img2Xls.xlsx')
    ws = wb.add_worksheet()
    for row in range(rows):
        xls_row = row*3
        for col in range(cols):
            r, g, b = pixels[row, col]
            cell_format = wb.add_format()
            
            rgb_v = RGB(r, 0, 0)
            if( row == 0 and col == 0):
                print (rgb_v)
            cell_format.set_bg_color(rgb_v)
            ws.write(xls_row,     col, r, cell_format)

            rgb_v = RGB(0, g, 0)
            cell_format.set_bg_color(rgb_v)
            ws.write(xls_row + 1, col, g, cell_format)

            rgb_v = RGB(0, 0, b)
            cell_format.set_bg_color(rgb_v)
            ws.write(xls_row + 2, col, b, cell_format)
    wb.close()

if __name__ == "__main__":
    img = 'img/hkrmn.jpg'
    pixels, size = getImageDetails(img)
    populateXls( pixels, size )
    # print (pixels[0, 0], size)