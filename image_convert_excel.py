from PIL import Image
import win32com.client

def rgb_to_hex(rgb):
    bgr = (rgb[2], rgb[1], rgb[0])
    strValue = '%02x%02x%02x' % bgr
    iValue = int(strValue, 16)
    return iValue


xl = win32com.client.Dispatch("Excel.Application")
wb = xl.Workbooks.open('C:/Users/weijie/PycharmProjects/convert_excel/output.xlsx')
ws = wb.Sheets(1)

im = Image.open('image.jpg')
px = im.load()
width, height = im.size

for i in range(width):
    for j in range(height):
        ws.Cells(j + 1, i + 1).Interior.color = rgb_to_hex(px[i, j])

wb.Close(True)
xl.Application.Quit()
del xl