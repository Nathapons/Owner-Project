from openpyxl import Workbook
from openpyxl.drawing.image import Image
from openpyxl.drawing.xdr import XDRPoint2D, XDRPositiveSize2D
from openpyxl.utils.units import pixels_to_EMU, cm_to_EMU
from openpyxl.drawing.spreadsheet_drawing import OneCellAnchor, AnchorMarker

wb = Workbook()
ws = wb.active
img = Image('fujikura_logo.ico')

h, w = 30, 50

p2e = pixels_to_EMU
c2e = cm_to_EMU

# Calculated number of cells width or height from cm into EMUs
cellh = lambda x: c2e((x * 49.77)/99)
cellw = lambda x: c2e((x * (18.65-1.71))/10)

# Want to place image in row 5 (6 in excel), column 2 (C in excel)
# Also offset by half a column.
column = 2
coloffset = cellw(0.5)
row = 5
rowoffset = cellh(0.5)

size = XDRPositiveSize2D(p2e(h), p2e(w))
# Offset Case
# marker = AnchorMarker(col=column, colOff=coloffset, row=row, rowOff=rowoffset)
marker = AnchorMarker(col=column, row=row)
img.anchor = OneCellAnchor(_from=marker, ext=size)
ws.add_image(img) 
wb.save('test.xlsx')