from zpl import Label
from PIL import Image

l = Label(65,60)
height = 0


height += 4
image_width = 64
l.origin((l.width-image_width)/2, height)
image_height = l.write_graphic(Image.open('nota.png'),
image_width)
l.endorigin()



open ('zplexport.zpl', "w").write(l.dumpZPL())
#print(l.dumpZPL())
