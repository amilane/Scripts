from PIL import Image
import os
os.chdir('C:\\Users\\Alisha\\Desktop\\123')

for p in range(11, 14):
  imgname = 'SCH - ' + str(p) 
  imgfile = imgname + '.bmp'
  img = Image.open(imgfile)
  img3 = img.crop( (125, 40, 2437, 1388) ) #A3
  #img3 = img.crop( (130, 40, 4925, 3470) ) #A1
  img3.save(imgname + '_crop' + '.bmp')