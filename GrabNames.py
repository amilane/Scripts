import os

directory = r'C:\Users\a-suchkova\Desktop\ВЫПУСК'

files = os.listdir(directory)

import pprint
pp = pprint.PrettyPrinter(indent=0)
pp.pprint(files)

f = open(directory + '\Список файлов.txt', 'w')
for i in files:
  f.write(i + '\n')
f.close()