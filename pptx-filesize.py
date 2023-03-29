#!/usr/bin/python3

import pptx
import os


sizedict = {}

start_path = "/dtop/hack"
for (path,dir,files) in os.walk(start_path):
    for file in files:
        fstat = os.stat(os.path.join(path, file))
        # print(f"path {path} File {file} is {fstat.st_size} bytes big")
        k = path + file
        sizedict[k] = fstat.st_size


sorted_x = sorted(sizedict.items(), key=lambda kv: kv[1], reverse=True)

for f in sorted_x[0:9]:
    print(f"Big file: {f[0]}")
exit(0)
# unzip  <file> -d <dir>
