#!/usr/bin/env python3

import hashlib
import sys
import openpyxl
import pptx
import os




def hexstuff():
    prs = pptx.Presentation('hack.pptx')
    newImgFilename = "gray.jpg"
    img2 = pptx.parts.image.Image.from_file(newImgFilename)

    print(hashlib.sha224(prs.slides[0].shapes[2].image._blob).hexdigest())
    print(hashlib.sha224(img2._blob).hexdigest())

    prs.slides[0].shapes[2].image._blob = img2._blob
    print(hashlib.sha224(prs.slides[0].shapes[2].image._blob).hexdigest())



def ImageListToXls(presentation, output_xls):
    prs = pptx.Presentation(presentation)
    wb = openpyxl.Workbook()
    ws = wb.active


    slide_num = 0
    for slide in prs.slides:
        slide_num = slide_num + 1
        for shape in slide.shapes:
            try:
                imgsize = sys.getsizeof(shape.image.blob)
            except AttributeError as e:
                continue
            element_rid = shape._element.blip_rId
            related_part = shape.part.related_part(element_rid)
            source_filename = related_part.partname
            print(f"Slide {slide_num:>3} shape  {shape.shape_id:>5}, partname {shape.name:>15}, size {imgsize:>11} hash {shape.image.sha1:<4} file {source_filename:<20}")
            # xl_row = [slide_num, relpart.partname, imgsize, relpart.sha1]
            # ws.append(xl_row)
    wb.save(output_xls)







os.chdir('./temp')
ImageListToXls('big05.pptx', 'images_out.xlsx')

