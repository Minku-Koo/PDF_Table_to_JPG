
"""
-Table Extract from PDF Document
-Crop PDF Table
-Convert to JPG File ans Save

Project Name : PDF Table to JPG
Create Date : 19/Nov/2020
Author : Minkuk Koo
E-Mail : corleone@kakao.com
Version : 1.0.0
Keyword : 'PDF', 'Table', 'Camelot' ,'PDF Extract', 'PYPDF', 'pdf2jpg'

* This Project used Excalibur Library
"""

import os, traceback, logging
from tkinter import Tk
from tkinter.filedialog import askopenfilename
from PyPDF2 import PdfFileReader, PdfFileWriter
from camelot.utils import get_page_layout, get_text_objects, get_rotation
from camelot.parsers import Lattice, Stream
from camelot.core import TableList
from camelot.io import read_pdf
import camelot
import pandas as df
import PyPDF2 
import os
from PIL import Image
import io
from pdf2image import convert_from_path, convert_from_bytes
from pdf2image.exceptions import (
    PDFInfoNotInstalledError,
    PDFPageCountError,
    PDFSyntaxError
)
import tempfile
import glob
from tqdm import tqdm
from pdf2jpg import pdf2jpg
from pdf2image import convert_from_path
import shutil
import time
temp = None

def file_path_select():
    root = Tk() # GUI로 파일경로를 선택하기위해 tkinter 라는 라이브러리 사용
    # GUI로 파일 선택 창이 나옴
    root.fileName = askopenfilename( filetypes=(("pdf file", "*.pdf"), ("all", "*.*")) ) #pdf 파일만 불러옴
    pdf_path = root.fileName
    root.destroy()  #gui 창 닫기
    return pdf_path #원하는 파일 경로 반환

def get_pages(filename, pages): #PDF 문서를 각 페이지를 구분하여 저장
    """Converts pages string to list of ints.

    Parameters
    ----------
    filename : str
        Path to PDF file.
    pages : str, optional (default: '1')
        Comma-separated page numbers.
        Example: 1,3,4 or 1,4-end.

    Returns
    -------
    N : int
        Total pages.
    P : list
        List of int page numbers.

    """
    page_numbers = []
    inputstream = open(filename, "rb")
    infile = PdfFileReader(inputstream, strict=False)
    N = infile.getNumPages()
    if pages == "1":
        page_numbers.append({"start": 1, "end": 1})
    else:
        if infile.isEncrypted:
            return "Encrypted", "error"
            
        if pages == "all":
            page_numbers.append({"start": 1, "end": infile.getNumPages()})
        else:
            for r in pages.split(","):
                if "-" in r:
                    a, b = r.split("-")
                    if b == "end":
                        b = infile.getNumPages()
                    page_numbers.append({"start": int(a), "end": int(b)})
                else:
                    page_numbers.append({"start": int(r), "end": int(r)})
    inputstream.close()
    P = []
    for p in page_numbers:
        P.extend(range(p["start"], p["end"] + 1))
    return sorted(set(P)), N

def save_page(filepath, page_number):
    filename = os.path.split( os.path.splitext(filepath)[0] )[1]
    dirname = os.path.join(os.path.dirname(filepath),  filename).replace('\\',"/")
    
    
    if not os.path.isdir(dirname):
        os.mkdir(dirname)
    
    infile = PdfFileReader(open(filepath, "rb"), strict=False)
    page = infile.getPage(page_number - 1)
    outfile = PdfFileWriter()
    outfile.addPage(page)
    outpath = os.path.join(os.path.dirname(filepath), "./{}/page-{}.pdf".format(filename, page_number))
    with open(outpath, "wb") as f:
        outfile.write(f)
    froot, fext = os.path.splitext(outpath)
    layout, __ = get_page_layout(outpath)
    
    chars = get_text_objects(layout, ltype="char")
    horizontal_text = get_text_objects(layout, ltype="horizontal_text")
    vertical_text = get_text_objects(layout, ltype="vertical_text")
    rotation = get_rotation(chars, horizontal_text, vertical_text)
    if rotation != "":
        outpath_new = "".join([froot.replace("page", "p"), "_rotated", fext])
        os.rename(outpath, outpath_new)
        infile = PdfFileReader(open(outpath_new, "rb"), strict=False)
        if infile.isEncrypted:
            infile.decrypt("")
        outfile = PdfFileWriter()
        p = infile.getPage(0)
        if rotation == "anticlockwise":
            p.rotateClockwise(90)
        elif rotation == "clockwise":
            p.rotateCounterClockwise(90)
        outfile.addPage(p)
        with open(outpath, "wb") as f:
            outfile.write(f)

def excalibur(filepath, pages): # excalibur 모듈 활용, PDF 각 페이지에서 Table 추출
    global temp
    pages_dic = {}
    filepaths = {}
    detected_areas = {}
    
    extract_pages, total_pages = get_pages(filepath, pages)
    if total_pages == "error":
        if extract_pages == "Encrypted":
            return "error:FIle is Encrypted"
    
    for page in extract_pages:
        new_filepath = filepath.replace(".pdf","/")+"page-{}".format(page)+".pdf"
        
        filepaths[page] = new_filepath
        save_page(filepath, page)
        
        lattice_areas, stream_areas = (None for i in range(2))
        
        page_file = os.path.join(
            os.path.dirname(filepath),
            os.path.basename(filepath).split(".")[0],
            "page-"+str(page)+".pdf"
            ).replace("\\", "/")
        
               
        # stream 방식
        try:
            parser = Stream()
            tables = parser.extract_tables(page_file)
            if len(tables):
                stream_areas = []
                for table in tables:
                    x1, y1, x2, y2 = table._bbox
                    stream_areas.append((x1, y2, x2, y1))
    
            temp = tables
    
        except Exception as e:
            print("\nStream Error")
            logging.error(traceback.format_exc())
            print(e)
            
        # Lattice  방식
        try:
            parser = Lattice()
            tables = parser.extract_tables(page_file)
            if len(tables):
                lattice_areas = []
                for table in tables:
                    x1, y1, x2, y2 = table._bbox
                    lattice_areas.append((x1, y2, x2, y1))
                
        except Exception as e:
            print("\nLattice Error")
            print(e)
            
        detected_areas[page] = {"lattice": lattice_areas, "stream": stream_areas}
        
        if (detected_areas[page]["stream"] is not None) and (detected_areas[page]["lattice"] is not None):
            parser = ("stream" if  len(detected_areas[page]["stream"]) > len(detected_areas[page]["lattice"]) else "lattice")
            
        else:
            parser = ("lattice" if detected_areas[page]["stream"] is None else "stream")
            
        pages_dic[page] = detected_areas[page][parser]
    
    if (detected_areas[page]["lattice"] is None) and (detected_areas[page]["stream"] is None):
        return "error:Table is not detected"
    
    tables = []
    
    for p in pages_dic.keys():
        
        Parser = (Lattice()if  detected_areas[p]["stream"] is None else Stream())
        t = Parser.extract_tables(filepaths[p])
        
        for _t in t:
            _t.page = int(p)
        tables.extend(t)
        
    tables = TableList(tables)
    
    return tables, detected_areas
    

# Table Crop PDF를 JPG 이미지로 변환
def save_images(source):
    os.chdir(source) #파라미터로 받은 경로로 설정
    pdfs = glob.glob("*.pdf") #PDF 파일만 추출
    
    new_folder = source+"/crop-jpg" #JPG 파일 저장할 폴더명 설정
    if os.path.exists(new_folder): #이미 존재하면 삭제하고 다시 생성
        shutil.rmtree(new_folder)
        os.makedirs(new_folder)
    else: # 존재하지 않으면 폴더 생성
        os.makedirs(new_folder)
        
    for it in tqdm(pdfs):
        if "crop" not in it: #PDF 파일명에 crop이 존재하지 않으면 > 기존 pdf file
            continue
        else: print("Convert jpg >>",it) # crop tablle pdf file
        render = pdf2jpg.convert_pdf2jpg(it, source, dpi=300, pages='ALL')[0] #JPG로 변환
        time.sleep(0.2) # JPG 변환을 위해 대기
        
        old_folder = "\\".join(render["output_jpgfiles"][0].split("\\")[:-1])
        pdf_file = render["output_jpgfiles"][0].split("\\")[-1]
        shutil.move(old_folder+"/"+pdf_file ,new_folder+"/"+pdf_file ) #JPG 파일 이동
        
        shutil.rmtree(render["output_pdfpath"])

# Camelot으로 인식한 Table 좌표를 통해 PDF Crop
def pdf_crop(filepath, x, y, w, h,pdf_num): # 파일 경로, 최초 x, y 값, 가로 세로 길이, PDF 번호
    with open(filepath,'rb') as fin:
        pdf = PyPDF2.PdfFileReader(fin)
        page = pdf.getPage(0)

        # Coordinates found by inspection.
        # Can these coordinates be found automatically?
        
        # page.cropBox.lowerLeft=(88,322)
        # page.cropBox.upperRight = (508,602)
        
        page.cropBox.lowerLeft=(x,y)
        page.cropBox.upperRight = (w,h)

        output = PyPDF2.PdfFileWriter()
        output.addPage(page)
        name = filepath.split("/")[-1].split(".")[0]
        path_  = "/".join(filepath.split("/")[:-1])+'/'+name+'-crop-'+str(pdf_num)
        with open(path_+'.pdf','wb') as fo:
            output.write(fo) # Crop PDF 저장
        time.sleep(0.1)
        """
        # 여기서 바로 JPG 변환해도 무방
        # save_images("/".join(filepath.split("/")[:-1]))
        """
    return 0
    
    
if __name__ == '__main__':
    pdf_path = file_path_select() #PDF 파일 지정
    result, detected_areas = excalibur(pdf_path, "all") #해당 PDF에서 추출한 Table

    for page in range(len(detected_areas)):
        page_file = os.path.join(
                    os.path.dirname(pdf_path),
                    os.path.basename(pdf_path).split(".")[0],
                    "page-"+str(page+1)+".pdf"
                    ).replace("\\", "/")
        
        n=1
        folder_path = pdf_path.split(".")[0]
        for x, y, w, h in detected_areas[page+1]['stream']:
            pdf_crop(page_file, x, y, w, h, n) #PDF Crop
            n+=1
            
    select = input("Image Convert Start? [y/n] :") # JPG 변환 여부
    if select.upper()=="Y":
        save_images(folder_path) # JPG로 변환하여 저장





