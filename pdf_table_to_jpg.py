
"""
-Table Extract from PDF Document
-Crop PDF Table
-Convert to JPG File ans Save

Project Name : PDF Table to JPG
Create Date : 19/Nov/2020
Update Date : 11/Mar/2021
Author : Minkuk Koo
E-Mail : corleone@kakao.com
Version : 1.2.0
Keyword : 'PDF', 'Table', 'Camelot' ,'PDF Extract', 'PyPDF', 'pdf2jpg'

* If you get Error message like 'utf-8' encoding~
    you should update pdf2jpg library.
    > You can see variable named 'output' in 'pdf2jpg' library
    > You must decode that : output = output.decode() ==> output = output.decode("cp949")
* This Project used Camelot
"""

import os, traceback, logging, glob, tempfile
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
from tqdm import tqdm
from pdf2jpg import pdf2jpg
import shutil, time

def file_path_select():
    root = Tk() # GUI로 파일경로를 선택하기위해 tkinter 라는 라이브러리 사용
    # GUI로 파일 선택 창이 나옴
    root.fileName = askopenfilename( filetypes=(("pdf file", "*.pdf"), ("all", "*.*")) ) #pdf 파일만 불러옴
    pdf_path = root.fileName
    root.destroy()  #gui 창 닫기
    return pdf_path #원하는 파일 경로 반환

# PDF 페이지 개수를 count
def get_pages(filename, pages): 
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

#PDF 문서 각 페이지를 구분하여 저장
def save_page(filepath, page_number):
    """
    Parameters
    ----------
    filepath : str
        Path to PDF file.
    page_number : int
        PDF page number

    Returns
    -------
    (Nothing)
    """
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

# camelot 모듈 활용, PDF 각 페이지에서 Table 추출
def table_extract(filepath, pages): 
    """
    Parameters
    ----------
    filepath : str
        Path to PDF file.
    pages : int
        PDF page number

    Returns
    -------
    detected_areas : list
        table in PDF coordinates list
    """
    
    pages_dic, filepaths = {}, {}
    detected_areas = {}
    
    extract_pages, total_pages = get_pages(filepath, pages)
    if total_pages == "error":
        if extract_pages == "Encrypted": return "error:FIle is Encrypted"
    
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
                for table in tables: # table object
                    x1, y1, x2, y2 = table._bbox
                    stream_areas.append((x1, y1, x2, y2))
    
        except Exception as e:
            print("\nStream Error!!")
            logging.error(traceback.format_exc())
            print(e)
            
        # Lattice  방식
        try:
            parser = Lattice()
            tables = parser.extract_tables(page_file)
            if len(tables):
                lattice_areas = []
                for table in tables: # table object
                    x1, y1, x2, y2 = table._bbox
                    lattice_areas.append((x1, y1, x2, y2))
                
        except Exception as e:
            print("\nLattice Error!!")
            print(e)
            
        # dirctionary { parser method : table coordinates list } 
        detected_areas[page] = {"lattice": lattice_areas, "stream": stream_areas}
        
        if (detected_areas[page]["stream"] is not None) and \
            (detected_areas[page]["lattice"] is not None):
            parser = ("stream" if  len(detected_areas[page]["stream"]) > len(detected_areas[page]["lattice"]) else "lattice")
            
        else:
            parser = ("lattice" if detected_areas[page]["stream"] is None else "stream")
            
        pages_dic[page] = detected_areas[page][parser]
    
    if (detected_areas[page]["lattice"] is None) and (detected_areas[page]["stream"] is None):
        return "error:Table is not detected"
    
    '''
    tables = []
    
    for p in pages_dic.keys():
        
        Parser = (Lattice()if  detected_areas[p]["stream"] is None else Stream())
        t = Parser.extract_tables(filepaths[p])
        
        for _t in t:
            _t.page = int(p)
        tables.extend(t)
        
    tables = TableList(tables)
    '''
    
    return detected_areas
    
# Table Cropped PDF를 JPG 이미지로 변환
def save_images(dir_path):
    """
    Parameters
    ----------
    dir_path : str
        Path that save image files

    Returns
    -------
    (Nothing)
    """
    os.chdir(dir_path) #파라미터로 받은 경로로 설정
    pdfs = glob.glob("*.pdf") #PDF 파일만 추출
    
    new_dir = dir_path+"/crop-jpg" #JPG 파일 저장할 폴더명 설정
    if os.path.exists(new_dir): #이미 존재하면 삭제하고 다시 생성
        shutil.rmtree(new_dir)
        os.makedirs(new_dir)
    else: # 존재하지 않으면 폴더 생성
        os.makedirs(new_dir)
    
    pdf_name = dir_path.split('/')[-1]
    
    for file in tqdm(pdfs):
        if "crop" not in file: continue #PDF 파일명에 crop이 존재하지 않으면 > 기존 pdf file
        else: print("\nConvert jpg >>", file) # crop tablle pdf file
        
        render = pdf2jpg.convert_pdf2jpg(file, dir_path, dpi=400, pages='ALL')[0] #JPG로 변환
        #time.sleep(0.1) # JPG 변환을 위해 대기
        
        old_dir = "\\".join(render["output_jpgfiles"][0].split("\\")[:-1])
        pdf_file = render["output_jpgfiles"][0].split("\\")[-1]
        shutil.move(old_dir+"/"+pdf_file ,new_dir+"/"+ pdf_name+"-" +pdf_file.split("_")[1] ) #JPG 파일 이동
        
        shutil.rmtree(render["output_pdfpath"])
        
    return 0

# Camelot으로 인식한 Table 좌표를 통해 PDF Crop
def pdf_crop(filepath, x1, y1, x2, y2, pdf_num, parser): # 파일 경로, 최초 x, y 값, 가로 세로 길이, PDF 번호, parser method
    with open(filepath,'rb') as f:
        pdf = PyPDF2.PdfFileReader(f)
        page = pdf.getPage(0)

        page.cropBox.upperLeft=(x1,y1)
        page.cropBox.lowerRight = (x2,y2)
        
        output = PyPDF2.PdfFileWriter()
        output.addPage(page)
        name = filepath.split("/")[-1].split(".")[0]
        path_  = "/".join(filepath.split("/")[:-1])+'/'+name+"-"+parser+'-crop-'+str(pdf_num)
        
        with open(path_+'.pdf','wb') as f:
            output.write(f) # Cropped PDF 저장
        
        """
        # 여기서 바로 JPG 변환할 수도 있음
        # save_images("/".join(filepath.split("/")[:-1]))
        """
    return 0
    
# main function
def main():
    pdf_path = file_path_select() #PDF 파일 지정
    detected_areas = table_extract(pdf_path, "all") #해당 PDF에서 추출한 Table
    
    for page in range(len(detected_areas)):
        page_file = os.path.join(
                    os.path.dirname(pdf_path),
                    os.path.basename(pdf_path).split(".")[0],
                    "page-"+str(page+1)+".pdf"
                    ).replace("\\", "/")
        
        num=1
        dir_path = pdf_path.split(".")[0]
        parser = "stream"
        if detected_areas[page+1][parser] != None:
            for x1, y1, x2, y2 in detected_areas[page+1][parser]:
                pdf_crop(page_file, x1, y1, x2, y2, num, parser) #PDF Crop
                num+=1
        parser ="lattice"
        if detected_areas[page+1][parser] != None:
            for x1, y1, x2, y2 in detected_areas[page+1][parser]:
                pdf_crop(page_file, x1, y1, x2, y2, num, parser) #PDF Crop
                num+=1
            
    select = input("Start Image Converting? [y/n] :") # JPG 변환 여부
    if select.upper()=="Y":
        save_images(dir_path) # JPG로 변환하여 저장


if __name__ == '__main__':
    
    main()


