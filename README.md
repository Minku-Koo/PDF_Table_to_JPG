# PDF_Table_to_JPG
## Extract table from PDF document, Crop and Convert to JPG file


**Table Extract from PDF Document**

**Crop PDF Table**

**Convert to JPG File and Save**


* Project Name : PDF Table to JPG

* Create Date : 19/Nov/2020

* Update Date : 11/Mar/2021

* Author : Minkuk Koo

* E-Mail : corleone@kakao.com

* Version : 1.2.0

* Keyword : 'PDF', 'Table', 'Camelot' ,'PDF Extract', 'PYPDF', 'pdf2jpg'


---------------------------------------------------------------------------------

***Warning***

1. If you get Error message like 'utf-8' encoding~ ,than 
    you should update pdf2jpg library.
    
    > You can see variable named 'output' in 'pdf2jpg' library
    
    > You must decode that : output = output.decode() ==> output = output.decode("cp949")
2. This Project used Camelot Library


***경 고***

PDF 파일 경로에 한글이 포함되어 있는 경우, 에러가 발생할 수 있습니다.

이럴 경우에는 별도로 pdf2jpg 라이브러리 파일을 열어서

output 변수를 따로 decoding 해주어야 합니다.



기존 output = output.decode() 으로 되어있는 부분을

output = output.decode("cp949") 로 바꾸어주면, 한글 폴더명이 포함되어도 해결됩니다.


---------------------------------------------------------------------------------

***Notice***
1. 이 프로젝트는 Camelot를 활용하여 제작되었습니다.
1. Camelot이 테이블을 추출하는 것이 100% 정확하지 않습니다.
1. Stream 방식의 경우, Table이 온전히 Crop되지 않습니다. 테이블의 일부가 잘려나갈 수 있습니다.



---------------------------------------------------------------------------------



## Example

**Original PDF Document**

![git1](https://user-images.githubusercontent.com/25974226/99659903-ed597c80-2aa4-11eb-921d-10cd975db817.PNG)


**Cropped Table JPG (Lattice Method)**
***table 1***
![sample-page-2-lattice-crop-3 pdf](https://user-images.githubusercontent.com/25974226/110661046-843e7600-8207-11eb-9568-3583734da546.jpg)
***table 2***
![sample-page-2-lattice-crop-4 pdf](https://user-images.githubusercontent.com/25974226/110661033-8274b280-8207-11eb-8b31-4a739346fa2e.jpg)


**Cropped Table JPG (Stream Method)**
***table 1 (Failed to detect table)***
![sample-page-2-stream-crop-1 pdf](https://user-images.githubusercontent.com/25974226/110661038-83a5df80-8207-11eb-8fd3-46ecdf33c034.jpg)
***table 2***
![sample-page-2-stream-crop-2 pdf](https://user-images.githubusercontent.com/25974226/110661043-843e7600-8207-11eb-95a5-f7921abb72a6.jpg)


