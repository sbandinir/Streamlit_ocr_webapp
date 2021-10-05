#librerie
import streamlit as st
import cv2
import numpy as np
import pandas as pd  
import pytesseract
import pdf2image
import matplotlib.pyplot as pyplot
from matplotlib import patches
import base64
from io import BytesIO
import xlsxwriter

#definisco coordinate
coordinates= {
    'PA Media 24h - ABPM: PAS media 24h' : (453,1163,512,1196),
    'PA Media 24h - ABPM: PAD media 24h' : (459,1213,511,1246),
    'PA Media 24h - ABPM: PAM media 24h' : (459,1263,515,1298),
    'PA Media 24h - ABPM: FC media 24h' : (456,1314,511,1348),
    'PA Media 24h - ABPM: PP media 24h' : (461,1365,514,1399),
    'PA Media 24h - ABPM: PAS media 24h STD' : (452,1466,514,1503),
    'PA Media 24h - ABPM: PAD media 24h STD' : (452,1518,515,1554),
    'PA Media 24h - ABPM: PAM media 24h STD' : (452,1569,512,1605),
    'PA Media 24h - ABPM: FC media 24h STD' : (454,1618,513,1654),
    'PA Media 24h - ABPM: PP media 24h STD' : (453,1670,513,1707),
    'PA Media 24h - ABPM: PAS letture oltre limite 24h' :(468,1772,518,1805),
    'PA Media 24h - ABPM: PAD letture oltre limite 24h' : (470,1820,518,1859),
    'PA Media Giorno - ABPM: PAS media GIORNO' : (684,1163,751,1195),
    'PA Media Giorno - ABPM: PAD media GIORNO' : (688,1212,784,1246),
    'PA Media Giorno - ABPM: PAM media GIORNO' : (694,1261,748,1297),
    'PA Media Giorno - ABPM: FC media GIORNO' : (690,1312,747,1350),
    'PA Media Giorno - ABPM: PP media GIORNO' : (695,1361,748,1399),
    'PA Media Giorno - ABPM: PAS media GIORNO STD' : (689,1468,744,1503),
    'PA Media Giorno - ABPM: PAD media GIORNO STD' : (691,1518,747,1555),
    'PA Media Giorno - ABPM: PAM media GIORNO STD' : (691,1569,748,1604),
    'PA Media Giorno - ABPM: FC media GIORNO STD' : (690,1619,751,1655),
    'PA Media Giorno - ABPM: PP media GIORNO STD' : (689,1670,749,1708),
    'PA Media Giorno - ABPM: PAS letture oltre limite GIORNO' :(710,1771,751,1807),
    'PA Media Giorno - ABPM: PAD letture oltre limite GIORNO' : (709,1821,749,1858),
    'PA Media Notte - ABPM: PAS media NOTTE' : (917,1162,985,1195),
    'PA Media Notte - ABPM: PAD media NOTTE' : (921,1210,988,1249),
    'PA Media Notte - ABPM: PAM media NOTTE' : (924,1262,994,1296),
    'PA Media Notte - ABPM: FC media NOTTE' : (925,1310,985,1350),
    'PA Media Notte - ABPM: PP media NOTTE' : (927,1363,988,1402),
    'PA Media Notte - ABPM: PAS media NOTTE STD' : (923,1465,986,1504),
    'PA Media Notte - ABPM: PAD media NOTTE STD' : (924,1516,986,1553),
    'PA Media Notte - ABPM: PAM media NOTTE STD' : (932,1571,980,1601),
    'PA Media Notte - ABPM: FC media NOTTE STD' : (913,1615,1001,1657),
    'PA Media Notte - ABPM: PP media NOTTE STD' : (926,1668,987,1706),
    'PA Media Notte - ABPM: PAS letture oltre limite NOTTE' :(931,1772,985,1806),
    'PA Media Notte - ABPM: PAD letture oltre limite NOTTE' : (938,1821,986,1857)}

#pathtesseract
#in locale, decommenta: pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'.

#definisco funzioni
def crop_save_test(image: str, text: str, coordinates: set, language: str = 'ita') -> dict:
    img = image
    img_cropped = img.crop(coordinates) 
    #preprocessing
    img_cropped = cv2.cvtColor(np.array(img_cropped), cv2.COLOR_BGR2GRAY)
    kernel = np.ones((1, 1), np.uint8)
    img_cropped = cv2.dilate(img_cropped, kernel, iterations=1)
    img_cropped = cv2.erode(img_cropped, kernel, iterations=1)
    img_cropped=cv2.threshold(cv2.bilateralFilter(img_cropped, 5, 75, 75), 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)[1]
    img_ocr = pytesseract.image_to_string(img_cropped, lang='ita',config='--psm 10 --oem 3 -c tessedit_char_whitelist=0123456789,')
    img_ocr = img_ocr.replace("\n\x0c", "")
    img_ocr = img_ocr.replace(",", ".")
        
    return img_ocr

def to_excel(df):
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    df.to_excel(writer, index=False, sheet_name='Sheet1') # <--- here
    writer.save()
    processed_data = output.getvalue()
    return processed_data

def get_table_download_link(df):
    val = to_excel(df)
    b64 = base64.b64encode(val)  
    return f'<a href="data:application/octet-stream;base64,{b64.decode()}" download="Risultato_OCR.xlsx">Download xlsx file</a>'


#applicazione web
st.title("OCR referto e export in Excel")
st.text("Selezionare il Pdf per cui si vuole fare l'ocr")
uploaded_file = st.file_uploader("Upload Pdf",type=["pdf"])

if uploaded_file is not None:
  #in locale, decommenta:images = pdf2image.convert_from_bytes(uploaded_file.read() ,poppler_path=r'C:\Program Files\poppler-0.68.0\bin')
  images = pdf2image.convert_from_bytes(uploaded_file.read())
  page1=images[0] #interessa solo la prima pagina del pdf 
  #plot
  #st.image(page1, use_column_width=True) ####MODIFICA QUI PLOT
  figure, ax = pyplot.subplots(figsize=(15,15))

  rect1 = patches.Rectangle((453,1163),60,30, edgecolor='r', facecolor="none")
  rect2 = patches.Rectangle((453,1213),60,30, edgecolor='r', facecolor="none")
  rect3 = patches.Rectangle((453,1263),60,30, edgecolor='r', facecolor="none")
  rect4 = patches.Rectangle((453,1314),60,30, edgecolor='r', facecolor="none")
  rect5 = patches.Rectangle((453,1365),60,30, edgecolor='r', facecolor="none")
  rect6 = patches.Rectangle((453,1466),60,30, edgecolor='r', facecolor="none")
  rect7 = patches.Rectangle((453,1518),60,30, edgecolor='r', facecolor="none")
  rect8 = patches.Rectangle((453,1569),60,30, edgecolor='r', facecolor="none")
  rect9 = patches.Rectangle((453,1618),60,30, edgecolor='r', facecolor="none")
  rect10 = patches.Rectangle((453,1670),60,30, edgecolor='r', facecolor="none")
  rect11 = patches.Rectangle((470,1772),60,30, edgecolor='r', facecolor="none")
  rect12 = patches.Rectangle((470,1820),60,30, edgecolor='r', facecolor="none")
  rect13 = patches.Rectangle((690,1163),60,30, edgecolor='r', facecolor="none")
  rect14 = patches.Rectangle((690,1212),60,30, edgecolor='r', facecolor="none")
  rect15 = patches.Rectangle((690,1261),60,30, edgecolor='r', facecolor="none")
  rect16 = patches.Rectangle((690,1312),60,30, edgecolor='r', facecolor="none")
  rect17 = patches.Rectangle((690,1361),60,30, edgecolor='r', facecolor="none")
  rect18 = patches.Rectangle((690,1468),60,30, edgecolor='r', facecolor="none")
  rect19 = patches.Rectangle((690,1518),60,30, edgecolor='r', facecolor="none")
  rect20 = patches.Rectangle((690,1569),60,30, edgecolor='r', facecolor="none")
  rect21 = patches.Rectangle((690,1619),60,30, edgecolor='r', facecolor="none")
  rect22 = patches.Rectangle((690,1670),60,30, edgecolor='r', facecolor="none")
  rect23 = patches.Rectangle((710,1771),60,30, edgecolor='r', facecolor="none")
  rect24 = patches.Rectangle((710,1821),60,30, edgecolor='r', facecolor="none")
  rect25 = patches.Rectangle((925,1162),60,30, edgecolor='r', facecolor="none")
  rect26 = patches.Rectangle((925,1210),60,30, edgecolor='r', facecolor="none")
  rect27 = patches.Rectangle((925,1262),60,30, edgecolor='r', facecolor="none")
  rect28 = patches.Rectangle((925,1310),60,30, edgecolor='r', facecolor="none")
  rect29 = patches.Rectangle((925,1363),60,30, edgecolor='r', facecolor="none")
  rect30 = patches.Rectangle((925,1465),60,30, edgecolor='r', facecolor="none")
  rect31= patches.Rectangle((925,1516),60,30, edgecolor='r', facecolor="none")
  rect32= patches.Rectangle((925,1569),60,30, edgecolor='r', facecolor="none")
  rect33= patches.Rectangle((925,1618),60,30, edgecolor='r', facecolor="none")
  rect34= patches.Rectangle((925,1668),60,30, edgecolor='r', facecolor="none")
  rect35= patches.Rectangle((938,1772),60,30, edgecolor='r', facecolor="none")
  rect36= patches.Rectangle((938,1821),60,30, edgecolor='r', facecolor="none")
  ax.imshow(page1)
  ax.add_patch(rect1)
  ax.add_patch(rect2)
  ax.add_patch(rect3)
  ax.add_patch(rect4)
  ax.add_patch(rect5)
  ax.add_patch(rect6)
  ax.add_patch(rect7)
  ax.add_patch(rect8)
  ax.add_patch(rect9)
  ax.add_patch(rect10)
  ax.add_patch(rect11)
  ax.add_patch(rect12)
  ax.add_patch(rect13)
  ax.add_patch(rect14)
  ax.add_patch(rect15)
  ax.add_patch(rect16)
  ax.add_patch(rect17)
  ax.add_patch(rect18)
  ax.add_patch(rect19)
  ax.add_patch(rect20)
  ax.add_patch(rect21)
  ax.add_patch(rect22)
  ax.add_patch(rect23)
  ax.add_patch(rect24)
  ax.add_patch(rect25)
  ax.add_patch(rect26)
  ax.add_patch(rect27)
  ax.add_patch(rect28)
  ax.add_patch(rect29)
  ax.add_patch(rect30)
  ax.add_patch(rect31)
  ax.add_patch(rect32)
  ax.add_patch(rect33)
  ax.add_patch(rect34)
  ax.add_patch(rect35)
  ax.add_patch(rect36)
  st.pyplot(figure)

  i=0
  dict_val={}
  for key, value in coordinates.items():
    x = crop_save_test(image = page1, text = key, coordinates = value)
    dict_val[list(coordinates.keys())[i]]=x
    i+=1
  
  df_output = pd.DataFrame.from_dict(dict_val.items())
  df_output = df_output.set_index(0)
  df_output=df_output.transpose()
  df_output.reset_index(drop=True)
  st.text("Risultato dell'OCR in forma tabellare")
  st.write(df_output)
  st.text("Export in excel")
  st.markdown(get_table_download_link(df_output), unsafe_allow_html=True)
