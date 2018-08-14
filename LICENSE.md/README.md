# -*- coding: utf-8 -*-
"""
Editor de Spyder

Este es un archivo temporal
"""
# -*- coding: utf-8 -*-
"""
Created on Sun Aug  5 20:12:40 2018

@author: LINDA
"""
import urllib
import zipfile
import pandas as pd
from os import scandir
from os.path import abspath
import tempfile
from selenium import webdriver
import os
import bs4
import re
import requests
from timeit import default_timer as timer
from PIL import Image
import pytesseract
pytesseract.pytesseract.tesseract_cmd = r'D:\Program Files (x86)\Tesseract-OCR\tesseract.exe'
os.environ['http_proxy'] = ""

class Sunat:
    def __init__(self,web_driver):
        self.web_driver = web_driver
        self.url = 'http://e-consultaruc.sunat.gob.pe/cl-ti-itmrconsmulruc/jrmS00Alias'
    
    def get_information(self,file_zip,contador_error_captcha):
        try:
            df = None
            #Inicializando captcha
            captcha_text=""
            #Bucle para validar captcha
            captcha_text_contador=1
            
            while (captcha_text=="" or len(captcha_text)!=4) and captcha_text_contador<4:
            
                self.web_driver.get(self.url)
                #print('0')
                captcha_text = self.solve_captcha(self.web_driver)
                #print('1')
                
                if captcha_text=="" or len(captcha_text)!=4:
                    print("Captcha resuelto incorrectamente")
                    captcha_text_contador=captcha_text_contador+1
                else:
                    print("Captcha resuelto correctamente")
                    resultado_1=self.submit(file_zip,captcha_text)
                    error=self.get_error_captcha(resultado_1)
                    
                    if error=='SI' and contador_error_captcha<4:
                        print("Error_Captcha: El Codigo ingresado es incorrecto: ",contador_error_captcha)
                        contador_error_captcha = contador_error_captcha + 1
                        df = self.get_information(file_zip,contador_error_captcha)
                    elif error == "SI" and contador_error_captcha==4:
                        return None
                    else:
                        print('Hay resultados')
                        my_folder=r'D:\Users\C24778\ConsultaRuc\archivo_down_temp'
                        rucs_val_file_name,rucs_inval_file_name=self.get_file_result(my_folder)
                        
                        df=pd.read_csv(rucs_val_file_name,delimiter="|")
                        
                        #df['trabajadores']= df.apply(lambda row: self.get_trabajadores(row['NumeroRuc'],row['Nombre ó RazonSocial']),axis=1)
                        #df['representantes']= df.apply(lambda row: self.get_representantes(row['NumeroRuc'],row['Nombre ó RazonSocial']),axis=1)
                        
                        print(rucs_val_file_name)
                        
                        os.remove(rucs_val_file_name)
                        #os.remove(rucs_inval_file_name)
                        break
            return df
        
        except:
            return None
        
    def solve_captcha(self,web_driver):
        img_xpath = '//form[@name="frmConsMulRucArchivo"]//img[@src="captcha?accion=image"]'
        print("entro al captcha")
        #Tomar un screencshot
        screenshot_temp_name = tempfile.NamedTemporaryFile().name+'.png'
        print(screenshot_temp_name)
        web_driver.save_screenshot(screenshot_temp_name, delete=False)
        #Encontrar la imagen del captcha
        img_frame = web_driver.find_element_by_xpath(img_xpath)
        print(img_frame)
        loc = img_frame.location
        size = img_frame.size
        
        
        #Recortar la imagen
        img = Image.open(screenshot_temp_name)
        
        left = loc['x']
        top = loc['y']
        right = loc['x'] + size['width']
        bottom = loc['y'] + size['height']
        
        img = img.crop((left,top,right,bottom))
        #print('2')
        img.save(r'D:\Users\C24778\ConsultaRuc\Captcha\captcha.png')
        os.environ['TESSDATA_PREFIX']=r"D:\Program Files (x86)\Tesseract-OCR\tessdata"
        captcha_text=pytesseract.image_to_string(img,config="-c tessedit_char_whitelist=abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ -psm 6").replace(' ', '').upper() 
        #print(captcha_text)
        os.remove(screenshot_temp_name)
        os.environ['TESSDATA_PREFIX']=""
        return captcha_text
    
    def submit(self,file_zip,captcha_text):
        box_file=self.web_driver.find_element_by_xpath('//input[@name="archivo"]')
        box_captcha=self.web_driver.find_element_by_xpath('//input[@name="codigoA"]')
        
        box_file.send_keys(file_zip)
        box_captcha.send_keys(captcha_text)
        
        boton=self.web_driver.find_element_by_xpath('//input[@onclick="enviarArchivo()"]')
        boton.click()
        
        html = bs4.BeautifulSoup(self.web_driver.page_source,'lxml')
        
        return html
    
    def get_error_captcha(self, html):
        tag = html.find(
            'p',
            {'class':'error'},
            text=re.compile('el\s+c[ó|o]digo\s+ingresado\s+es\s+incorrecto', re.IGNORECASE))
        
        if tag is None:
            error = "NO"
        else:
            error = "SI"
        return error 
    
    def get_file_result(self,my_folder):
        
        href_element=self.web_driver.find_element_by_partial_link_text('.zip')
        link=href_element.get_attribute('href')
        
        name_file_zip='file_sunat.zip'
        
        os.environ['http_proxy'] = 'claro-proxy'
        urllib.request.urlretrieve(link, os.path.join(my_folder,name_file_zip))
        file_zip = zipfile.ZipFile(os.path.join(my_folder, name_file_zip), 'r')
        file_zip.extractall(my_folder)
        file_zip.close()
        os.remove(os.path.join(my_folder, name_file_zip))
        
        rucs_val_file_name, rucs_inval_file_name = self.archivo_ruc(my_folder)
        os.environ['http_proxy'] = ""
        return os.path.join(my_folder,rucs_val_file_name), os.path.join(my_folder,rucs_inval_file_name)
        
    
    def archivo_ruc(self,ruta):
        rucs_val_file_name=""
        rucs_inval_file_name=""
        for arch in scandir(ruta):
            #los archivos de rucs validos empiezan con RM
            if arch.is_file() and list(arch.name)[0] == 'R' and list(arch.name)[1] == 'M':
                rucs_val_file_name=abspath(arch.path)
            elif arch.is_file():
                rucs_inval_file_name=abspath(arch.path)
                
        return rucs_val_file_name, rucs_inval_file_name
     
    def get_trabajadores(self,ruc, nombre):
        params = {
            'accion': 'getCantTrab',
            'nroRuc': ruc,
            'desRuc': nombre,
        }
        
        trabajadores={}
        print(ruc)
        try:
            #print(ruc)
            http_proxy  = 'http:claro-proxy'
            proxyDict = {'http': http_proxy}
            res = requests.get('http://e-consultaruc.sunat.gob.pe:80/cl-ti-itmrconsruc/jcrS00Alias', params,proxies=proxyDict) #time=5
            soup = bs4.BeautifulSoup(res.text, 'lxml')
            #print(soup)
        
            tables=soup.find_all('table') #lista de tablas
            results_table = tables[1] # tabla con el titulo o mensaje de la tabla
    
            existe_data=results_table.find('td')
            
            if existe_data.get_text().strip().startswith('No'):
                return None
            else:
                results_table_2=tables[2] # tabla de trabajadores
                #print(results_table_2)
                rows=results_table_2.find_all('tr')[1:] #devuelve una lista de filas de la tabla de trabajadores
                last_position=len(rows)-1
                cantidad=0
                
                while cantidad==0 and last_position>1:
                    last_row=rows[last_position]
                    cantidad=self.validar_int(last_row.find_all('td')[1].get_text())
                    trabajadores['cantidad']=cantidad
                    trabajadores['periodo']=last_row.find_all('td')[0].get_text()
                    last_position=last_position-1
                    #print('Cantidad de trabajadores: ',trabajadores['cantidad'], '  Periodo: ',trabajadores['periodo'])
                
                
                return trabajadores
        except:
            return None
        
        
    
    def validar_int(self,num):
        try:
            int(num)
            return int(num)
        except:
            return 0
        
    def get_representantes(self,ruc, nombre):
        params = {
            'accion': 'getRepLeg',
            'nroRuc': ruc,
            'desRuc': nombre,
        }
        
        try:
            http_proxy  = 'http:claro-proxy'
            proxyDict = {'http': http_proxy}
            res = requests.get('http://e-consultaruc.sunat.gob.pe:80/cl-ti-itmrconsruc/jcrS00Alias', params, proxies=proxyDict )
            soup = bs4.BeautifulSoup(res.text, 'lxml')
        
        
            tables=soup.find_all('table') #lista de tablas
            
            if len(tables)==0:
                representantes = None
                return representantes
            else:
                results_table_2=tables[2] 
                #print(tables[2] )
                rows=results_table_2.find_all('tr')[1:]
                i=1
                representantes = {}
                while i<len(rows):
                    datos={}
                    row = rows[i].find_all('td')
                    datos['nombre']=row[2].get_text().strip()
                    datos['cargo']=row[3].get_text().strip()
                    datos['tipo_documento']=row[0].get_text().strip()
                    datos['nro_documento']=int(row[1].get_text().strip())
                    representantes['R'+str(i)]=datos
                    #print(representantes)
                    i=i+1
            #print(ruc)
            return representantes
        except:
            representantes = None
            return representantes
        
        
        
start=timer()    


driver = webdriver.PhantomJS()
sunat= Sunat(driver)
ruta_archivos=r"D:\Users\C24778\ConsultaRuc\archivos"
df_new=None
i=1
for arch in scandir(ruta_archivos):
   
    if arch.is_file():
        print(i)
        print(arch.name)
        df=sunat.get_information(os.path.join(ruta_archivos,arch.name),1)
        columns=['NumeroRuc', 'Nombre ó RazonSocial', 'Tipo de Contribuyente',
       'Profesion u Oficio', 'Nombre Comercial', 'Condicion del Contribuyente',
       'Estado del Contribuyente', 'Fecha de Inscripcion',
       'Fecha de Inicio de Actividades', 'Departamento', 'Provincia',
       'Distrito', 'Direccion', 'Telefono', 'Fax',
       'Actividad de Comercio Exterior', 'Principal- CIIU',
       'Secundario 1- CIIU', 'Secundario 2- CIIU', 'Afecto Nuevo RUS',
       'Buen Contribuyente', 'Agente de Retencion',
       'Agente de Percepcion VtaInt', 'Agente de Percepcion ComLiq']
        #, 'trabajadores', 'representantes'
        
        df_new=pd.concat([df_new,df[columns]])
        i=i+1
        print("----------------------------------------------------------")
#print(sunat.get_information(file_zip,1))
end=timer()
print(end-start)
