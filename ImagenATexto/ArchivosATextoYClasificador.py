import os
import sys
import fitz  # PyMuPDF para procesar PDFs
import shutil
import pytesseract
from PIL import Image
from datetime import datetime


######## FUNCIONALIDAD ########

# Este codigo transforma los ficheros con extension .png .jpg .jpeg .pdf a texto, que encuentre en la carpeta de la variable "sRutaImg" y 
# despues tambien podemos clasificar esos ficheros en carpetas dependiendo del texto que contengan.
# El texto transformado se guarda en ResTextoObtenido.txt


######## CONFIGURACION ########

sRutaImg = r'C:\Users\lucas\VisualStudio\CosasVarias\ImageToText' # Ubicacion de la carpeta donde queremos buscar los ficheros
pytesseract.pytesseract.tesseract_cmd = r'C:\Users\lucas\AppData\Local\Programs\Tesseract-OCR\tesseract.exe' # Ubicacion de tesseract
sys.stdout.reconfigure(encoding='utf-8') # Configurar la codificación de la salida estándar a UTF-8

###############################

def fObtenerDiayHoraActual():
    # Obtener la fecha y hora actual
    HoraActual = datetime.now()
    # Formatear la fecha y hora actual como un solo string
    sFechaActual = HoraActual.strftime("%Y-%m-%d %H:%M:%S")
    return sFechaActual



def fAñadirLog(sTextoLog):
    with open(r'C:\Users\lucas\VisualStudio\CosasVarias\ImageToText\ResTextoObtenido.txt', 'a') as ArchivoLog:
        ArchivoLog.write(sTextoLog + '\n')



def fTransformarImagen(sRutaCompleta):
    try:
        # Realizar OCR en la imagen
        sTexto = pytesseract.image_to_string(sRutaCompleta)
        #print(sTexto)
        return True, sTexto
    
    except pytesseract.TesseractError:
        print(f'UPS!! Error al realizar OCR \n')
        sTexto = ""
        return False, sTexto



def fClasificarDocu(sRutaImg):

    bTexto = False
    sTextoCompleto = ""


    sTextoIni = f'\n ···················· Inicio pytesseract: {fObtenerDiayHoraActual()} ····················'
    
    # Obtener la lista de archivos en la carpeta
    lListaArchivos = os.listdir(sRutaImg)

    if lListaArchivos:
        # Filtrar archivos de imagen
        lArchivosFiltrado = [sArchivo for sArchivo in lListaArchivos if sArchivo.lower().endswith(('.png', '.jpg', '.jpeg', '.pdf'))]

        if lArchivosFiltrado:
            for sNombreImg in lArchivosFiltrado:

                sRutaCompleta = os.path.join(sRutaImg, sNombreImg)

                # Verificar si el archivo es un .png .jpg .jpeg
                if sNombreImg.lower().endswith(('.png', '.jpg', '.jpeg')):
                    sLogNombreImg = f'{sTextoIni} \n #### Archivo Imagen: {sNombreImg} ####\n'
                    print(sLogNombreImg)
                    #Transformar Imagen
                    bTexto, sTexto = fTransformarImagen(sRutaCompleta)
                    sTextoLog = f'{sLogNombreImg} \n {sTexto}'
                    fAñadirLog(sTextoLog)


                # Verificar si el archivo es un PDF
                if sNombreImg.lower().endswith('.pdf'):
                    sLogNombreImg = f'{sTextoIni} \n #### Archivo PDF: {sNombreImg} ####\n'
                    print(sLogNombreImg)
                    try:
                        # Abrir el archivo PDF
                        doc = fitz.open(sRutaCompleta)
  
                        # Iterar sobre todas las páginas del PDF
                        for i in range(doc.page_count):
                            # Obtener la página actual del PDF
                            page = doc.load_page(i)

                            # Convertir la página a una imagen (pixmap)
                            pixmap = page.get_pixmap()

                            # Convertir el pixmap a un objeto de imagen de Pillow
                            image = Image.frombytes("RGB", [pixmap.width, pixmap.height], pixmap.samples)

                            # Guardar la imagen en formato PNG
                            sRutaPngPdf = os.path.join(sRutaImg, f'tmp_{sNombreImg.replace('.pdf', '')}_{i+1}.png')
                            image.save(sRutaPngPdf, format="PNG")

                            #Transformar Imagen
                            bTexto, sTexto = fTransformarImagen(sRutaPngPdf) 

                            # Agregar el texto de la página al texto completo (a paginas anteriores)
                            sTextoCompleto = sTextoCompleto + sTexto + "\n"
                            print(f'-- PaginaPDF a PNG: {sRutaPngPdf}\n')
                            # Eliminar la imagen temporal
                            os.remove(sRutaPngPdf)
                        
                        # Cerrar el archivo PDF después de terminar
                        doc.close()
                        
                        sTextoLog = f'{sLogNombreImg} \n {sTextoCompleto}'
                        fAñadirLog(sTextoLog)

                    except Exception as e:
                        print(f'UPS!! Error al procesar PDF: {e}\n')


                if bTexto:
                    # Verificamos si el texto contiene ciertas palabras y movemos el archivo a la carpeta correspondiente
                    if 'Acrobat'.lower() in sTexto.lower():
                        sMoverACarpeta = '\\Acrobat'
                    elif 'Discord'.lower() in sTexto.lower():
                        sMoverACarpeta = '\\Discord'
                    elif 'México'.lower() in sTexto.lower():
                        sMoverACarpeta = '\\Mexico'
                    else:
                        sMoverACarpeta = '\\imgProcesado'

                    sCarpetaDestino = sRutaImg + sMoverACarpeta


                    # Verificamos si la carpeta de destino existe
                    if os.path.exists(sCarpetaDestino):
                        # Obtener la ruta completa del archivo en la carpeta de destino
                        sCompArchRepe = os.path.join(sCarpetaDestino, sNombreImg)
                        
                        # Verificar si el archivo que queremos mover ya existe en la carpeta de destino
                        if os.path.exists(sCompArchRepe):

                            # Dividir el nombre del archivo y su extensión
                            sNombreBase, sExtension = os.path.splitext(sNombreImg)

                            iContador = 1
                            
                            while True:
                                # Construir el nuevo nombre de archivo con el iContador
                                sNombreNuevo = f"{sNombreBase} ({iContador}){sExtension}"
                                if not os.path.exists(os.path.join(sCarpetaDestino, sNombreNuevo)):
                                    break
                                iContador += 1

                            # Mover el archivo con el nuevo nombre a la carpeta de destino
                            shutil.move(sRutaCompleta, os.path.join(sCarpetaDestino, sNombreNuevo))
                            print(f"-- El archivo se ha movido a la carpeta: '{sMoverACarpeta}' con el nuevo nombre '{sNombreNuevo}'\n")

                        else:
                            # Mover el archivo a la carpeta de destino
                            shutil.move(sRutaCompleta, sCarpetaDestino)
                            print(f"-- El archivo se ha movido a la carpeta: '{sMoverACarpeta}'\n")
                    else:
                        print(f"UPS!! La carpeta '{sMoverACarpeta}' no existe en el sistema. \n")
                else:
                    print(f"UPS!! Ocurrio un error al procesar la imagen.\n")
        else:
            print(f"UPS!! No hay ningun archivo con las extensiones ('.png', '.jpg', '.jpeg', '.pdf') en la carpeta: {sRutaImg} \n")
    else:
        print(f"UPS!! No hay ningun archivo en esta carpeta. \n")

fClasificarDocu(sRutaImg)


print(f" -- Ejecucion Finalizada -- ")