import shutil
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import pandas as pd
from PIL import Image, ImageTk
import easyocr
import fitz  
import os

def crear_archivo_excel_si_no_existe(nombre_archivo):
    dir_actual = os.path.dirname(os.path.abspath(__file__))
    ruta_archivo = os.path.join(dir_actual, nombre_archivo)
    if not os.path.isfile(ruta_archivo):
        df_vacio = pd.DataFrame()
        try:
            df_vacio.to_excel(ruta_archivo, index=False)
            print(f"Se ha creado el archivo '{nombre_archivo}' en el directorio actual.")
        except Exception as e:
            print("Se produjo un error al intentar crear el archivo:", e)
    else:
        print(f"El archivo '{nombre_archivo}' ya existe en el directorio actual.")

crear_archivo_excel_si_no_existe("comprobantes.xlsx")

def calculate_window_position(ventana, wventana, hventana):
    wtotal = ventana.winfo_screenwidth()
    htotal = ventana.winfo_screenheight()
    pwidth = round(wtotal/2-wventana/2)
    pheight = round(htotal/2-hventana/2)
    return pwidth, pheight

def crear_interfaz():
    def on_cerrar():
        delete_files()
        ventana.destroy()
    ventana = tk.Tk()
    ventana.title("APCODA-SALTA CODERS")
    wventana = 800
    hventana = 650
    pwidth, pheight = calculate_window_position(ventana, wventana, hventana)
    ventana.geometry(f"{wventana}x{hventana}+{pwidth}+{pheight}")  # Modificar el tamaño de la ventana
    ventana.iconbitmap('logo.ico')
    # marco de la izquierda
    marco_izquierdo = tk.Frame(ventana, bg="#461e5a", width=250)
    marco_izquierdo.pack(side="left", fill="y")
    marco_izquierdo.pack_propagate(False)
    # opciones de comprobantes
    opciones = [
        "Banco/Billetera",
        "MACRO",
        "MERCADO PAGO",
        "NACION+",
        "SANTANDER",
        ]
    # ComboBox de opciones
    combo_comprobantes = ttk.Combobox(marco_izquierdo, state='readonly', width=50)
    combo_comprobantes.pack(pady=30, padx=20)
    combo_comprobantes['values'] = opciones
    combo_comprobantes.current(0)
    combo_comprobantes.bind("<<ComboboxSelected>>",
                            lambda event: select_opcion(combo_comprobantes.get(), ventana, marco_izquierdo))
    ventana.combo = combo_comprobantes
    # Reinicia el excel
    df = pd.read_excel("comprobantes.xlsx", engine='openpyxl')
    df = df.head(0)
    df.to_excel("comprobantes.xlsx", index=False)
    # contenedor label
    contenedor_lbl = tk.LabelFrame(marco_izquierdo,bg="snow2",width=230,borderwidth= 2,text="Datos")
    contenedor_lbl.pack(pady=10,padx=5, expand=True)
    # Label nombre
    contenedor_nombre = tk.LabelFrame(contenedor_lbl, bg="gainsboro", borderwidth=2, relief="groove")
    contenedor_nombre.pack(pady=10,padx=5)
    lbl_nombre_tit = tk.Label(contenedor_nombre, text="Nombre del Cliente: ",width=50, anchor="w", bg="gainsboro")
    lbl_nombre_tit.pack()
    entry_nombre = tk.Entry(contenedor_nombre, text="", bg="gainsboro",width=51)
    entry_nombre.pack(fill="both")
    entry_nombre.config(state="readonly")
    # Label numero
    contenedor_numero = tk.LabelFrame(contenedor_lbl, bg="gainsboro", borderwidth=2, relief="groove")
    contenedor_numero.pack(pady=10,padx=5)
    lbl_numero_tit = tk.Label(contenedor_numero, text= "CUIT:", anchor="w",width=50,bg="gainsboro")
    lbl_numero_tit.pack()
    entry_numero = tk.Entry(contenedor_numero, text= "",width=50,bg="gainsboro")
    entry_numero.pack(fill="both")
    entry_numero.config(state="readonly")   
    ventana.lblNumero = lbl_numero_tit
    # Label importe
    contenedor_importe = tk.LabelFrame(contenedor_lbl, bg="gainsboro", borderwidth=2, relief="groove")
    contenedor_importe.pack(pady=10,padx=5)
    lbl_importe_tit = tk.Label(contenedor_importe, text= "Importe Total:",width=50, anchor="w",bg="gainsboro")
    lbl_importe_tit.pack()
    entry_importe = tk.Entry(contenedor_importe, text= "",width=50,bg="gainsboro")
    entry_importe.pack(fill="both")
    entry_importe.config(state="readonly")
    # Label fecha
    contenedor_fecha = tk.LabelFrame(contenedor_lbl, bg="gainsboro", borderwidth=2, relief="groove")
    contenedor_fecha.pack(pady=10,padx=5)
    lbl_fecha_tit = tk.Label(contenedor_fecha, text= "Fecha:",width=50, anchor="w",bg="gainsboro")
    lbl_fecha_tit.pack()
    entry_fecha = tk.Entry(contenedor_fecha, text= "",width=50,bg="gainsboro")
    entry_fecha.pack(fill="both")
    entry_fecha.config(state="readonly")
    contenedor_lbl.nombre = entry_nombre
    contenedor_lbl.numero = entry_numero
    contenedor_lbl.importe = entry_importe
    contenedor_lbl.fecha = entry_fecha
    marco_izquierdo.contenedor = contenedor_lbl
    lbl_cant = tk.Label(marco_izquierdo,bg="#461e5a", text="")
    lbl_cant.pack(pady=10)
    ventana.cant = lbl_cant
    # Botones para pasar de imagen
    frame_siguiente = tk.Frame(marco_izquierdo)
    boton_ant = tk.Button(frame_siguiente,text="Anterior", command=lambda:next_previus(ventana,marco_izquierdo,"anterior"))
    boton_ant.pack(side=tk.LEFT, padx=5)
    boton_ant.config(state="disabled")
    boton_sig = tk.Button(frame_siguiente,text="Siguente", command=lambda:next_previus(ventana,marco_izquierdo,"siguiente"))
    boton_sig.pack(side=tk.LEFT, padx=5)
    boton_sig.config(state="disabled")
    frame_siguiente.sig = boton_sig
    frame_siguiente.ant = boton_ant
    ventana.frame_siguiente = frame_siguiente
    # Boton para guardar en excel
    boton_guardar = tk.Button(marco_izquierdo, text="Guardar en Excel", command=lambda: guardar_en_excel(ventana,marco_izquierdo),width=15)
    boton_guardar.pack(pady=(0, 10))
    boton_guardar.pack_forget()
    boton_guardar.config(state="disabled")
    ventana.guardar = boton_guardar
    # Botón para extraer en Excel
    boton_excel = tk.Button(marco_izquierdo, text="Extraer en Excel", command=lambda: extraer_en_excel(ventana),width=15)
    boton_excel.pack(pady=(0, 10))
    boton_excel.pack_forget()
    boton_excel.config(state="disabled")
    ventana.excel = boton_excel
    # contenedor de los botones que controlan la imagen
    contenedor_botones = tk.Frame(ventana)
    contenedor_botones.pack(pady=5)
    # Crear un objeto EasyOCR
    reader = easyocr.Reader(['en'])
    # Botón para cargar el comprobante en formato imagen
    boton_cargar_img = tk.Button(contenedor_botones, text="Cargar Imagen",
                             command=lambda: cargar_comprobante(ventana,marco_izquierdo, "img"))
    boton_cargar_img.pack(side="left", pady=5, padx=5)
    boton_cargar_img.config(state="disabled")
    ventana.btn_img = boton_cargar_img
    #Botón para cargar el comprobante en formato pdf
    boton_cargar_pdf = tk.Button(contenedor_botones, text="Cargar PDF",
                             command=lambda: cargar_comprobante(ventana,marco_izquierdo, "pdf"))
    boton_cargar_pdf.pack(side="left", pady=5, padx=5)
    boton_cargar_pdf.config(state="disabled")
    ventana.btn_pdf = boton_cargar_pdf
    # Botón para extraer la información
    boton_extraer = tk.Button(contenedor_botones, text="Extraer Información",
                          command=lambda: extraer_informacion(ventana, combo_comprobantes.get(), marco_izquierdo, reader))
    boton_extraer.pack(side="left", pady=5, padx=5)
    boton_extraer.config(state="disabled")
    ventana.extraer = boton_extraer
    # boton para volver atras (mi idea es que los botones de opciones se bloquen)
    boton_atras = tk.Button(contenedor_botones, text="Atras", command=lambda: atras(ventana,marco_izquierdo))
    ventana.atras = boton_atras
    boton_atras.pack(side="left",pady = 5, padx= 5)
    boton_atras.config(state="disabled")
    # mensaje que dice de donde es el comprobante
    lbl_opcion = tk.Label(ventana, text="Aquí se vera el comprobante")
    lbl_opcion.pack(padx=5, pady=5)
    ventana.lbl = lbl_opcion
    # Crear la previsualización de los comprobantes
    previsualizacion = tk.Canvas(ventana,width=530,height=553)
    previsualizacion.pack(padx=10, pady=10)
    ventana.previsualizacion = previsualizacion  # Guardar referencia a la etiqueta en la ventana
    ventana.name = ""
    ventana.formato = ""
    ventana.path = ""
    ventana.lista = []
    ventana.lista_excel = []
    ventana.banco = ""
    ventana.ext = False
    ventana.indice = 0
    temp_folder_path = 'temp_uploads'
    if not os.path.exists(temp_folder_path):
                os.makedirs(temp_folder_path)
    ruta_carpeta = 'temp_uploads'
    # Obtener la lista de archivos en la carpeta
    archivos = os.listdir(ruta_carpeta)
    # Iterar sobre cada archivo y borrarlo
    for archivo in archivos:
        ruta_archivo = os.path.join(ruta_carpeta, archivo)
        os.remove(ruta_archivo)     
    ventana.protocol("WM_DELETE_WINDOW", on_cerrar)
    # Iniciar el bucle de eventos
    ventana.mainloop()

def select_opcion(opcion, ventana,marco):
    # agregar la lógica para cargar los comprobantes
    if opcion != "Banco/Billetera":
        combo = ventana.combo
        combo.configure(state="disabled")
        cargar_img = ventana.btn_img
        cargar_pdf = ventana.btn_pdf
        boton_atras = ventana.atras
        siguiente = ventana.frame_siguiente
        cantidad = ventana.cant
        ventana.excel.pack_forget()
    boton_atras.config(state="normal") 
    name = None     
    match opcion:
        case "MERCADO PAGO":
            marco.config(bg="#02aff3")
            siguiente.configure(bg="#02aff3")
            cantidad.config(bg="#02aff3")
            cargar_img.config(state="normal")
            cargar_pdf.config(state="normal")
            name = "img\\mp.jpeg"
            
        case "NACION+":
            marco.config(bg="#017994")
            siguiente.configure(bg="#017994")
            cantidad.config(bg="#017994")
            cargar_img.config(state="normal")
            cargar_pdf.config(state="normal")
            name = "img\\nacion+.jpeg"
        case "MACRO":
            marco.config(bg="#002c53")
            siguiente.configure(bg="#002c53")
            cantidad.config(bg="#002c53")
            cargar_img.config(state="normal")
            cargar_pdf.config(state="normal")
            name = "img\\macro.jpg"    
        case "SANTANDER":
            marco.config(bg = "#ed1b24")
            siguiente.configure(bg ="#ed1b24")
            cantidad.config(bg="#ed1b24")
            cargar_img.config(state="normal")
            cargar_pdf.config(state="normal")
            name = "img\\santander.jpeg"

    if name != None:
        lbl_opcion = ventana.lbl
        lbl_opcion.config(text="Esta opcion acepta este tipo de comprobante")
        archivo = Image.open(name)
        if archivo:
            # Reescalo la imagen
            ancho_original, alto_original = archivo.size
            nuevo_ancho = int((ancho_original / alto_original) * 553)
            archivo = archivo.resize((nuevo_ancho, 553))
            imagen_tk = ImageTk.PhotoImage(archivo)
            #cargo la imagen
            previsualizacion = ventana.previsualizacion
            previsualizacion.create_image(530/2-(nuevo_ancho/2), 0, anchor=tk.NW, image=imagen_tk)
            previsualizacion.image = imagen_tk
            previsualizacion.pack()
  
def restablecer(ventana,marco):
    lbl_opcion = ventana.lbl
    lbl_opcion.config(text="Aquí se vera el comprobante")
    # Borro imagen cargada
    previsualizacion = ventana.previsualizacion
    previsualizacion.config(image = None)
    previsualizacion.image = None
    # Reinicio datos
    contenedor = marco.contenedor
    nombre = contenedor.nombre
    numero = contenedor.numero
    importe = contenedor.importe
    fecha = contenedor.fecha
    nombre.config(state="normal")
    numero.config(state="normal")
    importe.config(state="normal")
    fecha.config(state="normal")
    nombre.delete(0, tk.END)
    numero.delete(0, tk.END)
    importe.delete(0, tk.END)
    fecha.delete(0, tk.END)
    ventana.banco = ""
    nombre.config(text="")
    numero.config(text="")
    importe.config(text="")
    fecha.config(text="")    
    # oculto botones ant y sig
    ventana.frame_siguiente.pack_forget()
    ventana.frame_siguiente.sig.config(state="disabled")
    ventana.excel.config(state="disabled")
    ventana.ext = False
    ventana.cant.config(text = "")
    # Borro lista de datos
    ventana.lista = []
    delete_files()
    df = pd.read_excel("comprobantes.xlsx", engine='openpyxl')
    df = df.head(0)
    df.to_excel("comprobantes.xlsx", index=False)
                  
def guardar_comprobantes (ventana, marco, comando):
    ventana.lista = []
    temp_folder_path = 'temp_uploads'
    delete_files()
    if comando == "img": 
        # Abro muchos archivos y los cargo
        files = filedialog.askopenfilenames(title="Seleccionar archivos",filetypes=[("Archivos de Imagen", "*.jpg;*.jpeg;*.png;*.bmp")])
        if files:
            for file in files:
                filename = os.path.basename(file)
                # Mover el archivo a la carpeta tempora
                shutil.copy(file, os.path.join(temp_folder_path, filename))
            # consigo el path de la primera imagen
            name = os.listdir(temp_folder_path)[0]  
            ventana.name = name
            filePath  = os.path.join(temp_folder_path, name)
            archivo = Image.open(filePath)
            lbl_opcion = ventana.lbl
            lbl_opcion.config(text=" ")
            if archivo:
                # Reescalo la imagen
                ancho_original, alto_original = archivo.size
                nuevo_ancho = int((ancho_original / alto_original) * 553)
                archivo = archivo.resize((nuevo_ancho, 553))
                imagen_tk = ImageTk.PhotoImage(archivo)            
                #cargo la imagen
                previsualizacion = ventana.previsualizacion
                previsualizacion.create_image(530/2-(nuevo_ancho/2), 0, anchor=tk.NW, image=imagen_tk)
                previsualizacion.image = imagen_tk
                previsualizacion.pack()
                ventana.path = filePath
                messagebox.showinfo("Exito", f"Se han cargado {len(os.listdir(temp_folder_path))} archivos")
                
        else: 
            messagebox.showwarning("Cuidado","No se han cargado archivos")
            restablecer(ventana,marco)
        ventana.formato = "img" 
    else:
        files = filedialog.askopenfilenames(title="Seleccionar archivos",filetypes=[("Archivos PDF", "*.pdf")])
        if files: 
            for file in files:
                filename = os.path.basename(file)
                shutil.copy(file, os.path.join(temp_folder_path, filename))
            name = os.listdir(temp_folder_path)[0]  
            ventana.name = name
            filePath  = os.path.join(temp_folder_path, name)   
            if name:
                doc = fitz.open(filePath)
                for pagina in doc:
                    pix = pagina.get_pixmap()
                    imagen = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                    proporcion = 553 / imagen.height
                    ancho = imagen.width * proporcion
                    imagen = imagen.resize((int(ancho), 553))
                    imagen_tk = ImageTk.PhotoImage(imagen)
                    previsualizacion = ventana.previsualizacion
                    previsualizacion.create_image(530/2 - (ancho/2), 0, anchor=tk.NW, image=imagen_tk)
                    previsualizacion.image = imagen_tk
                    previsualizacion.pack()
                archivo = open("pagina.txt", 'w', encoding='utf-8')
                archivo.seek(0)
                for i in range(len(doc)):
                    page = doc.load_page(i)
                    archivo.write(page.get_text())
                doc.close() 
            ventana.formato = "pdf"
            messagebox.showinfo("Exito", f"Se han cargado {len(os.listdir(temp_folder_path))} archivos")
            
        else:
            messagebox.showwarning("Cuidado","No se han cargado archivos")
            restablecer(ventana,marco)     
    boton_extraer = ventana.extraer
    boton_extraer.config(state="normal")            
    guardar = ventana.guardar
    siguiente = ventana.frame_siguiente
    ventana.excel.pack_forget()
    guardar.pack_forget()
    siguiente.pack(pady= 10)
    guardar.pack(pady = 10)
    ventana.frame_siguiente.ant.config(state="disabled")
    ventana.frame_siguiente.sig.config(state="disabled")
    ventana.excel.config(state="disabled")
    # Reinicio datos
    contenedor = marco.contenedor
    nombre = contenedor.nombre
    numero = contenedor.numero
    importe = contenedor.importe
    fecha = contenedor.fecha
    nombre.config(state="normal")
    numero.config(state="normal")
    importe.config(state="normal")
    fecha.config(state="normal")
    nombre.delete(0, tk.END)
    numero.delete(0, tk.END)
    importe.delete(0, tk.END)
    fecha.delete(0, tk.END)
    ventana.banco = ""
    nombre.config(text="")
    numero.config(text="")
    importe.config(text="")
    fecha.config(text="")
    cant = ventana.cant
    cant.config(text=f"01 de {str(len(os.listdir(temp_folder_path))).zfill(2)} ")  
    contenedor = marco.contenedor
    nombre = contenedor.nombre
    numero = contenedor.numero
    importe = contenedor.importe
    fecha = contenedor.fecha
    nombre.config(state="disabled")
    numero.config(state="disabled")
    importe.config(state="disabled")
    fecha.config(state="disabled")
    
def cargar_comprobante(ventana, marco, comando):
    if ventana.ext:
        resultado = messagebox.askokcancel("Confirmar", "Si continua los datos NO guardados se perderan\n¿Estás seguro de que deseas continuar?")
        if resultado:
            guardar_comprobantes(ventana, marco, comando)
    else:
        guardar_comprobantes(ventana, marco, comando)
                          
def read_text_from_image(reader, filePath):
    with open("pagina.txt", 'w', encoding='utf-8') as archivo:
        resultado = reader.readtext(filePath)
        for detection in resultado:
            archivo.write(detection[1] + "\n")

def read_text_from_pdf(filePath):
    with open("pagina.txt", 'w', encoding='utf-8') as archivo:
        doc = fitz.open(filePath)
        for page in doc:
            archivo.write(page.get_text())

def extraer_informacion(ventana, opcion, marco, reader):    
    lista = []
    for name in os.listdir('temp_uploads'):
        filePath  = os.path.join('temp_uploads', name)
        if ventana.formato == "img":
            read_text_from_image(reader, filePath)
        else:
            read_text_from_pdf(filePath)
        datos = {
            "clientName":None,
            "numeroCuenta":None,
            "importe_total":None,
            "fecha" : None,
            "banco" : None
        }
        with open("pagina.txt", 'r', encoding='utf-8')as archivo:
            archivo.seek(0)
            lineas = archivo.readlines()        
        if opcion == "MERCADO PAGO":
            datos["clientName"], datos["numeroCuenta"], datos["importe_total"], datos["fecha"] = mercado_pago(lineas)
            datos["banco"] = "Mercado Pago"
            ventana.banco = "Mercado Pago"
        elif opcion == "NACION+":
            datos["clientName"],datos["numeroCuenta"], datos["importe_total"], datos["fecha"] = banco_nacion_plus(lineas)
            ventana.banco = "NACION+"
            datos["banco"] = "NACION+"
        elif opcion == "MACRO":
            datos["clientName"],datos["numeroCuenta"], datos["importe_total"], datos["fecha"]  = banco_macro(lineas,ventana.formato)
            ventana.banco = "Banco Macro"
            datos["banco"] = "Banco Macro"
        elif opcion == "SANTANDER":
            datos["clientName"],datos["numeroCuenta"], datos["importe_total"], datos["fecha"]  = banco_santander(lineas)
            ventana.banco = "Santander"
            datos["banco"] = "Santander"
        lista.append(datos)
    ventana.lista = lista
    ventana.ext = True
    if len(os.listdir('temp_uploads')) != 1: 
        ventana.frame_siguiente.sig.config(state="normal")
    ventana.guardar.config(state="normal")
    ventana.extraer.config(state="disabled")
    lbl_opcion = ventana.lbl  
    lbl_opcion.config(text="Si hay datos erróneos, puedes corregirlos manualmente.")
    cambiar_datos(ventana,marco, lista[0]["clientName"],lista[0]["numeroCuenta"], lista[0]["importe_total"],lista[0]["fecha"],0,"ini")


def next_previus(ventana, marco, command):
    if len(ventana.lista) == 0:
        return
    lista_archivos = os.listdir('temp_uploads')
    index_actual = lista_archivos.index(ventana.name)
    tam = len(lista_archivos)
    if command == "anterior":
        if index_actual == 0:
            return
        new_index = index_actual - 1
    elif command == "siguiente":
        if index_actual == tam - 1:
            return
        new_index = index_actual + 1
    else:
        return
    ventana.frame_siguiente.ant.config(state="normal")
    ventana.frame_siguiente.sig.config(state="normal")
    if new_index == 0:
        ventana.frame_siguiente.ant.config(state="disabled")
    elif new_index == tam - 1:
        ventana.frame_siguiente.sig.config(state="disabled")
    previsualizacion = ventana.previsualizacion
    previsualizacion.config(image=None)
    previsualizacion.image = None
    archivo_path = os.path.join('temp_uploads', lista_archivos[new_index])
    ventana.name = lista_archivos[new_index]
    ventana.path = archivo_path
    if ventana.formato == "img":
        archivo = Image.open(archivo_path)
        lbl_opcion = ventana.lbl
        lbl_opcion.config(text=" ")
        nuevo_ancho = int((archivo.width / archivo.height) * 553)
        archivo = archivo.resize((nuevo_ancho, 553))
        imagen_tk = ImageTk.PhotoImage(archivo)
        previsualizacion.create_image(530 / 2 - (nuevo_ancho / 2), 0, anchor=tk.NW, image=imagen_tk)
        previsualizacion.image = imagen_tk
    else:
        doc = fitz.open(archivo_path)
        for pagina in doc:
            pix = pagina.get_pixmap()
            imagen = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            proporcion = 553 / imagen.height
            ancho = imagen.width * proporcion
            imagen = imagen.resize((int(ancho), 553))
            imagen_tk = ImageTk.PhotoImage(imagen)
            previsualizacion.create_image(530 / 2 - (ancho / 2), 0, anchor=tk.NW, image=imagen_tk)
            previsualizacion.image = imagen_tk
    previsualizacion.pack()
    cambiar_datos(ventana, marco, ventana.lista[new_index]["clientName"], ventana.lista[new_index]["numeroCuenta"],
                   ventana.lista[new_index]["importe_total"], ventana.lista[new_index]["fecha"], new_index)
    ventana.cant.config(text=f"{str(new_index + 1).zfill(2)} de {str(tam).zfill(2)} ")


def reiniciarDatos(ventana,marco):
    if ventana.winfo_exists():
        lbl_opcion = ventana.lbl
        lbl_opcion.config(text="Aquí se vera el comprobante")
        ventana.lblNumero.config(text="CUIT:")
        # Descartivo botones derecha
        cargar_img = ventana.btn_img
        cargar_pdf = ventana.btn_pdf
        boton_extraer = ventana.extraer
        boton_atras = ventana.atras
        cargar_img.config(state="disabled")
        cargar_pdf.config(state="disabled")
        boton_extraer.config(state="disabled")
        boton_atras.config(state="disabled")
        # Borro imagen cargada
        previsualizacion = ventana.previsualizacion
        previsualizacion.config(image = None)
        previsualizacion.image = None
        # reinicio combo
        combo = ventana.combo
        combo.configure(state="readonly")
        combo.current(0)
        # Reinicio datos
        contenedor = marco.contenedor
        nombre = contenedor.nombre
        numero = contenedor.numero
        importe = contenedor.importe
        fecha = contenedor.fecha
        nombre.config(state="normal")
        numero.config(state="normal")
        importe.config(state="normal")
        fecha.config(state="normal")
        nombre.delete(0, tk.END)
        numero.delete(0, tk.END)
        importe.delete(0, tk.END)
        fecha.delete(0, tk.END)
        ventana.banco = ""
        nombre.config(text="")
        numero.config(text="")
        importe.config(text="")
        fecha.config(text="")    
        # oculto botones ant y sig
        ventana.frame_siguiente.pack_forget()
        ventana.frame_siguiente.sig.config(state="disabled")
        # ventana.modificar.config(state="disabled")
        ventana.guardar.pack_forget()
        ventana.guardar.config(state="disabled")
        if ventana.lista_excel:
            ventana.excel.pack(pady = 10)
            ventana.excel.config(state="normal")
        ventana.ext = False
        # Borro lista de datos
        ventana.lista = []
        delete_files()
        df = pd.read_excel("comprobantes.xlsx", engine='openpyxl')
        df = df.head(0)
        df.to_excel("comprobantes.xlsx", index=False)
        ventana.cant.config(text = "")

def atras(ventana, marco):
    siguiente = ventana.frame_siguiente
    cantidad = ventana.cant
    if ventana.ext:
        resultado = messagebox.askokcancel("Confirmar", "Si continua los datos NO guardados se perderan\n¿Estás seguro de que deseas continuar?")
        if resultado:
            reiniciarDatos(ventana, marco)
    else:
        reiniciarDatos(ventana, marco)
    marco.config(bg="#461e5a")
    siguiente.configure(bg="#461e5a")
    cantidad.config(bg="#461e5a")
        
    

def cambiar_datos(ventana, marco, name, cuenta, importe_tot, fecha_, i, loc=""):
    contenedor = marco.contenedor
    nombre = contenedor.nombre
    numero = contenedor.numero
    importe = contenedor.importe
    fecha = contenedor.fecha
    nombre.config(state="normal")
    numero.config(state="normal")
    importe.config(state="normal")
    fecha.config(state="normal")
    nombre.delete(0, tk.END)
    numero.delete(0, tk.END)
    importe.delete(0, tk.END)
    fecha.delete(0, tk.END)
    nombre.insert(0, str(name))
    numero.insert(0, str(cuenta))
    importe.insert(0, str(importe_tot))
    fecha.insert(0, str(fecha_))
    ventana.lista[i]["clientName"] = nombre.get()
    ventana.lista[i]["numeroCuenta"] = numero.get()
    ventana.lista[i]["importe_total"] = importe.get()
    ventana.lista[i]["fecha"] = fecha.get()
    ventana.lista[i]["banco"] = ventana.banco
    
def mercado_pago(lineas):
    clientName = None
    importe_total = None
    numeroCuenta = None
    fecha = None
    try:
        i = 0
        while i < len(lineas):
            if "hs" in lineas[i]:
                importe_total = str(lineas[i+1].replace("$","").replace(".","").replace(",",".") )
                break
            i += 1
    except Exception as e:
        print(e)
    try:
        for i in range(len(lineas)):
            if "CUITICUIL:" in lineas[i] or "CUIT/CUIL:" in lineas[i]:
                clientName = lineas[i - 1].strip() 
                numeroCuenta = lineas[i].split().pop().replace("-","")
                break
    except Exception as e:
        print(e)    
    try:
        for i in range(len(lineas)):
            if "transferencia" in lineas[i]:
                fecha_ = lineas[i+1].split()
                meses = {"enero": "01", "febrero": "02", "marzo": "03", "abril": "04", "mayo": "05", "junio": "06", "julio": "07", "agosto": "08", "septiembre": "09", "octubre": "10", "noviembre": "11", "diciembre": "12"} 
                mes = meses.get(str(fecha_[2]))
                fecha = f"{str(fecha_[0]).zfill(2)}/{mes}/{fecha_[3]}"
                break
    except Exception as e:
        print(e)       
    if clientName is None:
        clientName = "Nombre no encontrado"
    if numeroCuenta is None:
        numeroCuenta = "Numero de cuenta no encontrado"
    if importe_total is None:
        importe_total = "Importe no encontrado"
    if fecha is None:
        fecha = "Fecha no encontrada"
    return clientName, numeroCuenta, importe_total, fecha


def banco_macro(lineas, formato=None):
    datos = {
        "clientName": None,
        "numeroCuenta": None,
        "importe_total": None,
        "fecha": None,
    }
    datos["clientName"], datos["numeroCuenta"], datos["importe_total"], datos["fecha"] = banco_macro_c3(lineas, formato)
    if datos["fecha"] == "Fecha no encontrada":
        datos["clientName"], datos["numeroCuenta"], datos["importe_total"], datos["fecha"] = banco_macro_c2(lineas)
        if datos["importe_total"] == "Importe no encontrado":
            datos["clientName"], datos["numeroCuenta"], datos["importe_total"], datos["fecha"] = banco_macro_c1(lineas)
    return datos["clientName"], datos["numeroCuenta"], datos["importe_total"], datos["fecha"]
        
def banco_macro_c1(lineas):
    clientName = "Nombre no encontrado"
    importe_total = "Importe no encontrado"
    numeroCuenta = "Numero de cuenta no encontrado"
    fecha = "Fecha no encontrada"
    
    for i in range(len(lineas)):
        if "$" in lineas[i]:
            importe_total = lineas[i + 1].strip().rstrip("\n").replace(".", "").replace(",", ".")
            break
            
    for linea in lineas:
        if "Cuenta" in linea:
            numeroCuenta = linea.split()[-1]
            break
            
    for i in range(len(lineas)):
        if "Fecha" in lineas[i]:
            fecha = lineas[i].split()[1]
            break
            
    return clientName, numeroCuenta, importe_total, fecha

def banco_macro_c2(lineas):
    clientName = None
    importe_total = None
    numeroCuenta = None
    clientName = None
    fecha = None
    fecha = lineas[0][:10]
    if "Macro" in fecha:
        fecha = lineas[1][:10]
    for i in range(len(lineas)):
        if "Importe" in lineas[i]:
            importe_total= str(lineas[i+1]).rstrip().replace("$","").replace("S","").replace(",",".")
        if "debitar:" in lineas[i]:
            numeroCuenta = str(lineas[i+1])
    if clientName == None:
        clientName = "Nombre no encontrado"
    if numeroCuenta == None:
        numeroCuenta = "Numero de cuenta no encontrado"
    if importe_total == None:
        importe_total = "Importe no encontrado"
    if fecha == None:
        fecha = "Fecha no encontrada"
    return clientName, numeroCuenta, importe_total,fecha

def banco_macro_c3(lineas,formato):
    clientName = None
    importe_total = None
    numeroCuenta = None
    clientName = None
    fecha = None
    try:
        if formato == "img":
            for i in range(len(lineas)):
                if "Importe" in lineas[i]:
                    importe_total= str(lineas[i+1]).rstrip().replace("$","").replace("S","").replace(",","")
                if "Fecha" in lineas[i]:
                    fecha = lineas[i+1].split().pop()
                if "Ordenante" in lineas[i]:
                    clientName = lineas[i+1].rstrip()
            for i in range(len(lineas)): 
                if "CUITICUIL" in lineas[i] or "CUIT/CUIL" in lineas[i]:
                    numeroCuenta = lineas[i].split().pop()
                    break
        else:
            for i in range(len(lineas)):
                if "Importe" in lineas[i]:
                    importe_total= str(lineas[i+1]).rstrip().replace("$","").replace("S","").replace(",","")
                if "Importe" in lineas[i] and fecha == None:
                    fecha = lineas[i+1].split().pop()
                if "Ordenante" in lineas[i]:
                    clientName = lineas[i+1].rstrip()
            for i in range(len(lineas)): 
                if "CUITICUIL" in lineas[i] or "CUIT/CUIL" in lineas[i]:
                    numeroCuenta = lineas[i].split().pop()
                    break
    except Exception as e:
        print(e)
        
    if clientName == None:
        clientName = "Nombre no encontrado"
    if numeroCuenta == None:
        numeroCuenta = "Numero de cuenta no encontrado"
    if importe_total == None:
        importe_total = "Importe no encontrado"
    if fecha == None:
        fecha = "Fecha no encontrada"
    return clientName, numeroCuenta, importe_total,fecha

def banco_santander(lineas):
    clientName = None
    importe_total = None
    numeroCuenta = None 
    fecha = None   
    for i in range(len(lineas)):
        if "Importe" in lineas[i]:
            importe_total = lineas[i + 1].rstrip().replace("$", "").replace(".","").replace(",",".")
        if "Envia" in lineas[i]:
            clientName = str(lineas[i+1].replace(";","").replace(",","")).rstrip()
        if "debito" in lineas[i]:
            numeroCuenta = lineas[i+1].rstrip().split().pop().replace("-","").replace("/","")
        if "Fecha" in lineas [i]:
            fecha = lineas[i+1].rstrip()
    if clientName == None:
        clientName = "Nombre no encontrado"
    if numeroCuenta == None:
        numeroCuenta = "Numero de cuenta no encontrado"
    if importe_total == None:
        importe_total = "Importe no encontrado"
    if fecha == None:
        fecha = "Fecha no encontrada" 
    return clientName, numeroCuenta, importe_total, fecha

def banco_nacion_plus(lineas):
    clientName = None
    importe_total = None
    numeroCuenta = None 
    fecha = None   
    for i in range(len(lineas)):
        if "Monto" in lineas[i]:
            importe_total = lineas[i + 1].rstrip().replace("$", "").replace("S", "").replace(".", "").replace(",", ".")
        if "Envia" in lineas[i]:
            clientName = str(lineas[i+1].replace(";","").replace(",","")).rstrip()
        if "debito" in lineas[i]:
            numeroCuenta = lineas[i+1].rstrip().split().pop()
        if "Fecha" in lineas [i]:
            fecha = lineas[i+1].split()[0]
    if clientName == None:
        clientName = "Nombre no encontrado"
    if numeroCuenta == None:
        numeroCuenta = "Numero de cuenta no encontrado"
    if importe_total == None:
        importe_total = "Importe no encontrado"
    if fecha == None:
        fecha = "Fecha no encontrada" 
    return clientName, numeroCuenta, importe_total, fecha

def guardar_en_excel(ventana, marco):
    contenedor = marco.contenedor
    nombre = contenedor.nombre
    numero = contenedor.numero
    importe = contenedor.importe
    fecha = contenedor.fecha
    datos = {
        "clientName": nombre.get(),
        "numeroCuenta": numero.get(),
        "importe_total": importe.get(),
        "fecha": fecha.get(),
        "banco": ventana.banco
    }
    ventana.lista[ventana.indice] = datos
    ventana.lista_excel += ventana.lista
    ventana.guardar.config(state="disabled")
    ventana.ext = False
    
    messagebox.showinfo("Exito", "Datos guardados con exito")
    
def extraer_en_excel(ventana):
    df = pd.DataFrame(columns=["fecha", "clientName", "importe_total", "numeroCuenta", "banco"])
    df.to_excel("comprobantes.xlsx", index=False)
    for datos in ventana.lista_excel:
        if datos["importe_total"].replace('.', '', 1).isdigit():
            datos["importe_total"] = float(datos["importe_total"])
        df.loc[len(df)] = [datos["fecha"], datos["clientName"], datos["importe_total"], datos["numeroCuenta"], datos["banco"]]
        df.to_excel("comprobantes.xlsx", index=False)

    destino = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    if destino:
        # Copiar el archivo Excel a la nueva ubicación
        shutil.copy("comprobantes.xlsx", destino)
        ventana.excel.pack_forget()
        messagebox.showinfo("Exito", "Excel extraido con exito")
        ventana.ext = False
        ventana.lista_excel = []
    else:
        messagebox.showwarning("Cuidado", "No se ha seleccionado una ubicación de destino.")
        
def delete_files():
    temp_folder_path = 'temp_uploads'
    for filename in os.listdir(temp_folder_path):
        file_path = os.path.join(temp_folder_path, filename)
        try:
            if os.path.isfile(file_path):
                os.unlink(file_path)
        except Exception as e:
            print(f"No se pudo eliminar {file_path}: {e}")

crear_interfaz()
