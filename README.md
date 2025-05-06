# 🧾 Lector de Comprobantes de Transferencia (APCODA)

Aplicación de escritorio desarrollada en **Python** que permite leer comprobantes de transferencias bancarias y billeteras virtuales (en formato imagen o PDF), extraer información relevante y guardarla automáticamente en un archivo de texto.

## 📌 Descripción

Esta herramienta fue creada para facilitar la gestión de comprobantes de transferencias. Automatiza el proceso de lectura, extracción de datos clave (monto, fecha, entidad, CBU/alias, etc.) y almacenamiento en un archivo `.txt`.

Es útil para pequeñas empresas, profesionales independientes o cualquier persona que necesite registrar comprobantes de forma rápida y ordenada.

## ⚙️ Características

- 📄 Soporte para comprobantes en formato **imagen (JPG, PNG)** y **PDF**
- 🧠 Lectura inteligente mediante **OCR (Reconocimiento Óptico de Caracteres)**
- 🏦 Compatible con múltiples bancos y billeteras virtuales (Mercado Pago, Ualá, etc.)
- 💾 Almacenamiento automático de la información en un archivo `.txt`
- 💻 Interfaz sencilla y funcional
- 📘 Manual de uso incluido en el repositorio

## 🛠️ Tecnologías utilizadas

- Python 3
- [Tesseract OCR](https://github.com/tesseract-ocr/tesseract)
- PyPDF2 o pdfplumber (para lectura de PDFs)
- Pillow (para procesamiento de imágenes)
- Tkinter (para interfaz gráfica)
- OS / pathlib / datetime (manejo de archivos y rutas)

📄 Manual de uso
El manual completo con capturas de pantalla y explicación paso a paso se encuentra en el archivo manual_de_uso.pdf, dentro de este mismo repositorio.

👨‍💻 Autor
Desarrollado por [Lucas Vargas] – Año: 2025
Proyecto independiente con fines prácticos y educativos.
