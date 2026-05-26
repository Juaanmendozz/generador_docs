# 📄 Generador de Documentos Word

Aplicación de escritorio para generar documentos Word (.docx) profesionales en segundos.

## Documentos que puedes generar
- ✅ Certificado laboral
- ✅ Factura de venta
- ✅ Carta / Oficio
- ✅ Contrato (prestación de servicios u otro tipo)

---

## ▶️ Cómo ejecutar

### En Windows
1. Asegúrate de tener **Python 3.8+** instalado → https://python.org
   - Durante la instalación marca la opción **"Add Python to PATH"**
2. Haz doble clic en `INSTALAR_Y_EJECUTAR.bat`
3. La primera vez instalará las dependencias automáticamente y abrirá la app.

### En cualquier sistema (manual)
```bash
pip install python-docx pillow
python app.py
```

---

## 📋 Cómo usar la app
1. Selecciona el **tipo de documento** en el menú desplegable.
2. Llena los campos del formulario.
3. Haz clic en **"Generar .docx"**.
4. Elige dónde guardar el archivo.
5. ¡El documento se abre automáticamente en Word!

---

## 💡 Formato de ítems en Factura
Cada línea debe seguir este formato:
```
descripción|cantidad|valor_unitario
```
Ejemplo:
```
Diseño de logo|1|500000
Horas de soporte|3|80000
```

---

## 🛠️ Requisitos
- Python 3.8 o superior
- Librerías: `python-docx`, `pillow` (se instalan automáticamente con el .bat)

---

Desarrollado con 💙 usando Python + tkinter + python-docx
