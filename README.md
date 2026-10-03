# Sistema de Procesamiento de Facturas y Listas de Empaque

Este sistema procesa documentos logísticos (facturas y packing lists en formato PDF o Excel) aplicando limpieza de descripciones, agrupación por bulto/marca/detalle y sumatorias con cuadre matemático exacto.

## 🚀 Opciones de Uso

### Opción 1: Interfaz Gráfica Web (Recomendada)
Haz doble clic en:
👉 **`INICIAR_INTERFAZ_GRAFICA.bat`**
Se abrirá automáticamente una ventana en tu navegador web donde podrás:
1. Arrastrar y soltar uno o varios PDFs/Excels.
2. Hacer clic en **"Procesar Facturas y Consolidar"**.
3. Ver las métricas en tiempo real (Total Bultos, Piezas, Peso kg, CBM m³).
4. Ver la tabla completa en pantalla.
5. Hacer clic en el botón verde **"Descargar Excel Consolidado (.xlsx)"**.

---

### Opción 2: Proceso en Lote por Carpeta
1. Coloca tus archivos en la carpeta: **`Entrada/`**
2. Haz doble clic en: 👉 **`PROCESAR_CARPETA_ENTRADA.bat`**
3. El resultado se guardará en la carpeta: **`Salida/CONSOLIDADO_FACTURAS.xlsx`**

---

### Opción 3: Arrastrar y Soltar un Archivo Individual
Arrastra cualquier archivo PDF o Excel directamente sobre:
👉 **`PROCESAR_UN_ARCHIVO.bat`**

---

### Opción 4: Desde Antigravity
En cualquier chat de Antigravity puedes escribir:
- *"Procesa la carpeta de facturas y genera el Excel"*
- *"Procesa el archivo `nombre_archivo.pdf`"*
