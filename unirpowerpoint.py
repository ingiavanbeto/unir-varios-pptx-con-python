import os
import comtypes.client

def unir_presentaciones(carpeta, archivo_salida):
    ppt_app = comtypes.client.CreateObject("PowerPoint.Application")
    ppt_app.Visible = True

    archivos_pptx = [f for f in os.listdir(carpeta) if f.endswith(".pptx") and f != archivo_salida]

    if not archivos_pptx:
        print("No se encontraron archivos .pptx en la carpeta.")
        return

    print("Archivos encontrados:")
    for a in archivos_pptx:
        print(f"- {a}")

    archivo_base = os.path.join(carpeta, archivos_pptx[0])
    presentacion_destino = ppt_app.Presentations.Open(archivo_base, WithWindow=False)

    for archivo in archivos_pptx[1:]:
        ruta = os.path.join(carpeta, archivo)
        presentacion_temp = ppt_app.Presentations.Open(ruta, WithWindow=False)
        total_slides = presentacion_temp.Slides.Count
        slide_indices = list(range(1, total_slides + 1))
        slide_range = presentacion_temp.Slides.Range(slide_indices)
        slide_range.Copy()

        presentacion_destino.Slides.Paste(-1)
        presentacion_temp.Close()

    salida = os.path.join(carpeta, archivo_salida)
    ruta_pdf = os.path.join(carpeta, archivo_salida.replace(".pptx", ".pdf"))
    presentacion_destino.SaveAs(salida)
    presentacion_destino.SaveAs(ruta_pdf, FileFormat=32)  # 32 es el código para PDF

    presentacion_destino.Close()
    ppt_app.Quit()
    print(f"\n✅ Presentación final guardada como: {archivo_salida}")

# Ejecutar
carpeta_actual = os.getcwd()
nombre_salida = "presentacion_unida.pptx"
unir_presentaciones(carpeta_actual, nombre_salida)