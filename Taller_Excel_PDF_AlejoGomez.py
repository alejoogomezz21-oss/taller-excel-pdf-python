
#Importamos lo mismo que tiene la guia
import openpyxl
from tabulate import tabulate
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

# Encabezados y alineación de la tabla
encabezados = ["ID", "Producto", "Categoría", "Precio ($)", "Ventas (Unid)"]
alineacion = ("center", "left", "center", "right", "right")

# Precio mínimo para el filtro del Reto 1 (productos con precio > 100)
PRECIO_FILTRO = 100


# la funcion 1: cargar los datos desde Excel
def cargar_datos():
    wb = openpyxl.load_workbook("ventas_tecnologia_v2.xlsx", data_only=True)
    hoja = wb.active

    datos = []

   
    for fila in hoja.iter_rows(min_row=2, values_only=True):
        datos.append(list(fila))

    return datos


# la funcion 2: filtrar por precio (Reto 1)
def filtrar_por_precio(datos):
    filtrados = []
    for fila in datos:
        if fila[3] > PRECIO_FILTRO:
            filtrados.append(fila)
    return filtrados


# la funcion 3: calcular métricas (Reto 2)
# Aqui lo que se devuelve es el total de unidades vendidas y el precio promedio
def calcular_metricas(datos):
    total_unidades = 0
    suma_precios = 0

    for fila in datos:
        suma_precios = suma_precios + fila[3]
        total_unidades = total_unidades + fila[4]

    if len(datos) > 0:
        precio_promedio = suma_precios / len(datos)
    else:
        precio_promedio = 0

    return total_unidades, precio_promedio


# la funcion 4: mostrar la tabla en consola con tabulate
def mostrar_tabla(datos, titulo):
    print(f"\n{titulo} ({len(datos)} PRODUCTOS)")
    print(tabulate(datos, headers=encabezados, tablefmt="fancy_grid", colalign=alineacion))


# la funcion 5: mostrar las métricas al final de la consola (Reto 2)
def mostrar_metricas(datos):
    total_unidades, precio_promedio = calcular_metricas(datos)
    print("\n MÉTRICAS DE RESUMEN ")
    print(f"Total de unidades vendidas: {total_unidades:,}")
    print(f"Precio promedio de los productos: ${precio_promedio:,.2f}")


# la funcion 6: buscar un producto por ID (Reto 3)
def buscar_por_id(datos):
    texto_id = input("\nIngrese el ID del producto a buscar: ").strip()

    if not texto_id.isdigit():
        print("[ADVERTENCIA] Debe ingresar un número de ID válido.")
        return

    id_buscado = int(texto_id)

    resultado = []
    for fila in datos:
        if fila[0] == id_buscado:
            resultado.append(fila)

    if len(resultado) > 0:
        mostrar_tabla(resultado, f"RESULTADO DE BÚSQUEDA ID {id_buscado}")
    else:
        print(f"[ADVERTENCIA] El ID {id_buscado} no existe en el catálogo.")


# la funcion 7: generar el PDF con ReportLab
def generar_pdf(datos, titulo_informe):
    # 1. Documento base (tamaño carta)
    doc = SimpleDocTemplate("reporte_ejecutivo_ventas.pdf", pagesize=letter)
    estilos = getSampleStyleSheet()

    estilo_titulo = ParagraphStyle(
        "EstiloTitulo", parent=estilos["Title"],
        fontName="Helvetica-Bold", fontSize=18,
        textColor=colors.HexColor("#1A365D"), spaceAfter=10
    )
    estilo_cuerpo = ParagraphStyle("EstiloCuerpo", parent=estilos["Normal"], fontSize=9)
    estilo_encabezado = ParagraphStyle(
        "EstiloEncabezado", parent=estilos["Normal"],
        fontName="Helvetica-Bold", fontSize=10,
        textColor=colors.whitesmoke, alignment=1
    )

    story = [
        Paragraph(titulo_informe, estilo_titulo),
        Paragraph(f"<b>Total de Productos Procesados:</b> {len(datos)}", estilo_cuerpo),
        Spacer(1, 12)
    ]

    encabezados_pdf = [Paragraph(f"<b>{h}</b>", estilo_encabezado) for h in encabezados]

    filas_pdf = [
        [
            Paragraph(str(r[0]), ParagraphStyle("C", parent=estilo_cuerpo, alignment=1)),
            Paragraph(str(r[1]), ParagraphStyle("L", parent=estilo_cuerpo, alignment=0)),
            Paragraph(str(r[2]), ParagraphStyle("C", parent=estilo_cuerpo, alignment=1)),
            Paragraph(f"${r[3]:,.2f}", ParagraphStyle("R", parent=estilo_cuerpo, alignment=2)),
            Paragraph(str(r[4]), ParagraphStyle("R", parent=estilo_cuerpo, alignment=2))
        ]
        for r in datos
    ]

    total_unidades, precio_promedio = calcular_metricas(datos)
    fila_resumen = [
        Paragraph("", estilo_cuerpo),
        Paragraph("<b>RESUMEN (precio promedio / total unidades)</b>",
                  ParagraphStyle("L", parent=estilo_cuerpo, alignment=0)),
        Paragraph("", estilo_cuerpo),
        Paragraph(f"<b>${precio_promedio:,.2f}</b>",
                  ParagraphStyle("R", parent=estilo_cuerpo, alignment=2)),
        Paragraph(f"<b>{total_unidades:,}</b>",
                  ParagraphStyle("R", parent=estilo_cuerpo, alignment=2))
    ]

    anchos_cols = [40, 160, 100, 80, 80]
    tabla_pdf = Table([encabezados_pdf] + filas_pdf + [fila_resumen],
                      colWidths=anchos_cols, repeatRows=1)

    tabla_pdf.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1A365D")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#CBD5E0")),
        ("ROWBACKGROUNDS", (0, 1), (-1, -2), [colors.white, colors.HexColor("#F7FAFC")]),
        ("BACKGROUND", (0, -1), (-1, -1), colors.HexColor("#E2E8F0")),
    ]))

    story.append(tabla_pdf)

    doc.build(story)
    print("\n[ÉXITO] Archivo 'reporte_ejecutivo_ventas.pdf' generado correctamente.")


# la funcion 8: menú interactivo
def mostrar_menu():
    print("\n" + "=" * 40)
    print("        MENÚ - REPORTE DE VENTAS")
    print("=" * 40)
    print("1) Ver catálogo completo")
    print("2) Filtrar por precio (> $100)")
    print("3) Buscar por ID")
    print("4) Exportar informe PDF")
    print("5) Salir del programa")


def exportar_pdf(datos):
    print("\n¿Qué desea exportar?")
    print("  1) Catálogo completo")
    print("  2) Solo productos con precio > $100")
    opcion = input("Elija 1 o 2: ").strip()

    if opcion == "1":
        generar_pdf(datos, "Informe Ejecutivo de Ventas Tecno")
    elif opcion == "2":
        filtrados = filtrar_por_precio(datos)
        generar_pdf(filtrados, "Informe Ejecutivo de Ventas Tecno (Precio > $100)")
    else:
        print("[ADVERTENCIA] Opción no válida. No se exportó ningún PDF.")


def main():
    datos = cargar_datos()
    print(f"Total de registros leídos en memoria: {len(datos)}")

    while True:
        mostrar_menu()
        opcion = input("Seleccione una opción (1-5): ").strip()

        if opcion == "1":
            mostrar_tabla(datos, "REPORTE DE VENTAS EN CONSOLA")
            mostrar_metricas(datos)
        elif opcion == "2":
            filtrados = filtrar_por_precio(datos)
            mostrar_tabla(filtrados, "PRODUCTOS CON PRECIO MAYOR A $100")
            mostrar_metricas(filtrados)
        elif opcion == "3":
            buscar_por_id(datos)
        elif opcion == "4":
            exportar_pdf(datos)
        elif opcion == "5":
            print("\nSaliendo del programa. ¡Hasta pronto!")
            break
        else:
            print("[ADVERTENCIA] Opción no válida. Intente de nuevo.")


if __name__ == "__main__":
    main()
