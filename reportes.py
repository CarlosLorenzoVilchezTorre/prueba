import csv
import xlsxwriter
import sys
import re
import os
import glob

csv.field_size_limit(sys.maxsize)

def extraer_valor(texto):
    patron = r"======Current Value\(s\).*?======\s*([\s\S]+?)(?=\n======|\Z)"
    match = re.search(patron, texto, re.DOTALL)
    if match:
        return match.group(1).strip()
    else:
        return "No se encontró un valor debajo de 'Current Value(s)'"

def extraer_valor_requerido(texto):
    patron = r"======Expected Value\(s\).*?======\s*([\s\S]+?)(?=\n======|\Z)"
    match = re.search(patron, texto, re.DOTALL)
    if match:
        return match.group(1).strip()
    else:
        return "No se encontró un valor debajo de 'Expected Value(s)'"

def procesar_archivo_csv(ruta_archivo):
    controles = []
    equipos = []

    with open(ruta_archivo, mode='r', encoding='utf-8') as archivo:
        lector = list(csv.reader(archivo))

        empezar_a_leer = False
        Nombre_de_la_lbs = lector[5][0]
        for fila in lector:
            if not empezar_a_leer:
                if len(fila) > 0 and fila[0].strip() == "Host IP":
                    empezar_a_leer = True
                continue

            valor_actual = extraer_valor(fila[17])
            valor_requerido = extraer_valor_requerido(fila[17])

            try:
                controles.append([
                    fila[7],   # Control ID
                    fila[0],   # Equipos
                    fila[9],   # Statement
                    fila[10],  # Criticality Label
                    fila[16],  # Valor Requerido original
                    valor_actual,  # Valor Actual
                    valor_requerido, # Valor Requerido extraído
                    fila[14],  # Cumplimiento
                    fila[11],  # Criticality Value
                    fila[4],   # SO
                    fila[17],  # Detalle
                    Nombre_de_la_lbs  # Tipo
                ])
            except IndexError:
                continue

        inicio_equipos = None
        fin_equipos = None
        for i, fila in enumerate(lector):
            if len(fila) > 0:
                if "IP Address" in fila[0]:
                    inicio_equipos = i + 1
                elif "ASSET TAGS" in fila[0]:
                    fin_equipos = i
                    break

        if inicio_equipos is not None and fin_equipos is not None:
            subtabla = lector[inicio_equipos:fin_equipos]
            for fila in subtabla:
                if any(campo.strip() for campo in fila):
                    fila_con_tipo = fila + [Nombre_de_la_lbs]
                    equipos.append(fila_con_tipo)

    return controles, equipos

def consolidar_csv_en_excel(carpeta_csvs, archivo_salida):
    controles_total = []
    equipos_total = []

    archivos_csv = glob.glob(os.path.join(carpeta_csvs, "*.csv"))
    for archivo in archivos_csv:
        print(f"Procesando: {archivo}")
        controles, equipos = procesar_archivo_csv(archivo)
        controles_total.extend(controles)
        equipos_total.extend(equipos)

    workbook = xlsxwriter.Workbook(archivo_salida)
    hoja_equipos = workbook.add_worksheet("Equipos")
    hoja_resumen = workbook.add_worksheet("Resumen x LBS")

    # ---------------- HOJAS POR TIPO (LBS) ----------------
    controles_por_tipo = {}
    for fila in controles_total:
        tipo = fila[11]
        if tipo not in controles_por_tipo:
            controles_por_tipo[tipo] = []
        controles_por_tipo[tipo].append(fila)

    encabezados_controles = [
        "Control ID", "Equipos", "Statement", "Criticality Label",
        "Valor Requerido", "Valor Actual", "Valor Requerido Extraído", "Cumplimiento",
        "Criticality Value", "SO", "Detalle", "Tipo"
    ]

    for tipo, filas in controles_por_tipo.items():
        nombre_hoja = tipo[:31]
        hoja = workbook.add_worksheet(nombre_hoja)

        for col, encabezado in enumerate(encabezados_controles):
            hoja.write(0, col, encabezado)

        for row, fila in enumerate(filas, start=1):
            for col, valor in enumerate(fila):
                hoja.write(row, col, valor)

    # ---------------- HOJA EQUIPOS ----------------
    if equipos_total:
        num_columnas = len(equipos_total[0]) - 1
        encabezado_generico = [f"Columna {i+1}" for i in range(num_columnas)]
        encabezado_generico.append("Tipo")

        for col, valor in enumerate(encabezado_generico):
            hoja_equipos.write(0, col, valor)

        for row, fila in enumerate(equipos_total, start=1):
            for col, valor in enumerate(fila):
                hoja_equipos.write(row, col, valor)

    # ---------------- HOJA RESUMEN x LBS ----------------
    formato_titulo = workbook.add_format({'bold': True, 'bg_color': '#305496', 'font_color': 'white', 'align': 'center'})
    formato_encabezado = workbook.add_format({'bold': True, 'bg_color': '#D9E1F2', 'align': 'center'})
    formato_borde = workbook.add_format({'border': 1, 'align': 'center'})
    formato_gris = workbook.add_format({'bg_color': '#F2F2F2', 'align': 'center', 'border': 1})

    resumen = {}
    for fila in controles_total:
        tipo = fila[11]
        so = fila[9]
        cumplimiento = fila[7].strip().lower()

        if tipo not in resumen:
            resumen[tipo] = {}
        if so not in resumen[tipo]:
            resumen[tipo][so] = {"Cumple": 0, "No Cumple": 0}

        if cumplimiento == "cumple":
            resumen[tipo][so]["Cumple"] += 1
        else:
            resumen[tipo][so]["No Cumple"] += 1

    fila_excel = 0
    for tipo, sistemas in resumen.items():
        hoja_resumen.merge_range(fila_excel, 0, fila_excel, 4, f"Resumen de Cumplimiento LBS {tipo}", formato_titulo)
        fila_excel += 1

        hoja_resumen.write(fila_excel, 0, "Sistema Operativo", formato_encabezado)
        hoja_resumen.write(fila_excel, 1, "Cumple LBS", formato_encabezado)
        hoja_resumen.write(fila_excel, 2, "No Cumple LBS", formato_encabezado)
        hoja_resumen.write(fila_excel, 3, "% Cumplimiento", formato_encabezado)
        hoja_resumen.write(fila_excel, 4, "Total", formato_encabezado)
        fila_excel += 1

        total_cumple = 0
        total_no_cumple = 0

        for so, datos in sistemas.items():
            cumple = datos["Cumple"]
            no_cumple = datos["No Cumple"]
            total = cumple + no_cumple
            porcentaje = (cumple / total * 100) if total > 0 else 0

            hoja_resumen.write(fila_excel, 0, so, formato_borde)
            hoja_resumen.write(fila_excel, 1, cumple, formato_borde)
            hoja_resumen.write(fila_excel, 2, no_cumple, formato_borde)
            hoja_resumen.write(fila_excel, 3, f"{porcentaje:.0f}%", formato_borde)
            hoja_resumen.write(fila_excel, 4, total, formato_borde)
            fila_excel += 1

            total_cumple += cumple
            total_no_cumple += no_cumple

        hoja_resumen.write(fila_excel, 0, "Total de Servidores Cumplen LBS", formato_gris)
        hoja_resumen.write(fila_excel, 1, total_cumple, formato_gris)
        hoja_resumen.write(fila_excel, 2, total_no_cumple, formato_gris)
        hoja_resumen.write(fila_excel, 3, "", formato_gris)
        hoja_resumen.write(fila_excel, 4, total_cumple + total_no_cumple, formato_gris)
        fila_excel += 2

    workbook.close()
    print(f"✅ Consolidado generado correctamente: {archivo_salida}")

# =================== EJECUCIÓN ===================
carpeta_entrada = "./reportes"  # Cambia esto a tu ruta
archivo_salida = "reporte general.xlsx"
consolidar_csv_en_excel(carpeta_entrada, archivo_salida)
