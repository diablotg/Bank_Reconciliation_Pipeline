import pandas as pd
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side
import os
import sys
import json
import re

def load_config():
    if not os.path.exists('config.json'):
        print("ERROR: No se encuentra config.json")
        sys.exit(1)
    with open('config.json', 'r') as f:
        return json.load(f)

def clean_data(df):
    """Limpia espacios y maneja valores nulos."""
    return df.apply(lambda x: x.str.strip() if x.dtype == "object" else x)

def get_folio_number(val):
    """Extrae el número de folio de un string de referencia."""
    try:
        # Busca el primer número en la cadena
        match = re.search(r'\d+', str(val))
        return int(match.group()) if match else None
    except:
        return None

def process_conciliacion():
    config = load_config()
    input_banco = 'input/banco.xlsx'
    input_contabilidad = 'input/contabilidad.xlsx'
    output_file = 'output/conciliacion.xlsx'

    # Validaciones iniciales
    if not os.path.exists(input_banco) or not os.path.exists(input_contabilidad):
        print(f"ERROR: Asegúrese de que existan {input_banco} y {input_contabilidad}")
        input("Presione Enter para salir...")
        sys.exit(1)

    if not os.path.exists('output'):
        os.makedirs('output')

    print("Cargando datos...")
    try:
        # Banco: A-G (0-6)
        df_banco = pd.read_excel(input_banco, header=None)
        # Contabilidad: C=2, H=7 (0-indexed)
        df_conta = pd.read_excel(input_contabilidad, header=None)
    except Exception as e:
        print(f"ERROR al leer archivos: {e}")
        input("Presione Enter para salir...")
        sys.exit(1)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Conciliación"

    current_row = 1

    # Definir estilos
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = openpyxl.styles.PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
    border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    # Procesar Almacenes 1-7 + Almacén 8
    all_blocks = config['almacenes'] + [config['almacen_8_ingresos']]

    for block in all_blocks:
        is_alm_8 = 'rango_folios' in block
        nombre = block['nombre']
        print(f"Procesando {nombre}...")

        # --- FILTRADO BANCO ---
        # Filtro: Concepto (Col B=1) contiene ref_banco
        mask_banco = df_banco[1].astype(str).str.contains(block['ref_banco'], case=False, na=False)
        data_banco = df_banco[mask_banco].copy()
        
        # Ordenar por fecha descendente (Col A=0)
        try:
            data_banco[0] = pd.to_datetime(data_banco[0])
            data_banco = data_banco.sort_values(by=0, ascending=False)
        except:
            pass # Si la fecha no es válida, dejar como está

        total_banco = data_banco[5].sum() # Col F=5 (Abono)

        # --- FILTRADO CONTABILIDAD ---
        if is_alm_8:
            # Lógica Folios 1-1100
            def in_range(val):
                f = get_folio_number(val)
                return f is not None and block['rango_folios'][0] <= f <= block['rango_folios'][1]
            mask_conta = df_conta[2].apply(in_range)
        else:
            # Lógica Referencias 1-7
            mask_conta = df_conta[2].astype(str).isin(block['ref_contabilidad'])
        
        data_conta = df_conta[mask_conta].copy()
        total_conta = data_conta[7].sum() # Col H=7 (Importe)

        # --- ESCRITURA EN EXCEL ---
        # Título del bloque
        ws.cell(row=current_row, column=1, value=nombre).font = Font(bold=True, size=14)
        current_row += 1

        # Headers Banco (A-G)
        headers_banco = ["Fecha", "Concepto", "Ref", "Ref Amp", "Cargo", "Abono", "Saldo"]
        for i, h in enumerate(headers_banco, 1):
            cell = ws.cell(row=current_row, column=i, value=h)
            cell.font = header_font
            cell.fill = header_fill

        # Headers Contabilidad (I-...)
        ws.cell(row=current_row, column=9, value="Referencia Conta").font = header_font
        ws.cell(row=current_row, column=9).fill = header_fill
        ws.cell(row=current_row, column=10, value="Importe Conta").font = header_font
        ws.cell(row=current_row, column=10).fill = header_fill

        current_row += 1
        start_data_row = current_row

        # Escribir datos Banco
        for r_idx, row_data in enumerate(data_banco.values):
            for c_idx, value in enumerate(row_data[:7]): # Columnas A-G
                ws.cell(row=start_data_row + r_idx, column=c_idx + 1, value=value)

        # Escribir datos Contabilidad
        for r_idx, row_data in enumerate(data_conta.values):
            # Col C (2) y H (7)
            ws.cell(row=start_data_row + r_idx, column=9, value=row_data[2])
            ws.cell(row=start_data_row + r_idx, column=10, value=row_data[7])

        # Actualizar current_row al final del bloque más largo
        max_rows = max(len(data_banco), len(data_conta))
        current_row = start_data_row + max_rows + 1

        # Resumen del bloque
        ws.cell(row=current_row, column=5, value="TOTAL BANCO:").font = Font(bold=True)
        ws.cell(row=current_row, column=6, value=total_banco).font = Font(bold=True)
        
        ws.cell(row=current_row + 1, column=9, value="TOTAL CONTA:").font = Font(bold=True)
        ws.cell(row=current_row + 1, column=10, value=total_conta).font = Font(bold=True)

        diff = total_banco - total_conta
        current_row += 3
        ws.cell(row=current_row, column=1, value="DIFERENCIA:").font = Font(bold=True)
        ws.cell(row=current_row, column=2, value=diff).font = Font(bold=True)
        
        # Regla de color para diferencia
        if abs(diff) > 1:
            ws.cell(row=current_row, column=2).font = Font(bold=True, color="FF0000")
        else:
            ws.cell(row=current_row, column=2).font = Font(bold=True, color="00B050")

        # SEPARACIÓN VISUAL: 20 filas vacías
        current_row += 20

    # Ajustar ancho de columnas básico
    for col in ws.columns:
        max_length = 0
        column = col[0].column_letter
        for cell in col:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        ws.column_dimensions[column].width = max_length + 2

    print(f"Guardando reporte en {output_file}...")
    wb.save(output_file)
    print("¡Proceso completado exitosamente!")

if __name__ == "__main__":
    try:
        process_conciliacion()
    except Exception as e:
        print(f"\nFATAL ERROR: {e}")
    input("\nPresione Enter para cerrar...")
