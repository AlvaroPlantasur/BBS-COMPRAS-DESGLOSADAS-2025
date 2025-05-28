import os
import psycopg2
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from datetime import datetime
import sys
import copy

def main():
    # 1. Obtener credenciales y ruta del archivo
    db_name = os.environ.get('DB_NAME')
    db_user = os.environ.get('DB_USER')
    db_password = os.environ.get('DB_PASSWORD')
    db_host = os.environ.get('DB_HOST')
    db_port = os.environ.get('DB_PORT')
    file_path = os.environ.get('EXCEL_FILE_PATH')

    db_params = {
        'dbname': db_name,
        'user': db_user,
        'password': db_password,
        'host': db_host,
        'port': db_port
    }

    # 2. Definir fechas
    fecha_inicio_str = '2025-01-01'
    fecha_fin = datetime.now().date()
    fecha_fin_str = fecha_fin.strftime('%Y-%m-%d')

    # 3. Consulta SQL
    query = f""" select 
        ail.invoice_id as "ID FACTURA",
        ai.name as "REFERENCIA ALBARÁN",
        ai.internal_number as "CÓDIGO FACTURA",
        to_char(ai.date_invoice, 'DD/MM/YYYY') as "FECHA FACTURA",
        pp.default_code as "REFERENCIA PRODUCTO", 
        pp.name as "NOMBRE", 
        s.name AS "SECCION", 
        f.name as "FAMILIA", 
        sf.name as "SUBFAMILIA",
        rc.name as "COMPAÑÍA",
        ssp.name as "SEDE",
        (CASE WHEN stp.directo_cliente = true THEN 'Sí' ELSE 'No' END) AS "CAMIÓN DIRECTO",
        SUM(stp.cargos_extra_prorrateo) AS "CARGOS EXTRA",
        ai.portes as "PORTES",
        rp.nombre_comercial as "CLIENTE",
        rp.vat as "CIF CLIENTE",
        c.name as "PAÍS",
        extract(MONTH FROM ai.date_invoice) as "MES",
        extract(DAY FROM ai.date_invoice) as "DÍA",
        (
            case when ai.type = 'in_invoice' then ail.cantidad_pedida
            when ai.type = 'in_refund' then -ail.cantidad_pedida
            end
        ) as "UNIDADES COMPRA",
        (
            case when ai.type = 'in_invoice' then ail.price_subtotal
            when ai.type = 'in_refund' then -ail.price_subtotal
            end
        ) as "BASE COMPRA TOTAL",
        (
            case when ai.type = 'in_invoice' then sum(case when coalesce(at.amount,0) = 1 then 0.0 else (coalesce(at.amount,0)*ail.price_subtotal) end)
            when ai.type = 'in_refund' then -sum(case when coalesce(at.amount,0) = 1 then 0.0 else (coalesce(at.amount,0)*ail.price_subtotal) end)
            end
        ) as "IMPUESTOS",
        (
            case when ai.type = 'in_invoice' then ail.price_subtotal + sum(case when coalesce(at.amount,0) = 1 then 0.0 else (coalesce(at.amount,0)*ail.price_subtotal) end)
            when ai.type = 'in_refund' then -(ail.price_subtotal + sum(case when coalesce(at.amount,0) = 1 then 0.0 else (coalesce(at.amount,0)*ail.price_subtotal) end))
            end
        ) as "IMPORTE COMPRA TOTAL"
    
    from account_invoice_line ail
    inner join account_invoice ai ON ai.id = ail.invoice_id
    inner join product_product pp ON ail.product_id = pp.id
    inner join res_partner rp on rp.id = ai.partner_id
    inner join res_partner_address rpa ON rpa.id = ai.address_invoice_id
    inner join res_country c on c.id = rpa.pais_id
    left outer join stock_picking stp ON stp.name = split_part(ai.origin,':', 1)
    left outer join res_company rc on rc.id = ai.company_id
    left outer join stock_sede_ps ssp on ssp.id = ai.sede_id
    left outer join product_category s ON (s.id = pp.seccion)
    left outer join product_category f ON (f.id = pp.familia)
    left outer join product_category sf ON (sf.id = pp.subfamilia)
    left outer join account_invoice_line_tax ailt on ail.id = ailt.invoice_line_id
    left outer join account_tax at on ailt.tax_id = at.id
    where ai.state in ('open','paid') and ai.type in ('in_invoice','in_refund') and ai.date_invoice >= '{fecha_inicio_str}' and ai.date_invoice <= '{fecha_fin_str}' and ai.obsolescencia = false
    group by 
        ail.id,
        rp.id,
        ail.company_id,
        ai.sede_id,
        ai.date_invoice,
        to_char(ai.date_invoice, 'YYYY'),
        to_char(ai.date_invoice, 'MM'),
        to_char(ai.date_invoice, 'YYYY-MM-DD'),
        pp.seccion,
        pp.familia,
        pp.subfamilia,
        pp.default_code,
        pp.id,
        ai.partner_id,
        ai.anticipo,
        c.name,
        rpa.prov_id,
        rpa.state_id_2,
        ai.name,
        ai.internal_number,
        ai.origin,
        ail.cantidad_pedida,
        ail.price_subtotal,
        s.name,
        f.name,
        sf.name,
        rc.name,
        ssp.name,
        ai.directo_cliente,
        ai.portes,
        rp.nombre_comercial,
        ai.type,
        stp.directo_cliente,
        rp.vat; """  

    # 4. Ejecutar la consulta
    try:
        with psycopg2.connect(**db_params) as conn:
            with conn.cursor() as cur:
                cur.execute(query)
                resultados = cur.fetchall()
                headers = [desc[0] for desc in cur.description]
    except Exception as e:
        print(f"Error al conectar o ejecutar la consulta: {e}")
        sys.exit(1)

    if not resultados:
        print("No se obtuvieron resultados de la consulta.")
        return
    else:
        print(f"Se obtuvieron {len(resultados)} filas de la consulta.")

    # 5. Cargar el archivo Excel
    try:
        book = load_workbook(file_path)
        sheet = book.active
    except FileNotFoundError:
        print(f"No se encontró el archivo '{file_path}'.")
        return

    # 6. Detectar la fila de cabecera (la primera fila no vacía con más de 1 celda con contenido)
    header_row = 1
    for i, row in enumerate(sheet.iter_rows(min_row=1, max_row=10), start=1):
        values = [cell.value for cell in row if cell.value is not None]
        if len(values) > 1:
            header_row = i
            break
    print(f"Fila de cabecera detectada: {header_row}")

    # 7. Borrar todas las filas de datos dejando la cabecera
    if sheet.max_row > header_row:
        sheet.delete_rows(header_row + 1, sheet.max_row - header_row)

    # 8. Insertar nuevos datos debajo de la cabecera
    for row in resultados:
        sheet.append(row)

    # 9. Copiar estilo de la fila de cabecera (fila siguiente si existe)
    if sheet.max_row > header_row + 1:
        for col in range(1, sheet.max_column + 1):
            source_cell = sheet.cell(row=header_row + 1, column=col)
            for row in range(header_row + 2, sheet.max_row + 1):
                target_cell = sheet.cell(row=row, column=col)
                target_cell.font = copy.copy(source_cell.font)
                target_cell.fill = copy.copy(source_cell.fill)
                target_cell.border = copy.copy(source_cell.border)
                target_cell.alignment = copy.copy(source_cell.alignment)

    # 10. Actualizar tabla
    if "CompDesglosadas2025" in sheet.tables:
        tabla = sheet.tables["CompDesglosadas2025"]
        max_row = sheet.max_row
        max_col = sheet.max_column
        last_col_letter = get_column_letter(max_col)
        new_ref = f"A{header_row}:{last_col_letter}{max_row}"
        tabla.ref = new_ref
        print(f"Tabla 'CompDesglosadas2025' actualizada a rango: {new_ref}")
    else:
        print("No se encontró la tabla 'CompDesglosadas2025'. No se actualizará la referencia.")

    # 11. Guardar archivo
    book.save(file_path)
    print(f"Archivo actualizado correctamente: '{file_path}'.")

if __name__ == '__main__':
    main()
