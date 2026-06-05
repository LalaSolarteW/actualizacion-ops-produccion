from matplotlib.pylab import rint
import requests
from requests.auth import HTTPBasicAuth
import pandas as pd
import os


USER = os.environ.get("IMPEL_USER")
PASS = os.environ.get("IMPEL_PASS")
URL = "https://www.impeltechnology.com/rest/api/selectQuery"

# Cargue de archivos
archivo_final = "op_det_completo.xlsx"
archivo_final_os = "op_det_os.xlsx"
df_estados = pd.read_csv('estados_impel.csv', encoding='latin-1') # Estados de OP (anulado, terminado, reproceso...)
df_comerciales = pd.read_csv('comerciales.csv', encoding='utf-8-sig') # Comerciales con su id_comercial

# ----------------------------------------- FUNCIONES -----------------------------------------

def obtener_ultima_fecha(ruta_archivo, columna_fecha, fecha_default="2024-01-01"):
    """Obtener última fecha guardada en el archivo Excel."""
    if os.path.exists(ruta_archivo):
        df = pd.read_excel(ruta_archivo)
        if not df.empty and columna_fecha in df.columns:
            df[columna_fecha] = pd.to_datetime(df[columna_fecha], errors='coerce')
            ultima = df[columna_fecha].max()
            if pd.notna(ultima):
                # Restamos 1 día para asegurar overlap y no perder registros
                desde = ultima - pd.Timedelta(days=1)
                return desde.strftime("%Y-%m-%d")
    return fecha_default


def descargar_datos(query_base, campo_fecha, fecha_inicio=None):
    """
    Descargar datos de la API en bloques de 10k filas.
    query_base debe terminar en WHERE 1=1 (sin ORDER BY).
    """
    all_data = []
    start = 0
    limit = 10000

    while True:
        if fecha_inicio:
            query = f"""
            {query_base}
            AND {campo_fecha} >= '{fecha_inicio}'
            ORDER BY {campo_fecha} ASC
            """
        else:
            query = f"""
            {query_base}
            ORDER BY {campo_fecha} ASC
            """

        params = {
            "output": "json",
            "query": query,
            "startRow": start,
            "maxRows": limit
        }

        response = requests.get(
            URL,
            params=params,
            auth=HTTPBasicAuth(USER, PASS),
            timeout=60
        )
        response.raise_for_status()

        data = response.json()

        if not data:
            break

        all_data.extend(data)
        print(f"  Descargadas: {len(all_data)} filas")

        if len(data) < limit:
            break

        start += limit

    return pd.DataFrame(all_data)


def guardar_datos(df_nuevo, ruta_archivo, clave_dedup=None):
    """Unir datos nuevos con existentes, eliminar duplicados y guardar."""
    if df_nuevo.empty:
        print("  ⚠️ No hay datos nuevos")
        return

    if os.path.exists(ruta_archivo):
        df_existente = pd.read_excel(ruta_archivo)
        df_final = pd.concat([df_existente, df_nuevo], ignore_index=True)
        df_final.drop_duplicates(subset=clave_dedup, inplace=True)
    else:
        os.makedirs(os.path.dirname(ruta_archivo), exist_ok=True) if os.path.dirname(ruta_archivo) else None
        df_final = df_nuevo

    df_final.to_excel(ruta_archivo, index=False)
    print(f"  ✅ Archivo guardado ({ruta_archivo}): {len(df_final)} filas")


# ----------------------------------------- MAIN -----------------------------------------

def main():

    if not USER or not PASS:
        raise ValueError("Faltan credenciales.")

    # ===== OP DET =====
    print("\n===== OP DET =======")
    fecha_op_det = obtener_ultima_fecha(archivo_final, "Fecha Op-Det")
    print(f"  📅 Última fecha op-det: {fecha_op_det}")

    query_opdet = """
    SELECT Num_OP, name, id, Producir, Detalle, ClienteNombre, Estado, Total_Precio, createdAt, R49573112
    FROM productionorder 
    WHERE Estado IN (14149160, 14149163, 15549065) 
    """
    df_opdet = descargar_datos(query_opdet, "createdAt", fecha_op_det)
    df_opdet.columns = ['Num-OP', 'OP Det', 'Id op-det','Cantidad OP-D', 'Detalle', 'Cliente', 'status_id', 'Total Precio', 'Fecha Op-Det', 'Id Costeo Producto']
    # Merge para obtener el nombre del estado
    df_opdet = df_opdet.merge(
        df_estados[['status_id', 'Estado']],
        on='status_id',
        how='left'
    )
    df_opdet = df_opdet.drop(columns=['status_id'])
    
    # ===== TALLAS =====
    print("\n===== TALLAS =======")
    fecha_tallas = obtener_ultima_fecha(archivo_final, "Fecha")
    print(f"  📅 Última fecha tallas: {fecha_tallas}")

    query_tallas = """
    SELECT R13490599, R11199080, name, Ordenada, createdAt
    FROM Produccion_Prenda
    WHERE 1=1
    """
    df_tallas = descargar_datos(query_tallas, "createdAt", fecha_tallas)
    df_tallas.columns = ['Id op-det', 'Id producto', 'Talla', 'Cantidad', 'Fecha']


    # ===== PRODUCTOS =====
    print("\n===== PRODUCTOS =======")

    query_productos = """
    SELECT OBJ_NAME, id, createdAt
    FROM Producto3
    WHERE 1=1
    """
    df_productos = descargar_datos(query_productos, "createdAt")
    df_productos.columns = ['Producto', 'Id producto', 'Fecha_productos']
    
    # ===== COSTEO PRODUCTO =====
    print("\n===== COSTEO PRODUCTO =======")
    fecha_costeoprod = obtener_ultima_fecha(archivo_final, "Fecha Costeo Producto")
    print(f"  📅 Última fecha costeo producto: {fecha_costeoprod}")

    query_costeoprod = """
    SELECT name, id, R11292088, CREATED_AT
    FROM  Costeo_Producto
    WHERE 1=1
    """
    df_costeoprod = descargar_datos(query_costeoprod, "CREATED_AT", fecha_costeoprod)
    df_costeoprod.columns = ['Costeo Producto', 'Id Costeo Producto', 'Id Costeo', 'Fecha Costeo Producto']
    
    # ===== COSTEO =====
    print("\n===== COSTEO =======")
    fecha_costeo = obtener_ultima_fecha(archivo_final, "Fecha Costeo")
    print(f"  📅 Última fecha costeo: {fecha_costeo}")

    query_costeo = """
    SELECT name, id, createdBy, CREATED_AT 
    FROM Costeo
    WHERE 1=1
    """
    df_costeo = descargar_datos(query_costeo, "CREATED_AT", fecha_costeo)
    df_costeo.columns = ['Costeo', 'Id Costeo','id_comercial', 'Fecha Costeo']
    
    # Merge para obtener el nombre del comercial
    df_costeo = df_costeo.merge(
        df_comerciales,
        on='id_comercial',
        how='left'
    )
    print(df_costeo.columns.tolist())
    
    df_costeo = df_costeo[[
        'Costeo', 'Id Costeo','Comercial', 'Fecha Costeo'
    ]]
    
    # ===== PROCESOS DE PRODUCCIÓN DE OS =====
    print("\n===== SATELITE PROCESOS =======")
    fecha_sp = obtener_ultima_fecha(archivo_final_os, "Fecha SP")
    print(f"  📅 Última fecha sp: {fecha_sp}")

    query_sp = """
    SELECT R11266072, id, name, createdAt 
    FROM Satelite_Proceso
    WHERE 1=1
    """
    df_sp = descargar_datos(query_sp, "createdAt", fecha_sp)
    df_sp.columns = ['Id_OS', 'Id_Servicio', 'Servicio', 'Fecha SP']
    
    # ===== ORDENES DE SATELITE =====
    print("\n===== OS =======")
    fecha_os = obtener_ultima_fecha(archivo_final_os, "Fecha OS")
    print(f"  📅 Última fecha os: {fecha_os}")

    query_os = """
    SELECT Num_OS, name, Cantidad, id, Fecha_de_Entrega 
    FROM Orden_de_Satelite 
    WHERE 1=1
    """
    df_os = descargar_datos(query_os, "Fecha_de_Entrega", fecha_os)
    df_os.columns = ['Num OS', 'OP Det', 'Cantidad OS', 'Id_OS', 'Fecha OS']
    
    df_os["OP Det"] = df_os["OP Det"].str.replace(r'^[^-]+-[^-]+-', '', regex=True)
    
    # Unimos con satelite procesos para obtener el servicio asociado a cada OS
    df_os = df_os.merge(
        df_sp,
        on='Id_OS',
        how='left'
    )
    
    # ===== MERGES =====
    print("\n===== MERGES =======")

    # 1. Merge para obtener nombre del producto
    df_prendas = df_tallas.merge(
        df_productos[['Id producto', 'Producto']],
        on='Id producto',
        how='left'
    )
    print(f"  df_prendas: {len(df_prendas)} filas")

    # 2. Unimos las prendas con todo y talla a los OP-Det
    df_opd = df_opdet.merge(
        df_prendas,
        on='Id op-det',
        how='left'
    )
    print(f"  df_opd: {len(df_opd)} filas")
    
    # 3. Unimos el costeo x pto al costeo general para obtener el comercial
    df_costeo= df_costeo.merge(
        df_costeoprod,
        on='Id Costeo',
        how='left'
    )
    print(df_costeo.columns.tolist())
    
    df_costeo = df_costeo[[
        'Costeo', 'Costeo Producto', 'Comercial','Fecha Costeo', 'Fecha Costeo Producto', 'Id Costeo Producto'
    ]]
    
    # Finalmente unimos la tabla de costeo a la de OP-Det
    df_final = df_opd.merge(
        df_costeo,
        on='Id Costeo Producto',
        how='left'
    )
    
    # ===== Tabla OS ======
    # Unimos los datos de OP-Det a las ordenes de satelite
    df_os = df_opdet.merge(
        df_os,
        on='OP Det',
        how='left'
    )
    
    df_final = df_final [[
        'Num-OP', 'OP Det', 'Producto', 'Detalle', 'Cliente', 'Cantidad', 'Talla', 'Total Precio',
        'Estado', 'Comercial', 'Costeo', 'Costeo Producto', 'Fecha Costeo', 'Fecha Costeo Producto',
        'Id Costeo Producto', 'Fecha Op-Det', 'Fecha', 'Id op-det'
    ]]
    df_final['Detalle'] = df_final['Detalle'].str.replace('\n', ' ').str.replace('\r', ' ')
    
    df_os = df_os [[
        'Num-OP', 'OP Det', 'Cantidad OP-D', 'Detalle',  'Cliente', 'Total Precio',
        'Estado', 'Id_OS', 'Cantidad OS', 'Servicio',
        'Num OS', 'Fecha OS', 'Fecha Op-Det', 'Id op-det'
    ]]
    df_os['Detalle'] = df_os['Detalle'].str.replace('\n', ' ').str.replace('\r', ' ')
    
    # ===== GUARDAR RESULTADO FINAL =====
    guardar_datos(df_final, archivo_final, clave_dedup=['Id op-det', 'Talla'])
    print("\n✅ Proceso completado")

    guardar_datos(df_os, archivo_final_os, clave_dedup=['Id op-det', 'Num OS'])
    print("\n✅ Proceso OS completado")
if __name__ == "__main__":
    main()