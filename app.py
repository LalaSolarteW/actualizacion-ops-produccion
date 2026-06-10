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
df_estados = pd.read_csv('estados_impel.csv', encoding='latin-1')
df_comerciales = pd.read_csv('comerciales.csv', encoding='utf-8-sig')


# ----------------------------------------- FUNCIONES -----------------------------------------

def descargar_datos(query_base, campo_fecha):
    """Descargar datos de la API en bloques de 10k filas."""
    all_data = []
    start = 0
    limit = 10000

    while True:
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


def limpiar_texto(serie):
    """Limpiar caracteres especiales que corrompen el Excel."""
    return serie.str.replace('\n', ' ', regex=False) \
                .str.replace('\r', ' ', regex=False) \
                .str.replace('\t', ' ', regex=False) \
                .str.replace(r'[^\x20-\x7E\xC0-\xFF]', '', regex=True)


def guardar_datos(df, ruta_archivo):
    """Guardar DataFrame en Excel, sobreescribiendo el archivo anterior."""
    if df.empty:
        print("  ⚠️ No hay datos")
        return
    df.to_excel(ruta_archivo, index=False)
    print(f"  ✅ Archivo guardado ({ruta_archivo}): {len(df)} filas")


# ----------------------------------------- MAIN -----------------------------------------

def main():

    if not USER or not PASS:
        raise ValueError("Faltan credenciales.")

    # ===== OP DET =====
    print("\n===== OP DET =======")
    query_opdet = """
    SELECT Num_OP, name, id, Producir, Detalle, ClienteNombre, Estado, Total_Precio, createdAt, R49573112
    FROM productionorder 
    WHERE Estado IN (14149160, 14149163, 15549065, 14149164)
    AND 1=1
    """
    df_opdet = descargar_datos(query_opdet, "createdAt")
    df_opdet.columns = ['Num-OP', 'OP Det', 'Id op-det', 'Cantidad OP-D', 'Detalle', 'Cliente', 'status_id', 'Total Precio', 'Fecha Op-Det', 'Id Costeo Producto']
    df_opdet = df_opdet.merge(
        df_estados[['status_id', 'Estado']],
        on='status_id',
        how='left'
    )
    df_opdet = df_opdet.drop(columns=['status_id'])
    print(f"  df_opdet: {len(df_opdet)} filas")

    # ===== TALLAS =====
    print("\n===== TALLAS =======")
    query_tallas = """
    SELECT R13490599, R11199080, name, Ordenada, createdAt
    FROM Produccion_Prenda
    WHERE 1=1
    """
    df_tallas = descargar_datos(query_tallas, "createdAt")
    df_tallas.columns = ['Id op-det', 'Id producto', 'Talla', 'Cantidad', 'Fecha']
    print(f"  df_tallas: {len(df_tallas)} filas")

    # ===== PRODUCTOS =====
    print("\n===== PRODUCTOS =======")
    query_productos = """
    SELECT OBJ_NAME, id, createdAt
    FROM Producto3
    WHERE 1=1
    """
    df_productos = descargar_datos(query_productos, "createdAt")
    df_productos.columns = ['Producto', 'Id producto', 'Fecha_productos']
    print(f"  df_productos: {len(df_productos)} filas")

    # ===== COSTEO PRODUCTO =====
    print("\n===== COSTEO PRODUCTO =======")
    query_costeoprod = """
    SELECT name, id, R11292088, CREATED_AT
    FROM Costeo_Producto
    WHERE 1=1
    """
    df_costeoprod = descargar_datos(query_costeoprod, "CREATED_AT")
    df_costeoprod.columns = ['Costeo Producto', 'Id Costeo Producto', 'Id Costeo', 'Fecha Costeo Producto']
    print(f"  df_costeoprod: {len(df_costeoprod)} filas")

    # ===== COSTEO =====
    print("\n===== COSTEO =======")
    query_costeo = """
    SELECT name, id, createdBy, CREATED_AT 
    FROM Costeo
    WHERE 1=1
    """
    df_costeo = descargar_datos(query_costeo, "CREATED_AT")
    df_costeo.columns = ['Costeo', 'Id Costeo', 'id_comercial', 'Fecha Costeo']
    df_costeo = df_costeo.merge(
        df_comerciales,
        on='id_comercial',
        how='left'
    )
    df_costeo = df_costeo[['Costeo', 'Id Costeo', 'Comercial', 'Fecha Costeo']]
    print(f"  df_costeo: {len(df_costeo)} filas")

    # ===== SATELITE PROCESOS =====
    print("\n===== SATELITE PROCESOS =======")
    query_sp = """
    SELECT R11266072, id, name, createdAt 
    FROM Satelite_Proceso
    WHERE 1=1
    """
    df_sp = descargar_datos(query_sp, "createdAt")
    df_sp.columns = ['Id_OS', 'Id_Servicio', 'Servicio', 'Fecha SP']
    print(f"  df_sp: {len(df_sp)} filas")

    # ===== ORDENES DE SATELITE =====
    print("\n===== OS =======")
    query_os = """
    SELECT Num_OS, name, Cantidad, id, Fecha_de_Entrega 
    FROM Orden_de_Satelite 
    WHERE 1=1
    """
    df_os = descargar_datos(query_os, "Fecha_de_Entrega")
    df_os.columns = ['Num OS', 'OP Det', 'Cantidad OS', 'Id_OS', 'Fecha OS']
    df_os["OP Det"] = df_os["OP Det"].str.replace(r'^[^-]+-[^-]+-', '', regex=True)
    df_os = df_os.merge(df_sp, on='Id_OS', how='left')
    print(f"  df_os: {len(df_os)} filas")

    # ===== MERGES =====
    print("\n===== MERGES =======")

    # 1. Tallas + Productos
    df_prendas = df_tallas.merge(
        df_productos[['Id producto', 'Producto']],
        on='Id producto',
        how='left'
    )
    print(f"  df_prendas: {len(df_prendas)} filas")

    # 2. OP Det + Prendas
    df_opd = df_opdet.merge(
        df_prendas,
        on='Id op-det',
        how='left'
    )
    print(f"  df_opd: {len(df_opd)} filas")

    # 3. Costeo + Costeo Producto
    df_costeo = df_costeo.merge(
        df_costeoprod,
        on='Id Costeo',
        how='left'
    )
    df_costeo = df_costeo[[
        'Costeo', 'Costeo Producto', 'Comercial', 'Fecha Costeo', 'Fecha Costeo Producto', 'Id Costeo Producto'
    ]]

    # 4. OP Det + Costeo
    df_final = df_opd.merge(
        df_costeo,
        on='Id Costeo Producto',
        how='left'
    )

    # 5. OS + OP Det
    df_os_final = df_opdet.merge(
        df_os,
        on='OP Det',
        how='left'
    )

    # ===== SELECCION DE COLUMNAS =====
    df_final = df_final[[
        'Num-OP', 'OP Det', 'Producto', 'Detalle', 'Cliente', 'Cantidad', 'Talla', 'Total Precio',
        'Estado', 'Comercial', 'Costeo', 'Costeo Producto', 'Fecha Costeo', 'Fecha Costeo Producto',
        'Id Costeo Producto', 'Fecha Op-Det', 'Fecha', 'Id op-det'
    ]]

    df_os_final = df_os_final[[
        'Num-OP', 'OP Det', 'Cantidad OP-D', 'Detalle', 'Cliente', 'Total Precio',
        'Estado', 'Id_OS', 'Cantidad OS', 'Servicio',
        'Num OS', 'Fecha OS', 'Fecha Op-Det', 'Id op-det'
    ]]

    # ===== LIMPIEZA DE TEXTO =====
    df_final['Detalle'] = limpiar_texto(df_final['Detalle'])
    df_os_final['Detalle'] = limpiar_texto(df_os_final['Detalle'])

    # ===== GUARDAR =====
    guardar_datos(df_final, archivo_final)
    print("\n✅ Proceso completado")

    guardar_datos(df_os_final, archivo_final_os)
    print("\n✅ Proceso OS completado")


if __name__ == "__main__":
    main()