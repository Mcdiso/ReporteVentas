import pyodbc
from tkinter import messagebox

# Variables de conexión (ajustar según tu configuración)
SERVER = "MSI\\COMPAC"
USER = "sa"
PASSWORD = "compac$1"

def test_connection():
    """Probar la conexión a la base de datos"""
    try:
        conn = pyodbc.connect(f'DRIVER={{SQL Server}};SERVER={SERVER};UID={USER};PWD={PASSWORD}')
        print("Conexión exitosa")
        conn.close()
    except Exception as e:
        print(f"Error de conexión: {e}")
        messagebox.showerror("Error de conexión", f"No se pudo conectar al servidor: {e}")

def get_databases():
    """Obtenemos las bases de datos permitidas en el servidor SQL"""
    allowed_databases = {
        "adEGA_INDUSTRIAL_ZONA": "BAJIO (LA FIERA)",
        "adEGA_INDUSTRIAL_ZONA_NORTE": "NORTE (DORADOS)",
    }

    try:
        conn = pyodbc.connect(f'DRIVER={{SQL Server}};SERVER={SERVER};UID={USER};PWD={PASSWORD}')
        query = "SELECT name FROM sys.databases"
        cursor = conn.cursor()
        cursor.execute(query)
        all_databases = [row[0] for row in cursor.fetchall()]
        conn.close()

        filtered_databases_tech = [
            db for db in all_databases if db in allowed_databases
        ]

        filtered_databases = [
            allowed_databases[db] for db in filtered_databases_tech
        ]

        return filtered_databases

    except Exception as e:
        print(f"Error al obtener bases de datos: {e}")
        messagebox.showerror("Error de conexión", f"No se pudo conectar al servidor: {e}")
        return []

def get_agents(database):
    """Obtener los agentes de ventas de la base de datos"""
    database_mapping = {
        "BAJIO (LA FIERA)": "adEGA_INDUSTRIAL_ZONA",
        "NORTE (DORADOS)": "adEGA_INDUSTRIAL_ZONA_NORTE",
    }

    try:
        tech_db_name = database_mapping.get(database)

        if not tech_db_name:
            raise Exception(f"La base de datos '{database}' no está configurada correctamente.")

        conn = pyodbc.connect(f'DRIVER={{SQL Server}};SERVER={SERVER};DATABASE={tech_db_name};UID={USER};PWD={PASSWORD}')
        query = "SELECT DISTINCT CNOMBREAGENTE FROM admAgentes"
        cursor = conn.cursor()
        cursor.execute(query)
        agents = [row[0] for row in cursor.fetchall()]
        conn.close()
        return agents
    except Exception as e:
        print(f"Error al obtener agentes: {e}")
        messagebox.showerror("Error de consulta", f"No se pudieron obtener los agentes: {e}")
        return []

def execute_query(query, params):
    """Ejecutar una consulta SQL con parámetros"""
    try:
        conn = pyodbc.connect(f'DRIVER={{SQL Server}};SERVER={SERVER};UID={USER};PWD={PASSWORD}')
        cursor = conn.cursor()
        cursor.execute(query, params)
        rows = cursor.fetchall()
        conn.close()
        return rows
    except Exception as e:
        print(f"Error al ejecutar la consulta: {e}")
        messagebox.showerror("Error de consulta", f"No se pudo ejecutar la consulta: {e}")
        return []
    
def execute_global_report_query(database, selected_agent, start_date, end_date):
    """Ejecutar la consulta SQL para obtener los datos del reporte global sin incluir 'Flete'"""
    # Mapeo de nombres amigables a nombres técnicos para la conexión
    database_mapping = {
        "BAJIO (LA FIERA)": "adEGA_INDUSTRIAL_ZONA",
        "NORTE (DORADOS)": "adEGA_INDUSTRIAL_ZONA_NORTE",
    }

    try:
        # Obtener el nombre técnico de la base de datos
        tech_db_name = database_mapping.get(database)

        if not tech_db_name:
            raise Exception(f"La base de datos '{database}' no está configurada correctamente.")

        # Conectar usando el nombre técnico de la base de datos
        conn = pyodbc.connect(f'DRIVER={{SQL Server}};SERVER={SERVER};DATABASE={tech_db_name};UID={USER};PWD={PASSWORD}')
        cursor = conn.cursor()

        # Convertir las fechas a formato YYYY-MM-DD para la consulta
        if isinstance(start_date, str):
            start_date_str = start_date
        else:
            start_date_str = start_date.strftime('%Y-%m-%d')

        if isinstance(end_date, str):
            end_date_str = end_date
        else:
            end_date_str = end_date.strftime('%Y-%m-%d')

        query = f"""
        WITH CostoHistorico AS (
            SELECT 
                ch.CIDPRODUCTO,
                ch.CULTIMOCOSTOH,
                ch.CIDALMACEN,
                ch.CFECHACOSTOH,
                m.CIDMOVIMIENTO,
                ROW_NUMBER() OVER (
                    PARTITION BY ch.CIDPRODUCTO, ch.CIDALMACEN, m.CIDMOVIMIENTO
                    ORDER BY 
                        CASE
                            WHEN ch.CFECHACOSTOH = m.CFECHA THEN 0
                            WHEN ch.CFECHACOSTOH < m.CFECHA THEN 1
                            ELSE 2
                        END,
                        ABS(DATEDIFF(DAY, ch.CFECHACOSTOH, m.CFECHA))
                ) AS rn
            FROM admCostosHistoricos ch
            INNER JOIN admMovimientos m ON ch.CIDPRODUCTO = m.CIDPRODUCTO AND ch.CIDALMACEN = m.CIDALMACEN
        )
        SELECT 
            d.CSERIEDOCUMENTO AS Serie,
            d.CFOLIO AS Folio,
            d.CRAZONSOCIAL AS Cliente,
            d.CFECHA AS Fecha_Elaboracion,
            SUM(ch.CULTIMOCOSTOH * m.CUNIDADES) AS Costo_Material,
            SUM(
                CASE 
                    WHEN d.CIDMONEDA = 1 THEN m.CPRECIO * m.CUNIDADES
                    WHEN d.CIDMONEDA = 2 THEN m.CPRECIO * m.CUNIDADES * d.CTIPOCAMBIO
                    ELSE 0
                END
            ) AS Neto_Documento,
            (SUM(
                CASE 
                    WHEN d.CIDMONEDA = 1 THEN m.CPRECIO * m.CUNIDADES
                    WHEN d.CIDMONEDA = 2 THEN m.CPRECIO * m.CUNIDADES * d.CTIPOCAMBIO
                    ELSE 0
                END
            ) - SUM(ch.CULTIMOCOSTOH * m.CUNIDADES)) AS Utilidad_Final,
            CONCAT(
                ROUND(
                    (SUM(
                        CASE 
                            WHEN d.CIDMONEDA = 1 THEN m.CPRECIO * m.CUNIDADES
                            WHEN d.CIDMONEDA = 2 THEN m.CPRECIO * m.CUNIDADES * d.CTIPOCAMBIO
                            ELSE 0
                        END
                    ) - SUM(ch.CULTIMOCOSTOH * m.CUNIDADES)) 
                    / NULLIF(SUM(
                        CASE 
                            WHEN d.CIDMONEDA = 1 THEN m.CPRECIO * m.CUNIDADES
                            WHEN d.CIDMONEDA = 2 THEN m.CPRECIO * m.CUNIDADES * d.CTIPOCAMBIO
                            ELSE 0
                        END
                    ), 0) * 100, 2
                ), ' %'
            ) AS Margen,
            ROUND((SUM(
                CASE 
                    WHEN d.CIDMONEDA = 1 THEN m.CPRECIO * m.CUNIDADES
                    WHEN d.CIDMONEDA = 2 THEN m.CPRECIO * m.CUNIDADES * d.CTIPOCAMBIO
                    ELSE 0
                END
            ) - SUM(ch.CULTIMOCOSTOH * m.CUNIDADES)) * 0.05, 2) AS Comision,
            ag.CNOMBREAGENTE AS Agente_Venta
        FROM admDocumentos d
        INNER JOIN admMovimientos m ON d.CIDDOCUMENTO = m.CIDDOCUMENTO
        LEFT JOIN (
            SELECT 
                CIDPRODUCTO,
                CULTIMOCOSTOH,
                CIDALMACEN,
                CIDMOVIMIENTO
            FROM CostoHistorico
            WHERE rn = 1
        ) ch ON m.CIDPRODUCTO = ch.CIDPRODUCTO AND m.CIDALMACEN = ch.CIDALMACEN AND m.CIDMOVIMIENTO = ch.CIDMOVIMIENTO
        LEFT JOIN admAgentes ag ON d.CIDAGENTE = ag.CIDAGENTE
        WHERE 
            d.CIDDOCUMENTODE = 4
            AND d.CCANCELADO = 0
            AND ag.CNOMBREAGENTE = ?
            AND d.CFECHA BETWEEN ? AND ?
            AND d.CIDCONCEPTODOCUMENTO IN (4, 5, 3003, 3035)
        GROUP BY 
            d.CSERIEDOCUMENTO, 
            d.CFOLIO, 
            d.CRAZONSOCIAL, 
            d.CFECHA, 
            ag.CNOMBREAGENTE
        ORDER BY d.CFECHA ASC, d.CFOLIO;
        """

        # Ejecutar la consulta
        cursor.execute(query, (selected_agent, start_date_str, end_date_str))
        rows = cursor.fetchall()
        conn.close()
        return rows
    except Exception as e:
        print(f"Error en la consulta global: {e}")
        messagebox.showerror("Error de consulta", f"No se pudo ejecutar la consulta global: {e}")
        return []

def execute_client_report_query(database, selected_agent, start_date, end_date):
    """Ejecutar la consulta SQL para obtener los datos del reporte por cliente"""
    # Mapeo de nombres amigables a nombres técnicos para la conexión
    database_mapping = {
        "BAJIO (LA FIERA)": "adEGA_INDUSTRIAL_ZONA",
        "NORTE (DORADOS)": "adEGA_INDUSTRIAL_ZONA_NORTE",
    }

    try:
        # Obtener el nombre técnico de la base de datos
        tech_db_name = database_mapping.get(database)

        if not tech_db_name:
            raise Exception(f"La base de datos '{database}' no está configurada correctamente.")

        # Conectar usando el nombre técnico de la base de datos
        conn = pyodbc.connect(f'DRIVER={{SQL Server}};SERVER={SERVER};DATABASE={tech_db_name};UID={USER};PWD={PASSWORD}')
        cursor = conn.cursor()

        # Convertir las fechas a formato YYYY-MM-DD para la consulta
        if isinstance(start_date, str):
            start_date_str = start_date
        else:
            start_date_str = start_date.strftime('%Y-%m-%d')

        if isinstance(end_date, str):
            end_date_str = end_date
        else:
            end_date_str = end_date.strftime('%Y-%m-%d')
        query = """
        WITH CostoHistorico AS (
            SELECT 
                ch.CIDPRODUCTO,
                ch.CULTIMOCOSTOH,
                ch.CIDALMACEN,
                ch.CFECHACOSTOH,
                m.CIDMOVIMIENTO,
                ROW_NUMBER() OVER (
                    PARTITION BY ch.CIDPRODUCTO, ch.CIDALMACEN, m.CIDMOVIMIENTO
                    ORDER BY 
                        CASE
                            WHEN ch.CFECHACOSTOH = m.CFECHA THEN 0 -- Igual a la fecha del movimiento
                            WHEN ch.CFECHACOSTOH < m.CFECHA THEN 1 -- Menor o igual, preferido
                            ELSE 2 -- Mayor, en último lugar
                        END,
                        ABS(DATEDIFF(DAY, ch.CFECHACOSTOH, m.CFECHA)) -- Diferencia de días
                ) AS rn
            FROM admCostosHistoricos ch
            INNER JOIN admMovimientos m ON ch.CIDPRODUCTO = m.CIDPRODUCTO AND ch.CIDALMACEN = m.CIDALMACEN
        )
        SELECT 
            d.CRAZONSOCIAL AS Cliente,
            SUM(ch.CULTIMOCOSTOH * m.CUNIDADES) AS Costo_Material,
            SUM(
                CASE 
                    WHEN d.CIDMONEDA = 1 THEN m.CPRECIO * m.CUNIDADES
                    WHEN d.CIDMONEDA = 2 THEN m.CPRECIO * m.CUNIDADES * d.CTIPOCAMBIO
                    ELSE 0
                END
            ) AS Neto_Documento,
            (SUM(
                CASE 
                    WHEN d.CIDMONEDA = 1 THEN m.CPRECIO * m.CUNIDADES
                    WHEN d.CIDMONEDA = 2 THEN m.CPRECIO * m.CUNIDADES * d.CTIPOCAMBIO
                    ELSE 0
                END
            ) - SUM(ch.CULTIMOCOSTOH * m.CUNIDADES)) AS Utilidad_Final,
            CONCAT(
                ROUND(
                    (SUM(
                        CASE 
                            WHEN d.CIDMONEDA = 1 THEN m.CPRECIO * m.CUNIDADES
                            WHEN d.CIDMONEDA = 2 THEN m.CPRECIO * m.CUNIDADES * d.CTIPOCAMBIO
                            ELSE 0
                        END
                    ) - SUM(ch.CULTIMOCOSTOH * m.CUNIDADES)) 
                    / NULLIF(SUM(
                        CASE 
                            WHEN d.CIDMONEDA = 1 THEN m.CPRECIO * m.CUNIDADES
                            WHEN d.CIDMONEDA = 2 THEN m.CPRECIO * m.CUNIDADES * d.CTIPOCAMBIO
                            ELSE 0
                        END
                    ), 0) * 100, 2
                ), ' %'
            ) AS Margen,
            ROUND((SUM(
                CASE 
                    WHEN d.CIDMONEDA = 1 THEN m.CPRECIO * m.CUNIDADES
                    WHEN d.CIDMONEDA = 2 THEN m.CPRECIO * m.CUNIDADES * d.CTIPOCAMBIO
                    ELSE 0
                END
            ) - SUM(ch.CULTIMOCOSTOH * m.CUNIDADES)) * 0.05, 2) AS Comision,
            ag.CNOMBREAGENTE AS Agente_Venta
        FROM admDocumentos d
        INNER JOIN admMovimientos m ON d.CIDDOCUMENTO = m.CIDDOCUMENTO
        LEFT JOIN (
            SELECT 
                CIDPRODUCTO,
                CULTIMOCOSTOH,
                CIDALMACEN,
                CIDMOVIMIENTO
            FROM CostoHistorico
            WHERE rn = 1
        ) ch ON m.CIDPRODUCTO = ch.CIDPRODUCTO AND m.CIDALMACEN = ch.CIDALMACEN AND m.CIDMOVIMIENTO = ch.CIDMOVIMIENTO
        LEFT JOIN admAgentes ag ON d.CIDAGENTE = ag.CIDAGENTE
        WHERE 
            d.CIDDOCUMENTODE = 4
            AND d.CCANCELADO = 0
            AND ag.CNOMBREAGENTE = ?
            AND d.CFECHA BETWEEN ? AND ?
            AND d.CIDCONCEPTODOCUMENTO IN (4, 5, 3003, 3035)
        GROUP BY 
            d.CRAZONSOCIAL, 
            ag.CNOMBREAGENTE
        ORDER BY 
            d.CRAZONSOCIAL, 
            ag.CNOMBREAGENTE;
        """
        # Ejecutar la consulta
        cursor.execute(query, (selected_agent, start_date_str, end_date_str))
        rows = cursor.fetchall()
        conn.close()
        return rows
    except Exception as e:
        print(f"Error en la consulta por cliente: {e}")
        messagebox.showerror("Error de consulta", f"No se pudo ejecutar la consulta por cliente: {e}")
        return []
    
def execute_partida_report_query(database, selected_agent, start_date, end_date):
    """Ejecutar la consulta SQL para obtener los datos del reporte por partida"""
    database_mapping = {
        "BAJIO (LA FIERA)": "adEGA_INDUSTRIAL_ZONA",
        "NORTE (DORADOS)": "adEGA_INDUSTRIAL_ZONA_NORTE",
    }

    try:
        # Obtener el nombre técnico de la base de datos
        tech_db_name = database_mapping.get(database)

        if not tech_db_name:
            raise Exception(f"La base de datos '{database}' no está configurada correctamente.")

        # Conectar usando el nombre técnico de la base de datos
        conn = pyodbc.connect(f'DRIVER={{SQL Server}};SERVER={SERVER};DATABASE={tech_db_name};UID={USER};PWD={PASSWORD}')
        cursor = conn.cursor()

        # Convertir las fechas a formato YYYY-MM-DD para la consulta
        if isinstance(start_date, str):
            start_date_str = start_date
        else:
            start_date_str = start_date.strftime('%Y-%m-%d')

        if isinstance(end_date, str):
            end_date_str = end_date
        else:
            end_date_str = end_date.strftime('%Y-%m-%d')

        query = """
        WITH CostoHistorico AS (
            SELECT 
                ch.CIDPRODUCTO,
                ch.CULTIMOCOSTOH,
                ch.CIDALMACEN,
                ch.CFECHACOSTOH,
                m.CIDMOVIMIENTO,
                ROW_NUMBER() OVER (
                    PARTITION BY ch.CIDPRODUCTO, ch.CIDALMACEN, m.CIDMOVIMIENTO
                    ORDER BY 
                        CASE
                            WHEN ch.CFECHACOSTOH = m.CFECHA THEN 0
                            WHEN ch.CFECHACOSTOH < m.CFECHA THEN 1
                            ELSE 2
                        END,
                        ABS(DATEDIFF(DAY, ch.CFECHACOSTOH, m.CFECHA))
                ) AS rn
            FROM admCostosHistoricos ch
            INNER JOIN admMovimientos m ON ch.CIDPRODUCTO = m.CIDPRODUCTO AND ch.CIDALMACEN = m.CIDALMACEN
        )
        SELECT 
            d.CSERIEDOCUMENTO AS Serie,
            d.CFOLIO AS Folio,
            m.CIDPRODUCTO AS Partida,
            d.CFECHA AS Fecha_Elaboracion,
            SUM(ch.CULTIMOCOSTOH * m.CUNIDADES) AS Costo_Material,
            SUM(
                CASE 
                    WHEN d.CIDMONEDA = 1 THEN m.CPRECIO * m.CUNIDADES
                    WHEN d.CIDMONEDA = 2 THEN m.CPRECIO * m.CUNIDADES * d.CTIPOCAMBIO
                    ELSE 0
                END
            ) AS Neto_Documento,
            (SUM(
                CASE 
                    WHEN d.CIDMONEDA = 1 THEN m.CPRECIO * m.CUNIDADES
                    WHEN d.CIDMONEDA = 2 THEN m.CPRECIO * m.CUNIDADES * d.CTIPOCAMBIO
                    ELSE 0
                END
            ) - SUM(ch.CULTIMOCOSTOH * m.CUNIDADES)) AS Utilidad_Final,
            CONCAT(
                ROUND(
                    (SUM(
                        CASE 
                            WHEN d.CIDMONEDA = 1 THEN m.CPRECIO * m.CUNIDADES
                            WHEN d.CIDMONEDA = 2 THEN m.CPRECIO * m.CUNIDADES * d.CTIPOCAMBIO
                            ELSE 0
                        END
                    ) - SUM(ch.CULTIMOCOSTOH * m.CUNIDADES)) 
                    / NULLIF(SUM(
                        CASE 
                            WHEN d.CIDMONEDA = 1 THEN m.CPRECIO * m.CUNIDADES
                            WHEN d.CIDMONEDA = 2 THEN m.CPRECIO * m.CUNIDADES * d.CTIPOCAMBIO
                            ELSE 0
                        END
                    ), 0) * 100, 2
                ), ' %'
            ) AS Margen,
            ROUND((SUM(
                CASE 
                    WHEN d.CIDMONEDA = 1 THEN m.CPRECIO * m.CUNIDADES
                    WHEN d.CIDMONEDA = 2 THEN m.CPRECIO * m.CUNIDADES * d.CTIPOCAMBIO
                    ELSE 0
                END
            ) - SUM(ch.CULTIMOCOSTOH * m.CUNIDADES)) * 0.05, 2) AS Comision,
            ag.CNOMBREAGENTE AS Agente_Venta
        FROM admDocumentos d
        INNER JOIN admMovimientos m ON d.CIDDOCUMENTO = m.CIDDOCUMENTO
        LEFT JOIN (
            SELECT 
                CIDPRODUCTO,
                CULTIMOCOSTOH,
                CIDALMACEN,
                CIDMOVIMIENTO
            FROM CostoHistorico
            WHERE rn = 1
        ) ch ON m.CIDPRODUCTO = ch.CIDPRODUCTO AND m.CIDALMACEN = ch.CIDALMACEN AND m.CIDMOVIMIENTO = ch.CIDMOVIMIENTO
        LEFT JOIN admAgentes ag ON d.CIDAGENTE = ag.CIDAGENTE
        WHERE 
            d.CIDDOCUMENTODE = 4
            AND d.CCANCELADO = 0
            AND d.CIDCONCEPTODOCUMENTO IN (4, 5, 3003, 3035)
            AND ag.CNOMBREAGENTE = ?
            AND d.CFECHA BETWEEN ? AND ?
        GROUP BY 
            d.CSERIEDOCUMENTO, 
            d.CFOLIO, 
            m.CIDPRODUCTO, 
            d.CFECHA, 
            ag.CNOMBREAGENTE
        ORDER BY d.CFECHA ASC, d.CFOLIO;
        """

        # Ejecutar la consulta
        cursor.execute(query, (selected_agent, start_date_str, end_date_str))
        rows = cursor.fetchall()
        conn.close()
        return rows
    except Exception as e:
        print(f"Error en la consulta por partida: {e}")
        messagebox.showerror("Error de consulta", f"No se pudo ejecutar la consulta por partida: {e}")
        return []
