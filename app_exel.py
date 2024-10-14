import os
import psycopg2
from datetime import datetime
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd

def obtener_ultimo_id_consejos(cursor):
    cursor.execute("SELECT MAX(id_consejos) FROM tbl_circuito_comunal_consejos1")
    result = cursor.fetchone()
    return result[0] if result[0] is not None else 0

def post_registrar_comuna(data):
    conn = psycopg2.connect(
        host="", 
        port="5432", 
        dbname="", 
        user="postgres", 
        password=""
    )
    cursor = conn.cursor()

    try:
        conn.autocommit = False
        ultimo_id_brigada = obtener_ultimo_id_consejos(cursor)
        nuevo_id_brigada = ultimo_id_brigada + 1

        ultimo_id_consejos = obtener_ultimo_id_consejos(cursor)
        nuevo_id_consejos = ultimo_id_consejos + 1

        data['id_brigada'] = nuevo_id_brigada
        data['id_consejos'] = nuevo_id_consejos
        
        brigada_data = (
            data['id_consejos'],
            data['id_brigada'],
            data['nombre_consejos'],
            data['nombre_consejos'],
            data['id_rol_estructura'],
            data['codigoestado'],
            data['codigomunicipio'],
            data['codigoparroquia'],
            data['codigo'],
            data['activo'],
            datetime.now().strftime('%Y-%m-%d'),
        )

        query_consejos = """
            INSERT INTO tbl_circuito_comunal_consejos1 (
                id_consejos, id_brigada, nombre_consejos, nombre_sector, id_rol_estructura, 
                codigoestado, codigomunicipio, codigoparroquia, 
                codigo, activo, fecha
            ) VALUES (%s,%s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
        """
        cursor.execute(query_consejos, brigada_data)

        if cursor.rowcount == 0:
            conn.rollback()
            return {'status': False, 'error': 'Error en la inserción de datos en tbl_circuito_comunal_consejos1.'}

        conn.commit()
        return {'status': True, 'message': 'Registro exitoso'}

    except Exception as e:
        conn.rollback()
        return {'status': False, 'error': str(e)}

    finally:
        cursor.close()
        conn.close()

def registrar_comunas_desde_excel(ruta_excel, progress, ventana):
    df = pd.read_excel(ruta_excel, engine='openpyxl')
    df.columns = df.columns.str.strip()
    
    total_filas = len(df)
    progress['maximum'] = total_filas

    for index, row in df.iterrows():
        data = {
            'nombre_consejos': row['nombre_consejos'],
            'codigoestado': row['codigoestado'],
            'codigomunicipio': row['codigomunicipio'],
            'codigoparroquia': row['codigoparroquia'],
            'id_rol_estructura': row['id_rol_estructura'],
            'codigo': row['codigo'],
            'activo': row['activo'],
        }

        resultado = post_registrar_comuna(data)
        if not resultado['status']:
            messagebox.showerror("Error", resultado['error'])
            return

        progress['value'] = index + 1
        ventana.update_idletasks()

    messagebox.showinfo("Éxito", "Todos los registros se han insertado correctamente.")
    progress['value'] = 0

def seleccionar_archivo(progress, ventana):
    ruta_excel = filedialog.askopenfilename(
        title="Seleccionar archivo Excel",
        filetypes=[("Archivos Excel", "*.xlsx *.xls")]
    )
    if ruta_excel:
        registrar_comunas_desde_excel(ruta_excel, progress, ventana)

def descargar_formato_excel():
    columnas = [
        'nombre_consejos', 'codigoestado', 'codigomunicipio', 'codigoparroquia',
        'id_rol_estructura', 'codigo', 'activo'
    ]
    df = pd.DataFrame(columns=columnas)

    ruta_documentos = os.path.join(os.path.expanduser("~"), "Documents", "formato_comunas.xlsx")
    df.to_excel(ruta_documentos, index=False, engine='openpyxl')

    messagebox.showinfo("Formato Descargado", f"El formato Excel se ha guardado en {ruta_documentos}")

def crear_interfaz():
    ventana = tk.Tk()
    ventana.title("Registro de Comunas desde Excel")
    ventana.geometry("400x300")

    etiqueta = tk.Label(ventana, text="Seleccione el archivo Excel para registrar comunas:")
    etiqueta.pack(pady=10)

    progress = ttk.Progressbar(ventana, orient="horizontal", length=300, mode="determinate")
    progress.pack(pady=10)

    boton_seleccionar = tk.Button(ventana, text="Seleccionar Archivo", command=lambda: seleccionar_archivo(progress, ventana))
    boton_seleccionar.pack(pady=5)

    etiqueta2 = tk.Label(ventana, text="O descargue el formato de archivo Excel:")
    etiqueta2.pack(pady=10)

    boton_descargar = tk.Button(ventana, text="Descargar Formato Excel", command=descargar_formato_excel)
    boton_descargar.pack(pady=5)

    ventana.mainloop()

if __name__ == "__main__":
    crear_interfaz()