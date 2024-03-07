import pandas as pd

def obtener_condiciones_por_teclado():
    print("Ingresa las condiciones de búsqueda:")
    fac_gr_nombre = input("FacGrNombre: ")
    ruta_nombre = input("RutaNombre: ")

    return {'FacGrNombre': fac_gr_nombre, 'RutaNombre': ruta_nombre}

def buscar_usuarios_sin_repeticiones(SistemaConsultas_RutasLectores):
    # Cargar el archivo Excel
    try:
        xl = pd.ExcelFile(SistemaConsultas_RutasLectores)
    except FileNotFoundError:
        print("¡Error! Archivo no encontrado.")
        return

    # Leer la hoja de datos
    df = xl.parse()

    # Obtener condiciones de búsqueda por teclado
    condiciones = obtener_condiciones_por_teclado()

    # Filtrar datos según las condiciones
    resultado = df[(df['FacGrNombre'] == condiciones['FacGrNombre']) &
                   (df['RutaNombre'] == condiciones['RutaNombre'])]

    # Eliminar duplicados de UsrNom y UsrPersona
    resultado_sin_repeticiones = resultado[['UsrNom', 'UsrPersona']].drop_duplicates()

    # Mostrar resultados sin repeticiones
    if not resultado_sin_repeticiones.empty:
        print("Resultados de la búsqueda:")
        print(resultado_sin_repeticiones.to_string(index=False))

        # Guardar resultados en un archivo CSV
        archivo_destino = './resultados_busqueda.csv'
        resultado_sin_repeticiones.to_csv(archivo_destino, index=False)
        print(f"Resultados guardados en {archivo_destino}")
    else:
        print("No se encontraron resultados para las condiciones dadas.")

# Llama a la función con tu archivo Excel
archivo_excel = './SistemaConsultas_RutasLectores.xlsx'
buscar_usuarios_sin_repeticiones(archivo_excel)
