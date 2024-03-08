import pandas as pd
from pathlib import Path

def obtener_condiciones_por_teclado():
    print("Ingresa las condiciones de búsqueda:")
    fac_gr_nombre = input("FacGrNombre: ")
    ruta_nombre = input("RutaNombre: ")

    return {'FacGrNombre': fac_gr_nombre, 'RutaNombre': ruta_nombre}

def buscar_usuarios_sin_repeticiones(df):
    # Obtener condiciones de búsqueda por teclado
    condiciones = obtener_condiciones_por_teclado()

    # Filtrar datos según las condiciones
    resultado = df[(df['FacGrNombre'] == condiciones['FacGrNombre']) &
                   (df['RutaNombre'] == condiciones['RutaNombre'])]

    # Eliminar duplicados de UsrNom
    resultado_sin_repeticiones = resultado[['RutaNombre', 'UsrNom', 'UsrPersona']].drop_duplicates(subset='UsrNom')

    # Mostrar resultados en consola
    if not resultado_sin_repeticiones.empty:
        print("Resultados de la búsqueda:")
        print(resultado_sin_repeticiones.to_string(index=False))

        # Exportar resultados al archivo de Excel sin preguntar al usuario
        guardar_resultados_en_excel(resultado_sin_repeticiones, 'FiltroRutaUnica')
    else:
        print("No se encontraron resultados para las condiciones dadas.")

def guardar_resultados_en_excel(df_resultados, nombre_archivo):
    archivo_excel = Path(f'./{nombre_archivo}.xlsx')

    if not archivo_excel.is_file():
        opcion = input("La hoja 'Sheet1' no existe. ¿Qué desea hacer?\n1. Agregar\n2. Cancelar\nIngrese el número de la opción deseada: ")
        if opcion == '1':
            with pd.ExcelWriter(archivo_excel, engine='xlsxwriter') as writer:
                df_resultados.to_excel(writer, sheet_name='Sheet1', index=False)
            print(f"Datos trasladados con éxito al archivo Excel '{archivo_excel}'.")
        elif opcion == '2':
            print("Operación cancelada.")
        else:
            print("Opción no válida. Intente de nuevo.")
    else:
        opcion = input("La hoja 'Sheet1' ya existe. ¿Qué desea hacer?\n1. Sobrescribir\n2. Cancelar\nIngrese el número de la opción deseada: ")
        if opcion == '1':
            with pd.ExcelWriter(archivo_excel, engine='xlsxwriter', mode='w') as writer:
                df_resultados.to_excel(writer, sheet_name='Sheet1', index=False)
            print(f"Datos trasladados con éxito al archivo Excel '{archivo_excel}' (hoja sobrescrita).")
        elif opcion == '2':
            print("Operación cancelada.")
        else:
            print("Opción no válida. Intente de nuevo.")





def obtener_rutas_desde_hoja2(archivo_excel):
    try:
        xl = pd.ExcelFile(archivo_excel)
        hoja2 = xl.parse(sheet_name='Hoja2')  # Ajusta el nombre de la hoja si es diferente
        rutas = hoja2['RutaNombre'].tolist()
        return rutas
    except FileNotFoundError:
        print("¡Error! Archivo no encontrado.")
        return []

def buscar_usrnoms_por_rutas(df, rutas):
    resultados = []

    # Realizar la búsqueda por cada RutaNombre
    for ruta in rutas:
        resultado_ruta = df[df['RutaNombre'] == ruta][['RutaNombre', 'UsrNom', 'UsrPersona']].drop_duplicates(subset='UsrNom')
        resultado_ruta['RutaNombre'] = ruta
        resultados.append(resultado_ruta)

    # Concatenar todos los resultados en un solo DataFrame
    df_resultados = pd.concat(resultados)

    # Mostrar resultados en consola y guardar en nuevo Excel
    if not df_resultados.empty:
        print("\nResultados para todas las RutaNombre:")
        print(df_resultados.to_string(index=False))

        guardar_resultados_en_excel(df_resultados, 'FiltroRutasMultiples')
    else:
        print("No se encontraron resultados para las RutaNombre dadas.")

# Menú de opciones
def menu(df):
    print("Seleccione una opción:")
    print("1. Filtrar por FacGrNombre y RutaNombre")
    print("2. Mostrar UsrNom por RutaNombre desde Hoja2")
    opcion = input("Ingrese el número de la opción deseada: ")

    if opcion == '1':
        # Opción 1: Filtrar por FacGrNombre y RutaNombre
        buscar_usuarios_sin_repeticiones(df)
    elif opcion == '2':
        # Opción 2: Mostrar UsrNom por RutaNombre desde Hoja2
        rutas = obtener_rutas_desde_hoja2('./SistemaConsultas_RutasLectores.xlsx')
        if rutas:
            buscar_usrnoms_por_rutas(df, rutas)
        else:
            print("No se encontraron RutaNombre en la hoja2.")
    else:
        print("Opción no válida. Intente de nuevo.")

# Lógica principal
def main():
    # Llama a la función con tu archivo Excel
    archivo_excel = './SistemaConsultas_RutasLectores.xlsx'
    df = pd.read_excel(archivo_excel)

    while True:
        menu(df)

if __name__ == "__main__":
    main()
