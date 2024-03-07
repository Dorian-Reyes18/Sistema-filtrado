import pandas as pd
import xlsxwriter
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

    # Mostrar resultados sin repeticiones
    if not resultado_sin_repeticiones.empty:
        print("Resultados de la búsqueda:")
        print(resultado_sin_repeticiones.to_string(index=False))

        # Preguntar si desea trasladar los datos al archivo de Excel
        respuesta = input("¿Desea trasladar estos datos al archivo de Excel? (sí/no): ").lower()
        if respuesta == 'si':
            guardar_resultados_en_excel(resultado_sin_repeticiones)
    else:
        print("No se encontraron resultados para las condiciones dadas.")

def guardar_resultados_en_excel(df_resultados):
    # Obtener el nombre del archivo Excel
    archivo_excel = Path('./DatosTraslados.xlsx')

    # Si el archivo no existe, crearlo y escribir los datos
    if not archivo_excel.is_file():
        with pd.ExcelWriter(archivo_excel, engine='xlsxwriter') as writer:
            df_resultados.to_excel(writer, sheet_name='Sheet1', index=False)
    else:
        # Si el archivo existe, abrirlo y agregar los datos al final de la hoja existente
        with pd.ExcelWriter(archivo_excel, mode='a', engine='openpyxl') as writer:
            # Verificar si la hoja 'Sheet1' ya existe
            if 'Sheet1' in pd.ExcelFile(archivo_excel).sheet_names:
                # Leer el DataFrame actual en el archivo Excel
                df_existente = pd.read_excel(archivo_excel, sheet_name='Sheet1')
                
                # Concatenar el nuevo DataFrame con el existente y eliminar duplicados
                df_resultados_final = pd.concat([df_existente, df_resultados], ignore_index=True).drop_duplicates(subset=['RutaNombre', 'UsrNom'])
                
                # Escribir el DataFrame final en el archivo Excel
                df_resultados_final.to_excel(writer, sheet_name='Sheet1', index=False)
            else:
                # Si la hoja no existe, simplemente escribir el nuevo DataFrame
                df_resultados.to_excel(writer, sheet_name='Sheet1', index=False)

    print(f"Datos trasladados con éxito al archivo Excel '{archivo_excel}'.")

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
        resultados.append(resultado_ruta)

    # Concatenar todos los resultados en un solo DataFrame
    df_resultados = pd.concat(resultados)

    # Mostrar resultados y guardar en nuevo Excel
    if not df_resultados.empty:
        print("\nResultados para todas las RutaNombre:")
        print(df_resultados.to_string(index=False))

        # Imprimir elementos duplicados eliminados
        duplicados_eliminados = len(df_resultados) - len(df_resultados.drop_duplicates(subset=['RutaNombre', 'UsrNom']))
        if duplicados_eliminados > 0:
            print(f"Se eliminaron {duplicados_eliminados} elementos duplicados.")

        guardar_resultados_en_excel(df_resultados)
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
