import pandas as pd
import xlsxwriter

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

    # Eliminar duplicados de UsrNom y UsrPersona
    resultado_sin_repeticiones = resultado[['UsrNom', 'UsrPersona']].drop_duplicates()

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
    # Crear un nuevo archivo Excel y guardar los datos
    with pd.ExcelWriter('./DatosTraslados.xlsx', engine='xlsxwriter') as writer:
        df_resultados.to_excel(writer, sheet_name='Sheet1', index=False)
    
    print("Datos trasladados con éxito al nuevo archivo Excel 'DatosTraslados.xlsx'.")


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
    resultados = {}

    # Realizar la búsqueda por cada RutaNombre
    for ruta in rutas:
        usrnom_set = set(df[df['RutaNombre'] == ruta]['UsrNom'])
        resultados[ruta] = usrnom_set

    # Mostrar resultados
    for ruta, usrnoms in resultados.items():
        print(f"\nResultados para RutaNombre: {ruta}")
        if usrnoms:
            for usrnom in usrnoms:
                print(usrnom)
        else:
            print("No se encontraron resultados para la RutaNombre dada.")

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
