import flet as ft
import pandas as pd
from flet import FilePicker, FilePickerResultEvent

def main(page: ft.Page):
    """Función principal para la interfaz gráfica"""
    page.title = "Combinador de Archivos Excel"
    page.theme_mode = "light"
    page.padding = 20

    # Variables para almacenar las rutas de los archivos y la carpeta de destino
    archivos = []
    carpeta_destino = None
    num_hojas = 0

    # Función para combinar archivos
    def combinar_archivos(e):
        if len(archivos) < num_hojas:
            mensaje.value = f"Por favor, selecciona {num_hojas} archivos."
            page.update()
            return
        if not carpeta_destino:
            mensaje.value = "Por favor, selecciona una carpeta de destino."
            page.update()
            return

        try:
            # Verificar si openpyxl está instalado
            try:
                import openpyxl
            except ImportError:
                mensaje.value = "Error: 'openpyxl' no está instalado. Ejecuta 'pip install openpyxl'."
                page.update()
                return

            # Cargar los archivos Excel
            dfs = []
            for archivo in archivos:
                try:
                    df = pd.read_excel(archivo, engine="openpyxl")
                    dfs.append(df)
                except Exception as e:
                    mensaje.value = f"Error al leer el archivo {archivo}: {str(e)}"
                    page.update()
                    return

            # Determinar el número mínimo de filas entre todos los DataFrames
            num_filas = min(len(df) for df in dfs)

            # Truncar todos los DataFrames para que tengan el mismo número de filas
            dfs = [df.head(num_filas) for df in dfs]

            # Combinar los DataFrames basados en el índice (orden de las filas)
            df_final = dfs[0]
            for df in dfs[1:]:
                # Renombrar columnas duplicadas para evitar conflictos
                columnas_duplicadas = set(df_final.columns) & set(df.columns)
                df = df.rename(columns={col: f"{col}_df{dfs.index(df) + 1}" for col in columnas_duplicadas})
                # Combinar los DataFrames
                df_final = df_final.join(df)

            # Guardar en un nuevo archivo Excel
            archivo_salida = f"{carpeta_destino}/archivo_combinado.xlsx"
            df_final.to_excel(archivo_salida, index=False, engine="openpyxl")
            mensaje.value = f"¡Archivo combinado guardado como '{archivo_salida}'!"
        except Exception as e:
            mensaje.value = f"Error: {str(e)}"
        page.update()

    # Función para manejar la selección de archivos
    def seleccionar_archivo(e: FilePickerResultEvent):
        if e.files:
            archivos.append(e.files[0].path)
            mensaje.value = f"Archivos seleccionados: {len(archivos)}/{num_hojas}"
            if len(archivos) == num_hojas:
                boton_seleccionar_archivo.disabled = True  # Deshabilitar el botón de selección
                boton_combinar.disabled = False  # Habilitar el botón de combinar
            page.update()
        else:
            mensaje.value = "Ningún archivo seleccionado."
        page.update()

    # Función para manejar la selección de la carpeta de destino
    def seleccionar_carpeta_destino(e: FilePickerResultEvent):
        nonlocal carpeta_destino
        if e.path:
            carpeta_destino = e.path
            mensaje.value = f"Carpeta de destino seleccionada: {carpeta_destino}"
            if len(archivos) == num_hojas:
                boton_combinar.disabled = False  # Habilitar el botón de combinar
        else:
            mensaje.value = "Ninguna carpeta seleccionada."
        page.update()

    # Función para retroceder
    def retroceder(e):
        nonlocal num_hojas, archivos
        num_hojas = 0
        archivos = []
        mensaje.value = "Selecciona cuántas hojas deseas unir."
        boton_seleccionar_archivo.disabled = False  # Habilitar el botón de selección
        boton_combinar.disabled = True  # Deshabilitar el botón de combinar
        page.controls.clear()
        page.add(contenido_inicial)
        page.update()

    # Configurar el FilePicker
    file_picker = FilePicker(on_result=seleccionar_archivo)
    folder_picker = FilePicker(on_result=seleccionar_carpeta_destino)
    page.overlay.extend([file_picker, folder_picker])

    # Elementos de la interfaz
    mensaje = ft.Text(value="Selecciona cuántas hojas deseas unir.", color="blue")
    input_hojas = ft.TextField(label="Número de hojas", width=200)
    boton_confirmar_hojas = ft.ElevatedButton(
        "Confirmar",
        on_click=lambda e: avanzar_seleccion_archivos(),
    )
    boton_retroceder = ft.ElevatedButton(
        "Retroceder",
        on_click=retroceder,
    )
    boton_seleccionar_archivo = ft.ElevatedButton(
        "Seleccionar Archivo",
        on_click=lambda _: file_picker.pick_files(),
        disabled=True,  # Deshabilitado inicialmente
    )
    boton_seleccionar_carpeta = ft.ElevatedButton(
        "Seleccionar Carpeta de Destino",
        on_click=lambda _: folder_picker.get_directory_path(),
    )
    boton_combinar = ft.ElevatedButton(
        "Combinar Archivos",
        on_click=combinar_archivos,
        disabled=True,  # Deshabilitado inicialmente
    )

    # Contenido inicial
    contenido_inicial = ft.Column(
        [
            mensaje,
            input_hojas,
            ft.Row([boton_confirmar_hojas, boton_retroceder], alignment=ft.MainAxisAlignment.CENTER),
        ],
        alignment=ft.MainAxisAlignment.CENTER,
        horizontal_alignment=ft.CrossAxisAlignment.CENTER,
    )

    # Función para avanzar a la selección de archivos
    def avanzar_seleccion_archivos():
        nonlocal num_hojas
        try:
            num_hojas = int(input_hojas.value)
            if num_hojas < 1:
                mensaje.value = "El número de hojas debe ser al menos 1."
                page.update()
                return
            mensaje.value = f"Selecciona {num_hojas} archivos."
            page.controls.clear()
            page.add(
                ft.Column(
                    [
                        mensaje,
                        ft.Row([boton_retroceder], alignment=ft.MainAxisAlignment.CENTER),
                        ft.Row([boton_seleccionar_archivo], alignment=ft.MainAxisAlignment.CENTER),
                        ft.Row([boton_seleccionar_carpeta], alignment=ft.MainAxisAlignment.CENTER),
                        ft.Row([boton_combinar], alignment=ft.MainAxisAlignment.CENTER),
                    ],
                    alignment=ft.MainAxisAlignment.CENTER,
                    horizontal_alignment=ft.CrossAxisAlignment.CENTER,
                )
            )
            boton_seleccionar_archivo.disabled = False  # Habilitar el botón de selección
        except ValueError:
            mensaje.value = "Por favor, ingresa un número válido."
        page.update()

    # Iniciar con el contenido inicial
    page.add(contenido_inicial)

# Iniciar la aplicación
ft.app(target=main)