import math
import numpy as np
import pandas as pd


def calcular_estadisticas(datos):
    n = len(datos)
    promedio = np.mean(datos)

    # Desviación estándar muestral (N - 1)
    desviacion_estandar = np.std(datos, ddof=1) if n > 1 else 0.0

    # Error absoluto de la media
    error_absoluto = (
        desviacion_estandar / math.sqrt(n) if n > 1 else 0.0
    )

    # Error relativo y porcentual
    if promedio != 0:
        error_relativo = error_absoluto / abs(promedio)
        error_porcentual = error_relativo * 100
    else:
        error_relativo = 0.0
        error_porcentual = 0.0

    return {
        "promedio": promedio,
        "desviacion_estandar": desviacion_estandar,
        "error_absoluto": error_absoluto,
        "error_relativo": error_relativo,
        "error_porcentual": error_porcentual,
    }


def solicitar_datos():
    print("=== PROGRAMA DE CÁLCULO DE ERRORES EXPERIMENTALES ===")

    magnitud = input(
        "Ingrese el nombre de la magnitud (ej. Masa, Tiempo, Longitud): "
    ).strip()
    unidad = input(
        "Ingrese la unidad de medida (ej. kg, s, m) [Opcional]: "
    ).strip()
    unidad_str = f" ({unidad})" if unidad else ""

    print("\nIngrese los datos numéricos separados por comas o espacios:")
    entrada = input("> ")

    datos_raw = entrada.replace(",", " ").split()
    datos = []

    for item in datos_raw:
        try:
            datos.append(float(item))
        except ValueError:
            print(
                f"⚠️ Advertencia: '{item}' no es un número válido y se omitirá."
            )

    return magnitud, unidad_str, datos


def exportar_a_excel(datos, resultados, magnitud, unidad_str):
    nombre_archivo = "reporte_calculo_errores.xlsx"

    df_datos = pd.DataFrame(
        {f"Muestras de {magnitud}{unidad_str}": datos}
    )

    df_resumen = pd.DataFrame(
        {
            "Métrica": [
                "Cantidad de Datos (N)",
                f"Promedio{unidad_str}",
                f"Desviación Estándar Muestral{unidad_str}",
                f"Error Absoluto (Media){unidad_str}",
                "Error Relativo",
                "Error Porcentual (%)",
            ],
            "Valor": [
                len(datos),
                f"{resultados['promedio']:.4f}",
                f"{resultados['desviacion_estandar']:.4f}",
                f"{resultados['error_absoluto']:.4f}",
                f"{resultados['error_relativo']:.6f}",
                f"{resultados['error_porcentual']:.2f}%",
            ],
        }
    )

    with pd.ExcelWriter(nombre_archivo, engine="openpyxl") as writer:
        df_datos.to_excel(writer, sheet_name="Datos", index=True)
        df_resumen.to_excel(writer, sheet_name="Resultados", index=False)

    print(f"\n📂 Archivo Excel generado con éxito: '{nombre_archivo}'")


def main():
    magnitud, unidad_str, datos = solicitar_datos()

    if not datos:
        print("\n❌ No se ingresaron datos válidos para procesar.")
        return

    res = calcular_estadisticas(datos)
    total_datos = len(datos)

    print("\n" + "=" * 50)
    print(f"   RESULTADOS: {magnitud.upper()}")
    print("=" * 50)
    print(f"Número de datos (N):       {total_datos}")
    print(
        f"Promedio:                  {res['promedio']:.4f}{unidad_str}"
    )
    print(
        f"Desviación Estándar:       {res['desviacion_estandar']:.4f}{unidad_str}"
    )
    print(
        f"Error Absoluto (Media):    {res['error_absoluto']:.4f}{unidad_str}"
    )
    print(f"Error Relativo:            {res['error_relativo']:.6f}")
    print(f"Error Porcentual:          {res['error_porcentual']:.2f}%")
    print("=" * 50)

    if total_datos > 10:
        print(f"\n💡 Se ingresaron {total_datos} datos (más de 10).")
        opcion = (
            input("¿Desea exportar los resultados a Excel? (s/n): ")
            .strip()
            .lower()
        )
        if opcion in ["s", "si", "sí", "y"]:
            try:
                exportar_a_excel(datos, res, magnitud, unidad_str)
            except Exception as e:
                print(f"❌ Error al exportar: {e}")
                print(
                    "Asegúrate de tener instalada la librería openpyxl (`pip install openpyxl`)."
                )


if __name__ == "__main__":
    main()
