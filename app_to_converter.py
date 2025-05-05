# app_to_converter.py - Proyecto completo con ajuste de encabezado y columna de FECHA DE COLOCACIÓN

from flask import Flask, render_template_string, request, send_file, redirect, url_for
import pandas as pd
import io
import datetime
import os

app = Flask(__name__)

# HTML embebido con Bootstrap
HTML = """
<!DOCTYPE html>
<html lang=\"es\">
<head>
    <meta charset=\"UTF-8\">
    <meta name=\"viewport\" content=\"width=device-width, initial-scale=1.0\">
    <title>Conversor TO MSC</title>
    <link href=\"https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css\" rel=\"stylesheet\">
</head>
<body>
<div class=\"container mt-5\">
    <h2 class=\"mb-4\">Conversor de Reporte a Formato TO - MSC</h2>
    <form method=\"POST\" enctype=\"multipart/form-data\">
        <div class=\"mb-3\">
            <label for=\"file\" class=\"form-label\">Subí tu archivo Excel del sistema:</label>
            <input type=\"file\" class=\"form-control\" name=\"file\" required>
        </div>
        <button type=\"submit\" class=\"btn btn-primary\">Procesar</button>
    </form>
    {% if outputs %}
        <div class=\"mt-5\">
            <h4>Archivos generados:</h4>
            <ul>
                {% for label, link in outputs.items() %}
                    <li><a href=\"{{ link }}\" class=\"btn btn-success my-1\">Descargar {{ label }}</a></li>
                {% endfor %}
            </ul>
        </div>
    {% endif %}
</div>
</body>
</html>
"""

if not os.path.exists("outputs"):
    os.makedirs("outputs")

EXPORT_SERVICIOS = [
    "Recoge contenedor para export",
    "Carrusel export",
    "Movimiento de exportacion",
    "Retira full export"
]
COYOL_SERVICIOS = ["Movimiento de export", "Retira full export"]
EXTRA_COSTOS_LIST = [
    "Demora de Chasis Sencillo",
    "Diesel Adicional",
    "Gen Set Diario",
    "Demora de Chasis Equipo Especial",
    "Choferes Quimiquero",
    "Diesel para viaje",
    "Estadía en Chasis 3 Ejes"  # ← nuevo agregado
]

def clasificar_to(fila):
    ruta = str(fila.get("Ruta", "")).strip()
    tipo_serv = str(fila.get("Tipo Servicio", "")).strip()

    if ruta.startswith("CAL") and tipo_serv in EXPORT_SERVICIOS:
        return "TO Exportación Caldera"
    elif ruta.startswith("SJO") and tipo_serv in COYOL_SERVICIOS:
        return "TO Exportación Coyol"
    elif ruta.startswith("Lio") and tipo_serv in EXPORT_SERVICIOS:
        return "TO Exportación Limón"
    elif ruta.startswith("CAL") and tipo_serv not in EXPORT_SERVICIOS:
        return "TO Importación Caldera"
    elif ruta.startswith("Lio") and tipo_serv not in EXPORT_SERVICIOS:
        return "TO Importación Limón"
    else:
        return None

@app.route('/', methods=['GET', 'POST'])
def index():
    outputs = {}
    if request.method == 'POST':
        file = request.files['file']
        if not file:
            return redirect(url_for('index'))

        df = pd.read_excel(file, header=1)
        print("Columnas cargadas:", df.columns.tolist())

        df['CUADRO TO'] = df.apply(clasificar_to, axis=1)

        for grupo in df['CUADRO TO'].dropna().unique():
            df_grupo = df[df['CUADRO TO'] == grupo]
            if not df_grupo.empty:
                resumen = df_grupo.groupby("Contenedor").apply(lambda x: pd.Series({
                    "BL o BOOKING": extraer_booking(x),
                    "PUERTO DE SALIDA": x.iloc[0]["Origen"],
                    "DIRECCIÓN DE COLOCACIÓN": x.iloc[0]["Ubicación Final"],
                    "ENTREGA DE VACIO": obtener_entrega_vacio(x),
                    "COSTO FLETE $": obtener_monto(x, tipo="Guía"),
                    "PATIO DE RETIRO $": obtener_patio_retiro(x),
                    "3 EJES $": obtener_monto(x, tipo="Cargo Adicional Guía", servicio="Sobre Peso 3 ejes"),
                    "RETORNO $": obtener_monto(x, tipo="Cargo Adicional Guía", servicio_prefix="SJO-RT"),
                    "EXTRA COSTOS $": obtener_extra_costos(x),
                    "MONTO TOTAL $": 0,
                    "FECHA DE COLOCACIÓN": pd.to_datetime(x.iloc[0]["Fecha y Hora Llegada"]).date()
                })).reset_index()

                resumen["MONTO TOTAL $"] = resumen[[
                    "COSTO FLETE $", "PATIO DE RETIRO $", "3 EJES $", "RETORNO $", "EXTRA COSTOS $"
                ]].sum(axis=1)

                resumen.insert(0, "CLIENTE", "")
                resumen.insert(2, "TAMAÑO", "")
                resumen.insert(3, "FECHA DE COLOCACIÓN", resumen.pop("FECHA DE COLOCACIÓN"))
                resumen["COMENTARIOS TTA"] = ""
                resumen["COMENTARIOS MSC"] = ""

                cols_finales = [
                    "CLIENTE", "BL o BOOKING", "TAMAÑO", "FECHA DE COLOCACIÓN", "Contenedor", "PUERTO DE SALIDA",
                    "DIRECCIÓN DE COLOCACIÓN", "ENTREGA DE VACIO", "COSTO FLETE $",
                    "3 EJES $", "PATIO DE RETIRO $", "RETORNO $", "EXTRA COSTOS $", "MONTO TOTAL $",
                    "COMENTARIOS TTA", "COMENTARIOS MSC"
                ]

                resumen = resumen[cols_finales]
                timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                nombre_archivo = f"outputs/{grupo.replace(' ', '_')}_{timestamp}.xlsx"
                with pd.ExcelWriter(nombre_archivo, engine='xlsxwriter') as writer:
                resumen.to_excel(writer, index=False, sheet_name='TO')
                workbook = writer.book
                worksheet = writer.sheets['TO']

                # Formato del encabezado
                format_header = workbook.add_format({
                    'bold': True,
                    'bg_color': '#000000',
                    'font_color': 'white',
                    'align': 'center',
                    'valign': 'vcenter',
                    'border': 1
                })

                # Formato celdas normales con borde
                format_border = workbook.add_format({'border': 1})

                # Formato dinero
                format_money = workbook.add_format({'border': 1, 'num_format': '$#,##0.00'})

                # Aplicar formato al encabezado
                worksheet.set_row(0, 20, format_header)

                # Ajuste de columnas
                for col_num, col_name in enumerate(resumen.columns):
                    ancho = max(resumen[col_name].astype(str).map(len).max(), len(str(col_name))) + 2
                    if "$" in col_name:
                        worksheet.set_column(col_num, col_num, ancho, format_money)
                    else:
                        worksheet.set_column(col_num, col_num, ancho, format_border) #Hasta aqui
                outputs[grupo] = url_for('download_file', filename=os.path.basename(nombre_archivo))

    return render_template_string(HTML, outputs=outputs)

@app.route('/download/<filename>')
def download_file(filename):
    path = os.path.join("outputs", filename)
    return send_file(path, as_attachment=True)

def extraer_booking(df):
    notas = df["Notas"].astype(str).tolist()
    for nota in notas:
        for palabra in nota.split():
            if any(char.isdigit() for char in palabra):
                return palabra
    return ""

def obtener_entrega_vacio(df):
    vacio = df[(df["Tipo"] == "Guía") & (df["Tipo Servicio"] == "Retira vacio export")]
    if not vacio.empty:
        return vacio.iloc[0]["Ubicación Final"]
    return ""

def obtener_monto(df, tipo, servicio=None, servicio_prefix=None):
    f = df[df["Tipo"] == tipo]
    if servicio:
        f = f[f["Tipo Servicio"] == servicio]
    if servicio_prefix:
        f = f[f["Tipo Servicio"].astype(str).str.startswith(servicio_prefix)]
    return f["Monto"].astype(float).sum()

def obtener_patio_retiro(df):
    f = df[(df["Tipo"] == "Guía") & (df["Tipo Servicio"] == "Retira vacio export")]
    return f["Monto"].astype(float).sum()

def obtener_extra_costos(df):
    f = df[df["Tipo Servicio"].isin(EXTRA_COSTOS_LIST)]
    return f["Monto"].astype(float).sum()

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 10000))
    app.run(host='0.0.0.0', port=port)
