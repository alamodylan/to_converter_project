from flask import Flask, render_template_string, request, send_file, redirect, url_for
import pandas as pd
import io
import datetime
import os
import re

app = Flask(__name__)

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
    "Estadía en Chasis 3 Ejes"
]

TARIFAS_DIARIAS = {
    "Demora de Chasis Sencillo": 45,
    "Demora de Chasis Equipo Especial": 75,
    "Estadía en Chasis 3 Ejes": 75,
    "Gen Set Diario": 40
}

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
                    "DIRECCIÓN DE COLOCACIÓN": obtener_direccion_colocacion(x),
                    "ENTREGA DE VACIO": x.iloc[0]["Ubicación Final"],
                    "PATIO DE RETIRO $": obtener_monto(x, servicio="RETIRA VACIO EXPORT"),
                    "COSTO FLETE $": obtener_monto(x, tipo="Guía", exclude_servicio="RETIRA VACIO EXPORT"),
                    "3 EJES $": obtener_monto(x, tipo="Cargo Adicional Guía", servicio="Sobre Peso 3 ejes"),
                    "RETORNO $": obtener_monto(x, tipo="Cargo Adicional Guía", servicio_prefix="SJO-RT"),
                    "EXTRA COSTOS $": obtener_extra_costos(x),
                    "MONTO TOTAL $": 0,
                    "FECHA DE COLOCACIÓN": pd.to_datetime(x[x["Tipo"] == "Guía"].iloc[0]["Fecha y Hora Llegada"], dayfirst=True).date(),
                    "COMENTARIOS TTA": obtener_comentarios_tta(x)
                })).reset_index()

                resumen["MONTO TOTAL $"] = resumen[[
                    "COSTO FLETE $", "PATIO DE RETIRO $", "3 EJES $", "RETORNO $", "EXTRA COSTOS $"
                ]].sum(axis=1)

                resumen.insert(0, "CLIENTE", "")
                resumen.insert(2, "TAMAÑO", "")
                resumen.insert(3, "FECHA DE COLOCACIÓN", resumen.pop("FECHA DE COLOCACIÓN"))
                resumen["COMENTARIOS MSC"] = ""

                cols_finales = [
                    "CLIENTE", "BL o BOOKING", "TAMAÑO", "FECHA DE COLOCACIÓN", "Contenedor", "PUERTO DE SALIDA",
                    "DIRECCIÓN DE COLOCACIÓN", "ENTREGA DE VACIO", "COSTO FLETE $",
                    "3 EJES $", "PATIO DE RETIRO $", "RETORNO $", "EXTRA COSTOS $", "MONTO TOTAL $",
                    "COMENTARIOS TTA", "COMENTARIOS MSC"
                ]

                resumen = resumen[cols_finales]
                resumen = resumen.sort_values(by="FECHA DE COLOCACIÓN", ascending=True)
                timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                nombre_archivo = f"outputs/{grupo.replace(' ', '_')}_{timestamp}.xlsx"
                with pd.ExcelWriter(nombre_archivo, engine='xlsxwriter') as writer:
                    resumen.to_excel(writer, index=False, sheet_name='TO')
                    workbook = writer.book
                    worksheet = writer.sheets['TO']

                    header_format = workbook.add_format({
                        'bold': True,
                        'text_wrap': True,
                        'valign': 'center',
                        'align': 'center',
                        'fg_color': '#000000',
                        'font_color': '#FFFFFF',
                        'border': 1
                    })

                    for col_num, value in enumerate(resumen.columns.values):
                        worksheet.write(0, col_num, value, header_format)

                    for i, col in enumerate(resumen.columns):
                        max_len = max(resumen[col].astype(str).map(len).max(), len(col)) + 2
                        worksheet.set_column(i, i, max_len)

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

def obtener_monto(df, tipo=None, servicio=None, servicio_prefix=None, exclude_servicio=None):
    f = df.copy()
    if exclude_servicio:
        f = f[f["Tipo Servicio"] != exclude_servicio]
    if tipo:
        f = f[f["Tipo"] == tipo]
    if servicio:
        f = f[f["Tipo Servicio"] == servicio]
    if servicio_prefix:
        f = f[f["Tipo Servicio"].astype(str).str.startswith(servicio_prefix)]
    return f["Monto"].astype(float).sum()

def obtener_direccion_colocacion(df):
    notas_guia = df[df["Tipo"] == "Guía"]["Notas"].astype(str)

    for nota in notas_guia:
        match = re.search(r'Descarga(.*?)\*', nota, re.IGNORECASE)
        if match:
            return match.group(1).strip()
    return ""

def obtener_patio_retiro(df):
    f = df[(df["Tipo"] == "Guía") & (df["Tipo Servicio"] == "RETIRA VACIO EXPORT")]
    return f["Monto"].astype(float).sum()

def obtener_extra_costos(df):
    f = df[df["Tipo Servicio"].isin(EXTRA_COSTOS_LIST)]
    return f["Monto"].astype(float).sum()

def obtener_comentarios_tta(df):
    comentarios = []

    if obtener_monto(df, tipo="Guía", exclude_servicio="RETIRA VACIO EXPORT") > 0:
        comentarios.append("Flete")
    if obtener_patio_retiro(df) > 0:
        comentarios.append("Patio de Retiro")
    if obtener_monto(df, tipo="Cargo Adicional Guía", servicio="Sobre Peso 3 ejes") > 0:
        comentarios.append("3 Ejes")
    if obtener_monto(df, tipo="Cargo Adicional Guía", servicio_prefix="SJO-RT") > 0:
        comentarios.append("Retorno")

    adicionales = df[df["Tipo Servicio"].isin(TARIFAS_DIARIAS.keys())]
    for _, row in adicionales.iterrows():
        servicio = row["Tipo Servicio"]
        monto = float(row["Monto"])
        dias = int(round(monto / TARIFAS_DIARIAS[servicio]))
        comentarios.append(f"{servicio} ({dias} días)")

    # ⬇️ Agregamos IMO si aplica
    imo_monto = df[
    (df["Tipo Servicio"].astype(str).str.strip().str.lower() == "choferes quimiquero")
]["Monto"].astype(float).sum()
    if imo_monto > 0:
        comentarios.append("IMO")

    return " | ".join(comentarios)

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 10000))
    app.run(host='0.0.0.0', port=port)
