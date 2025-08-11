from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
from procesador import procesar_excel, detectar_html_y_procesar
from io import BytesIO

app = Flask(__name__)

# CORS explícito y robusto (para preflight de FormData)
CORS(
    app,
    resources={r"/*": {"origins": "*"}},
    methods=["GET", "POST", "OPTIONS"],
    allow_headers=["Content-Type"],
    expose_headers=["Content-Type"],
    supports_credentials=False,
)

@app.after_request
def add_cors_headers(resp):
    # Refuerzo por si algún proxy elimina cabeceras
    resp.headers["Access-Control-Allow-Origin"] = "*"
    resp.headers["Access-Control-Allow-Methods"] = "GET, POST, OPTIONS"
    resp.headers["Access-Control-Allow-Headers"] = "Content-Type"
    return resp


@app.route("/", methods=["GET"])
def home():
    return "🟢 Backend Reloj Control activo"


@app.route("/procesar", methods=["POST", "OPTIONS"])
def procesar_archivo():
    # Preflight
    if request.method == "OPTIONS":
        return ("", 204)

    if 'archivo' not in request.files:
        return jsonify({"error": "No se envió ningún archivo"}), 400

    file = request.files['archivo']
    if file.filename == '':
        return jsonify({"error": "El nombre del archivo está vacío"}), 400

    contenido = file.read()
    if not contenido:
        return jsonify({"error": "El archivo está vacío"}), 400

    try:
        # ¿Es un .xls “HTML”?
        cab = contenido[:200].lstrip().lower()
        es_html = cab.startswith(b"<html") or b"<table" in cab

        if es_html:
            # Parseo HTML + cálculo
            output = detectar_html_y_procesar(contenido)
        else:
            # Excel real (e.g., .xlsx)
            output = procesar_excel(BytesIO(contenido))

        return send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name='resultado.xlsx'
        )

    except Exception as e:
        msg = str(e)
        if "Unsupported format" in msg or "BOF" in msg or "xlrd" in msg:
            return jsonify({
                "error": (
                    "El archivo no parece ser un Excel 97-2003 binario (.xls). "
                    "Se intentó tratar como HTML y falló. "
                    f"Detalle: {msg}"
                )
            }), 400
        return jsonify({"error": f"Error al procesar archivo: {msg}"}), 500


if __name__ == "__main__":
    # Render usa gunicorn via Procfile; esto es para pruebas locales
    app.run(host="0.0.0.0", port=5000, debug=False)
