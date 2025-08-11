from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
from io import BytesIO
from procesador import procesar_excel, detectar_html_y_procesar

app = Flask(__name__)
# CORS abierto (si quieres, limita con origins=["https://tudominio"])
CORS(app)


@app.route("/", methods=["GET"])
def health():
    """Ruta de salud para que el frontend despierte/verifique el servidor."""
    return "🟢 Backend Reloj Control activo", 200


@app.route("/procesar", methods=["POST"])
def procesar_archivo():
    if "archivo" not in request.files:
        return jsonify({"error": "No se envió ningún archivo"}), 400

    file = request.files["archivo"]
    if file.filename.strip() == "":
        return jsonify({"error": "El nombre del archivo está vacío"}), 400

    # Leemos en memoria (permite inspección + reuso)
    contenido = file.read()
    if not contenido:
        return jsonify({"error": "El archivo está vacío"}), 400

    try:
        # Detecta si es XLS-HTML (muchas plataformas exportan HTML con extensión .xls)
        cab = contenido[:200].lstrip().lower()
        es_html = cab.startswith(b"<html") or b"<table" in cab

        if es_html:
            output = detectar_html_y_procesar(contenido)
        else:
            # .xls binario real o .xlsx
            output = procesar_excel(BytesIO(contenido))

        return send_file(
            output,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            as_attachment=True,
            download_name="resultado.xlsx",
        )

    except Exception as e:
        msg = str(e)
        # Mensajes más claros para el caso XLS-HTML o formato inválido
        pistas = ("Unsupported format", "BOF", "xlrd", "read_html", "No tables found")
        if any(p in msg for p in pistas):
            return (
                jsonify(
                    {
                        "error": (
                            "No se pudo interpretar el archivo como Excel. "
                            "Si tu sistema exporta HTML con extensión .xls, se intentó leerlo como HTML "
                            "pero falló. Detalle: " + msg
                        )
                    }
                ),
                400,
            )
        return jsonify({"error": f"Error al procesar archivo: {msg}"}), 500


if __name__ == "__main__":
    # Desarrollo local (en Render se usa gunicorn vía Procfile)
    app.run(host="0.0.0.0", port=5000, debug=False)
