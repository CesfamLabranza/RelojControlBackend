from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
from procesador import procesar_excel
from io import BytesIO

app = Flask(__name__)
CORS(app)  # permite que el frontend (tu página) pueda llamar al backend

@app.route("/")
def home():
    return "🟢 Backend Reloj Control activo"

@app.route("/procesar", methods=["POST"])
def procesar_archivo():
    # Verifica que venga el archivo en el campo 'archivo'
    if "archivo" not in request.files:
        return jsonify({"error": "No se envió ningún archivo"}), 400

    file = request.files["archivo"]

    if file.filename == "":
        return jsonify({"error": "El nombre del archivo está vacío"}), 400

    # Sólo aceptamos .xls (Excel 97-2003 binario)
    if not file.filename.lower().endswith(".xls"):
        return jsonify({"error": "Sólo se aceptan archivos .xls"}), 400

    try:
        # Leemos todos los bytes del archivo
        data = file.read()
        if not data:
            return jsonify({"error": "El archivo llegó vacío"}), 400

        # Detección de HTML “disfrazado” de .xls (muchos sistemas exportan así)
        head = data[:32].lstrip()
        # Si empieza con <html, <!doctype html, etc., es HTML y no un .xls binario
        if head.startswith(b"<") or head.lower().startswith(b"<!doctype html") or b"<html" in head.lower():
            return jsonify({
                "error": (
                    "El archivo no es un Excel 97-2003 real (.xls). "
                    "Parece ser un HTML exportado con extensión .xls. "
                    "Por favor exporta como Excel 97-2003 auténtico (archivo binario .xls) "
                    "o envíame un ejemplo para adaptar el parser."
                )
            }), 400

        # Procesamos con la función principal (usa openpyxl/xlrd según tu implementación)
        resultado = procesar_excel(BytesIO(data))

        # Devolvemos el Excel resultante (xlsx)
        return send_file(
            resultado,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            as_attachment=True,
            download_name="resultado.xlsx",
        )

    except Exception as e:
        # Cualquier otra excepción se reporta con detalle
        return jsonify({"error": f"Error al procesar archivo: {str(e)}"}), 500


if __name__ == "__main__":
    # Para despliegue local; en producción (Render) se usa gunicorn con Procfile
    app.run(host="0.0.0.0", port=5000)
