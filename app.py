from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
from io import BytesIO

from procesador import procesar_excel

app = Flask(__name__)
CORS(app)  # permitir CORS desde tu frontend

@app.route("/")
def home():
    return "🟢 Backend Reloj Control activo"

@app.route("/procesar", methods=["POST"])
def procesar_archivo():
    # ⚠️ el campo se llama 'archivo' (coincide con el frontend)
    if "archivo" not in request.files:
        return jsonify({"error": "No se envió ningún archivo"}), 400

    f = request.files["archivo"]
    if f.filename == "":
        return jsonify({"error": "El nombre del archivo está vacío"}), 400

    # Sólo aceptar .xls (no .xlsx)
    if not f.filename.lower().endswith(".xls"):
        return jsonify({"error": "Sólo se aceptan archivos .xls"}), 400

    try:
        contenido = f.read()
        salida = procesar_excel(BytesIO(contenido))

        return send_file(
            salida,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            as_attachment=True,
            download_name="resultado.xlsx",
        )
    except Exception as e:
        return jsonify({"error": f"Error al procesar archivo: {str(e)}"}), 500

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
