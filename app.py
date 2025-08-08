from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
from procesador import procesar_excel
from io import BytesIO

app = Flask(__name__)
CORS(app)  # Permite CORS para que el frontend pueda comunicarse

@app.route("/")
def home():
    return "🟢 Backend Reloj Control activo"

@app.route("/procesar", methods=["POST"])
def procesar_archivo():
    if 'file' not in request.files:
        return jsonify({"error": "No se envió ningún archivo"}), 400

    file = request.files['file']

    if file.filename == '':
        return jsonify({"error": "El nombre del archivo está vacío"}), 400

    if not file.filename.endswith(".xlsx"):
        return jsonify({"error": "Solo se aceptan archivos .xlsx"}), 400

    try:
        output = procesar_excel(file.stream)
        return send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name="resultado.xlsx"
        )
    except Exception as e:
        return jsonify({"error": f"Error al procesar archivo: {str(e)}"}), 500

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
