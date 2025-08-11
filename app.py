from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
from procesador import procesar_excel
from io import BytesIO

app = Flask(__name__)
CORS(app)  # permitir CORS

@app.route("/")
def home():
    return "🟢 Backend Reloj Control activo"

@app.route("/procesar", methods=["POST"])
def procesar_archivo():
    if "archivo" not in request.files:
        return jsonify({"error": "No se envió ningún archivo"}), 400

    file = request.files["archivo"]

    if file.filename == "":
        return jsonify({"error": "El nombre del archivo está vacío"}), 400

    # ✅ Aceptamos .xls (lo que tú subirás)
    if not file.filename.lower().endswith(".xls"):
        return jsonify({"error": "Sólo se aceptan archivos .xls"}), 400

    try:
        # Pasamos el stream al procesador
        stream = BytesIO(file.read())
        output = procesar_excel(stream)

        return send_file(
            output,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            as_attachment=True,
            download_name="resultado.xlsx",
        )
    except Exception as e:
        return jsonify({"error": f"Error al procesar archivo: {str(e)}"}), 500

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
