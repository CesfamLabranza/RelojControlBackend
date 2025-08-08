from flask import Flask, request, send_file, render_template, jsonify
from procesador import procesar_excel
import os

app = Flask(__name__)

@app.route("/")
def index():
    return "Backend RelojControl OK"

@app.route("/procesar", methods=["POST"])
def procesar():
    if 'archivo' not in request.files:
        return "No se envió ningún archivo", 400

    archivo = request.files['archivo']
    if archivo.filename == '':
        return "Nombre de archivo vacío", 400

    # Guardar temporalmente como .xls o .xlsx
    extension = os.path.splitext(archivo.filename)[1]
    ruta_entrada = f"entrada{extension}"
    ruta_salida = "salida.xlsx"
    archivo.save(ruta_entrada)

    try:
        procesar_excel(ruta_entrada, ruta_salida)
        return send_file(ruta_salida, as_attachment=True)
    except Exception as e:
        return f"Error procesando el archivo: {str(e)}", 500

if __name__ == "__main__":
    app.run(debug=True)
