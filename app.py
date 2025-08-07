from flask import Flask, request, send_file, render_template
from procesador import procesar_excel
import os

app = Flask(__name__)

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/procesar", methods=["POST"])
def procesar():
    if 'archivo' not in request.files:
        return "No se envió ningún archivo", 400

    archivo = request.files['archivo']
    if archivo.filename == '':
        return "Nombre de archivo vacío", 400

    ruta_entrada = "entrada.xlsx"
    ruta_salida = "salida.xlsx"
    archivo.save(ruta_entrada)

    procesar_excel(ruta_entrada, ruta_salida)
    return send_file(ruta_salida, as_attachment=True)

if __name__ == "__main__":
    app.run(debug=True)
