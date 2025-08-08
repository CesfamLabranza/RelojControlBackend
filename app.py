from flask import Flask, request, send_file, render_template, jsonify
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

    ruta_entrada = "entrada.xlsx" if archivo.filename.endswith(".xlsx") else "entrada.xls"
    ruta_salida = "salida.xlsx"
    archivo.save(ruta_entrada)

    try:
        procesar_excel(ruta_entrada, ruta_salida)
        return send_file(ruta_salida, as_attachment=True)
    except Exception as e:
        print(f"Error: {e}")
        return "Ocurrió un error al procesar el archivo.", 500

# Claves y archivos permitidos
CLAVES_VALIDAS = {
    "infolabranza": "archivos/labranza.xlsx",
    # Puedes agregar más claves y rutas aquí
}

@app.route("/descargar", methods=["POST"])
def descargar_archivo():
    data = request.get_json()
    clave = data.get("clave", "")

    if clave in CLAVES_VALIDAS:
        ruta_archivo = CLAVES_VALIDAS[clave]
        try:
            return send_file(ruta_archivo, as_attachment=True)
        except FileNotFoundError:
            return jsonify({"error": "Archivo no encontrado"}), 404
    else:
        return jsonify({"error": "Clave inválida"}), 403

if __name__ == "__main__":
    app.run(debug=True)
