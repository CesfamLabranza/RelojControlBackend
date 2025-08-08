from io import StringIO
import sys

@app.route("/procesar", methods=["POST"])
def procesar():
    if 'archivo' not in request.files:
        return jsonify({"error": "No se envió ningún archivo"}), 400

    archivo = request.files['archivo']
    if archivo.filename == '':
        return jsonify({"error": "Nombre de archivo vacío"}), 400

    ruta_entrada = "entrada.xlsx" if archivo.filename.endswith(".xlsx") else "entrada.xls"
    ruta_salida = "salida.xlsx"
    archivo.save(ruta_entrada)

    # Capturar logs en un buffer
    buffer = StringIO()
    sys.stdout = buffer

    try:
        procesar_excel(ruta_entrada, ruta_salida)
        sys.stdout = sys.__stdout__  # Restaurar salida
        return send_file(ruta_salida, as_attachment=True)
    except Exception as e:
        sys.stdout = sys.__stdout__
        log = buffer.getvalue()
        error_msg = f"❌ Error al procesar: {str(e)}\n\n{log}"
        return jsonify({"error": error_msg}), 500
