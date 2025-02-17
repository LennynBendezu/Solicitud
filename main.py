from flask import Flask, request, jsonify, send_file
import openpyxl
import os

app = Flask(__name__)

# Ruta del archivo Excel
EXCEL_FILE = "solicitudes.xlsx"

# Crear archivo Excel si no existe
if not os.path.exists(EXCEL_FILE):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Solicitudes"
    ws.append(["Nombres", "Apellidos", "Correo Electrónico", "Teléfono", "Servicio", "Mensaje"])
    wb.save(EXCEL_FILE)

@app.route('/webhook', methods=['POST'])
def dialogflow_webhook():
    req = request.get_json()

    # Manejar diferentes formatos de JSON
    params = req.get('sessionInfo', {}).get('parameters', req)  # Usa 'req' directo si no viene desde Dialogflow

    nombres = params.get("nombres", "")
    apellidos = params.get("apellidos", "")
    correo = params.get("correoelectronico", "")
    telefono = params.get("telefono", "")
    servicio = params.get("servicio", "")
    mensaje = params.get("mensaje", "")

    # Agregar la solicitud al archivo Excel
    try:
        wb = openpyxl.load_workbook(EXCEL_FILE)
        ws = wb.active
        ws.append([nombres, apellidos, correo, telefono, servicio, mensaje])
        wb.save(EXCEL_FILE)
    except Exception as e:
        return jsonify({"error": f"No se pudo escribir en el Excel: {str(e)}"}), 500

    # Respuesta a Dialogflow CX
    response = {
        "fulfillment_response": {
            "messages": [
                {
                    "text": {
                        "text": ["Tu solicitud ha sido guardada correctamente en el archivo Excel."]
                    }
                }
            ]
        }
    }
    return jsonify(response)

# Ruta para descargar el archivo Excel
@app.route('/descargar_excel', methods=['GET'])
def descargar_excel():
    return send_file(EXCEL_FILE, as_attachment=True)

if __name__ == '__main__':
    app.run(host="0.0.0.0", port=5000, debug=True)
