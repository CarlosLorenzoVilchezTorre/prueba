from flask import Flask, request, redirect, url_for, send_file, render_template_string
import os
from werkzeug.utils import secure_filename
from reportes import consolidar_csv_en_excel

UPLOAD_FOLDER = 'uploads'
OUTPUT_FILE = 'reporte_general.xlsx'

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

HTML = """
<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="UTF-8">
  <title>Consolidador de Reportes CSV</title>
  <script src="https://cdn.tailwindcss.com"></script>
</head>
<body class="bg-gray-100 min-h-screen flex items-center justify-center">
  <div class="bg-white shadow-lg rounded-lg p-8 max-w-xl w-full text-center">
    <h1 class="text-3xl font-bold mb-6 text-blue-600">Subir Archivos CSV</h1>
    <form method="post" enctype="multipart/form-data" class="space-y-4">
      <input type="file" name="files" multiple required
        class="block w-full text-sm text-gray-500
        file:mr-4 file:py-2 file:px-4
        file:rounded-full file:border-0
        file:text-sm file:font-semibold
        file:bg-blue-50 file:text-blue-700
        hover:file:bg-blue-100">
      <button type="submit"
        class="bg-blue-600 hover:bg-blue-700 text-white font-bold py-2 px-6 rounded-full">
        Procesar Archivos
      </button>
    </form>
    {% if link %}
      <div class="mt-6">
        <a href="{{ link }}"
          class="text-green-600 font-semibold hover:underline">
          📥 Descargar reporte generado
        </a>
      </div>
    {% endif %}
  </div>
</body>
</html>
"""

@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        # Eliminar archivos antiguos
        for f in os.listdir(UPLOAD_FOLDER):
            os.remove(os.path.join(UPLOAD_FOLDER, f))
        
        # Guardar archivos nuevos
        files = request.files.getlist('files')
        for file in files:
            if file.filename.endswith('.csv'):
                filename = secure_filename(file.filename)
                file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))

        # Procesar y generar Excel
        consolidar_csv_en_excel(app.config['UPLOAD_FOLDER'], OUTPUT_FILE)

        return render_template_string(HTML, link=url_for('download_file'))
    return render_template_string(HTML)

@app.route('/download')
def download_file():
    return send_file(OUTPUT_FILE, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
