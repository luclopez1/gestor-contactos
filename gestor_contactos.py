# ============================================================
#  GESTOR DE CONTACTOS WEB - VERSION RAILWAY
#  Funciona en internet sin necesidad de guardar archivos
#
#  REQUIERE:
#  pip install flask openpyxl gunicorn
# ============================================================

from flask import Flask, request, redirect, url_for, send_file
import openpyxl
import io

app = Flask(__name__)

# Almacenamiento en memoria
contactos_db = []

def generar_excel():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Contactos"
    ws.append(["nombre", "empresa", "email"])
    for c in contactos_db:
        ws.append([c["nombre"], c["empresa"], c["email"]])
    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer

@app.route("/")
def index():
    total = len(contactos_db)
    filas = ""
    for i, c in enumerate(contactos_db):
        filas += f"""
        <tr>
            <td>{i + 1}</td>
            <td>{c['nombre']}</td>
            <td>{c['empresa'] or '-'}</td>
            <td>{c['email']}</td>
            <td>
                <a href="/eliminar/{i}"
                   onclick="return confirm('Eliminar?')"
                   style="color:#e74c3c;text-decoration:none;font-weight:bold">
                   X Eliminar
                </a>
            </td>
        </tr>
        """

    html = f"""
    <!DOCTYPE html>
    <html lang="es">
    <head>
        <meta charset="UTF-8">
        <title>Gestor de Contactos</title>
        <style>
            * {{ margin: 0; padding: 0; box-sizing: border-box; }}
            body {{ font-family: Arial, sans-serif; background: #f5f7fa; color: #333; }}
            header {{ background: #2F75B6; color: white; padding: 20px 40px; }}
            header h1 {{ font-size: 22px; }}
            .container {{ max-width: 900px; margin: 40px auto; padding: 0 20px; }}
            .card {{ background: white; border-radius: 10px; padding: 30px; margin-bottom: 30px; box-shadow: 0 2px 8px rgba(0,0,0,0.08); }}
            .card h2 {{ font-size: 18px; color: #2F75B6; margin-bottom: 20px; padding-bottom: 10px; border-bottom: 2px solid #e8f0fe; }}
            .form-row {{ display: flex; gap: 16px; flex-wrap: wrap; }}
            .form-group {{ flex: 1; min-width: 200px; display: flex; flex-direction: column; gap: 6px; }}
            label {{ font-size: 13px; font-weight: bold; color: #555; }}
            input {{ padding: 10px 14px; border: 1px solid #ddd; border-radius: 6px; font-size: 14px; outline: none; }}
            input:focus {{ border-color: #2F75B6; }}
            button {{ margin-top: 20px; background: #2F75B6; color: white; border: none; padding: 12px 30px; border-radius: 6px; font-size: 15px; cursor: pointer; font-weight: bold; }}
            .badge {{ background: #e8f0fe; color: #2F75B6; padding: 4px 12px; border-radius: 20px; font-size: 13px; font-weight: bold; margin-left: 10px; }}
            table {{ width: 100%; border-collapse: collapse; font-size: 14px; }}
            th {{ background: #2F75B6; color: white; padding: 12px 16px; text-align: left; }}
            td {{ padding: 11px 16px; border-bottom: 1px solid #f0f0f0; }}
            tr:nth-child(even) td {{ background: #f9fbff; }}
            .empty {{ text-align: center; color: #aaa; padding: 40px; font-size: 15px; }}
            .download-btn {{ background: #27ae60; color: white; padding: 10px 20px; border-radius: 6px; text-decoration: none; font-size: 14px; font-weight: bold; display: inline-block; }}
            .aviso {{ background: #fff8e1; border: 1px solid #f0c040; color: #7a5c00; padding: 10px 16px; border-radius: 6px; margin-bottom: 16px; font-size: 13px; }}
        </style>
    </head>
    <body>
        <header><h1>Gestor de Contactos</h1></header>
        <div class="container">
            <div class="card">
                <h2>Añadir nuevo contacto</h2>
                <p class="aviso">Los contactos se guardan temporalmente. Descarga el Excel antes de cerrar la sesion.</p>
                <form method="POST" action="/aniadir">
                    <div class="form-row">
                        <div class="form-group">
                            <label>Nombre *</label>
                            <input type="text" name="nombre" placeholder="Ej: Carlos Garcia" required>
                        </div>
                        <div class="form-group">
                            <label>Empresa</label>
                            <input type="text" name="empresa" placeholder="Ej: Tech Solutions SL">
                        </div>
                        <div class="form-group">
                            <label>Email *</label>
                            <input type="email" name="email" placeholder="Ej: carlos@empresa.com" required>
                        </div>
                    </div>
                    <button type="submit">+ Añadir contacto</button>
                </form>
            </div>
            <div class="card">
                <h2>
                    Contactos guardados
                    <span class="badge">{total}</span>
                    {"&nbsp;&nbsp;<a href='/descargar' class='download-btn'>Descargar Excel</a>" if total > 0 else ""}
                </h2>
                {"<table><thead><tr><th>#</th><th>Nombre</th><th>Empresa</th><th>Email</th><th>Accion</th></tr></thead><tbody>" + filas + "</tbody></table>" if total > 0 else '<p class="empty">No hay contactos todavia.</p>'}
            </div>
        </div>
    </body>
    </html>
    """
    return html

@app.route("/aniadir", methods=["POST"])
def aniadir():
    nombre = request.form.get("nombre", "").strip()
    empresa = request.form.get("empresa", "").strip()
    email = request.form.get("email", "").strip()
    if nombre and email:
        contactos_db.append({"nombre": nombre, "empresa": empresa, "email": email})
    return redirect(url_for("index"))

@app.route("/eliminar/<int:indice>")
def eliminar(indice):
    if 0 <= indice < len(contactos_db):
        contactos_db.pop(indice)
    return redirect(url_for("index"))

@app.route("/descargar")
def descargar():
    buffer = generar_excel()
    return send_file(
        buffer,
        as_attachment=True,
        download_name="contactos.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

if __name__ == "__main__":
    print("\nAbre tu navegador en: http://localhost:5000\n")
    app.run(debug=False)
