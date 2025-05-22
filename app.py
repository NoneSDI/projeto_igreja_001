from flask import Flask, render_template, request, redirect, url_for, session, send_file
from werkzeug.security import generate_password_hash, check_password_hash
import sqlite3
from datetime import datetime
import calendar
import pandas as pd
import tempfile
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl import load_workbook
import locale
from datetime import datetime
from openpyxl.utils import get_column_letter
import tempfile
from flask import send_file
import locale
from datetime import datetime
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

app = Flask(__name__)
app.secret_key = 'chave_secreta'

def init_db():
    with sqlite3.connect("usuarios.db") as conn:
        c = conn.cursor()
        c.execute('''CREATE TABLE IF NOT EXISTS usuarios (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        cargo TEXT,
                        nome TEXT,
                        email TEXT UNIQUE,
                        senha TEXT,
                        telefone TEXT)''')
        c.execute('''CREATE TABLE IF NOT EXISTS opcoes_lista (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        opcao TEXT)''')
        conn.commit()

@app.route("/", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        email = request.form["email"]
        senha = request.form["senha"]
        with sqlite3.connect("usuarios.db") as conn:
            c = conn.cursor()
            c.execute("SELECT * FROM usuarios WHERE email = ?", (email,))
            user = c.fetchone()
            if user and check_password_hash(user[4], senha):
                session["user_id"] = user[0]
                session["nome"] = user[2]
                return redirect(url_for("painel"))
            else:
                return "Usuário ou senha inválidos"
    return render_template("login.html")

@app.route("/cadastro", methods=["GET", "POST"])
def cadastro():
    if request.method == "POST":
        cargo = request.form["cargo"]
        nome = request.form["nome"]
        email = request.form["email"]
        senha = request.form["senha"]
        confirmacao = request.form["confirmacao"]
        telefone = request.form.get("telefone", "")

        if senha != confirmacao:
            return "As senhas não coincidem!"

        senha_hash = generate_password_hash(senha)

        try:
            with sqlite3.connect("usuarios.db") as conn:
                c = conn.cursor()
                c.execute("INSERT INTO usuarios (cargo, nome, email, senha, telefone) VALUES (?, ?, ?, ?, ?)",
                          (cargo, nome, email, senha_hash, telefone))
                conn.commit()
            return redirect(url_for("login"))
        except sqlite3.IntegrityError:
            return "E-mail já cadastrado!"

    return render_template("cadastro.html")

import locale
from datetime import datetime
import calendar

@app.route("/painel", methods=["GET"])
def painel():
    if "user_id" not in session:
        return redirect(url_for("login"))

    # Configura o locale para português do Brasil, compatível com Windows e Linux
    import locale
    import platform
    try:
        if platform.system() == 'Windows':
            locale.setlocale(locale.LC_TIME, 'Portuguese_Brazil.1252')
        else:
            locale.setlocale(locale.LC_TIME, 'pt_BR.UTF-8')
    except locale.Error:
        print("Aviso: Locale português não pôde ser configurado. Datas podem aparecer em inglês.")

    hoje = datetime.now()
    dias_mes = calendar.monthrange(hoje.year, hoje.month)[1]

    # Lista com datas e dias da semana em português, corrigindo encoding
    dias_semana_pt = {
        "Monday": "Segunda-feira",
        "Tuesday": "Terça-feira",
        "Wednesday": "Quarta-feira",
        "Thursday": "Quinta-feira",
        "Friday": "Sexta-feira",
        "Saturday": "Sábado",
        "Sunday": "Domingo"
    }

    datas = []
    for d in range(1, dias_mes + 1):
        data_obj = datetime(hoje.year, hoje.month, d)
        dia_str = f"{hoje.year}-{hoje.month:02d}-{d:02d}"
        dia_semana_eng = data_obj.strftime('%A')
        dia_semana = dias_semana_pt.get(dia_semana_eng, dia_semana_eng)
        datas.append({
            "dia": dia_str,
            "dia_semana": dia_semana
        })

    with sqlite3.connect("usuarios.db") as conn:
        c = conn.cursor()
        c.execute("SELECT opcao FROM opcoes_lista")
        opcoes = [row[0] for row in c.fetchall()]

    meses_pt = {
        "January": "Janeiro",
        "February": "Fevereiro",
        "March": "Março",
        "April": "Abril",
        "May": "Maio",
        "June": "Junho",
        "July": "Julho",
        "August": "Agosto",
        "September": "Setembro",
        "October": "Outubro",
        "November": "Novembro",
        "December": "Dezembro"
    }
    nome_mes_ingles = hoje.strftime('%B')
    nome_mes = meses_pt.get(nome_mes_ingles, nome_mes_ingles)

    return render_template("painel.html", datas=datas, opcoes=opcoes, nome_mes=nome_mes)

@app.route("/add_opcao", methods=["POST"])
def add_opcao():
    nova_opcao = request.form["nova_opcao"]
    with sqlite3.connect("usuarios.db") as conn:
        c = conn.cursor()
        c.execute("INSERT INTO opcoes_lista (opcao) VALUES (?)", (nova_opcao,))
        conn.commit()
    return redirect(url_for("painel"))

from collections import defaultdict

@app.route("/gerar_excel", methods=["POST"])
def gerar_excel():
    dados = request.get_json()
    if not dados:
        return "Dados inválidos", 400

    hoje = datetime.now()

    meses_pt = {
        "January": "Janeiro",
        "February": "Fevereiro",
        "March": "Março",
        "April": "Abril",
        "May": "Maio",
        "June": "Junho",
        "July": "Julho",
        "August": "Agosto",
        "September": "Setembro",
        "October": "Outubro",
        "November": "Novembro",
        "December": "Dezembro"
    }
    nome_mes_ingles = hoje.strftime('%B')
    nome_mes = meses_pt.get(nome_mes_ingles, nome_mes_ingles)

    dias_semana_pt = {
        "Monday": "Segunda-feira",
        "Tuesday": "Terça-feira",
        "Wednesday": "Quarta-feira",
        "Thursday": "Quinta-feira",
        "Friday": "Sexta-feira",
        "Saturday": "Sábado",
        "Sunday": "Domingo"
    }

    # Agrupar atividades por dia (data_completa)
    agrupados = defaultdict(list)
    for item in dados:
        try:
            dt = datetime.strptime(item.get('dia', ''), '%Y-%m-%d')
            data_formatada = dt.strftime('%d-%m-%Y')
            dia_semana_ingles = dt.strftime('%A')
            dia_semana_pt = dias_semana_pt.get(dia_semana_ingles, dia_semana_ingles)
        except Exception:
            data_formatada = item.get('dia', '')
            dia_semana_pt = item.get('dia_semana', '')

        data_completa = f"{data_formatada}\n{dia_semana_pt}"

        atividade_raw = item.get('atividade', '')
        # Se for lista de atividades, juntar com quebras de linha e bolinhas
        if isinstance(atividade_raw, list):
            atividades_formatadas = [f"• {a}" for a in atividade_raw]
        else:
            atividades_formatadas = [f"• {atividade_raw}" if atividade_raw else '']

        hora_inicio = item.get('hora_inicio', '')
        hora_fim = item.get('hora_fim', '')
        horario = ''
        if hora_inicio or hora_fim:
            horario = f"{hora_inicio} às {hora_fim}".strip(" -")

        # Agora concatena: Atividade - Horário, se horário existir
        for atv in atividades_formatadas:
            if horario:
                agrupados[data_completa].append(f"    {atv}   (HORÁRIO:  {horario})")
            else:
                agrupados[data_completa].append(f"    {atv}")


    # Agora montar a lista final para o DataFrame
    dados_final = []
    for data_completa, lista_atividades in agrupados.items():
        
        # Junta todas as atividades do mesmo dia em uma string só, separadas por quebra de linha
        atividades_concatenadas = "\n".join(lista_atividades)
        dados_final.append({
            "Data": data_completa,
            "Atividades - Horário": atividades_concatenadas
        })

    df = pd.DataFrame(dados_final)[['Data', 'Atividades - Horário']]

    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        caminho = tmp.name

    df.to_excel(caminho, index=False)

    wb = load_workbook(caminho)
    ws = wb.active

    ws.insert_rows(1, amount=3)
    ws.insert_cols(1, amount=3)

    start_row = 4
    start_col = 4

    titulo = f"AGENDA MENSAL IGREJA NOSSA SRA. DAS GRAÇAS   -  Mês de {nome_mes}"
    ws.merge_cells(start_row=2, start_column=start_col, end_row=2, end_column=start_col + 1)
    cell_titulo = ws.cell(row=2, column=start_col)
    cell_titulo.value = titulo
    cell_titulo.alignment = Alignment(horizontal='center', vertical='center')
    cell_titulo.font = Font(size=14, bold=True)
    ws.row_dimensions[2].height = 35

    header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF")
    largura_colunas = [25, 100]  # Ajuste a largura conforme desejar
    for i, col_offset in enumerate(range(start_col, start_col + 2)):
        cell = ws.cell(row=start_row, column=col_offset)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        ws.column_dimensions[get_column_letter(col_offset)].width = largura_colunas[i]

    ws.row_dimensions[start_row].height = 40

    thin_border = Border(
        left=Side(border_style="thin", color="A9B7C6"),
        right=Side(border_style="thin", color="A9B7C6"),
        top=Side(border_style="thin", color="A9B7C6"),
        bottom=Side(border_style="thin", color="A9B7C6")
    )

    fill_light = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    fill_none = PatternFill(fill_type=None)

    max_row = ws.max_row
    max_col = start_col + 1

    for row_idx in range(start_row + 1, max_row + 1):
        fill = fill_light if (row_idx - start_row) % 2 == 1 else fill_none
        for col_idx in range(start_col, max_col + 1):
            cell = ws.cell(row=row_idx, column=col_idx)
    
    # Centraliza apenas a primeira coluna ("Data")
            if col_idx == start_col:
                cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            else:
                cell.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
    
            cell.border = thin_border
            cell.fill = fill
            ws.row_dimensions[row_idx].height = 80

    wb.save(caminho)
    return send_file(caminho, as_attachment=True, download_name="relatorio.xlsx")

@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))

if __name__ == "__main__":
    init_db()
    app.run(host="0.0.0.0", port=5000, debug=True)
