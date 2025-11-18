# ============================================
# AGENTE CINTHYA — VERSÃO CLOUD (Render.com)
# ============================================

import requests
import os
import tempfile
import shutil
from pathlib import Path
from flask import Flask, request, render_template_string, send_from_directory
from datetime import datetime
import pandas as pd
from unidecode import unidecode

# ============================================================
# CONFIGURAÇÃO TRELO
# ============================================================

API_KEY = os.getenv("TRELLO_API_KEY")
TOKEN = os.getenv("TRELLO_TOKEN")
BOARD_SHORTLINK = "qYOL5Nyj"


# ============================================================
# FUNÇÕES BASE DO TRELLO
# ============================================================

def trello_get(url, params=None):
    if params is None:
        params = {}
    params["key"] = API_KEY
    params["token"] = TOKEN
    return requests.get(url, params=params).json()

def trello_post(url, params=None):
    if params is None:
        params = {}
    params["key"] = API_KEY
    params["token"] = TOKEN
    return requests.post(url, params=params).json()

def trello_put(url, params=None):
    if params is None:
        params = {}
    params["key"] = API_KEY
    params["token"] = TOKEN
    return requests.put(url, params=params).json()

def trello_delete(url, params=None):
    if params is None:
        params = {}
    params["key"] = API_KEY
    params["token"] = TOKEN
    return requests.delete(url, params=params).json()


# ============================================================
# FUNÇÕES DO TRELLO — LISTAS / CARTÕES / COMENTÁRIOS / ETC
# ============================================================

def listar_listas():
    return trello_get(f"https://api.trello.com/1/boards/{BOARD_SHORTLINK}/lists")

def listar_cartoes(lista_id):
    return trello_get(f"https://api.trello.com/1/lists/{lista_id}/cards")

def criar_cartao(nome, lista_id, desc=""):
    return trello_post("https://api.trello.com/1/cards",
                       {"name": nome, "idList": lista_id, "desc": desc})

def mover_cartao(cartao_id, nova_lista):
    return trello_put(f"https://api.trello.com/1/cards/{cartao_id}",
                      {"idList": nova_lista})

def atualizar_descricao(cartao_id, texto):
    return trello_put(f"https://api.trello.com/1/cards/{cartao_id}/desc",
                      {"value": texto})


def listar_comentarios(cartao_id):
    actions = trello_get(
        f"https://api.trello.com/1/cards/{cartao_id}/actions?filter=commentCard"
    )
    lista = []
    for c in actions:
        lista.append({"id": c["id"], "texto": c["data"]["text"]})
    return lista


def adicionar_comentario(cid, txt):
    return trello_post(
        f"https://api.trello.com/1/cards/{cid}/actions/comments",
        {"text": txt}
    )


def editar_comentario(comment_id, novo_txt):
    return trello_put(
        f"https://api.trello.com/1/actions/{comment_id}/text",
        {"value": novo_txt}
    )


def deletar_comentario(comment_id):
    return trello_delete(f"https://api.trello.com/1/actions/{comment_id}")


# ============================================================
# WORD — PESQUISA
# ============================================================

from docx import Document

WORD_PATH = Path(
    r"C:\Users\Cinthya\OneDrive - Fiap-Faculdade de Informática e Administração Paulista\Python\Diário Cinthya - IA.docx"
)

def tmp_copy(path):
    t = Path(tempfile.gettempdir()) / f"tmp_{int(datetime.now().timestamp())}.docx"
    shutil.copy2(path, t)
    return t

def load_word_lines():
    temp = tmp_copy(WORD_PATH)
    doc = Document(temp)
    return [p.text.strip() for p in doc.paragraphs]

def norm(s):
    return unidecode(str(s)).lower()

def search_word(termo):
    t = norm(termo)
    return [l for l in load_word_lines() if t in norm(l)]


# ============================================================
# EXCEL — PESQUISA
# ============================================================

XLSX_PATH = Path(
    r"C:\Users\Cinthya\OneDrive - Fiap-Faculdade de Informática e Administração Paulista\Documentos - Financeiro - Tesouraria\Tesouraria\Conciliação Bancária\Automações\Bases de Vendas Atualizadas Diariamente.xlsx"
)

def search_excel(termo):
    temp = tmp_copy(XLSX_PATH)
    dfs = pd.read_excel(temp, sheet_name=None, dtype=str, keep_default_na=False)
    t = norm(termo)
    results = []
    for sheet, df in dfs.items():
        df_norm = df.applymap(lambda x: norm(x))
        mask = df_norm.applymap(lambda v: t in v)
        rows = df[mask.any(axis=1)]
        if not rows.empty:
            results.append({"sheet": sheet, "html": rows.to_html(index=False)})
    return results


# ============================================================
# FLASK APP
# ============================================================
app = Flask(__name__, static_folder="static")


# ============================================================
# INTERFACE (HTML BASE)
# ============================================================

BASE_HTML = """
<!doctype html>
<html>
<head>
<meta charset="utf-8">
<title>Agente da Cinthya 💖</title>

<style>
:root {
    --bg:#f7eefe;
    --card:#ffffff;
    --text:#000;
    --pink:#ff2ac5;
    --pink-hover:#ff0095;
}
body.dark {
    --bg:#1f1f1f;
    --card:#2a2a2a;
    --text:#f2f2f2;
    --pink:#ff4ccc;
    --pink-hover:#ff7ee3;
}
body { margin:0;background:var(--bg);font-family:Arial;color:var(--text); }
.topbar { display:flex;justify-content:space-between;align-items:center;padding:20px 30px; }
.logo { height:125px; }
.btn { background:var(--pink);color:white;padding:12px 24px;border:none;border-radius:14px;cursor:pointer;font-size:18px;margin:4px;font-weight:bold; }
.btn:hover { background:var(--pink-hover); }
h1 { text-align:center;color:var(--pink);font-size:36px;margin-top:-10px;margin-bottom:15px; }
.wrap { max-width:1200px;margin:0 auto;padding:10px; }
.section { background:var(--card);padding:25px;border-radius:20px;box-shadow:0 3px 12px rgba(0,0,0,0.25);margin-bottom:25px; }
.searchbox { width:100%;padding:11px;border-radius:12px;border:1px solid #bbb;margin-top:8px;font-size:16px; }
.bigbox { width:100%;height:85px;padding:12px;border-radius:12px;border:1px solid #aaa;font-size:16px;white-space:pre-wrap; }
.cardtitle { font-size:21px;padding:8px 6px;cursor:pointer; }
.dot { width:12px;height:12px;background:var(--pink);border-radius:50%;display:inline-block;margin-right:10px; }
.card-box { display:none; }
</style>

<script>
function mudarTema(){ document.body.classList.toggle("dark"); }
function abrirCard(id){ document.querySelectorAll(".card-box").forEach(b=>b.style.display="none"); document.getElementById(id).style.display="block"; }
function enviar(form_id){ 
    const f=document.getElementById(form_id); 
    fetch(f.action,{method:"POST",body:new FormData(f)})
        .then(r=>r.text())
        .then(t=>alert(t));
}
function atualizarDashboard(){
    fetch('/dashboard_refresh',{method:"POST"})
    .then(r=>r.text())
    .then(h=>{document.getElementById("dashboard_area").innerHTML=h;});
}
</script>

</head>

<body>

<div class="topbar">
    <img src="/static/alun.png" class="logo">
    <div>
        <button class="btn" onclick="mudarTema()">Tema 🌗</button>
        <button class="btn" onclick="atualizarDashboard()">Atualizar Dashboard 📊</button>
        <button class="btn" onclick="location.reload()">Atualizar Página 🔄</button>
    </div>
</div>

<h1>Agente da Cinthya 💖</h1>

<div class="wrap">

    <div style="text-align:center;">
        <button class="btn" onclick="window.location='/pesquisa_menu'">Pesquisa</button>
        <button class="btn" onclick="window.location='/cartoes_menu'">Cartões</button>
        <button class="btn" onclick="window.location='/criar_cartao_menu'">Criar Cartão</button>
    </div>

    <div class="section">
        {{conteudo|safe}}
    </div>

    <div id="dashboard_area" class="section">
        {{dashboard_html|safe}}
    </div>

</div>

</body>
</html>
"""


# ============================================================
# ROTAS DO SISTEMA
# ============================================================

@app.route("/")
def home():
    return render_template_string(BASE_HTML,
                                  conteudo="<p>Selecione uma opção acima 💖</p>",
                                  dashboard_html=gerar_dashboard_html())


# ------------ PESQUISA MENU -------------
@app.route("/pesquisa_menu")
def pesquisa_menu():
    listas = listar_listas()
    html = """
    <h2>Pesquisa no Word e Excel</h2>

    <form method="POST" action="/pesquisa">
        <label><input type="checkbox" name="fonte" value="word" checked> Word</label><br>
        <label><input type="checkbox" name="fonte" value="excel" checked> Excel</label><br><br>

        <input class="searchbox" name="q" placeholder="Digite uma palavra..."><br><br>
        <button class="btn">Pesquisar</button>
    </form>
    """
    return render_template_string(BASE_HTML, conteudo=html,
                                  dashboard_html=gerar_dashboard_html())


# ------------ EXECUTAR PESQUISA -------------
@app.route("/pesquisa", methods=["POST"])
def pesquisa():
    termo = request.form["q"]
    fontes = request.form.getlist("fonte")

    html = "<h2>Resultados:</h2>"
    achou = False

    if "word" in fontes:
        for l in search_word(termo):
            achou = True
            html += f"<p>{l}</p>"

    if "excel" in fontes:
        res = search_excel(termo)
        if res:
            achou = True
            for bloco in res:
                html += f"<h3>{bloco['sheet']}</h3>{bloco['html']}<br>"

    if not achou:
        html += "<p>Nada encontrado.</p>"

    return render_template_string(BASE_HTML, conteudo=html,
                                  dashboard_html=gerar_dashboard_html())


# ------------ MENU DE CARTÕES -------------
@app.route("/cartoes_menu")
def cartoes_menu():
    listas = listar_listas()
    op = "".join([f"<option value='{l['id']}'>{l['name']}</option>" for l in listas])

    html = f"""
    <h2>Escolha uma lista</h2>
    <form method="POST" action="/carregar_lista">
        <select class="searchbox" name="lista">{op}</select><br><br>
        <button class="btn">Carregar Cartões</button>
    </form>
    """
    return render_template_string(BASE_HTML, conteudo=html,
                                  dashboard_html=gerar_dashboard_html())


# -------------------------------------------------------
#  CARREGAR CARTÕES + DESCRIÇÃO + COMENTÁRIOS
# -------------------------------------------------------
@app.route("/carregar_lista", methods=["POST"])
def carregar_lista():
    lista = request.form["lista"]
    cards = listar_cartoes(lista)

    html = "<h2>Cartões:</h2>"

    for c in cards:
        cid = c["id"]
        titulo = c["name"]
        desc = c.get("desc", "")
        comentarios = listar_comentarios(cid)

        bloco_com = "\n".join([c["texto"] for c in comentarios])

        html += f"""
        <div class="cardtitle" onclick="abrirCard('box_{cid}')">
            <span class="dot"></span> {titulo}
        </div>

        <div id="box_{cid}" class="card-box section">

            <h3>📄 Descrição</h3>
            <form id="desc_{cid}"
                  action="/salvar_descricao/{cid}"
                  onsubmit="event.preventDefault(); enviar('desc_{cid}');">
                <textarea class="bigbox" name="descricao">{desc}</textarea><br>
                <button class="btn">Salvar Descrição</button>
            </form>

            <h3>📝 Comentários</h3>
            <form id="com_{cid}"
                  action="/salvar_comentarios/{cid}"
                  onsubmit="event.preventDefault(); enviar('com_{cid}');">
                <textarea class="bigbox" name="comentarios">{bloco_com}</textarea><br>
                <button class="btn">Salvar Comentários</button>
            </form>

            <h3>Mover cartão para:</h3>
            <form id="move_{cid}"
                  action="/mover_cartao/{cid}"
                  onsubmit="event.preventDefault(); enviar('move_{cid}');">
                <select class="searchbox" name="lista">
        """

        for l in listar_listas():
            html += f"<option value='{l['id']}'>{l['name']}</option>"

        html += """
                </select><br><br>
                <button class="btn">Mover</button>
            </form>
        </div><br>
        """

    return render_template_string(BASE_HTML, conteudo=html,
                                  dashboard_html=gerar_dashboard_html())


# ------------ SALVAR DESCRIÇÃO -------------
@app.route("/salvar_descricao/<cid>", methods=["POST"])
def salvar_descricao(cid):
    atualizar_descricao(cid, request.form["descricao"])
    return "Descrição atualizada 💖"


# ------------ SALVAR COMENTÁRIOS -------------
@app.route("/salvar_comentarios/<cid>", methods=["POST"])
def salvar_comentarios(cid):
    digitado = request.form["comentarios"]
    linhas = [l.strip() for l in digitado.split("\n") if l.strip()]

    comentarios = listar_comentarios(cid)
    originais = [c["texto"] for c in comentarios]
    mapa_id = {c["texto"]: c["id"] for c in comentarios}

    novas = []
    for l in linhas:
        if (
            len(l) > 16 and
            l[4] == "-" and
            l[7] == "-" and
            l[13] == ":"
        ):
            novas.append(l)
        else:
            d = datetime.now().strftime("%Y-%m-%d %H:%M")
            novas.append(f"{d} — {l}")

    removidos = [o for o in originais if o not in novas]
    novos = [n for n in novas if n not in originais]

    def limpar(t):
        if "(editado em" in t:
            return t.split("(editado em")[0].strip()
        return t

    editados = []
    for o in originais:
        base_o = limpar(o)
        for n in novas:
            base_n = limpar(n)
            if base_o == base_n and o != n:
                editados.append((o, n))

    for r in removidos:
        deletar_comentario(mapa_id[r])

    for o, n in editados:
        cid_com = mapa_id[o]
        data_original = o.split("—")[0].strip()
        base_txt = limpar(n)
        data_edit = datetime.now().strftime("%Y-%m-%d %H:%M")
        txt_final = f"{data_original} — {base_txt} (editado em {data_edit})"
        editar_comentario(cid_com, txt_final)

    for n in novos:
        adicionar_comentario(cid, n)

    return "Comentários atualizados 💖"


# ------------ MOVER CARTÃO -------------
@app.route("/mover_cartao/<cid>", methods=["POST"])
def mover_cartao_rota(cid):
    mover_cartao(cid, request.form["lista"])
    return "Cartão movido ✔"


# ------------ CRIAR CARTÃO (MENU) -------------
@app.route("/criar_cartao_menu")
def criar_cartao_menu():
    listas = listar_listas()
    op = "".join([f"<option value='{l['id']}'>{l['name']}</option>" for l in listas])

    html = f"""
    <h2>Criar Novo Cartão</h2>

    <form method="POST" action="/criar_cartao">
        <label>Nome do cartão:</label>
        <input class="searchbox" name="nome" required><br><br>

        <label>Descrição inicial:</label>
        <textarea class="bigbox" name="desc"></textarea><br><br>

        <label>Lista:</label>
        <select class="searchbox" name="lista">{op}</select><br><br>

        <button class="btn">Criar</button>
    </form>
    """

    return render_template_string(BASE_HTML, conteudo=html,
                                  dashboard_html=gerar_dashboard_html())


# ------------ CRIAR CARTÃO (AÇÃO) -------------
@app.route("/criar_cartao", methods=["POST"])
def criar_cartao_rota():
    criar_cartao(
        request.form["nome"],
        request.form["lista"],
        request.form["desc"]
    )
    return "Cartão criado com sucesso 💖"


# ------------ DASHBOARD -------------
def gerar_dashboard_html():
    listas = listar_listas()
    if not listas:
        return "<p>Dashboard indisponível.</p>"

    html = "<h2>📊 Dashboard</h2>"

    qtds = {l["name"]: len(listar_cartoes(l["id"])) for l in listas}
    maior = max(qtds.values()) if qtds else 1

    for nome, total in qtds.items():
        largura = 40 + int((total / maior) * 260)
        html += f"""
        <p><b>{nome}</b>: {total} cartões</p>
        <div style="width:{largura}px;height:22px;background:var(--pink);
                    border-radius:12px;margin-bottom:15px;"></div>
        """

    return html


@app.route("/dashboard_refresh", methods=["POST"])
def dashboard_refresh():
    return gerar_dashboard_html()


# ============================================================
# ARQUIVOS ESTÁTICOS
# ============================================================
@app.route("/static/<path:filename>")
def static_files(filename):
    return send_from_directory(app.static_folder, filename)


# ============================================================
# EXECUÇÃO
# ============================================================

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
