from datetime import date
from io import BytesIO
import sqlite3
from pathlib import Path
import uuid
import os

from flask import (
    Flask, render_template, request, redirect, url_for,
    flash, send_file, send_from_directory
)
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo
from dotenv import load_dotenv

APP_DIR = Path(__file__).resolve().parent
load_dotenv(APP_DIR / ".env")
DB_PATH = APP_DIR / "registros.db"
UPLOAD_FOLDER = APP_DIR / "fotos_avarias"
UPLOAD_FOLDER.mkdir(exist_ok=True)

ALLOWED_EXTENSIONS = {".jpg", ".jpeg", ".png", ".webp"}
MAX_PHOTO_SIZE = 10 * 1024 * 1024  # 10 MB

app = Flask(__name__)
app.secret_key = os.getenv("SECRET_KEY")
SENHA_EXCLUSAO = os.getenv("SENHA_EXCLUSAO")
app.config["MAX_CONTENT_LENGTH"] = MAX_PHOTO_SIZE


def conectar():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def preparar_banco():
    with conectar() as conn:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS lancamentos (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                produto TEXT NOT NULL,
                quantidade REAL NOT NULL,
                unidade TEXT NOT NULL CHECK(unidade IN ('kg', 'un')),
                data_lancamento TEXT NOT NULL,
                responsavel TEXT NOT NULL,
                tipo TEXT NOT NULL CHECK(tipo IN ('Avaria', 'Consumo')),
                foto_avaria TEXT,
                criado_em TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
            )
        """)

        # Se o banco já existia antes da função de fotos,
        # adiciona somente a nova coluna sem apagar os registros.
        colunas = [
            coluna["name"]
            for coluna in conn.execute("PRAGMA table_info(lancamentos)").fetchall()
        ]

        if "foto_avaria" not in colunas:
            conn.execute("ALTER TABLE lancamentos ADD COLUMN foto_avaria TEXT")


def salvar_foto(arquivo):
    """Salva a foto com nome único e retorna o nome salvo."""
    if not arquivo or not arquivo.filename:
        return None

    extensao = Path(arquivo.filename).suffix.lower()

    if extensao not in ALLOWED_EXTENSIONS:
        raise ValueError("Formato de foto inválido. Use JPG, JPEG, PNG ou WEBP.")

    nome = f"avaria_{date.today().strftime('%Y%m%d')}_{uuid.uuid4().hex[:12]}{extensao}"
    destino = UPLOAD_FOLDER / nome
    arquivo.save(destino)

    return nome


def excluir_foto(nome_foto):
    if not nome_foto:
        return

    caminho = UPLOAD_FOLDER / Path(nome_foto).name

    try:
        if caminho.is_file():
            caminho.unlink()
    except OSError:
        pass


@app.errorhandler(413)
def arquivo_muito_grande(_erro):
    flash("A foto é muito grande. O tamanho máximo é 10 MB.", "erro")
    return redirect(url_for("inicio"))


@app.route("/", methods=["GET", "POST"])
def inicio():

    # =====================================================
    # NOVO LANÇAMENTO
    # =====================================================

    if request.method == "POST":

        produto = request.form.get("produto", "").strip()

        quantidade_txt = (
            request.form.get("quantidade", "")
            .strip()
            .replace(",", ".")
        )

        unidade = request.form.get("unidade", "")

        data_lancamento = request.form.get(
            "data_lancamento",
            ""
        )

        responsavel = request.form.get(
            "responsavel",
            ""
        ).strip()

        tipo = request.form.get("tipo", "")

        arquivo_foto = request.files.get("foto")

        try:
            quantidade = float(quantidade_txt)
        except ValueError:
            quantidade = 0

        # =================================================
        # VALIDAÇÃO
        # =================================================

        if (
            not produto
            or not responsavel
            or quantidade <= 0
            or unidade not in ("kg", "un")
            or tipo not in ("Avaria", "Consumo")
            or not data_lancamento
        ):

            flash(
                "Confira os campos. "
                "A quantidade deve ser maior que zero.",
                "erro"
            )

            return redirect(
                url_for("inicio")
            )

        # =================================================
        # VALIDA DATA
        # =================================================

        try:

            date.fromisoformat(
                data_lancamento
            )

        except ValueError:

            flash(
                "A data informada não é válida.",
                "erro"
            )

            return redirect(
                url_for("inicio")
            )

        # =================================================
        # FOTO OBRIGATÓRIA PARA AVARIA
        # =================================================

        if tipo == "Avaria" and (
            not arquivo_foto
            or not arquivo_foto.filename
        ):

            flash(
                "É obrigatório adicionar uma foto "
                "para registrar uma avaria.",
                "erro"
            )

            return redirect(
                url_for("inicio")
            )

        foto_salva = None

        # =================================================
        # SALVAR NOVO LANÇAMENTO
        # =================================================

        try:

            if tipo == "Avaria":

                foto_salva = salvar_foto(
                    arquivo_foto
                )

            with conectar() as conn:

                conn.execute("""
                    INSERT INTO lancamentos (
                        produto,
                        quantidade,
                        unidade,
                        data_lancamento,
                        responsavel,
                        tipo,
                        foto_avaria
                    )
                    VALUES (?, ?, ?, ?, ?, ?, ?)
                """, (
                    produto,
                    quantidade,
                    unidade,
                    data_lancamento,
                    responsavel,
                    tipo,
                    foto_salva
                ))

            flash(
                "Lançamento salvo com sucesso!",
                "sucesso"
            )

        except ValueError as erro:

            flash(
                str(erro),
                "erro"
            )

        except Exception:

            if foto_salva:

                excluir_foto(
                    foto_salva
                )

            flash(
                "Não foi possível salvar o lançamento.",
                "erro"
            )

        return redirect(
            url_for("inicio")
        )

    # =====================================================
    # PAGINAÇÃO
    # =====================================================

    try:

        pagina = int(
            request.args.get(
                "pagina",
                1
            )
        )

    except ValueError:

        pagina = 1

    if pagina < 1:
        pagina = 1

    # Quantidade de registros por página
    por_pagina = 10

    # =====================================================
    # FILTROS
    # =====================================================

    filtro_data = request.args.get(
        "data",
        ""
    ).strip()

    filtro_tipo = request.args.get(
        "tipo_filtro",
        ""
    ).strip()

    # =====================================================
    # MONTA O FILTRO SQL
    # =====================================================

    condicoes = []
    parametros = []

    # Filtro por data
    if filtro_data:

        # Confirma se a data é válida
        try:

            date.fromisoformat(
                filtro_data
            )

            condicoes.append(
                "data_lancamento = ?"
            )

            parametros.append(
                filtro_data
            )

        except ValueError:

            filtro_data = ""

    # Filtro por tipo
    if filtro_tipo in (
        "Avaria",
        "Consumo"
    ):

        condicoes.append(
            "tipo = ?"
        )

        parametros.append(
            filtro_tipo
        )

    else:

        filtro_tipo = ""

    # =====================================================
    # WHERE
    # =====================================================

    where = ""

    if condicoes:

        where = (
            "WHERE "
            + " AND ".join(condicoes)
        )

    # =====================================================
    # CONSULTA
    # =====================================================

    with conectar() as conn:

        # -------------------------------------------------
        # TOTAL DE REGISTROS COM OS FILTROS
        # -------------------------------------------------

        total = conn.execute(
            f"""
            SELECT COUNT(*)
            FROM lancamentos
            {where}
            """,
            parametros
        ).fetchone()[0]

        # -------------------------------------------------
        # TOTAL DE PÁGINAS
        # -------------------------------------------------

        total_paginas = max(
            1,
            (total + por_pagina - 1)
            // por_pagina
        )

        # -------------------------------------------------
        # CORRIGE PÁGINA INVÁLIDA
        # -------------------------------------------------

        if pagina > total_paginas:

            pagina = total_paginas

        # -------------------------------------------------
        # OFFSET
        # -------------------------------------------------

        offset = (
            (pagina - 1)
            * por_pagina
        )

        # -------------------------------------------------
        # BUSCA SOMENTE OS 10 DA PÁGINA
        # -------------------------------------------------

        recentes = conn.execute(
            f"""
            SELECT
                id,
                produto,
                quantidade,
                unidade,
                data_lancamento,
                responsavel,
                tipo,
                foto_avaria

            FROM lancamentos

            {where}

            ORDER BY
                data_lancamento DESC,
                id DESC

            LIMIT ?
            OFFSET ?
            """,
            parametros + [
                por_pagina,
                offset
            ]
        ).fetchall()

    # =====================================================
    # ENVIA PARA O INDEX.HTML
    # =====================================================

    return render_template(
        "index.html",

        hoje=date.today().isoformat(),

        recentes=recentes,

        total=total,

        pagina=pagina,

        total_paginas=total_paginas,

        filtro_data=filtro_data,

        filtro_tipo=filtro_tipo
    )
    if request.method == "POST":
        produto = request.form.get("produto", "").strip()
        quantidade_txt = request.form.get("quantidade", "").strip().replace(",", ".")
        unidade = request.form.get("unidade", "")
        data_lancamento = request.form.get("data_lancamento", "")
        responsavel = request.form.get("responsavel", "").strip()
        tipo = request.form.get("tipo", "")
        arquivo_foto = request.files.get("foto")

        try:
            quantidade = float(quantidade_txt)
        except ValueError:
            quantidade = 0

        if (
            not produto
            or not responsavel
            or quantidade <= 0
            or unidade not in ("kg", "un")
            or tipo not in ("Avaria", "Consumo")
            or not data_lancamento
        ):
            flash("Confira os campos. A quantidade deve ser maior que zero.", "erro")
            return redirect(url_for("inicio"))

        try:
            date.fromisoformat(data_lancamento)
        except ValueError:
            flash("A data informada não é válida.", "erro")
            return redirect(url_for("inicio"))

        # Foto obrigatória somente para avaria.
        if tipo == "Avaria" and (
            not arquivo_foto or not arquivo_foto.filename
        ):
            flash("É obrigatório adicionar uma foto para registrar uma avaria.", "erro")
            return redirect(url_for("inicio"))

        foto_salva = None

        try:
            if tipo == "Avaria":
                foto_salva = salvar_foto(arquivo_foto)

            with conectar() as conn:
                conn.execute("""
                    INSERT INTO lancamentos (
                        produto, quantidade, unidade, data_lancamento,
                        responsavel, tipo, foto_avaria
                    )
                    VALUES (?, ?, ?, ?, ?, ?, ?)
                """, (
                    produto,
                    quantidade,
                    unidade,
                    data_lancamento,
                    responsavel,
                    tipo,
                    foto_salva
                ))

            flash("Lançamento salvo com sucesso!", "sucesso")

        except ValueError as erro:
            flash(str(erro), "erro")
        except Exception:
            if foto_salva:
                excluir_foto(foto_salva)
            flash("Não foi possível salvar o lançamento.", "erro")

        return redirect(url_for("inicio"))

    with conectar() as conn:
        recentes = conn.execute("""
            SELECT id, produto, quantidade, unidade,
                   data_lancamento, responsavel, tipo, foto_avaria
            FROM lancamentos
            ORDER BY data_lancamento DESC, id DESC
        """).fetchall()

        total = conn.execute(
            "SELECT COUNT(*) FROM lancamentos"
        ).fetchone()[0]

    return render_template(
        "index.html",
        hoje=date.today().isoformat(),
        recentes=recentes,
        total=total
    )


@app.route("/editar/<int:id>", methods=["GET", "POST"])
def editar(id):
    with conectar() as conn:
        registro = conn.execute(
            "SELECT * FROM lancamentos WHERE id = ?", (id,)
        ).fetchone()

        if registro is None:
            flash("Lançamento não encontrado.", "erro")
            return redirect(url_for("inicio"))

        if request.method == "POST":
            produto = request.form.get("produto", "").strip()
            quantidade_txt = request.form.get("quantidade", "").strip().replace(",", ".")
            unidade = request.form.get("unidade", "")
            data_lancamento = request.form.get("data_lancamento", "")
            responsavel = request.form.get("responsavel", "").strip()
            tipo = request.form.get("tipo", "")
            arquivo_foto = request.files.get("foto")

            try:
                quantidade = float(quantidade_txt)
            except ValueError:
                quantidade = 0

            if (
                not produto
                or not responsavel
                or quantidade <= 0
                or unidade not in ("kg", "un")
                or tipo not in ("Avaria", "Consumo")
            ):
                flash("Confira os campos preenchidos.", "erro")
                return redirect(url_for("editar", id=id))

            try:
                date.fromisoformat(data_lancamento)
            except ValueError:
                flash("Data inválida.", "erro")
                return redirect(url_for("editar", id=id))

            foto_atual = registro["foto_avaria"]
            nova_foto = None

            try:
                if tipo == "Avaria":
                    # Se estiver editando uma avaria e não escolher outra foto,
                    # mantém a foto que já estava cadastrada.
                    if arquivo_foto and arquivo_foto.filename:
                        nova_foto = salvar_foto(arquivo_foto)
                    elif not foto_atual:
                        flash(
                            "Esta avaria ainda não possui foto. "
                            "É obrigatório adicionar uma foto.",
                            "erro"
                        )
                        return redirect(url_for("editar", id=id))

                    foto_final = nova_foto if nova_foto else foto_atual

                else:
                    # Se mudar de Avaria para Consumo, remove a referência da foto.
                    foto_final = None

                conn.execute("""
                    UPDATE lancamentos
                    SET produto = ?, quantidade = ?, unidade = ?,
                        data_lancamento = ?, responsavel = ?,
                        tipo = ?, foto_avaria = ?
                    WHERE id = ?
                """, (
                    produto,
                    quantidade,
                    unidade,
                    data_lancamento,
                    responsavel,
                    tipo,
                    foto_final,
                    id
                ))

                # Depois de atualizar o banco, remove a foto antiga
                # se uma nova foi escolhida ou se virou Consumo.
                if foto_atual and (
                    (nova_foto and nova_foto != foto_atual)
                    or (tipo == "Consumo")
                ):
                    excluir_foto(foto_atual)

                flash("Lançamento atualizado com sucesso!", "sucesso")
                return redirect(url_for("inicio"))

            except ValueError as erro:
                flash(str(erro), "erro")
                return redirect(url_for("editar", id=id))
            except Exception:
                if nova_foto:
                    excluir_foto(nova_foto)
                flash("Não foi possível atualizar o lançamento.", "erro")
                return redirect(url_for("editar", id=id))

    return render_template("editar.html", registro=registro)


@app.route("/baixar-foto/<int:id>")
def baixar_foto(id):
    with conectar() as conn:
        registro = conn.execute(
            "SELECT foto_avaria, tipo FROM lancamentos WHERE id = ?",
            (id,)
        ).fetchone()

    if registro is None:
        flash("Lançamento não encontrado.", "erro")
        return redirect(url_for("inicio"))

    if registro["tipo"] != "Avaria" or not registro["foto_avaria"]:
        flash("Este lançamento não possui uma foto de avaria.", "erro")
        return redirect(url_for("inicio"))

    nome_foto = Path(registro["foto_avaria"]).name
    caminho = UPLOAD_FOLDER / nome_foto

    if not caminho.is_file():
        flash("A foto não foi encontrada no computador.", "erro")
        return redirect(url_for("inicio"))

    return send_from_directory(
        UPLOAD_FOLDER,
        nome_foto,
        as_attachment=True
    )



@app.route("/excluir/<int:id>", methods=["POST"])
def excluir(id):
    senha = request.form.get("senha", "")

    if senha != SENHA_EXCLUSAO:
        flash("Senha incorreta. O registro não foi excluído.", "erro")
        return redirect(url_for("inicio"))

    with conectar() as conn:
        registro = conn.execute(
            "SELECT foto_avaria FROM lancamentos WHERE id = ?",
            (id,)
        ).fetchone()

        if registro is None:
            flash("Lançamento não encontrado.", "erro")
            return redirect(url_for("inicio"))

        conn.execute(
            "DELETE FROM lancamentos WHERE id = ?",
            (id,)
        )

    if registro["foto_avaria"]:
        excluir_foto(registro["foto_avaria"])

    flash("Lançamento excluído com sucesso!", "sucesso")
    return redirect(url_for("inicio"))


@app.route("/limpar-todos", methods=["POST"])
def limpar_todos():
    senha = request.form.get("senha", "")

    if senha != SENHA_EXCLUSAO:
        flash("Senha incorreta. Os registros não foram excluídos.", "erro")
        return redirect(url_for("inicio"))

    with conectar() as conn:
        registros = conn.execute(
            "SELECT foto_avaria FROM lancamentos"
        ).fetchall()

        conn.execute("DELETE FROM lancamentos")

    for registro in registros:
        if registro["foto_avaria"]:
            excluir_foto(registro["foto_avaria"])

    flash("Todos os lançamentos foram excluídos com sucesso!", "sucesso")
    return redirect(url_for("inicio"))



@app.route("/exportar")
def exportar():
    with conectar() as conn:
        linhas = conn.execute("""
            SELECT data_lancamento, produto, quantidade,
                   unidade, tipo, responsavel
            FROM lancamentos
            ORDER BY data_lancamento ASC, id ASC
        """).fetchall()

    wb = Workbook()

    # =====================================================
    # ABA 1 - LANÇAMENTOS
    # =====================================================

    ws = wb.active
    ws.title = "Avarias e Consumo"

    cabecalhos = [
        "Data", "Produto", "Quantidade",
        "Unidade", "Tipo", "Lançado por"
    ]
    ws.append(cabecalhos)

    total_por_unidade = {
        "kg": 0,
        "un": 0
    }

    for linha in linhas:
        data_formatada = date.fromisoformat(
            linha["data_lancamento"]
        ).strftime("%d/%m/%Y")

        quantidade = float(linha["quantidade"])
        unidade = linha["unidade"]

        if unidade in total_por_unidade:
            total_por_unidade[unidade] += quantidade

        ws.append([
            data_formatada,
            linha["produto"],
            quantidade,
            unidade,
            linha["tipo"],
            linha["responsavel"]
        ])

    verde = PatternFill("solid", fgColor="237A4B")
    verde_claro = PatternFill("solid", fgColor="D9EAD3")
    cinza = PatternFill("solid", fgColor="E7E6E6")
    borda = Border(
        bottom=Side(style="thin", color="B7B7B7")
    )

    # Cabeçalho da tabela principal
    for celula in ws[1]:
        celula.font = Font(bold=True, color="FFFFFF")
        celula.fill = verde
        celula.alignment = Alignment(horizontal="center")
        celula.border = borda

    # Formatação das linhas
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        row[0].alignment = Alignment(horizontal="center")
        row[2].alignment = Alignment(horizontal="right")
        row[3].alignment = Alignment(horizontal="center")
        row[4].alignment = Alignment(horizontal="center")

        # Mostra quantidades com até 3 casas, sem zeros desnecessários.
        row[2].number_format = '#,##0.###'

    # Cria uma Tabela do Excel na primeira aba.
    if len(linhas) > 0:
        tabela = Table(
            displayName="TabelaLancamentos",
            ref=f"A1:F{ws.max_row}"
        )

        estilo_tabela = TableStyleInfo(
            name="TableStyleMedium4",
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=True,
            showColumnStripes=False
        )

        tabela.tableStyleInfo = estilo_tabela
        ws.add_table(tabela)

    # =====================================================
    # TOTAIS DA PRIMEIRA ABA
    # =====================================================

    linha_total = ws.max_row + 2

    ws.cell(linha_total, 1, "TOTAIS DE QUANTIDADE")
    ws.cell(linha_total, 1).font = Font(
        bold=True,
        color="FFFFFF"
    )
    ws.cell(linha_total, 1).fill = verde

    ws.merge_cells(
        start_row=linha_total,
        start_column=1,
        end_row=linha_total,
        end_column=2
    )

    ws.cell(linha_total + 1, 1, "Total em kg")
    ws.cell(linha_total + 1, 2, total_por_unidade["kg"])

    ws.cell(linha_total + 2, 1, "Total em unidades")
    ws.cell(linha_total + 2, 2, total_por_unidade["un"])

    # Remove as casas decimais apenas das unidades
    ws.cell(linha_total + 2, 2).number_format = '0'

    for linha in range(linha_total + 1, linha_total + 3):
        ws.cell(linha, 1).font = Font(bold=True)
        ws.cell(linha, 1).fill = verde_claro
        ws.cell(linha, 2).font = Font(bold=True)
        ws.cell(linha, 2).fill = verde_claro
        ws.cell(linha, 2).number_format = '#,##0.###'

    # =====================================================
    # LARGURAS E FILTRO DA PRIMEIRA ABA
    # =====================================================

    larguras = [14, 32, 16, 12, 16, 24]

    for i, largura in enumerate(larguras, 1):
        ws.column_dimensions[get_column_letter(i)].width = largura

    ws.freeze_panes = "A2"

    # =====================================================
    # ABA 2 - RESUMO / TABELA DINÂMICA
    # =====================================================

    resumo = wb.create_sheet("Resumo")

    resumo["A1"] = "RESUMO DOS PRODUTOS"
    resumo["A1"].font = Font(
        bold=True,
        size=16,
        color="FFFFFF"
    )
    resumo["A1"].fill = verde
    resumo["A1"].alignment = Alignment(
        horizontal="center",
        vertical="center"
    )

    resumo.merge_cells("A1:D1")
    resumo.row_dimensions[1].height = 28

    resumo["A2"] = (
        "Produtos ordenados pela quantidade total, "
        "do maior para o menor."
    )
    resumo["A2"].font = Font(italic=True)
    resumo.merge_cells("A2:D2")

    # Agrupa por produto + unidade.
    # Isso evita somar, por exemplo, 5 kg com 3 unidades.
    agrupado = {}

    for linha in linhas:
        produto = str(linha["produto"]).strip()
        unidade = linha["unidade"]
        chave = (produto.casefold(), unidade)

        if chave not in agrupado:
            agrupado[chave] = {
                "produto": produto,
                "unidade": unidade,
                "vezes": 0,
                "quantidade": 0
            }

        agrupado[chave]["vezes"] += 1
        agrupado[chave]["quantidade"] += float(linha["quantidade"])

    resumo_linhas = sorted(
        agrupado.values(),
        key=lambda item: (
            -item["quantidade"],
            -item["vezes"],
            item["produto"].casefold()
        )
    )

    resumo_cabecalhos = [
        "Produto",
        "Unidade",
        "Nº de vezes lançado",
        "Quantidade total"
    ]

    for coluna, titulo in enumerate(resumo_cabecalhos, 1):
        celula = resumo.cell(4, coluna, titulo)
        celula.font = Font(bold=True, color="FFFFFF")
        celula.fill = verde
        celula.alignment = Alignment(horizontal="center")

    for item in resumo_linhas:
        resumo.append([
            item["produto"],
            item["unidade"],
            item["vezes"],
            item["quantidade"]
        ])

    # Tabela real do Excel com filtro no cabeçalho.
    if resumo_linhas:
        ultima_linha = resumo.max_row

        tabela_resumo = Table(
            displayName="TabelaResumoProdutos",
            ref=f"A4:D{ultima_linha}"
        )

        estilo_resumo = TableStyleInfo(
            name="TableStyleMedium4",
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=True,
            showColumnStripes=False
        )

        tabela_resumo.tableStyleInfo = estilo_resumo
        resumo.add_table(tabela_resumo)

        for row in resumo.iter_rows(
            min_row=5,
            max_row=ultima_linha
        ):
            row[1].alignment = Alignment(horizontal="center")
            row[2].alignment = Alignment(horizontal="center")
            row[3].alignment = Alignment(horizontal="right")
            row[3].number_format = '#,##0.###'

    resumo.column_dimensions["A"].width = 34
    resumo.column_dimensions["B"].width = 14
    resumo.column_dimensions["C"].width = 22
    resumo.column_dimensions["D"].width = 20

    resumo.freeze_panes = "A5"

    # =====================================================
    # SALVA O EXCEL
    # =====================================================

    arquivo = BytesIO()
    wb.save(arquivo)
    arquivo.seek(0)

    return send_file(
        arquivo,
        as_attachment=True,
        download_name="controle_avarias.xlsx",
        mimetype=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        )
    )


if __name__ == "__main__":
    preparar_banco()
    app.run(host="0.0.0.0", port=5000, debug=False)
