import streamlit as st
import pandas as pd
import sqlite3
import math
import io
import hashlib
import os
import pathlib
import streamlit.components.v1 as components

st.set_page_config(page_title="Vésper | Tarifas", page_icon="✈️")

def load_js(path):
    try:
        js = pathlib.Path(path).read_text()
        st.html(f"<script>{js}</script>")
    except FileNotFoundError:
        pass

# ------------------------------------------------------------
# CONFIGURAÇÃO DE USUÁRIOS
# Altere as senhas aqui. Os valores são hashes SHA-256.
# Para gerar um novo hash: hashlib.sha256("sua_senha".encode()).hexdigest()
# ------------------------------------------------------------
def make_hash(password: str) -> str:
    return hashlib.sha256(password.encode()).hexdigest()

USUARIOS = {
    "operacao": {
        "senha_hash": os.getenv("SENHA_ADMIN", make_hash("operacao7474*")),
        "perfil": "adm",
        "nome": "Administrador"
    },
    "comercial": {
        "senha_hash": os.getenv("SENHA_USUARIO", make_hash("comercial2026*")),
        "perfil": "usuario",
        "nome": "Usuário"
    },
}

# ------------------------------------------------------------
# AUTENTICAÇÃO
# ------------------------------------------------------------
def verificar_login(login: str, senha: str):
    """Retorna o dict do usuário se credenciais válidas, senão None."""
    user = USUARIOS.get(login)
    if user and user["senha_hash"] == make_hash(senha):
        return user
    return None

def tela_login():
    st.markdown("<h1 style='font-size: 28px;'>Vésper | Comparador de Tarifas Aéreas</h1>", unsafe_allow_html=True)

    col_centro = st.columns([1, 1.2, 1])[1]
    with col_centro:
        login_input = st.text_input("Usuário", placeholder="Digite seu usuário", key="login_input")
        senha_input = st.text_input("Senha", type="password", placeholder="Digite sua senha", key="senha_input")
        entrar = st.button("Entrar", use_container_width=True)

        if entrar:
            user = verificar_login(login_input.strip(), senha_input)
            if user:
                st.session_state["autenticado"] = True
                st.session_state["perfil"] = user["perfil"]
                st.session_state["nome_usuario"] = user["nome"]
                st.session_state["login"] = login_input.strip()
                st.rerun()
                st.session_state["ultimo_acesso"] = time.time()
            else:
                st.error("Usuário ou senha incorretos.")

def logout():
    for key in ["autenticado", "perfil", "nome_usuario", "login"]:
        st.session_state.pop(key, None)
    st.rerun()

# ------------------------------------------------------------
# CSS
# ------------------------------------------------------------
def load_css(path):
    try:
        with open(path) as f:
            st.markdown(f"<style>{f.read()}</style>", unsafe_allow_html=True)
    except FileNotFoundError:
        pass  # CSS opcional — não quebra se não existir

load_css("style.css")

# ------------------------------------------------------------
# CONTROLE DE SESSÃO — redireciona para login se não autenticado
# ------------------------------------------------------------
import time

if "autenticado" not in st.session_state:
    st.session_state["autenticado"] = False

if st.session_state["autenticado"]:
    # Verifica timeout de 1 hora (3600 segundos)
    ultimo_acesso = st.session_state.get("ultimo_acesso", time.time())
    if time.time() - ultimo_acesso > 3600:
        for key in ["autenticado", "perfil", "nome_usuario", "login", "ultimo_acesso"]:
            st.session_state.pop(key, None)
        st.warning("Sua sessão expirou. Faça login novamente.")
        st.rerun()
    else:
        st.session_state["ultimo_acesso"] = time.time()

if not st.session_state["autenticado"]:
    tela_login()
    st.stop()

# A partir daqui o usuário está autenticado
perfil = st.session_state["perfil"]          # "adm" ou "usuario"
nome_usuario = st.session_state["nome_usuario"]

# ------------------------------------------------------------
# CABEÇALHO COM INFO DE SESSÃO
# ------------------------------------------------------------
col_titulo, col_sessao = st.columns([3, 1])
with col_titulo:
    st.markdown("<h1 style='font-size: 28px;'>Vésper | Comparador de Tarifas Aéreas</h1>", unsafe_allow_html=True)
with col_sessao:
    st.markdown("<div style='height: 20px;'></div>", unsafe_allow_html=True)
    if st.button("Sair", key="btn_logout"):
        logout()

# ------------------------------------------------------------
# BANCO DE DADOS - INICIALIZAÇÃO
# ------------------------------------------------------------
# ------------------------------------------------------------
# BANCO DE DADOS - INICIALIZAÇÃO
# ------------------------------------------------------------
def init_db(conn):
    """Garante que as tabelas base e bloqueios existam."""
    conn.execute("""
        CREATE TABLE IF NOT EXISTS base (
            cia TEXT, categoria TEXT, origem TEXT, destino TEXT,
            juncao TEXT, peso REAL, tarifa_minima REAL,
            valor_kg REAL, valor REAL, excedente REAL
        )
    """)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS bloqueios (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            praca TEXT NOT NULL, cia TEXT NOT NULL,
            categoria TEXT NOT NULL,
            tipo TEXT NOT NULL CHECK(tipo IN ('bloqueado', 'restrito')),
            novo_valor_kg REAL
        )
    """)
    conn.commit()

def get_bloqueios(conn):
    return pd.read_sql("SELECT * FROM bloqueios ORDER BY praca, cia, categoria", conn)

def add_bloqueio(conn, praca, cia, categoria, tipo, novo_valor_kg=None):
    existing = pd.read_sql("""
        SELECT id FROM bloqueios WHERE praca = ? AND cia = ? AND categoria = ?
    """, conn, params=[praca, cia, categoria])
    if not existing.empty:
        return False, "Já existe um bloqueio para essa praça/cia/categoria."
    conn.execute("""
        INSERT INTO bloqueios (praca, cia, categoria, tipo, novo_valor_kg) VALUES (?, ?, ?, ?, ?)
    """, (praca, cia, categoria, tipo, novo_valor_kg))
    conn.commit()
    return True, "Bloqueio registrado com sucesso!"

def remove_bloqueio(conn, bloqueio_id):
    conn.execute("DELETE FROM bloqueios WHERE id = ?", (bloqueio_id,))
    conn.commit()

# ------------------------------------------------------------
# EXPORT / IMPORT BANCO COMPLETO
# ------------------------------------------------------------
def exportar_banco(conn):
    df = pd.read_sql("SELECT * FROM base", conn)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for cia, grupo in df.groupby("cia"):
            grupo.to_excel(writer, sheet_name=str(cia)[:31], index=False)
    return output.getvalue()

def importar_banco_completo(conn, arquivo_bytes):
    COLUNAS_OBRIGATORIAS = {"cia", "categoria", "origem", "destino", "peso"}
    xls = pd.ExcelFile(io.BytesIO(arquivo_bytes))
    registros = []
    erros = []

    for aba in xls.sheet_names:
        try:
            df = pd.read_excel(xls, sheet_name=aba)
            colunas_presentes = set(df.columns.str.lower().str.strip())
            faltando = COLUNAS_OBRIGATORIAS - colunas_presentes
            if faltando:
                erros.append(f"Aba '{aba}': colunas obrigatórias ausentes → {', '.join(sorted(faltando))}")
                continue
            registros.append(df)
        except Exception as e:
            erros.append(f"Erro na aba '{aba}': {e}")

    # Aborta sem tocar no banco se qualquer aba apresentou erro
    if erros:
        return 0, erros

    if not registros:
        return 0, ["Nenhum dado válido encontrado no arquivo."]

    df_total = pd.concat(registros, ignore_index=True)
    conn.execute("DELETE FROM base")
    df_total.to_sql("base", conn, if_exists="append", index=False)
    conn.commit()
    return len(df_total), []

# ------------------------------------------------------------
# FUNÇÃO PRINCIPAL DE CÁLCULO
# ------------------------------------------------------------
def buscar_linha(conn, origem, destino, cia, categoria, peso=None, peso_exact=False):
    if peso is not None:
        operador = "=" if peso_exact else "<="
        query = f"""
            SELECT * FROM base
            WHERE origem = ? AND destino = ? AND cia = ? AND categoria = ?
            AND peso {operador} ?
            ORDER BY peso DESC LIMIT 1
        """
        params = [origem, destino, cia, categoria, peso]
    else:
        query = "SELECT * FROM base WHERE origem = ? AND destino = ? AND cia = ? AND categoria = ? LIMIT 1"
        params = [origem, destino, cia, categoria]
    resultado = pd.read_sql(query, conn, params=params)
    return None if resultado.empty else resultado.iloc[0]

def aplicar_tarifa_minima(resultado, tarifa_min):
    if tarifa_min is not None and not pd.isna(tarifa_min) and resultado < tarifa_min:
        return tarifa_min
    return resultado

def calcular_tarifa(origem, destino, cia, categoria, peso, conn, novo_valor_kg_override=None):
    cia = cia.upper().strip()
    categoria = categoria.upper().strip().replace(" ", "")

    if categoria == "PROXVOOLATAM":
        peso_cobrado = math.ceil(peso * 2) / 2  # arredonda para cima em 0.5kg

        if peso_cobrado <= 30:
            linha = buscar_linha(conn, origem, destino, cia, categoria, peso=peso_cobrado)
            if linha is not None and not pd.isna(linha["valor"]):
                return aplicar_tarifa_minima(linha["valor"], linha["tarifa_minima"])
        else:
        # Busca o valor fixo da faixa de 30kg
            linha_30 = buscar_linha(conn, origem, destino, cia, categoria, peso=30, peso_exact=True)
            if linha_30 is not None and not pd.isna(linha_30["valor"]) and not pd.isna(linha_30["excedente"]):
                valor_calculado = linha_30["valor"] + (peso_cobrado - 30) * linha_30["excedente"]
                return aplicar_tarifa_minima(valor_calculado, linha_30["tarifa_minima"])

        return None

    if cia == "AZUL" and categoria == "STANDARD":
        linha = buscar_linha(conn, origem, destino, cia, categoria)
        if linha is not None:
            valor_kg = novo_valor_kg_override or linha["valor_kg"]
            if not pd.isna(valor_kg):
                peso_cobrado = max(peso, 30)
                return aplicar_tarifa_minima(valor_kg * peso_cobrado, linha["tarifa_minima"])
        return None

    if cia == "GOLLOG" or (cia == "LATAM" and categoria == "ECONOMICLATAM"):
        linha = buscar_linha(conn, origem, destino, cia, categoria)
        if linha is not None:
            valor_kg = novo_valor_kg_override or linha["valor_kg"]
            if not pd.isna(valor_kg):
                return aplicar_tarifa_minima(valor_kg * peso, linha["tarifa_minima"])
        return None

    pesos = pd.read_sql("""
        SELECT peso FROM base
        WHERE origem = ? AND destino = ? AND cia = ? AND categoria = ?
        ORDER BY peso
    """, conn, params=[origem, destino, cia, categoria])

    if pesos.empty:
        return None

    peso_max = pesos["peso"].max()

    if peso > peso_max:
        linha = buscar_linha(conn, origem, destino, cia, categoria, peso=peso_max, peso_exact=True)
        if linha is not None:
            if novo_valor_kg_override is not None:
                return novo_valor_kg_override * peso
            if not pd.isna(linha["excedente"]):
                return linha["valor"] + (peso - peso_max) * linha["excedente"]
            return linha["valor"]
    else:
        linha = buscar_linha(conn, origem, destino, cia, categoria, peso=peso)
        if linha is not None:
            if novo_valor_kg_override is not None:
                return novo_valor_kg_override * peso
            return linha["valor"]

    return None

# ------------------------------------------------------------
# INTERFACE PRINCIPAL
# ------------------------------------------------------------
conn = sqlite3.connect("base.db")
init_db(conn)

try:
    # ── Campos de consulta ──────────────────────────────────
    col1, col2, col3, col4 = st.columns([1, 1, 1, 1.8])
    with col1:
        origem = st.text_input("Origem (ex: CAX)", key="origem").upper()

    with col2:
        destino = st.text_input("Destino (ex: GRU)", key="destino").upper()
        destino = st.session_state.get("destino", "").upper().strip()
    with col3:
        peso = st.number_input("Peso (kg)", min_value=0.5, step=0.5, key="peso")
        peso = st.session_state.get("peso", 1.0)
    with col4:
        st.markdown("<div style='height: 28px;'></div>", unsafe_allow_html=True)
        executar = st.button("Comparar Tarifas")

    # ── Seções exclusivas do ADM ────────────────────────────
    if perfil == "adm":

        # ── EXPANDER: GERENCIAR BLOQUEIOS ───────────────────
        with st.expander("⚙️ Gerenciar Bloqueios de Praças"):
            st.markdown("####  Adicionar bloqueio / restrição")
            col_a, col_b, col_c = st.columns(3)

            # Carrega CIAs disponíveis no banco
            cias_banco = pd.read_sql(
                "SELECT DISTINCT UPPER(TRIM(cia)) as cia FROM base ORDER BY cia", conn
            )["cia"].tolist()

            with col_a:
                b_praca = st.text_input("Praça (Destino)", key="b_praca").upper().strip()
                b_cia = st.selectbox("Companhia", options=[""] + cias_banco, key="b_cia")

            # Categorias filtradas pela CIA selecionada
            if b_cia:
                cats_banco = pd.read_sql("""
                    SELECT DISTINCT UPPER(TRIM(categoria)) as categoria
                    FROM base WHERE UPPER(TRIM(cia)) = ?
                    ORDER BY categoria
                """, conn, params=[b_cia])["categoria"].tolist()
            else:
                cats_banco = []

            with col_b:
                b_categoria = st.selectbox(
                    "Categoria",
                    options=[""] + cats_banco,
                    key="b_categoria",
                    disabled=(not b_cia),
                    help="Selecione uma Companhia primeiro." if not b_cia else ""
                )
                b_tipo = st.selectbox("Tipo de Bloqueio", ["bloqueado", "restrito"], key="b_tipo")

            with col_c:
                b_novo_valor = None
                if b_tipo == "restrito":
                    b_novo_valor = st.number_input("Novo valor_kg (R$)", min_value=0.01, step=0.01, key="b_novo_valor")
                st.markdown("<div style='height: 28px;'></div>", unsafe_allow_html=True)
                salvar_bloqueio = st.button("💾 Salvar Bloqueio")

            if salvar_bloqueio:
                if not b_praca or not b_cia or not b_categoria:
                    st.warning("Preencha todos os campos: Praça, Companhia e Categoria.")
                elif b_tipo == "restrito" and (b_novo_valor is None or b_novo_valor <= 0):
                    st.warning("Para bloqueio restrito, informe o novo valor_kg.")
                else:
                    ok, msg = add_bloqueio(conn, b_praca, b_cia, b_categoria, b_tipo, b_novo_valor if b_tipo == "restrito" else None)
                    if ok:
                        st.success(msg)
                    else:
                        st.error(msg)

            st.markdown("---")
            st.markdown("#### Bloqueios ativos")
            bloqueios_df = get_bloqueios(conn)

            if bloqueios_df.empty:
                st.info("Nenhuma praça bloqueada no momento.")
            else:
                for _, row in bloqueios_df.iterrows():
                    col_info, col_btn = st.columns([5, 1])
                    with col_info:
                        tipo_badge = "🚫 Bloqueado" if row["tipo"] == "bloqueado" else f"⚠️ Restrito (R$ {row['novo_valor_kg']:.2f}/kg)"
                        st.markdown(
                            f"**Praça:** {row['praca']} &nbsp;|&nbsp; "
                            f"**Cia:** {row['cia']} &nbsp;|&nbsp; "
                            f"**Categoria:** {row['categoria']} &nbsp;|&nbsp; "
                            f"{tipo_badge}"
                        )
                    with col_btn:
                        if st.button("Remover", key=f"del_{row['id']}"):
                            remove_bloqueio(conn, row["id"])
                            st.rerun()

        # ── EXPANDER: ATUALIZAÇÃO DE TARIFAS ────────────────
        with st.expander("📦 Base de dados"):

            # ── EXPORTAR ────────────────────────────────────────
            st.markdown("#### 📤 Exportar Banco de Dados")
            st.caption("Exporta todos os registros para Excel, com uma aba por cia.")

            col_exp1, col_exp2 = st.columns([2, 1])
            with col_exp1:
                senha_export = st.text_input(
                    "Senha do administrador", type="password",
                    key="senha_export", placeholder="Digite a senha para exportar"
                )
            with col_exp2:
                st.markdown("<div style='height: 28px;'></div>", unsafe_allow_html=True)
                exportar_btn = st.button("📤 Exportar Banco", key="btn_exportar")

            if exportar_btn:
                if make_hash(senha_export) != USUARIOS["operacao"]["senha_hash"]:
                    st.error("❌ Senha incorreta. Exportação não autorizada.")
                else:
                    with st.spinner("Gerando arquivo..."):
                        excel_bytes = exportar_banco(conn)
                    st.download_button(
                        label="⬇️ Baixar Excel",
                        data=excel_bytes,
                        file_name="base_tarifas.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="download_banco"
                    )

            st.markdown("---")

            # ── IMPORTAR ────────────────────────────────────────
            st.markdown("#### 📥 Importar Banco Atualizado")
            st.caption("⚠️ Isso substituirá **todo** o banco atual. Use somente com o arquivo exportado por este app.")

            arquivo_banco = st.file_uploader(
                "Selecione o arquivo Excel exportado",
                type=["xlsx"], key="upload_banco"
            )

            if arquivo_banco is not None:
                st.info(f"Arquivo carregado: **{arquivo_banco.name}**")
                col_imp1, col_imp2 = st.columns([2, 1])
                with col_imp1:
                    senha_import = st.text_input(
                        "Senha do administrador", type="password",
                        key="senha_import", placeholder="Digite a senha para importar"
                    )
                with col_imp2:
                    st.markdown("<div style='height: 28px;'></div>", unsafe_allow_html=True)
                    importar_btn = st.button("Confirmar Importação", key="confirmar_import_banco")

                if importar_btn:
                    if make_hash(senha_import) != USUARIOS["operacao"]["senha_hash"]:
                        st.error("❌ Senha incorreta. Importação não autorizada.")
                    else:
                        st.warning("⚠️ Esta ação substituirá todos os dados do banco.")
                        with st.spinner("Importando..."):
                            linhas, erros = importar_banco_completo(conn, arquivo_banco.read())
                        if erros:
                            for e in erros:
                                st.error(e)
                        else:
                            st.success(f"✅ Banco atualizado com sucesso! {linhas:,} registros importados.")

    # ── COMPARAÇÃO DE TARIFAS (disponível para todos) ───────
    if executar:
        if not origem or not destino:
            st.warning("Preencha os campos de Origem e Destino antes de comparar.")
        else:
            opcoes = pd.read_sql("""
                SELECT DISTINCT cia, categoria FROM base
                WHERE origem = ? AND destino = ? ORDER BY cia, categoria
            """, conn, params=[origem, destino])
            opcoes["cia"] = opcoes["cia"].str.upper().str.strip()
            opcoes["categoria"] = opcoes["categoria"].str.upper().str.strip()

            resultados = []
            bloqueios_destino = pd.read_sql("SELECT * FROM bloqueios WHERE praca = ?", conn, params=[destino])
            bloqueios_dict = {(r["cia"], r["categoria"]): r for _, r in bloqueios_destino.iterrows()}

            for _, row in opcoes.iterrows():
                cia_row, cat_row = row["cia"], row["categoria"]
                bloqueio = bloqueios_dict.get((cia_row, cat_row))

                if bloqueio is not None:
                    if bloqueio["tipo"] == "bloqueado":
                        resultados.append({"Companhia": cia_row, "Categoria": cat_row, "Valor (R$)": None, "Status": "🚫 Bloqueado"})
                        continue
                    elif bloqueio["tipo"] == "restrito":
                        valor = calcular_tarifa(origem, destino, cia_row, cat_row, peso, conn, novo_valor_kg_override=bloqueio["novo_valor_kg"])
                        resultados.append({
                            "Companhia": cia_row,
                            "Categoria": cat_row,
                            "Valor (R$)": round(valor, 2) if valor is not None else None,
                            "Status": "⚠️ Restrito"
                        })
                        continue

                valor = calcular_tarifa(origem, destino, cia_row, cat_row, peso, conn)
                if valor is not None:
                    resultados.append({"Companhia": cia_row, "Categoria": cat_row, "Valor (R$)": round(valor, 2), "Status": "✅ Disponível"})

            if resultados:
                df = pd.DataFrame(resultados)
                df_disp = df[df["Status"] != "🚫 Bloqueado"].sort_values("Valor (R$)")
                df_bloq = df[df["Status"] == "🚫 Bloqueado"]
                df = pd.concat([df_disp, df_bloq], ignore_index=True)

                valor_min = df_disp["Valor (R$)"].min() if not df_disp.empty else None

                def highlight_row(row):
                    if row["Status"] == "🚫 Bloqueado":
                        return ['background-color: #ffe0e0; color: #cc0000;'] * len(row)
                    elif row["Status"] == "⚠️ Restrito":
                        return ['background-color: #fff8e1; color: #996600;'] * len(row)
                    elif row["Status"] == "✅ Disponível" and valor_min is not None and row["Valor (R$)"] == valor_min:
                        return ['background-color: #7ED957; color: black;'] * len(row)
                    return [''] * len(row)

                styled = (
                    df.style
                    .apply(highlight_row, axis=1)
                    .format({"Valor (R$)": lambda x: f"{x:,.2f}" if pd.notna(x) else "—"})
                )
                st.success("Comparação concluída!")
                st.dataframe(styled)
            else:
                st.warning("Nenhuma tarifa encontrada para essa rota.")

finally:
    conn.close()
