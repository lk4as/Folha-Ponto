import streamlit as st
from docx import Document
from docx.shared import Pt, Inches, Cm
from docx.oxml.ns import qn
from datetime import date, datetime, timedelta
import io
import os
from PIL import Image

# --- CONFIGURAÇÕES VISUAIS ---
st.set_page_config(page_title="Gerador de Folha de Ponto", layout="centered")

# --- INSERÇÃO DA LOGOMARCA (Baseado no seu exemplo) ---
caminho_logo = "Logo tradicional.png"

try:
    # Verifica se o arquivo existe antes de tentar abrir
    if os.path.exists(caminho_logo):
        imagem = Image.open(caminho_logo)
        # Utilizando colunas para centralizar a imagem
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            st.image(imagem, use_container_width=True)
    else:
        st.warning(f"Arquivo '{caminho_logo}' não encontrado. O sistema funcionará sem a logo.")
except Exception as e:
    st.error(f"Erro ao carregar a logo: {e}")

st.markdown("<h1 style='text-align: center;'>Gerador de Folha de Ponto</h1>", unsafe_allow_html=True)
st.markdown("<p style='text-align: center;'>Bram Offshore | Uma empresa do grupo Chouest</p>", unsafe_allow_html=True)
st.markdown("---")

# --- ARQUIVOS PADRÃO ---
ARQUIVO_MODELO = "Folha de Ponto.docx"
NOME_DA_FONTE = 'Arial'

# Símbolos
CHECKBOX_MARCADO = "\u2612"
CHECKBOX_VAZIO = "\u2610"

# --- FUNÇÕES UTILITÁRIAS ---
def configurar_fonte(run, tamanho=9):
    run.font.name = NOME_DA_FONTE
    run.font.size = Pt(tamanho)
    r = run._element
    r.rPr.rFonts.set(qn('w:eastAsia'), NOME_DA_FONTE)

def limpar_e_escrever(celula, texto, tamanho=9, alinhamento=1):
    celula._element.clear_content()
    p = celula.add_paragraph()
    p.alignment = alinhamento
    run = p.add_run(str(texto))
    configurar_fonte(run, tamanho)

def inserir_assinatura(celula, imagem_obj):
    """
    Insere a imagem na célula. Aceita objeto de memória (upload).
    """
    celula._element.clear_content()
    p = celula.add_paragraph()
    p.alignment = 1
    run = p.add_run()
    try:
        # Rebobina o arquivo para leitura múltipla
        if hasattr(imagem_obj, 'seek'):
            imagem_obj.seek(0)
        run.add_picture(imagem_obj, width=Inches(0.9))
    except Exception as e:
         run.add_text("[Assinatura]")
         configurar_fonte(run, 8)

def iterar_todas_as_tabelas(doc_obj):
    for table in doc_obj.tables:
        yield table
        for row in table.rows:
            for cell in row.cells:
                if cell.tables:
                    yield from iterar_todas_as_tabelas(cell)

def preencher_campo_seguro(doc, rotulo_busca, valor_preencher, tamanho_fonte=10):
    rotulo_limpo = rotulo_busca.lower()
    for table in iterar_todas_as_tabelas(doc):
        for row in table.rows:
            for i, cell in enumerate(row.cells):
                if rotulo_limpo in cell.text.lower():
                    for j in range(i + 1, len(row.cells)):
                        cand_cell = row.cells[j]
                        if cand_cell != cell and rotulo_limpo not in cand_cell.text.lower():
                            limpar_e_escrever(cand_cell, valor_preencher, tamanho=tamanho_fonte, alinhamento=0)
                            return True
    return False

def marcar_base_preservando_linhas(doc, texto_base_alvo):
    for table in iterar_todas_as_tabelas(doc):
        for row in table.rows:
            for cell in row.cells:
                if "(Rio)" in cell.text and "CNPJ" not in cell.text: 
                    linhas_originais = [p.text for p in cell.paragraphs if p.text.strip()]
                    cell._element.clear_content()
                    for linha in linhas_originais:
                        p = cell.add_paragraph()
                        p.alignment = 0 
                        p.paragraph_format.space_after = Pt(0)
                        texto_limpo = linha.replace(CHECKBOX_MARCADO, "").replace(CHECKBOX_VAZIO, "").replace("☒", "").replace("☐", "").strip()
                        if texto_base_alvo in linha:
                            run = p.add_run(f"{CHECKBOX_MARCADO} {texto_limpo}")
                            configurar_fonte(run, 8)
                        else:
                            run = p.add_run(f"{CHECKBOX_VAZIO} {texto_limpo}")
                            configurar_fonte(run, 8)
                    return True
    return False

def sanitarizar_hora(hora_input):
    if not hora_input: return "00:00"
    h = str(hora_input).strip()
    if ":" not in h:
        try:
            return f"{int(h):02d}:00"
        except ValueError:
            return h
    return h

def calcular_quarteto(hora_entrada, hora_almoco_ida):
    fmt = "%H:%M"
    hora_entrada = sanitarizar_hora(hora_entrada)
    hora_almoco_ida = sanitarizar_hora(hora_almoco_ida)
    
    try:
        dt_ent = datetime.strptime(hora_entrada, fmt)
        dt_alm_ida = datetime.strptime(hora_almoco_ida, fmt)
        dt_alm_volta = dt_alm_ida + timedelta(hours=1)
        dt_saida = dt_ent + timedelta(hours=7)
        return [hora_entrada, hora_almoco_ida, dt_alm_volta.strftime(fmt), dt_saida.strftime(fmt)]
    except ValueError:
        return ["--:--", "--:--", "--:--", "--:--"]

def forcar_uma_pagina(doc):
    for section in doc.sections:
        section.bottom_margin = Cm(0.5)
        section.footer_distance = Cm(0.5)
    if doc.paragraphs:
        p = doc.paragraphs[-1]
        p.paragraph_format.space_after = Pt(0)
        for run in p.runs: run.font.size = Pt(1)

# --- FORMULÁRIO PRINCIPAL ---

st.subheader("Dados Cadastrais")

col1, col2 = st.columns(2)
with col1:
    mes_final = st.selectbox("Mês de Referência", options=list(range(1, 13)), index=11, format_func=lambda x: f"{x:02d}")
with col2:
    ano_final = st.number_input("Ano", value=2025, step=1)

col3, col4 = st.columns(2)
with col3:
    nome_func = st.text_input("Nome do Funcionário", "Lukas Souza Henriques Crespo")
with col4:
    cargo_func = st.text_input("Função / Cargo", "Estagiário - Dep. DP Assurance")

col5, col6 = st.columns(2)
with col5:
    base_opt = st.selectbox("Base Operacional", ["Rio", "Açu", "Macaé", "Guaxindiba"], index=1)
    mapa_bases = {"Rio": "(Rio)", "Açu": "(Açu)", "Macaé": "(Macaé)", "Guaxindiba": "(Guax.)"}
    texto_base = mapa_bases[base_opt]
with col6:
    data_emissao = st.text_input("Data de Emissão (CTPS)", "04/02/2020")

# --- SEÇÃO DE ASSINATURA INTEGRADA ---
st.markdown("---")
st.subheader("Assinatura Digital")
st.markdown("Faça o upload da imagem da sua assinatura (formato PNG ou JPG) para ser inserida no documento.")
assinatura_upload = st.file_uploader("Selecionar arquivo de assinatura", type=["png", "jpg", "jpeg"])

if assinatura_upload:
    st.success("Assinatura carregada com sucesso.")
else:
    st.info("Nenhuma assinatura carregada. O campo será preenchido apenas com texto.")

# --- SEÇÃO DE HORÁRIOS ---
st.markdown("---")
st.subheader("Configuração de Horários (Jornada 7h)")
st.caption("Dica: Digite apenas '8' para 08:00 ou '12' para 12:00.")

tipo_horario = st.radio("Tipo de Escala", ["Fixo (Mesmo horário todos os dias)", "Variável (Horário muda durante a semana)"], horizontal=True)

escala_semanal = {}
txt_horario_cabecalho = ""

if tipo_horario.startswith("Fixo"):
    c1, c2 = st.columns(2)
    ent = c1.text_input("Horário de Entrada", "08:00")
    alm = c2.text_input("Saída para Almoço", "12:00")
    
    quarteto = calcular_quarteto(ent, alm)
    st.info(f"Escala calculada: {quarteto[0]} - {quarteto[3]} (Intervalo: {quarteto[1]} às {quarteto[2]})")
    
    for d in range(5): escala_semanal[d] = quarteto
    txt_horario_cabecalho = f"{quarteto[0]} - {quarteto[3]}"

else:
    st.write("Defina o horário de entrada para cada dia da semana:")
    cols = st.columns(5)
    dias = ["Seg", "Ter", "Qua", "Qui", "Sex"]
    entradas = []
    
    for i, dia in enumerate(dias):
        val = cols[i].text_input(f"{dia}", "08:00")
        entradas.append(val)
        
    alm_var = st.text_input("Horário de Almoço (Padrão para a semana)", "12:00")
    
    for i, ent_dia in enumerate(entradas):
        escala_semanal[i] = calcular_quarteto(ent_dia, alm_var)
    
    txt_horario_cabecalho = "Variável (7h)"

st.markdown("---")
st.subheader("Registro de Feriados")
feriados_str = st.text_input("Dias de feriado no mês (separe os dias por vírgula, ex: 15, 25)", "")
feriados = []
if feriados_str:
    try:
        feriados = [int(x.strip()) for x in feriados_str.split(",") if x.strip().isdigit()]
    except:
        st.error("Formato inválido. Use apenas números separados por vírgula.")

st.markdown("---")

# --- PROCESSAMENTO ---
if st.button("Gerar Documento", type="primary", use_container_width=True):
    try:
        doc = Document(ARQUIVO_MODELO)
        
        # 1. Preenchimento de Cabeçalho
        preencher_campo_seguro(doc, "Funcionário", nome_func)
        preencher_campo_seguro(doc, "Função", cargo_func)
        preencher_campo_seguro(doc, "Emissão", data_emissao, tamanho_fonte=9)
        preencher_campo_seguro(doc, "Horário", txt_horario_cabecalho, tamanho_fonte=9)
        
        # 2. Marcar Mês
        mapa_meses = {
            1: ["dez", "jan"], 2: ["jan", "fev"], 3: ["fev", "mar"], 4: ["mar", "abr"],
            5: ["abr", "mai"], 6: ["mai", "jun"], 7: ["jun", "jul"], 8: ["jul", "ago"],
            9: ["ago", "set"], 10: ["set", "out"], 11: ["out", "nov"], 12: ["nov", "dez"]
        }
        mapa_texto = {
             1: "Dez/Jan", 2: "Jan/Fev", 3: "Fev/Mar", 4: "Mar/Abr", 5: "Abr/Mai", 6: "Mai/Jun",
             7: "Jun/Jul", 8: "Jul/Ago", 9: "Ago/Set", 10: "Set/Out", 11: "Out/Nov", 12: "Nov/Dez"
        }
        
        if mes_final in mapa_meses:
            termos = mapa_meses[mes_final]
            txt_mes = mapa_texto[mes_final]
            
            for table in iterar_todas_as_tabelas(doc):
                for row in table.rows:
                    for cell in row.cells:
                        if all(t in cell.text.lower() for t in termos):
                            limpar_e_escrever(cell, f"{CHECKBOX_MARCADO} {txt_mes}", tamanho=9, alinhamento=0)
        
        # 3. Marcar Base
        marcar_base_preservando_linhas(doc, texto_base)
        
        # 4. Preencher Tabela de Dias
        if mes_final == 1: mi, ai = 12, ano_final - 1
        else: mi, ai = mes_final - 1, ano_final
        
        C_DIA, C_INI, C_ALM_IDA, C_ALM_VOL, C_FIM, C_ASS = 0, 2, 3, 4, 6, 7
        
        for table in doc.tables:
            for row in table.rows:
                try:
                    txt_dia = row.cells[C_DIA].text.strip()
                    if not txt_dia.isdigit(): continue
                    dia = int(txt_dia)
                    
                    mc, ac = (mi, ai) if dia > 15 else (mes_final, ano_final)
                    try: dt = date(ac, mc, dia)
                    except: continue
                    
                    wd = dt.weekday()
                    cells_alvo = [C_INI, C_ALM_IDA, C_ALM_VOL, C_FIM]
                    if len(row.cells) <= C_FIM: cells_alvo = [1,2,3,4]

                    # Lógica de preenchimento (Feriado > FDS > Útil)
                    if dia in feriados:
                        for i in cells_alvo: 
                            if i < len(row.cells): limpar_e_escrever(row.cells[i], "FERIADO")
                        if C_ASS < len(row.cells): limpar_e_escrever(row.cells[C_ASS], "")
                    elif wd == 5:
                        for i in cells_alvo: 
                            if i < len(row.cells): limpar_e_escrever(row.cells[i], "SÁBADO")
                        if C_ASS < len(row.cells): limpar_e_escrever(row.cells[C_ASS], "")
                    elif wd == 6:
                        for i in cells_alvo: 
                            if i < len(row.cells): limpar_e_escrever(row.cells[i], "DOMINGO")
                        if C_ASS < len(row.cells): limpar_e_escrever(row.cells[C_ASS], "")
                    else:
                        horarios = escala_semanal.get(wd)
                        if horarios:
                            for k, idx in enumerate(cells_alvo):
                                if idx < len(row.cells): limpar_e_escrever(row.cells[idx], horarios[k])
                            
                            # INSERÇÃO DA ASSINATURA (SE HOUVER UPLOAD)
                            if C_ASS < len(row.cells):
                                if assinatura_upload:
                                    inserir_assinatura(row.cells[C_ASS], assinatura_upload)
                                else:
                                    limpar_e_escrever(row.cells[C_ASS], "[Assinatura]")
                                
                except IndexError: pass
        
        forcar_uma_pagina(doc)
        
        # SALVAR EM MEMÓRIA
        buffer = io.BytesIO()
        doc.save(buffer)
        buffer.seek(0)
        
        st.success("Documento gerado com sucesso!")
        st.download_button(
            label="Baixar Folha de Ponto",
            data=buffer,
            file_name=f"Folha_Ponto_{mes_final}_{ano_final}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True
        )
        
    except Exception as e:
        st.error(f"Erro ao gerar documento: {e}")
