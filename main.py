import streamlit as st
import pandas as pd
import zipfile
import smtplib
from email.message import EmailMessage
import time
from datetime import datetime
import os

# Configurações da página
st.set_page_config(page_title="Certificador ISPCAALA", page_icon="🎓")

# Criar diretório logs caso não exista
LOG_DIR = "logs"
os.makedirs(LOG_DIR, exist_ok=True)

# Criar nome do arquivo log baseado em data/hora
LOG_FILE = os.path.join(
    LOG_DIR,
    f"{datetime.now().strftime('%d-%m-%Y_%H-%M-%S')}_envio_certificados.txt"
)

def escrever_log(mensagem):
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(f"[{datetime.now().strftime('%d/%m/%Y %H:%M:%S')}] {mensagem}\n")


# --- CONTROLE DE SESSÃO (LOGIN) ---
if 'autenticado' not in st.session_state:
    st.session_state.autenticado = False

def validar_conexao(email, senha):
    try:
        with smtplib.SMTP_SSL("smtp.titan.email", 465) as server:
            server.login(email, senha)
            return True
    except:
        return False


# --- INTERFACE ---
if not st.session_state.autenticado:
    st.title("🔐 Acesso ao Sistema")
    st.info("Insira suas credenciais do e-mail institucional (Titan/Hostinger).")
    
    email_user = st.text_input("E-mail (@ispcaala.com)")
    senha_user = st.text_input("Senha", type="password")
    
    if st.button("Entrar"):
        if validar_conexao(email_user, senha_user):
            st.session_state.autenticado = True
            st.session_state.email = email_user
            st.session_state.senha = senha_user

            escrever_log(f"LOGIN REALIZADO COM SUCESSO: {email_user}")

            st.rerun()
        else:
            escrever_log(f"FALHA NO LOGIN: {email_user}")
            st.error("Falha na autenticação. Verifique seu e-mail e senha.")


else:
    # --- TELA DE UPLOAD E ENVIO ---
    st.title("📤 Envio de Certificados")
    st.sidebar.success(f"Conectado como: {st.session_state.email}")

    if st.sidebar.button("Sair/Trocar Conta"):
        escrever_log(f"LOGOUT: {st.session_state.email}")
        st.session_state.autenticado = False
        st.rerun()

    col1, col2 = st.columns(2)
    with col1:
        uploaded_excel = st.file_uploader("1. Planilha Excel", type=["xlsx"])
    with col2:
        uploaded_zip = st.file_uploader("2. Arquivos ZIP", type=["zip"])

    if st.button("🚀 Iniciar Envio em Massa"):
        if uploaded_excel and uploaded_zip:
            try:
                df = pd.read_excel(uploaded_excel)
                df.columns = df.columns.str.strip().str.lower()

                escrever_log("========================================")
                escrever_log(f"INÍCIO DO ENVIO EM MASSA - Total: {len(df)}")
                escrever_log(f"Usuário logado: {st.session_state.email}")
                escrever_log("========================================")

                with zipfile.ZipFile(uploaded_zip, 'r') as z:
                    arquivos_no_zip = z.namelist()

                    progresso = st.progress(0)
                    status_text = st.empty()

                    with smtplib.SMTP_SSL("smtp.titan.email", 465) as server:
                        server.login(st.session_state.email, st.session_state.senha)

                        total = len(df)

                        for i, linha in df.iterrows():
                            nome = linha['nome']
                            email_dest = linha['e-mail']
                            nome_pdf = linha['arquivo']

                            try:
                                if nome_pdf in arquivos_no_zip:
                                    with z.open(nome_pdf) as f:
                                        pdf_data = f.read()

                                    msg = EmailMessage()
                                    msg['Subject'] = f"Certificado - {nome}"
                                    msg['From'] = st.session_state.email
                                    msg['To'] = email_dest
                                    msg.set_content(
                                        f"Olá {nome},\n\nSegue em anexo o seu certificado.\n\nAtenciosamente."
                                    )

                                    msg.add_attachment(
                                        pdf_data,
                                        maintype='application',
                                        subtype='pdf',
                                        filename=nome_pdf
                                    )

                                    server.send_message(msg)

                                    status_text.text(f"✅ Enviado: {email_dest}")
                                    escrever_log(f"ENVIADO: {email_dest} | Nome: {nome} | Arquivo: {nome_pdf}")

                                else:
                                    st.warning(f"⚠️ {nome_pdf} não encontrado no ZIP.")
                                    escrever_log(f"ARQUIVO NÃO ENCONTRADO: {nome_pdf} | Destino: {email_dest} | Nome: {nome}")

                            except Exception as erro_envio:
                                st.error(f"Erro ao enviar para {email_dest}: {erro_envio}")
                                escrever_log(f"ERRO AO ENVIAR: {email_dest} | Arquivo: {nome_pdf} | Erro: {erro_envio}")

                            progresso.progress((i + 1) / total)
                            time.sleep(1)  # Delay de segurança Titan

                escrever_log("========================================")
                escrever_log("FIM DO ENVIO EM MASSA")
                escrever_log("========================================")

                st.balloons()
                st.success("✅ Todos os e-mails foram processados!")

                # Botão pra baixar log no Streamlit
                with open(LOG_FILE, "r", encoding="utf-8") as f:
                    st.download_button("📄 Baixar Log (.txt)", f, file_name=os.path.basename(LOG_FILE))

            except Exception as e:
                st.error(f"Ocorreu um erro: {e}")
                escrever_log(f"ERRO GERAL: {e}")

        else:
            st.warning("Selecione os dois arquivos antes de continuar.")
