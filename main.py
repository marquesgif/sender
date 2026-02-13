import streamlit as st
import pandas as pd
import zipfile
import smtplib
from email.message import EmailMessage
import time
from datetime import datetime
import os

# Configurações da página
st.set_page_config(page_title="Certificador ISPCAALA", page_icon="🎓", layout="wide")

# --- SISTEMA DE LOGS ---
LOG_DIR = "logs"
os.makedirs(LOG_DIR, exist_ok=True)
LOG_FILE = os.path.join(LOG_DIR, f"{datetime.now().strftime('%d-%m-%Y_%H-%M-%S')}_envio.txt")

def escrever_log(mensagem):
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(f"[{datetime.now().strftime('%d/%m/%Y %H:%M:%S')}] {mensagem}\n")

# --- LOGIN ---
if 'autenticado' not in st.session_state:
    st.session_state.autenticado = False

def validar_conexao(email, senha):
    try:
        with smtplib.SMTP_SSL("smtp.titan.email", 465, timeout=10) as server:
            server.login(email, senha)
            return True
    except:
        return False

# --- INTERFACE ---
if not st.session_state.autenticado:
    st.title("🔐 Acesso ao Sistema")
    email_user = st.text_input("E-mail (@ispcaala.com)")
    senha_user = st.text_input("Senha", type="password")
    
    if st.button("Entrar"):
        if validar_conexao(email_user, senha_user):
            st.session_state.autenticado = True
            st.session_state.email = email_user
            st.session_state.senha = senha_user
            st.rerun()
        else:
            st.error("Falha na autenticação.")

else:
    st.title("📤 Envio de Certificados (Pausa Inteligente)")
    st.sidebar.success(f"Conectado: {st.session_state.email}")

    if st.sidebar.button("Sair"):
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
                
                resumo_data = []
                total = len(df)
                envios_no_bloco = 0 
                
                with zipfile.ZipFile(uploaded_zip, 'r') as z:
                    arquivos_no_zip = z.namelist()
                    progresso = st.progress(0)
                    status_text = st.empty()
                    timer_text = st.empty()

                    def conectar_smtp():
                        srv = smtplib.SMTP_SSL("smtp.titan.email", 465)
                        srv.login(st.session_state.email, st.session_state.senha)
                        return srv

                    server = conectar_smtp()
                    i = 0  # Usaremos um controle manual do índice para permitir re-tentativa

                    while i < total:
                        linha = df.iloc[i]
                        nome = linha.get('nome', 'N/A')
                        email_dest = linha.get('e-mail', linha.get('email', 'N/A'))
                        nome_pdf = linha.get('arquivo', 'N/A')

                        # 1. Verificação Preventiva de Cota
                        if envios_no_bloco >= 48:
                            server.quit()
                            for m in range(80, 0, -1):
                                timer_text.warning(f"☕ Limite de 50 atingido. Pausando por {m} min...")
                                time.sleep(60)
                            timer_text.empty()
                            server = conectar_smtp()
                            envios_no_bloco = 0

                        try:
                            if nome_pdf in arquivos_no_zip:
                                with z.open(nome_pdf) as f:
                                    pdf_data = f.read()

                                msg = EmailMessage()
                                msg['Subject'] = f"Certificado - {nome}"
                                msg['From'] = st.session_state.email
                                msg['To'] = email_dest
                                msg.set_content(f"Olá {nome},\n\nSegue o certificado em anexo.")
                                msg.add_attachment(pdf_data, maintype='application', subtype='pdf', filename=nome_pdf)

                                server.send_message(msg)
                                envios_no_bloco += 1
                                status_text.success(f"✅ {i+1}/{total} Enviado: {nome}")
                                resumo_data.append({"Nome": nome, "E-mail": email_dest, "Status": "Sucesso", "Info": "Enviado"})
                                
                                # SÓ AVANÇA PARA O PRÓXIMO SE ENVIAR COM SUCESSO
                                i += 1 
                            else:
                                status_text.warning(f"⚠️ PDF ausente para {nome}")
                                resumo_data.append({"Nome": nome, "E-mail": email_dest, "Status": "Falha", "Info": "PDF ausente no ZIP"})
                                i += 1 # Avança pois o erro não é do servidor, mas do arquivo
                        
                        except Exception as e:
                            erro_str = str(e)
                            # 2. Verificação de Erro em Tempo Real (Se o erro aparecer antes do contador)
                            if "Quota Exceeded" in erro_str or "550" in erro_str or "too many errors" in erro_str.lower():
                                status_text.error(f"🚨 O servidor bloqueou o envio agora! Iniciando pausa forçada...")
                                server.quit()
                                for m in range(82, 0, -1): # Pausa um pouco maior por segurança
                                    timer_text.info(f"🔄 Servidor saturado. Retentando o mesmo e-mail em {m} min...")
                                    time.sleep(60)
                                timer_text.empty()
                                server = conectar_smtp()
                                envios_no_bloco = 0
                                # NÃO incrementamos o 'i', então ele tentará a mesma linha novamente
                            else:
                                status_text.error(f"❌ Erro fatal no e-mail de {nome}")
                                resumo_data.append({"Nome": nome, "E-mail": email_dest, "Status": "Falha", "Info": erro_str})
                                i += 1 # Avança se for erro de e-mail inválido, por exemplo

                        progresso.progress(i / total)
                        time.sleep(1.5)

                    server.quit()

                # --- DASHBOARD FINAL ---
                st.divider()
                st.subheader("📊 Resumo da Operação")
                resumo_df = pd.DataFrame(resumo_data)
                
                m1, m2, m3 = st.columns(3)
                m1.metric("Total", len(resumo_df))
                m2.metric("Sucessos ✅", len(resumo_df[resumo_df['Status'] == "Sucesso"]))
                m3.metric("Falhas ❌", len(resumo_df[resumo_df['Status'] == "Falha"]))
                st.dataframe(resumo_df, use_container_width=True)
                
                output_file = "relatorio_final.xlsx"
                resumo_df.to_excel(output_file, index=False)
                with open(output_file, "rb") as f:
                    st.download_button("📥 Baixar Relatório (Excel)", f, file_name=output_file)
                st.balloons()

            except Exception as e_geral:
                st.error(f"Erro crítico: {e_geral}")
        else:
            st.warning("Selecione os arquivos.")