import os
from datetime import datetime

from selenium import webdriver
from selenium.webdriver.edge.options import Options as EdgeOptions
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException


def abrir_outlook():
    os.system(
        r'start msedge.exe "https://outlook.office.com/mail/" '
        r'--remote-debugging-port=9222 '
        r'--user-data-dir="C:\edge-temp" --start-fullscreen'
    )

    edge_options = EdgeOptions()
    edge_options.debugger_address = "127.0.0.1:9222"

    return webdriver.Edge(options=edge_options)


def preencher_destinatarios(driver, wait):
    print("➡️ Preenchendo campo 'Para'")
    campo_para = wait.until(
        EC.element_to_be_clickable((By.XPATH, "//div[@aria-label='Para' and @role='textbox']"))
    )
    emails_para = os.getenv("EMAILS_TO", "").split(",")
    if emails_para and emails_para[0]:
        campo_para.click()
        campo_para.send_keys(",".join(emails_para))

    print("➡️ Preenchendo campo 'Cc'")
    campo_cc = wait.until(
        EC.element_to_be_clickable((By.XPATH, "//div[@aria-label='Cc' and @role='textbox']"))
    )
    emails_cc = os.getenv("EMAILS_CC", "").split(",")
    if emails_cc and emails_cc[0]:
        campo_cc.click()
        campo_cc.send_keys(",".join(emails_cc))


def preencher_assunto(driver, wait):
    hoje = datetime.now()
    assunto = f"XXXX | PANORAMA | {hoje.day:02d}.{hoje.month:02d}"
    print(f"➡️ Preenchendo campo 'Assunto' ({assunto})")
    campo_assunto = wait.until(
        EC.presence_of_element_located((By.XPATH, "//input[@aria-label='Assunto']"))
    )
    campo_assunto.click()
    campo_assunto.send_keys(assunto)


def preencher_corpo(driver, wait):
    print("➡️ Preenchendo corpo do e-mail")
    corpo = wait.until(
        EC.presence_of_element_located(
            (By.XPATH, "//div[@aria-label='Corpo da mensagem' and @role='textbox']")
        )
    )
    driver.execute_script("arguments[0].innerHTML = arguments[1];", corpo, mensagem_panorama())


def anexar_arquivo(driver, wait):
    print("📎 Clicando em 'Anexar arquivo'")
    botao_anexar = wait.until(
        EC.element_to_be_clickable((By.XPATH, "//button[@aria-label='Anexar arquivo']"))
    )
    botao_anexar.click()

    print("🗂️ Clicando em 'Navegar neste computador'")
    botao_navegar = wait.until(
        EC.element_to_be_clickable(
            (By.XPATH, "//span[contains(text(), 'Navegar neste computador')]")
        )
    )
    botao_navegar.click()

    pasta_downloads = os.path.join(os.path.expanduser("~"), "Downloads")
    arquivo_para_anexar = arquivo_mais_recente(pasta_downloads)

    if arquivo_para_anexar:
        print(f"✅ Arquivo anexado: {arquivo_para_anexar}")
        # escreve no input de seleção de arquivo
        # Selenium interage com <input type="file">
        input_file = wait.until(EC.presence_of_element_located((By.XPATH, "//input[@type='file']")))
        input_file.send_keys(arquivo_para_anexar)
    else:
        print("⚠️ Nenhum arquivo encontrado na pasta de Downloads.")


def enviar_email(driver, wait):
    print("📤 Enviando e-mail")
    botao_enviar = wait.until(
        EC.element_to_be_clickable((By.XPATH, "//button[@title='Enviar' or @aria-label='Enviar']"))
    )
    botao_enviar.click()
    print("✅ E-mail enviado com sucesso!")


def mandar_email():
    driver = abrir_outlook()
    wait = WebDriverWait(driver, 30)

    try:
        print("➡️ Clicando em 'Novo email'")
        novo_email = wait.until(
            EC.element_to_be_clickable((By.XPATH, "//button[contains(@aria-label, 'Novo email')]"))
        )
        novo_email.click()

        preencher_destinatarios(driver, wait)
        preencher_assunto(driver, wait)
        preencher_corpo(driver, wait)
        anexar_arquivo(driver, wait)
        enviar_email(driver, wait)

    except TimeoutException:
        print("❌ Algum elemento não foi encontrado no tempo esperado.")
    finally:
        print("🛑 Finalizando script.")


def mensagem_panorama():
    """
    Gera uma mensagem resumida para panorama,
    formatada com título e informações principais.
    """
    titulo = "📊 Panorama Diário"
    corpo = (
        "Aqui está o resumo atualizado das métricas:\n\n"
        "Chamadas atendidas\n"
        "Chamadas abandonadas\n"
        "Tempo médio de atendimento\n"
        "Nível de serviço"
    )
    rodape = "🔄 Atualização automática - Sistema Monitoramento"

    return f"{titulo}\n\n{corpo}\n\n{rodape}"


def arquivo_mais_recente(pasta_downloads):
    arquivos = [os.path.join(pasta_downloads, f) for f in os.listdir(pasta_downloads)]
    arquivos = [f for f in arquivos if os.path.isfile(f)]
    if not arquivos:
        return None
    return max(arquivos, key=os.path.getctime)
