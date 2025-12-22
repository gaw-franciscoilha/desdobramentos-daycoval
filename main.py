from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
import os
import time
from dotenv import load_dotenv

def login(driver):
    try:
        print("Acessando o sistema...")
        driver.get('https://spjw.daycoval.com.br:8282')
        print("Sistema acessado com sucesso")
    except Exception as e:
        print(f"Erro ao acessar o sistema: {e}")

    try:
        print("Preenchenco user...")
        login_input = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, "//input[@id = 'ctl00_tbarMain_pgcMain_hucLogin_txtUser_I']")))
        login_input.send_keys(os.getenv('USER'))
        print("User preenchido com sucesso")
    except Exception as e:
        print(f"Erro ao preencher o user: {e}")

    try:
        print("Preenchendo senha...")
        password_input = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, "//input[@id = 'ctl00_tbarMain_pgcMain_hucLogin_txtPass_I']")))
        password_input.send_keys(os.getenv('PWD'))
        print("Senha preenchida com sucesso")
    except Exception as e:
        print(f"Erro ao preencher a senha: {e}")

    try:
        print("Clicando no botão de login...")
        login_button = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, "//div[@id = 'ctl00_tbarMain_pgcMain_hucLogin_btnLogin']")))
        login_button.click()
        print("Login realizado com sucesso")
    except Exception as e:
        print(f"Erro ao realizar o login: {e}")

def logout(driver):
    try:
        print("Acessando a página de logout...")
        login_section = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, "(//span[text() = 'Login'])[1]")))
        login_section.click()
        print("Página de logout acessada com sucesso")
    except Exception as e:
        print(f"Erro ao acessar a página de logout: {e}")

    try:
        print("Clicando no botão de logout...")
        logout_button = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, "//img[@id = 'ctl00_ctl00_ctl00_tbarMain_pgcMain_hucLogin_btnLogoutImg']")))
        logout_button.click()
        print("Logout realizado com sucesso")
    except Exception as e:
        print(f"Erro ao realizar o logout: {e}")

def voltar_contexto_padrao(driver):
    try:
        print("Voltando o contexto original do webdriver")
        driver.switch_to.default_content()
        print("Contexto original do webdriver obtido")
    except Exception as e:
        print(f"Erro ao voltar o contexto original do webdriver: {e}")

def navegar_pesquisas(driver):
    actions = ActionChains(driver)

    try:
        print("Hover sobre a opção de todos os processos...")
        all_li_item = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "//li[@id = 'ctl00_tbarMain_pgcMain_arpProcesso1_HTC_mnuProcesso_DXI4_']")))
        actions.move_to_element(all_li_item).perform()

        all_option = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, "//a[@id = 'ctl00_tbarMain_pgcMain_arpProcesso1_HTC_mnuProcesso_DXI4i0_T']")))
        all_option.click()
        print("Opção de todos os processos selecionada com sucesso")
    except Exception as e:
        print(f"Erro ao navegar nas pesquisas: {e}")

def processar_item(driver, item):
    actions = ActionChains(driver)

    try:
        print("Enviando número do dossiê...")
        dossie_input = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, "//input[@id = 'ctl00_ctl00_ctl00_cphMain_pgcFiltros_cphFiltros_HucFoneticoDossieProcesso1_txtCodDossie_I']")))
        dossie_input.clear()
        dossie_input.send_keys(item)
        print("Número do dossiê enviado com sucesso")
    except Exception as e:
        print(f"Erro ao enviar o número do dossiê: {e}")

    try:
        print("Clicando no botão de pesquisar...")
        search_button = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, "//li[@id = 'ctl00_ctl00_ctl00_tbarMain_pgcMain_pnlConsultar_HTC_mnuRibConsultar_DXI0_']")))
        search_button.click()
        print("Pesquisa realizada com sucesso")
    except Exception as e:
        print(f"Erro ao realizar a pesquisa: {e}")

    try:
        print("Aguardando terminar a busca...")
        WebDriverWait(driver, 20).until(EC.invisibility_of_element_located((By.ID, "ctl00_ctl00_ctl00_cphMain_cphGrid_pgcGrid_cphGrid_grdMain_LPV")))
        print("Busca concluída")
    except Exception as e:
        print(f"Erro ao aguardar a conclusão da busca: {e}")

    try:
        print("Aguardando o processo buscado estar visivel...")
        WebDriverWait(driver, 20).until(EC.visibility_of_element_located((By.XPATH, f"//td[text() = '{item}']")))
        print("Processo buscado visivel")
    except Exception as e:
        print(f"Erro ao aguardar o processo buscado estar visivel: {e}")

    try:
        print("Hover sobre a opção de detalhes...")
        details_li_item = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "//li[@id = 'ctl00_ctl00_ctl00_cphMain_cphGrid_dnavProc_DXI17_']")))
        actions.move_to_element(details_li_item).perform()
        print("Opção de detalhes selecionada com sucesso")
    except Exception as e:
        print(f"Erro ao selecionar a opção de detalhes: {e}")

    try:
        print("Clicando em desdobramentos...")
        desdobramentos_button = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, "//div[@id = 'ctl00_ctl00_ctl00_cphMain_cphGrid_dnavProc_DXI17i15_T']")))
        desdobramentos_button.click()
        print("Desdobramentos selecionados com sucesso")
    except Exception as e:
        print(f"Erro ao selecionar desdobramentos: {e}")

    try:
        print("Aguardando aparecer a tela de desdobramentos...")
        desdobramentos_header = WebDriverWait(driver, 20).until(EC.visibility_of_element_located((By.XPATH, "//span[@id = 'ctl00_ctl00_ctl00_popupMain_PWH0T']")))
        print("Tela de desdobramentos aberta")
    except Exception as e:
        print(f"Erro ao aguardar a abertura da tela de desdobramentos: {e}")

    try:
        print("\n\nTrocando o contexto do webdriver para o iframe de desdobramentos\n\n")
        WebDriverWait(driver, 60).until(EC.frame_to_be_available_and_switch_to_it((By.ID, "ctl00_ctl00_ctl00_popupMain_CIF0")))
        print("Contexto do webdriver trocado com sucesso")
    except Exception as e:
        print(f"Erro ao trocar o contexto do webdriver: {e}")

    # ELEMENTOS DAQUI PARA FRENTE DEVEM ESTAR DENTRO DO IFRAME DE DESDOBRAMENTOS

    try:
        print("Clicando em incluir desdobramentos...")
        add_button = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, "//li[@id = 'ctl00_ctl00_cphGridDetail_cphGridDetail_dnavDesdobr_DXI10_']")))
        add_button.click()
        print("Botão de incluir desdobramentos clicado com sucesso")
    except Exception as e:
        print(f"Erro ao clicar no botão de incluir desdobramentos: {e}")

    # VOLTAR O CONTEXTO PARA O PADRÃO PARA GARANTIR QUE A TELA DE ADICIONAR DESDOBRAMENTOS ESTÁ VISÍVEL

    voltar_contexto_padrao(driver)

    try:
        print("Aguardando aparecer a tela de adicionar desdobramentos...")
        WebDriverWait(driver, 20).until(EC.visibility_of_element_located((By. XPATH, "//span[@id = 'ctl00_ctl00_ctl00_popupMain_PWH1T']")))
        print("Tela de adicionar desdobramentos aberta")
    except Exception as e:
        print(f"Erro ao aguardar a abertura da tela de adicionar desdobramentos: {e}")

    # TROCAR O CONTEXTO PARA O IFRAME DE ADICIONAR DESDOBRAMENTOS

    try:
        print("\n\nTrocando o contexto do webdriver para o iframe de adicionar desdobramentos\n\n")
        WebDriverWait(driver, 60).until(EC.frame_to_be_available_and_switch_to_it((By.ID, "ctl00_ctl00_ctl00_popupMain_CIF1")))
        print("Contexto do webdriver trocado com sucesso")
    except Exception as e:
        print(f"Erro ao trocar o contexto do webdriver: {e}")

    # ELEMENTOS DAQUI PARA FRENTE DEVEM ESTAR DENTRO DO IFRAME DE ADICIONAR DESDOBRAMENTOS

    time.sleep(5)

    try:
        print("Preenchendo o campo data...")
        data_input = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.ID, "ctl00_ctl00_cphZoomMain_cphMain_fvwMain_dteDtDesdobr_I")))
        data_input.click()
        data_input.send_keys(Keys.CONTROL, "a")
        data_input.send_keys(Keys.DELETE)
        data_input.send_keys("06122000")  # sem barras
        data_input.send_keys(Keys.TAB)  
        print("Data preenchida com sucesso")
    except Exception as e:
        print(f"Erro ao preencher o campo data: {e}")

    try:
        print("Preenchendo o campo instância...")
        instancia_input = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.ID, "ctl00_ctl00_cphZoomMain_cphMain_fvwMain_speInstancia_I")))
        instancia_input.click()
        instancia_input.send_keys(Keys.CONTROL, "a")
        instancia_input.send_keys(Keys.DELETE)
        instancia_input.send_keys('99')
        print("Campo instância preenchido com sucesso")
    except Exception as e:
        print(f"Erro ao preencher o campo instância: {e}")

    try:
        print("Preenchendo o campo ação/recurso...")
        recurso_input = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.ID, "ctl00_ctl00_cphZoomMain_cphMain_fvwMain_cbxAcao_I")))
        recurso_input.click()
        recurso_input.send_keys("Agravo de Instrumento")
        print("Campo ação/recurso preenchido com sucesso")
    except Exception as e:
        print(f"Erro ao preencher o campo ação/recurso: {e}")

    ##########

    time.sleep(15)

    # Alternar para o contexto padrão
    voltar_contexto_padrao(driver)

    try:
        print("Clicando no botão para fechar o popup de adicionar desdobramentos")
        close_button = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, "//div[@id = 'ctl00_ctl00_ctl00_popupMain_HCB1']")))
        close_button.click()
        print("Popup fechado com sucesso")
    except Exception as e:
        print(f"Erro ao fechar o popup: {e}")

    try:
        print("Clicando no botão para fechar o popup de desdobramentos")
        close_button = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, "//div[@id = 'ctl00_ctl00_ctl00_popupMain_HCB0']")))
        close_button.click()
        print("Popup fechado com sucesso")
    except Exception as e:
        print(f"Erro ao fechar o popup: {e}")

def main():
    load_dotenv()

    options = Options()
    options.add_argument("--start-maximized")

    driver = webdriver.Chrome(options=options)

    # Fazer login
    login(driver)

    # Navegar até tela de pesquisas
    navegar_pesquisas(driver)

    # Processar os itens
    items = ["DAYCOVAL.0014"] #"DAYCOVAL.0031", "DAYCOVAL.0036"]

    for item in items:
        processar_item(driver, item)

    # Fazer logout
    input("Pressione Enter para sair...")
    logout(driver)

    driver.quit()

if __name__ == "__main__":
    main()