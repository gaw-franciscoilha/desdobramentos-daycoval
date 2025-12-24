from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from dotenv import load_dotenv
import os
import pandas as pd
import datetime

def fechar_popup_desdobramentos(driver):
    print("Clicando no botão para fechar o popup de desdobramentos")
    close_button = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, "//div[@id = 'ctl00_ctl00_ctl00_popupMain_HCB0']")))
    close_button.click()
    print("Popup fechado com sucesso")

def fechar_popup_adicionar_desdobramentos(driver):
    print("Clicando no botão para fechar o popup de adicionar desdobramentos")
    close_button = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, "//div[@id = 'ctl00_ctl00_ctl00_popupMain_HCB1']")))
    close_button.click()
    print("Popup fechado com sucesso")

def retornar(driver):
    voltar_contexto_padrao(driver)
    fechar_popup_adicionar_desdobramentos(driver)
    fechar_popup_desdobramentos(driver)

def criar_driver():
    options = Options()
    options.add_argument("--start-maximized")
    driver = webdriver.Chrome(options=options)
    return driver

def carregar_planilha(caminho_arquivo):
    df = pd.read_excel(caminho_arquivo)
    
    # Definição das colunas que o script precisa para rodar
    colunas_necessarias = [
        'COD. DOSSIE', 'DATA DESDOBRAMENTO', 'INSTÂNCIA',
        'NUM. RECURSO/INCIDENTE', 'AÇÃO/RECURSO', 'PARTE INTERESSE',
        'POSIC. INTERESSE', 'PARTE CONTRÁRIA', 'POSIC. CONTRA', 'RESUMO'
    ]

    colunas_faltantes = [col for col in colunas_necessarias if col not in df.columns]
    if colunas_faltantes:
        raise ValueError(f"A planilha está faltando as seguintes colunas obrigatórias: {', '.join(colunas_faltantes)}")

    if df.empty:
        raise ValueError("A planilha contém as colunas corretas, mas não possui dados (está vazia).")

    campos_obrigatorios = ['COD. DOSSIE', 'DATA DESDOBRAMENTO', 'INSTÂNCIA', 'NUM. RECURSO/INCIDENTE']
    dados = []

    for i, row in df.iterrows():
        # Validação de campos obrigatórios
        for campo in campos_obrigatorios:
            if pd.isna(row[campo]) or str(row[campo]).strip() == "":
                raise ValueError(f"O campo obrigatório '{campo}' está vazio na linha {i + 2}.")

        dados.append({
            'COD. DOSSIE':	str(row['COD. DOSSIE']).strip(),
            'DATA DESDOBRAMENTO': pd.to_datetime(row['DATA DESDOBRAMENTO'], dayfirst=True).strftime('%d%m%Y'),
            'INSTÂNCIA': str(row['INSTÂNCIA']),
            'NUM. RECURSO/INCIDENTE': row['NUM. RECURSO/INCIDENTE'],
            'AÇÃO/RECURSO': str(row['AÇÃO/RECURSO']) if pd.notna(row['AÇÃO/RECURSO']) else '',
            'PARTE INTERESSE': str(row['PARTE INTERESSE']) if pd.notna(row['PARTE INTERESSE']) else '',
            'POSIC. INTERESSE': str(row['POSIC. INTERESSE']) if pd.notna(row['POSIC. INTERESSE']) else '',
            'PARTE CONTRÁRIA': str(row['PARTE CONTRÁRIA']) if pd.notna(row['PARTE CONTRÁRIA']) else '',	
            'POSIC. CONTRA': str(row['POSIC. CONTRA']) if pd.notna(row['POSIC. CONTRA']) else '',
            'RESUMO': str(row['RESUMO']) if pd.notna(row['RESUMO']) else '' 
        })

    return dados

def login(driver):
    try:
        print("Acessando o sistema...")
        driver.get('https://spjw.daycoval.com.br:8282')
        print("Sistema acessado com sucesso")
    except Exception as e:
        print(f"Erro ao acessar o sistema: {e}")

    try:
        print("Preenchendo user...")
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
        dossie_input.send_keys(item['COD. DOSSIE'])
        print("Número do dossiê enviado com sucesso")
    except Exception as e:
        print(f"Erro ao enviar o número do dossiê: {e}")
        return {"status": "Erro", "detalhes": f"Erro ao enviar o número do dossiê: {e}"}

    try:
        print("Clicando no botão de pesquisar...")
        search_button = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, "//li[@id = 'ctl00_ctl00_ctl00_tbarMain_pgcMain_pnlConsultar_HTC_mnuRibConsultar_DXI0_']")))
        search_button.click()
        print("Pesquisa realizada com sucesso")
    except Exception as e:
        print(f"Erro ao realizar a pesquisa: {e}")
        return {"status": "Erro", "detalhes": f"Erro ao realizar a pesquisa: {e}"}
    
    try:
        print("Aguardando terminar a busca...")
        WebDriverWait(driver, 20).until(EC.invisibility_of_element_located((By.ID, "ctl00_ctl00_ctl00_cphMain_cphGrid_pgcGrid_cphGrid_grdMain_LPV")))
        print("Busca concluída")
    except Exception as e:
        print(f"Erro ao aguardar a conclusão da busca: {e}")
        return {"status": "Erro", "detalhes": f"Erro ao aguardar a conclusão da busca: {e}"}

    try:
        print("Aguardando o processo buscado estar visivel...")
        process = WebDriverWait(driver, 20).until(EC.visibility_of_element_located((By.XPATH, f"//td[text() = '{item['COD. DOSSIE']}']")))
        process.click()
        print("Processo buscado visivel")
    except Exception as e:
        print(f"Erro ao aguardar o processo buscado estar visivel: {e}")
        return {"status": "Erro", "detalhes": f"Erro ao aguardar o processo buscado estar visivel: {e}"}

    try:
        print("Hover sobre a opção de detalhes...")
        details_li_item = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "//li[@id = 'ctl00_ctl00_ctl00_cphMain_cphGrid_dnavProc_DXI17_']")))
        actions.move_to_element(details_li_item).perform()
        print("Opção de detalhes selecionada com sucesso")
    except Exception as e:
        print(f"Erro ao selecionar a opção de detalhes: {e}")
        return {"status": "Erro", "detalhes": f"Erro ao selecionar a opção de detalhes: {e}"}

    try:
        print("Clicando em desdobramentos...")
        desdobramentos_button = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, "//div[@id = 'ctl00_ctl00_ctl00_cphMain_cphGrid_dnavProc_DXI17i15_T']")))
        desdobramentos_button.click()
        print("Desdobramentos selecionados com sucesso")
    except Exception as e:
        print(f"Erro ao selecionar desdobramentos: {e}")
        return {"status": "Erro", "detalhes": f"Erro ao selecionar desdobramentos: {e}"}

    try:
        print("Aguardando aparecer a tela de desdobramentos...")
        WebDriverWait(driver, 20).until(EC.visibility_of_element_located((By.XPATH, "//span[@id = 'ctl00_ctl00_ctl00_popupMain_PWH0T']")))
        print("Tela de desdobramentos aberta")
    except Exception as e:
        print(f"Erro ao aguardar a abertura da tela de desdobramentos: {e}")
        return {"status": "Erro", "detalhes": f"Erro ao aguardar a abertura da tela de desdobramentos: {e}"}

    try:
        print("\n\nTrocando o contexto do webdriver para o iframe de desdobramentos\n\n")
        WebDriverWait(driver, 60).until(EC.frame_to_be_available_and_switch_to_it((By.ID, "ctl00_ctl00_ctl00_popupMain_CIF0")))
        print("Contexto do webdriver trocado com sucesso")
    except Exception as e:
        print(f"Erro ao trocar o contexto do webdriver: {e}")
        return {"status": "Erro", "detalhes": f"Erro ao trocar o contexto do webdriver: {e}"}

    # ELEMENTOS DAQUI PARA FRENTE DEVEM ESTAR DENTRO DO IFRAME DE DESDOBRAMENTOS

    try:
        print("Garantindo que o processo correto foi clicado...")
        WebDriverWait(driver, 20).until(EC.visibility_of_element_located((By.XPATH, "//span[@id = 'ctl00_ctl00_pnlDadosMaster_cphDadosMaster_HMasterHeader1_Label18']")))
        print("O processo correto foi escolhido")
    except Exception as e:
        print(f"O processo buscado não é o mesmo que foi aberto no popup!")
        return {"status": "Erro", "detalhes": f"Processo buscado não é o mesmo que foi aberto no popup."}
    
    try:
        print("Clicando em incluir desdobramentos...")
        add_button = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, "//li[@id = 'ctl00_ctl00_cphGridDetail_cphGridDetail_dnavDesdobr_DXI10_']")))
        add_button.click()
        print("Botão de incluir desdobramentos clicado com sucesso")
    except Exception as e:
        print(f"Erro ao clicar no botão de incluir desdobramentos: {e}")
        return {"status": "Erro", "detalhes": f"Erro ao clicar no botão de incluir desdobramentos: {e}"}

    # VOLTAR O CONTEXTO PARA O PADRÃO PARA GARANTIR QUE A TELA DE ADICIONAR DESDOBRAMENTOS ESTÁ VISÍVEL

    voltar_contexto_padrao(driver)

    try:
        print("Aguardando aparecer a tela de adicionar desdobramentos...")
        WebDriverWait(driver, 60).until(EC.visibility_of_element_located((By. XPATH, "//span[@id = 'ctl00_ctl00_ctl00_popupMain_PWH1T']")))
        print("Tela de adicionar desdobramentos aberta")
    except Exception as e:
        print(f"Erro ao aguardar a abertura da tela de adicionar desdobramentos: {e}")
        return {"status": "Erro", "detalhes": f"Erro ao aguardar a tela de adicionar desdobramentos: {e}"}

    # TROCAR O CONTEXTO PARA O IFRAME DE ADICIONAR DESDOBRAMENTOS

    try:
        print("\n\nTrocando o contexto do webdriver para o iframe de adicionar desdobramentos\n\n")
        WebDriverWait(driver, 60).until(EC.frame_to_be_available_and_switch_to_it((By.ID, "ctl00_ctl00_ctl00_popupMain_CIF1")))
        print("Contexto do webdriver trocado com sucesso")
    except Exception as e:
        print(f"Erro ao trocar o contexto do webdriver: {e}")
        return {"status": "Erro", "detalhes": f"Erro ao trocar o contexto do webdriver: {e}"}

    # ELEMENTOS DAQUI PARA FRENTE DEVEM ESTAR DENTRO DO IFRAME DE ADICIONAR DESDOBRAMENTOS

    try:
        print("Preenchendo o campo data...")
        data_input = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.ID, "ctl00_ctl00_cphZoomMain_cphMain_fvwMain_dteDtDesdobr_I")))
        data_input.click()
        data_input.send_keys(Keys.CONTROL, "a")
        data_input.send_keys(Keys.DELETE)
        data_input.send_keys(item['DATA DESDOBRAMENTO'])  # sem barras
        data_input.send_keys(Keys.TAB)  
        print("Data preenchida com sucesso")
    except Exception as e:
        print(f"Erro ao preencher o campo data: {e}")
        retornar(driver)
        return {"status": "Erro", "detalhes": f"Erro ao preencher o campo data: {e}"}

    try:
        print("Preenchendo o campo instância...")
        instancia_input = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.ID, "ctl00_ctl00_cphZoomMain_cphMain_fvwMain_speInstancia_I")))
        instancia_input.click()
        instancia_input.send_keys(Keys.CONTROL, "a")
        instancia_input.send_keys(Keys.DELETE)
        instancia_input.send_keys(item['INSTÂNCIA'])
        print("Campo instância preenchido com sucesso")
    except Exception as e:
        print(f"Erro ao preencher o campo instância: {e}")
        retornar(driver)
        return {"status": "Erro", "detalhes": f"Erro ao preencher o campo instância: {e}"}

    try:
        print("Preenchendo o campo núm. recurso/incidente...")
        recurso_input = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By. XPATH, "//input[@id = 'ctl00_ctl00_cphZoomMain_cphMain_fvwMain_txtNumJud_I']")))
        recurso_input.click()
        recurso_input.send_keys(item['NUM. RECURSO/INCIDENTE'])
        print("Campo Núm. Recurso/Incidente preenchido com sucesso")
    except Exception as e:
        print(f"Erro ao preencher o campo núm. recurso/incidente: {e}")
        retornar(driver)
        return {"status": "Erro", "detalhes": f"Erro ao preencher o campo núm. recurso/incidente: {e}"}
    
    try:
        if item['AÇÃO/RECURSO'] != '':

            print("Preenchendo o campo ação/recurso...")
            recurso_input = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.ID, "ctl00_ctl00_cphZoomMain_cphMain_fvwMain_cbxAcao_I")))
            recurso_input.click()
            recurso_input.send_keys(item['AÇÃO/RECURSO'])
            recurso_input.send_keys(Keys.TAB)
            print("Campo ação/recurso preenchido com sucesso")
        
        else:
            print("Campo ação/recurso não preenchido")
            pass

    except Exception as e:
        print(f"Erro ao preencher o campo ação/recurso: {e}")
        retornar(driver)
        return {"status": "Erro", "detalhes": f"Erro ao preencher o campo ação/recurso: {e}"}

    try:
        if item['PARTE INTERESSE'] != '':

            print("Preenchendo o campo parte de interesse...")
            parte_interesse_input = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, "//input[@id = 'ctl00_ctl00_cphZoomMain_cphMain_fvwMain_cbxLitisInter_I']")))
            parte_interesse_input.click()
            parte_interesse_input.send_keys(item['PARTE INTERESSE'])
            parte_interesse_input.send_keys(Keys.TAB)
            print("Campo parte de interesse preenchido com sucesso")
        
        else:
            print("Campo parte de interesse não preenchido")
            pass

    except Exception as e:
        print(f"Erro ao preencher o campo parte de interesse: {e}")
        retornar(driver)
        return {"status": "Erro", "detalhes": f"Erro ao preencher o campo parte de interesse: {e}"}

    try:
        if item['POSIC. INTERESSE'] != '':

            print("Preenchendo o campo posic. inter...")
            posic_interesse_input = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, "//input[@id = 'ctl00_ctl00_cphZoomMain_cphMain_fvwMain_cbxPosicInter_I']")))
            posic_interesse_input.click()
            posic_interesse_input.send_keys(item['POSIC. INTERESSE'])
            posic_interesse_input.send_keys(Keys.TAB)
            print("Campo posic. inter. preenchido com sucesso")

        else:
            print("Campo posic. inter. não preenchido")
            pass

    except Exception as e:
        print(f"Erro ao preencher o campo: {e}")
        retornar(driver)
        return {"status": "Erro", "detalhes": f"Erro ao preencher o campo: {e}"}

    try:
        if item['PARTE CONTRÁRIA'] != '':

            print("Preenchendo o campo parte contrária...")
            parte_contraria_input = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, "//input[@id = 'ctl00_ctl00_cphZoomMain_cphMain_fvwMain_cbxLitisContra_I']")))
            parte_contraria_input.click()
            parte_contraria_input.send_keys(item['PARTE CONTRÁRIA'])
            parte_contraria_input.send_keys(Keys.TAB)
            print("Campo parte contrária preenchido com sucesso")

        else:
            print("Campo parte contrária não preenchido")
            pass

    except Exception as e:
        print(f"Erro ao preencher o campo parte contrária: {e}")
        retornar(driver)
        return {"status": "Erro", "detalhes": f"Erro ao preencher o campo parte contrária: {e}"}

    try:
        if item['POSIC. CONTRA'] != '':

            print("Preenchendo o campo posic. contra...")
            posic_contra_input = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, "//input[@id = 'ctl00_ctl00_cphZoomMain_cphMain_fvwMain_cbxPosicContra_I']")))
            posic_contra_input.click()
            posic_contra_input.send_keys(item['POSIC. CONTRA'])
            posic_contra_input.send_keys(Keys.TAB)
            print("Campo posic. contra preenchido com sucesso")

        else:
            print("Campo posic. contra não preenchido")
            pass

    except Exception as e:
        print(f"Erro ao preencher o campo posic. contra: {e}")
        retornar(driver)
        return {"status": "Erro", "detalhes": f"Erro ao preencher o campo posic. contra: {e}"}

    try:
        if item['RESUMO'] != '':

            print("Preenchendo o campo resumo...")
            resumo_textarea = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, "//textarea[@id = 'ctl00_ctl00_cphZoomMain_cphMain_fvwMain_txtResumo_I']")))
            resumo_textarea.click()
            resumo_textarea.send_keys(item['RESUMO'])
            print("Campo resumo preenchido com sucesso")

        else:
            print("Campo resumo não preenchido")
            pass

    except Exception as e:
        print(f"Erro ao preencher o campo resumo: {e}")
        retornar(driver)
        return {"status": "Erro", "detalhes": f"Erro ao preencher o campo resumo: {e}"}

    # REMOVER QUANDO FOR TESTAR O SALVAMENTO DO DESDOBRAMENTO

    # try:
    #     print("Clicando no botão de salvar...")
    #     save_button = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, "//td[@id = 'ctl00_ctl00_cphZoomMain_tbarMain_tdConfirma']//a")))
    #     save_button.click()
    #     print("Botão de salvar clicado com sucesso")
    # except Exception as e:
    #     print(f"Erro ao clicar no botão de salvar: {e}")
    #     return {"status": "Erro", "detalhes": f"Erro ao clicar no botão de salvar: {e}"}

    ##########

    # Alternar para o contexto padrão
    voltar_contexto_padrao(driver)

    try:
        fechar_popup_adicionar_desdobramentos(driver)
    except Exception as e:
        print(f"Erro ao fechar o popup: {e}")
        return {"status": "Erro", "detalhes": f"Não foi possível fechar o pop de adicionar desdobramentos: {e}"}
    
    try:
        fechar_popup_desdobramentos(driver)
    except Exception as e:
        print(f"Erro ao fechar o popup: {e}")
        return {"status": "Erro", "detalhes": f"Não foi possível fechar o pop de desdobramentos: {e}"}

    return {"status": "Sucesso", "detalhes": ""}

def main():
    # Definir os caminhos
    PLANILHA_PATH = r"C:\Users\Francisco Ilha\Downloads\modelo.xlsx"

    # Carregar variáveis de ambiente
    load_dotenv()
    
    # Carregar dados da planilha
    dados = carregar_planilha(PLANILHA_PATH)

    # Inicializar o relatório
    relatorio_dados = []

    # Iniciar o driver
    driver = criar_driver()

    try:
        # Fazer login
        login(driver)

        # Navegar até tela de pesquisas
        navegar_pesquisas(driver)

        for item in dados:
            resultado = {
                "COD. DOSSIE": item['COD. DOSSIE'],
                "NUM. RECURSO/INCIDENTE": item['NUM. RECURSO/INCIDENTE'],
                "STATUS": "Não processado",
                "HORÁRIO": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "DETALHES": ""
            }

            try:
                print(f"\n--- Iniciando processamento do item: {item['COD. DOSSIE']} ---\n")

                status_processamento = processar_item(driver, item)

                resultado["STATUS"] = status_processamento["status"]
                resultado["DETALHES"] = status_processamento["detalhes"]

                print(f"\n--- Resultado para {item['COD. DOSSIE']}: {status_processamento['status']} ---\n")
            
            except Exception as e:
                print(f"Erro ao processar o item: {e}")
                resultado["STATUS"] = "Erro"
                resultado["DETALHES"] = str(e)

            finally:
                relatorio_dados.append(resultado)

        print("\nAutomação concluída para todos os itens.")

    except Exception as e:
        print(f"Erro durante a execução: {e}")

    finally:
        if relatorio_dados:
            df_relatorio = pd.DataFrame(relatorio_dados)
            nome_arquivo = f"relatorio_{datetime.datetime.now().strftime('%d_%m_%H_%M')}.csv"
            df_relatorio.to_csv(nome_arquivo, index=False, sep=';', encoding='utf-8-sig')

            print(f"Relatório salvo em: {os.path.abspath(nome_arquivo)}")

    # Fazer logout
    input("Pressione Enter para sair...")
    logout(driver)

    driver.quit()

if __name__ == "__main__":
    main()