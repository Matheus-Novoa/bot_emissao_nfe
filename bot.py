from botcity.web import WebBot, Browser, By
from botcity.core import DesktopBot
from botcity.web.browsers.chrome import default_options
from dados import *
import os
from unidecode import unidecode
from pyautogui import alert, confirm
from dotenv import load_dotenv
from pywinauto.application import Application
from pywinauto.findwindows import find_window



load_dotenv()

if arquivo_progresso.exists():
    with open(arquivo_progresso) as f:
        linha = int(f.read().split()[-1])
    dados = dados.iloc[linha:]

def not_found(label):
    print(f"Element not found: {label}")


def entrar(bot):
    botao_fechar_janela = bot.find_element('//*[@id="modal-one"]/div[2]/button', By.XPATH)
    botao_fechar_janela.click()
    botao_autenticar = bot.find_element('//*[@id="form"]/div[1]/a/img', By.XPATH)
    botao_autenticar.click()
    
    campo_user = bot.find_element('//*[@id="username"]', By.XPATH)
    campo_user.send_keys(os.getenv('LOGIN'))
    
    campo_senha = bot.find_element('//*[@id="password"]', By.XPATH)
    campo_senha.send_keys(os.getenv('SENHA'))
    
    botao_entrar = bot.find_element('//*[@id="mainContent"]/div/form/input[5]', By.XPATH)
    botao_entrar.click()


def sair(bot):
    botao_sair = bot.find_element('//*[@id="identificacao"]/table/tbody/tr/td[2]/a/img', By.XPATH)
    botao_sair.click()
    voltar_pag_inicial = bot.find_element('//*[@id="mensagem"]/div/p[4]/a', By.XPATH)
    voltar_pag_inicial.click()


def trazer_janela_para_frente(titulo):
    try:
        # Localiza a janela pelo título
        hwnd = find_window(title=titulo)
        app = Application().connect(handle=hwnd)
        janela = app.window(handle=hwnd)
        # Traz a janela para frente
        janela.set_focus()
        return janela
    except Exception as e:
        print(f"Erro ao trazer a janela para frente: {e}")
        return None


def main():

    data_geracao = '30062024'
    mes = 'JUNHO'
    
    ANO = '2024'
    
    download_folder_path=os.getenv('DOWNLOAD_FOLDER_PATH')
    # data_geracao = input('Data Geração: ')
    # mes = input('Mês: ')

    bot_desktop = DesktopBot()
    bot = WebBot()
    bot.headless = False
    bot.browser = Browser.CHROME
    bot.driver_path = os.getenv('BOT_DRIVER_PATH')
    profile_path = os.getenv('PROFILE_PATH')

    def_options = default_options(
        user_data_dir=profile_path,  # Inform the profile path that wants to start the browser
        download_folder_path=download_folder_path
    ) 
    def_options.add_argument('–profile-directory=Default')
    bot.options = def_options
    
    bot.browse("https://nfe.portoalegre.rs.gov.br")

    bot.maximize_window()
    
    entrar(bot)

    aba_geracao = bot.find_element('//*[@id="nav"]/div[1]/a/img', By.XPATH)
    aba_geracao.click()
    bot.wait(1000)
    campo_geracao = bot.find_element('//*[@id="MesReferenciaModalPanelSubview:formMesReferencia:dtCompetencia"]', By.XPATH)
    campo_geracao.send_keys(data_geracao)
    bot.enter()
    bot.wait(3000)
    
    for cliente in dados.itertuples():
        
        try:
            ################### CPF ###################
            campo_cpf = bot.find_element('//*[@id="form:numDocumento"]', By.XPATH)
            campo_cpf.click()
            campo_cpf.send_keys(cliente.CPF)
            botao_lupa = bot.find_element('//*[@id="form:btAutoCompleteTomador"]', By.XPATH)
            botao_lupa.click()
            # bot.wait(3000)
            
            if bot.find("erro_cpf_nao_encontrado", matching=0.97, waiting_time=1500):
                not_found("erro_cpf_nao_encontrado")
                alert(title='CPF não cadastrado. Preencha manualmente')

            nome_responsavel_dados = unidecode(cliente.ResponsávelFinanceiro.upper().strip())
            campo_razao_social = bot.find_element('//*[@id="form:dnomeRazaoSocial"]', By.XPATH)
            bot.wait_for_element_visibility(campo_razao_social)
            nome_responsavel_retornado = unidecode(campo_razao_social.get_attribute('value').upper().strip())

            while nome_responsavel_retornado != nome_responsavel_dados:
                resposta = confirm(text=f'{nome_responsavel_retornado} | {nome_responsavel_dados}',
                        title='Nome do responsável não correspondente', buttons=['Continuar', 'Interromper'])
                if resposta == 'Continuar':
                    break
                else:
                    raise
            ################### SERVIÇO ###################
            aba_identificacao_servico = bot.find_element('//*[@id="topo_aba2"]/a', By.XPATH)
            bot.wait_for_element_visibility(aba_identificacao_servico)
            aba_identificacao_servico.click()
            texto_descricao = f'PRESTAÇÃO DE SERVIÇO EDUCAÇÃO INFANTIL/FUNDAMENTAL MÊS {mes}/{ANO} - ALUNO {cliente.Aluno}'
            campo_descricao = bot.find_element('//*[@id="form:descriminacaoServico"]', By.XPATH)
            campo_descricao.send_keys(texto_descricao)

            if cliente.Acumulador == '1':
                opcao_turma = bot.find_element('//*[@id="form:codigoCnae"]', By.XPATH)
                opcao_turma.click()
                bot.type_up()
                bot.enter()

            ################### VALOR ###################
            # bot.wait(1000)
            aba_valores = bot.find_element('//*[@id="topo_aba3"]/a', By.XPATH)
            aba_valores.click()
            campo_valor_total = bot.find_element('//*[@id="form:valorServicos"]', By.XPATH)
            campo_valor_total.send_keys(cliente.ValorTotal)

            ################### DOWNLOAD NOTA ###################
            botao_gerar_nfs = bot.find_element('//*[@id="form:bt_emitir_NFS-e"]', By.XPATH)
            botao_gerar_nfs.click()
            bot.wait(2000)
            
            botao_confirmar_geracao = bot.find_element('//*[@id="appletAssinador:j_id766"]', By.XPATH)
            botao_confirmar_geracao.click()
            
            bot.wait(1000)

            janela = trazer_janela_para_frente('Logon do Token')
            while janela == None:
                janela = trazer_janela_para_frente('Logon do Token')

            bot_desktop.kb_type('123456') # matriz
            # bot_desktop.kb_type('1234')
            bot_desktop.enter()
            bot.wait(3000)
            
            botao_download = bot.find_element('//*[@id="form"]/input[2]', By.XPATH, ensure_clickable=True, ensure_visible=True)
            # bot.wait_for_element_visibility(botao_download)
            botao_download.click()
            bot.wait(1000)
            num_nota = bot.get_last_created_file(download_folder_path).split(os.sep)[-1].split('.')[0][-4:]

            with open(arquivo_notas, 'a') as f:
                f.write(f'{num_nota} {cliente.ResponsávelFinanceiro}')
                f.write('\n')
        except:
            with open(arquivo_progresso, 'w') as f:
                f.write(f'Erro {cliente.ResponsávelFinanceiro} linha {cliente.Index}')
                raise
        ################### RETORNA E LIMPA OS CAMPOS ###################
        try:    
            botao_retorno = bot.find_element('//*[@id="form:j_id9"]', By.XPATH)
            botao_retorno.click()
            bot.wait(3000)
            
            botao_limpar_digitacao = bot.find_element('//*[@id="form"]/input[3]', By.XPATH)
            botao_limpar_digitacao.click()
            bot.wait(2000)
        except:
            with open(arquivo_progresso, 'w') as f:
                f.write(f'Erro após {cliente.ResponsávelFinanceiro} linha {int(cliente.Index) + 1}')
                raise
    sair(bot)
    bot.stop_browser()


if __name__ == '__main__':
    main()
