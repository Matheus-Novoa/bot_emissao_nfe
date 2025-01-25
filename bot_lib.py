from botcity.web import Browser, By, WebBot
from botcity.web.browsers.chrome import default_options
from botcity.core import DesktopBot
from unidecode import unidecode
from pyautogui import confirm
from pywinauto.application import Application
from pywinauto.findwindows import find_window
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By as sBy
from selenium.common.exceptions import StaleElementReferenceException, ElementClickInterceptedException
import os



class Bot:

    def __init__(self):
        self.webBot = WebBot()
        self.desktopBot = DesktopBot()


    def bot_setup(self, caminhoPastaDownload, profile_path):
        self.caminhoPastaDownload = caminhoPastaDownload.replace('/', '\\')

        self.webBot.headless = False
        self.webBot.browser = Browser.CHROME
        self.webBot.driver_path = ChromeDriverManager().install()

        def_options = default_options(
            user_data_dir=profile_path,  # Inform the profile path that wants to start the browser
            download_folder_path=self.caminhoPastaDownload
        ) 
        def_options.add_argument('–profile-directory=Default')
        self.webBot.options = def_options
        
        self.webBot.browse("https://nfe.portoalegre.rs.gov.br")

        self.webBot.maximize_window()

    
    def entrar(self, sede: dict):
        # Botão fechar janela
        self.webBot.find_element('//*[@id="modal-one"]/div[2]/button', By.XPATH).click()
        # Botão autenticar
        self.webBot.find_element('//*[@id="form"]/div[1]/a/img', By.XPATH).click()
        # Campo usuário
        self.webBot.find_element('//*[@id="username"]', By.XPATH).send_keys(sede['LOGIN'])
        # Campo senha
        self.webBot.find_element('//*[@id="password"]', By.XPATH).send_keys(sede['SENHA'])
        # Botão entrar
        self.webBot.find_element('//*[@id="mainContent"]/div/form/input[5]', By.XPATH).click()


    def sair(self):
        # Botão sair
        self.webBot.find_element('//*[@id="identificacao"]/table/tbody/tr/td[2]/a/img', By.XPATH).click()
        # Botão voltar para a página inicial
        self.webBot.find_element('//*[@id="mensagem"]/div/p[4]/a', By.XPATH).click()


    def definir_data(self, dataGeracao):
        self.webBot.find_element('//*[@id="nav"]/div[1]/a/img', By.XPATH).click()
        self.webBot.wait(1000)
        # campo_geracao
        self.webBot.find_element('//*[@id="MesReferenciaModalPanelSubview:formMesReferencia:dtCompetencia"]',
                                        By.XPATH).send_keys(dataGeracao)
        self.webBot.enter()
        self.webBot.wait(3000)


    def trazer_janela_para_frente(self, titulo):
        try:
            # Localiza a janela pelo título
            self.hwnd = find_window(title=titulo)
            self.janela = Application().connect(handle=self.hwnd).window(handle=self.hwnd)
            # Traz a janela para frente
            self.janela.set_focus()
            return self.janela
        except Exception:
            return None
        

    def preencher_campos(self, dadoCliente, mes, ano):

        self.dadoCliente = dadoCliente
        ################### CPF ###################
        while True:
            try:
                self._wait = WebDriverWait(self.webBot.driver, 15)
                campo_cpf = self._wait.until(EC.element_to_be_clickable((sBy.XPATH, '//*[@id="form:numDocumento"]')))
                campo_cpf.click()
                campo_cpf.send_keys(dadoCliente.CPF)
                # botao_lupa
                self.webBot.find_element('//*[@id="form:btAutoCompleteTomador"]', By.XPATH).click()
                break
            except ElementClickInterceptedException:
                continue

        nome_responsavel_dados = unidecode(dadoCliente.ResponsávelFinanceiro.upper().strip())
        campo_razao_social = self.webBot.find_element('//*[@id="form:dnomeRazaoSocial"]',
                                                By.XPATH)

        nome_responsavel_retornado = None
        while not (nome_responsavel_retornado and campo_razao_social):
            campo_razao_social = self.webBot.find_element('//*[@id="form:dnomeRazaoSocial"]',
                                                By.XPATH)
            try:
                nome_responsavel_retornado = unidecode(
                    campo_razao_social.get_attribute('value').upper().strip())
            except StaleElementReferenceException:
                # print('Não foi possível realizar a leitura do nome')
                continue

        while nome_responsavel_retornado != nome_responsavel_dados:
            resposta = confirm(
                    text=f'{nome_responsavel_retornado} | {nome_responsavel_dados}',
                    title='Nome do responsável não correspondente',
                    buttons=['Continuar', 'Interromper']
                    )
            if resposta == 'Continuar':
                break
            else:
                raise
        ################### SERVIÇO ###################
        aba_identificacao_servico = self.webBot.find_element('//*[@id="topo_aba2"]/a', By.XPATH)
        self.webBot.wait_for_element_visibility(aba_identificacao_servico)
        aba_identificacao_servico.click()
        texto_descricao = f'PRESTAÇÃO DE SERVIÇO EDUCAÇÃO INFANTIL/FUNDAMENTAL MÊS {mes}/{ano} - ALUNO {dadoCliente.Aluno}'
        
        # campo_descricao
        self.webBot.find_element('//*[@id="form:descriminacaoServico"]', By.XPATH).send_keys(texto_descricao)

        if dadoCliente.Acumulador == '1':
            # opcao_turma
            self.webBot.find_element('//*[@id="form:codigoCnae"]', By.XPATH).click()
            self.webBot.type_up()
            self.webBot.enter()

        ################### VALOR ###################
        # aba_valores
        self.webBot.find_element('//*[@id="topo_aba3"]/a', By.XPATH).click()
        # campo_valor_total
        self.webBot.find_element('//*[@id="form:valorServicos"]', By.XPATH).send_keys(dadoCliente.ValorTotal)

    
    def gerar_nf(self, nomeJanela, senha):
        ################### DOWNLOAD NOTA ###################
        # botao_gerar_nfs
        self.webBot.find_element('//*[@id="form:bt_emitir_NFS-e"]', By.XPATH).click()
        # botao_confirmar_geracao
        self.webBot.find_element('//*[@id="appletAssinador:j_id766"]',
                            By.XPATH,
                            ensure_clickable=True).click()
        
        nome_janela = nomeJanela
        self.webBot.wait(500)
        janela = self.trazer_janela_para_frente(nome_janela)
        while janela == None:
            self.webBot.wait(500)
            janela = self.trazer_janela_para_frente(nome_janela)

        self.desktopBot.kb_type(senha)
        self.desktopBot.enter()


    def baixar_nf(self, arquivo_notas):
        while True:
            try:
                self.webBot.wait(500)
                botao_download = self._wait.until(EC.element_to_be_clickable((sBy.XPATH, '//*[@id="form"]/input[2]')))
                botao_download.click()
                break
            except Exception:
                # print('Esperando o carregamento da página de downloads...')
                continue

        self.webBot.wait(1000)
        num_nota = self.webBot.get_last_created_file(self.caminhoPastaDownload).split(os.sep)[-1].split('.')[0][-4:]

        with open(arquivo_notas, 'a') as f:
            f.write(f'{num_nota} {self.dadoCliente.ResponsávelFinanceiro}')
            f.write('\n')


    def retornar(self, arquivo_progresso):
        while True:
            try:    
                # botao_retorno
                self.webBot.find_element('//*[@id="form:j_id9"]',
                                                By.XPATH,
                                                ensure_clickable=True,
                                                ensure_visible=True).click()
                
                # botao_limpar_digitacao
                botao_limpar_digitacao = self._wait.until(
                    EC.element_to_be_clickable((sBy.XPATH, '//*[@id="form"]/input[3]'))
                    )
                self.webBot.wait(500)
                botao_limpar_digitacao.click()
                break
            except ElementClickInterceptedException as err:
                self.webBot.wait(500)
                continue
            except Exception as err:
                with open(arquivo_progresso, 'w') as f:
                    f.write(f'Erro após {self.dadoCliente.ResponsávelFinanceiro} linha {int(self.dadoCliente.Index) + 1}')
                    raise err
                

    def fechar_navegador(self):
        self.webBot.stop_browser()

if __name__ == '__main__':
    ...
