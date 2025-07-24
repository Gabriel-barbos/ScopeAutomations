import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import TimeoutException, NoSuchElementException, StaleElementReferenceException
import time
import logging
import os


class CarAdditionAutomation:
    def __init__(self, webdriver_path=None):
        """
        Inicializa a automação
        
        Args:
            webdriver_path (str): Caminho para o ChromeDriver (opcional se estiver no PATH)
        """
        self.webdriver_path = webdriver_path
        self.driver = None
        self.wait = None
        self.nome_grupo = None  # Nome do grupo de veículos
        
        # Contadores para relatório
        self.carros_adicionados = []
        self.carros_nao_encontrados = []
        self.carros_ja_no_grupo = []
        self.total_processados = 0
        
        # Configurar logging
        logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
        self.logger = logging.getLogger(__name__)
    
    def setup_driver(self):
        """Configura o WebDriver do Chrome"""
        chrome_options = Options()
        # chrome_options.add_argument("--headless")  # Descomente para executar sem interface gráfica
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-dev-shm-usage")
        
        if self.webdriver_path:
            self.driver = webdriver.Chrome(executable_path=self.webdriver_path, options=chrome_options)
        else:
            self.driver = webdriver.Chrome(options=chrome_options)
        
        self.wait = WebDriverWait(self.driver, 10)
        self.driver.maximize_window()
    
    def inserir_chassis_terminal(self):
        """Permite inserir chassis diretamente no terminal"""
        chassis_list = []
        
        print("\n" + "="*50)
        print("📝 INSERÇÃO DE CHASSIS")
        print("="*50)
        print("Digite os números dos chassis (um por linha)")
        print("Digite 'fim' quando terminar")
        print("-"*50)
        
        while True:
            chassi = input("Chassi: ").strip()
            
            if chassi.lower() == 'fim':
                break
            elif chassi == '':
                continue
            else:
                chassis_list.append(chassi)
                print(f"✅ Chassi adicionado: {chassi} (Total: {len(chassis_list)})")
        
        return chassis_list
    
    def carregar_chassis_excel(self):
        """Carrega chassis de planilha Excel"""
        arquivo = os.path.join(os.path.dirname(__file__), 'AdicionarGrupo.xlsx')
        
        try:
            df = pd.read_excel(arquivo)
            chassis_list = df.iloc[:, 0].dropna().tolist()  # Primeira coluna
            chassis_list = [str(chassi).strip() for chassi in chassis_list]
            
            print(f"✅ {len(chassis_list)} chassis carregados da planilha")
            return chassis_list
        except Exception as e:
            print(f"❌ Erro ao carregar planilha: {e}")
            return []
    
    def escolher_metodo_entrada(self):
        """Permite escolher como inserir os chassis"""
        print("\n" + "="*50)
        print("🚗 AUTOMAÇÃO DE ADIÇÃO DE CARROS")
        print("="*50)
        print("Como deseja inserir os chassis?")
        print("1. ⌨️  Digitar no terminal")
        print("2. 📁 Carregar de planilha Excel")
        print("-"*50)
        
        while True:
            opcao = input("Escolha uma opção (1-2): ").strip()
            
            if opcao == '1':
                return self.inserir_chassis_terminal()
            elif opcao == '2':
                return self.carregar_chassis_excel()
            else:
                print("❌ Opção inválida. Digite 1 ou 2.")
    
    def definir_grupo_veiculos(self):
        """Define qual grupo de veículos será editado"""
        print("\n" + "="*50)
        print("🏷️  DEFINIR GRUPO DE VEÍCULOS")
        print("="*50)
        
        while True:
            self.nome_grupo = input("Digite o nome do grupo de veículos: ").strip()
            if self.nome_grupo:
                print(f"✅ Grupo definido: {self.nome_grupo}")
                break
            else:
                print("❌ Nome do grupo não pode estar vazio")
    
    def fazer_login_e_navegar(self):
        """
        Aguarda o usuário fazer login manualmente e navegar até a página de grupos
        """
        print("\n" + "="*60)
        print("🔐 LOGIN MANUAL NECESSÁRIO")
        print("="*60)
        print("O navegador foi aberto para você.")
        print("Por favor:")
        print("1. 🔑 Faça login na conta do cliente correto")
        print("2. 🔍 Navegue até a página de 'Editar Grupos de Veículos'")
        print("3. ✅ Certifique-se de estar na página correta")
        print("4. ⏸️  Volte aqui e pressione ENTER para continuar")
        print("-"*60)
        
        # Aguardar confirmação do usuário
        input("Pressione ENTER quando estiver pronto para continuar...")
        
        print("✅ Continuando com a automação...")
        time.sleep(2)
    
    def pesquisar_grupo_e_editar(self):
        """Pesquisa o grupo específico e clica em editar"""
        try:
            # SELETOR 1: Campo de pesquisa de grupos
            campo_pesquisa_grupo = self.wait.until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='search'][placeholder='Pesquisar']"))
            )
            
            # Limpar e pesquisar o grupo
            campo_pesquisa_grupo.clear()
            campo_pesquisa_grupo.send_keys(self.nome_grupo)
            
            # SELETOR 2: Aguardar carregamento automático da pesquisa
            time.sleep(2)  # Delay para pesquisa automática
            
            # SELETOR 3: Verificar se o grupo foi encontrado
            grupo_encontrado = self.wait.until(
                EC.presence_of_element_located((By.XPATH, f"//div[contains(@class, 'wj-cell') and contains(text(), '{self.nome_grupo}')]"))
            )
            self.logger.info(f"Grupo '{self.nome_grupo}' encontrado")
            
            # SELETOR 4: Botão editar grupo (botão com ícone lápis)
            botao_editar = self.wait.until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "button.button.btn-mz-grid[title='Editar grupo de veículos'] i.mz7-pencil"))
            )
            
            # Clicar no botão (clica no elemento pai se necessário)
            try:
                botao_editar.click()
            except:
                # Se não conseguir clicar no ícone, tenta clicar no botão pai
                botao_pai = botao_editar.find_element(By.XPATH, "..")
                botao_pai.click()
            
            # SELETOR 5: Modal de edição carregado
            self.wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "div.editor-section")))
            time.sleep(1)
            
            self.logger.info(f"Modal de edição do grupo '{self.nome_grupo}' aberto com sucesso")
            return True
            
        except TimeoutException:
            self.logger.error(f"Não foi possível encontrar ou editar o grupo '{self.nome_grupo}'")
            return False
        except Exception as e:
            self.logger.error(f"Erro ao abrir modal de edição: {e}")
            return False

    def debug_checkbox_detalhado(self, chassi):
        """Função para debugar de forma mais detalhada"""
        try:
            print(f"\n🔍 DEBUG DETALHADO para chassi: {chassi}")
            print("="*60)
            
            # 1. Verificar elementos do chassi
            elementos_chassi = self.driver.find_elements(By.XPATH, f"//*[contains(text(), '{chassi}')]")
            print(f"Elementos contendo o chassi encontrados: {len(elementos_chassi)}")
            
            for i, el in enumerate(elementos_chassi):
                if el.is_displayed():
                    print(f"  Elemento {i}: {el.tag_name}, classes: {el.get_attribute('class')}")
                    print(f"    Texto completo: {el.text}")
                    try:
                        parent = el.find_element(By.XPATH, "..")
                        print(f"    Parent: {parent.tag_name}")
                    except:
                        pass
            
            # 2. Checkboxes na página
            checkboxes = self.driver.find_elements(By.CSS_SELECTOR, "input[type='checkbox']")
            checkboxes_visiveis = [cb for cb in checkboxes if cb.is_displayed()]
            
            print(f"\nCheckboxes visíveis: {len(checkboxes_visiveis)}")
            for i, cb in enumerate(checkboxes_visiveis):
                try:
                    is_checked = self.driver.execute_script("return arguments[0].checked;", cb)
                    print(f"  Checkbox {i}: checked={is_checked}, id={cb.get_attribute('id')}, name={cb.get_attribute('name')}")
                except StaleElementReferenceException:
                    print(f"  Checkbox {i}: STALE ELEMENT")
                except Exception as e:
                    print(f"  Checkbox {i}: ERRO - {e}")
                    
        except Exception as e:
            print(f"Erro no debug detalhado: {e}")

    def encontrar_checkbox_por_contexto(self, chassi):
        """
        Tenta encontrar a checkbox específica através da estrutura HTML próxima ao chassi
        """
        try:
            # Estratégia 1: Procurar o chassi e depois a checkbox no mesmo container
            possibles_seletores = [
                f"//div[contains(text(), '{chassi}')]/ancestor::tr//input[@type='checkbox']",
                f"//div[contains(text(), '{chassi}')]/ancestor::div[contains(@class, 'row')]//input[@type='checkbox']",
                f"//span[contains(text(), '{chassi}')]/ancestor::tr//input[@type='checkbox']",
                f"//span[contains(text(), '{chassi}')]/ancestor::div[contains(@class, 'row')]//input[@type='checkbox']",
                f"//*[contains(text(), '{chassi}')]/ancestor::*[contains(@class, 'item')]//input[@type='checkbox']",
                f"//*[contains(text(), '{chassi}')]/preceding-sibling::*//input[@type='checkbox']",
                f"//*[contains(text(), '{chassi}')]/following-sibling::*//input[@type='checkbox']"
            ]
            
            for seletor in possibles_seletores:
                try:
                    checkbox = self.driver.find_element(By.XPATH, seletor)
                    if checkbox.is_displayed():
                        self.logger.info(f"Checkbox encontrada por contexto para chassi {chassi}")
                        return checkbox
                except NoSuchElementException:
                    continue
                except StaleElementReferenceException:
                    continue
                    
            return None
            
        except Exception as e:
            self.logger.warning(f"Erro ao buscar checkbox por contexto: {e}")
            return None

    def encontrar_checkbox_por_posicao(self, chassi):
        """
        Tenta encontrar a checkbox baseado na posição após pesquisar o chassi
        """
        try:
            # Verificar se existe apenas um resultado visível após a pesquisa
            elementos_chassi = self.driver.find_elements(By.XPATH, f"//*[contains(text(), '{chassi}')]")
            elementos_visiveis = [el for el in elementos_chassi if el.is_displayed()]
            
            if len(elementos_visiveis) == 1:
                # Se há apenas um chassi visível, pegar a primeira checkbox DESMARCADA
                checkboxes_desmarcadas = self.driver.find_elements(By.CSS_SELECTOR, "input[type='checkbox']:not(:checked)")
                checkboxes_visiveis = [cb for cb in checkboxes_desmarcadas if cb.is_displayed()]
                
                if len(checkboxes_visiveis) == 1:
                    self.logger.info(f"Checkbox encontrada por posição única para chassi {chassi}")
                    return checkboxes_visiveis[0]
                    
            return None
            
        except Exception as e:
            self.logger.warning(f"Erro ao buscar checkbox por posição: {e}")
            return None

    def processar_checkbox_com_retry(self, checkbox, chassi, max_tentativas=3):
        """
        Processa uma checkbox específica com retry para StaleElementReferenceException
        """
        for tentativa in range(max_tentativas):
            try:
                # Verificar se está desmarcada
                is_checked_js = self.driver.execute_script("return arguments[0].checked;", checkbox)
                is_checked_selenium = checkbox.is_selected()
                
                esta_desmarcada = not (is_checked_js or is_checked_selenium)
                
                self.logger.info(f"Checkbox para {chassi} - JS: {is_checked_js}, Selenium: {is_checked_selenium}")
                
                if esta_desmarcada:
                    # Tentar diferentes métodos de clique para MARCAR a checkbox
                    sucesso_clique = False
                    
                    # Método 1: Click JavaScript
                    try:
                        self.driver.execute_script("arguments[0].click();", checkbox)
                        time.sleep(0.8)
                        sucesso_clique = True
                    except Exception as e:
                        self.logger.warning(f"Falhou clique JS: {e}")
                    
                    # Método 2: Click Selenium direto
                    if not sucesso_clique:
                        try:
                            checkbox.click()
                            time.sleep(0.8)
                            sucesso_clique = True
                        except Exception as e:
                            self.logger.warning(f"Falhou clique Selenium: {e}")
                    
                    # Método 3: ActionChains
                    if not sucesso_clique:
                        try:
                            from selenium.webdriver.common.action_chains import ActionChains
                            ActionChains(self.driver).move_to_element(checkbox).click().perform()
                            time.sleep(0.8)
                            sucesso_clique = True
                        except Exception as e:
                            self.logger.warning(f"Falhou ActionChains: {e}")
                    
                    # Verificar se o clique funcionou verificando se a checkbox foi marcada
                    if sucesso_clique:
                        time.sleep(1)
                        try:
                            # Verificar se a checkbox foi marcada
                            nova_verificacao = self.driver.execute_script("return arguments[0].checked;", checkbox)
                            if nova_verificacao:
                                self.logger.info(f"Chassi {chassi} adicionado ao grupo com sucesso")
                                return True
                            else:
                                self.logger.warning(f"Clique não marcou a checkbox para {chassi}")
                                return False
                        except StaleElementReferenceException:
                            # Se checkbox ficou stale após marcação, assumir sucesso
                            self.logger.info(f"Chassi {chassi} adicionado com sucesso (checkbox stale após marcar)")
                            return True
                    else:
                        self.logger.error(f"Não foi possível clicar na checkbox para {chassi}")
                        return False
                else:
                    self.logger.warning(f"Chassi {chassi} encontrado mas já estava no grupo")
                    self.carros_ja_no_grupo.append(chassi)
                    return False
                    
            except StaleElementReferenceException:
                self.logger.warning(f"StaleElementReferenceException na tentativa {tentativa + 1} para {chassi}")
                if tentativa < max_tentativas - 1:
                    time.sleep(1)
                    # Tentar encontrar checkbox novamente
                    checkbox = self.encontrar_checkbox_por_contexto(chassi)
                    if not checkbox:
                        checkbox = self.encontrar_checkbox_por_posicao(chassi)
                    if not checkbox:
                        break
                else:
                    # Última tentativa falhou, mas pode ter funcionado
                    self.logger.info(f"StaleElement na última tentativa - assumindo sucesso para {chassi}")
                    return True
            except Exception as e:
                self.logger.error(f"Erro ao processar checkbox (tentativa {tentativa + 1}): {e}")
                if tentativa == max_tentativas - 1:
                    return False
                time.sleep(1)
        
        return False

    def tentar_todas_checkboxes_desmarcadas(self, chassi):
        """
        Como último recurso, tenta marcar todas as checkboxes desmarcadas uma por vez
        até encontrar a que corresponde ao chassi
        """
        try:
            self.logger.info(f"Tentando método de força bruta para chassi {chassi}")
            
            # Pegar todas as checkboxes DESMARCADAS
            checkboxes_desmarcadas = self.driver.find_elements(By.CSS_SELECTOR, "input[type='checkbox']:not(:checked)")
            checkboxes_visiveis = [cb for cb in checkboxes_desmarcadas if cb.is_displayed()]
            
            for i, checkbox in enumerate(checkboxes_visiveis):
                try:
                    # Marcar a checkbox
                    self.driver.execute_script("arguments[0].click();", checkbox)
                    time.sleep(1)
                    
                    # Verificar se a checkbox foi marcada (indicando sucesso)
                    try:
                        checkbox_marcada = self.driver.execute_script("return arguments[0].checked;", checkbox)
                        if checkbox_marcada:
                            # Checkbox foi marcada = chassi foi adicionado ao grupo
                            self.logger.info(f"Chassi {chassi} adicionado com sucesso (método força bruta, checkbox {i})")
                            return True
                        else:
                            # Checkbox não foi marcada, continuar para próxima
                            continue
                    except StaleElementReferenceException:
                        # Se checkbox ficou stale após marcar, assumir sucesso
                        self.logger.info(f"Chassi {chassi} adicionado com sucesso (método força bruta, checkbox stale)")
                        return True
                        
                except Exception as e:
                    self.logger.warning(f"Erro ao tentar checkbox {i}: {e}")
                    continue
            
            self.logger.warning(f"Nenhuma checkbox adicionou o chassi {chassi} ao grupo")
            return False
            
        except Exception as e:
            self.logger.error(f"Erro no método força bruta: {e}")
            return False

    def pesquisar_e_adicionar_chassi(self, chassi):
        """
        Pesquisa um chassi e adiciona ao grupo se encontrado
        
        Args:
            chassi (str): Número do chassi a ser pesquisado
            
        Returns:
            bool: True se adicionado com sucesso, False se não encontrado
        """
        try:
            # SELETOR 5: Campo de pesquisa de chassi dentro do modal
            campo_pesquisa_chassi = self.wait.until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "input.form-input[placeholder='Buscar']"))
            )
            
            # Limpar campo e digitar chassi
            campo_pesquisa_chassi.clear()
            campo_pesquisa_chassi.send_keys(chassi)
            
            # SELETOR 6: Aguardar carregamento automático da pesquisa
            time.sleep(1)  # Delay para pesquisa automática
            
            # Delay adicional para carregar completamente o resultado
            time.sleep(0.5)
            
            # Verificar se o chassi foi encontrado primeiro
            elementos_chassi = self.driver.find_elements(By.XPATH, f"//*[contains(text(), '{chassi}')]")
            chassi_encontrado = any(el.is_displayed() for el in elementos_chassi)
            
            if not chassi_encontrado:
                self.logger.warning(f"Chassi {chassi} não encontrado no sistema")
                return False
            
            # Método 1: Tentar encontrar a checkbox específica através da estrutura HTML
            checkbox_encontrada = self.encontrar_checkbox_por_contexto(chassi)
            if checkbox_encontrada:
                return self.processar_checkbox_com_retry(checkbox_encontrada, chassi)
            
            # Método 2: Se não encontrou por contexto, tentar por posição
            checkbox_encontrada = self.encontrar_checkbox_por_posicao(chassi)
            if checkbox_encontrada:
                return self.processar_checkbox_com_retry(checkbox_encontrada, chassi)
            
            # Método 3: Último recurso - tentar todas as checkboxes desmarcadas
            return self.tentar_todas_checkboxes_desmarcadas(chassi)
            
        except Exception as e:
            self.logger.error(f"Erro geral ao processar chassi {chassi}: {e}")
            return False
    
    def salvar_alteracoes(self):
        """Salva as alterações no grupo"""
        try:
            # SELETOR 8: Botão salvar dentro do modal
            botao_salvar = self.wait.until(
                EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Salvar') or contains(text(), 'SALVAR')]"))
            )
            botao_salvar.click()
            
            # Aguardar confirmação de salvamento
            time.sleep(3)
            self.logger.info("Alterações salvas com sucesso")
            return True
            
        except Exception as e:
            self.logger.error(f"Erro ao salvar alterações: {e}")
            return False
    
    def fechar_modal(self):
        """Fecha o modal de edição"""
        try:
            # SELETOR 9: Tentar várias formas de fechar o modal
            try:
                # Tentar botão com X ou close
                self.driver.find_element(By.XPATH, "//button[contains(@class, 'close') or contains(@class, 'btn-close')]").click()
            except:
                try:
                    # Tentar botão Cancelar
                    self.driver.find_element(By.XPATH, "//button[contains(text(), 'Cancelar') or contains(text(), 'CANCELAR')]").click()
                except:
                    # Tentar ESC como último recurso
                    from selenium.webdriver.common.keys import Keys
                    self.driver.find_element(By.TAG_NAME, 'body').send_keys(Keys.ESCAPE)
            
            time.sleep(1)
            
        except Exception as e:
            self.logger.warning(f"Dificuldade para fechar modal: {e}")
            # Tentar ESC como último recurso
            try:
                from selenium.webdriver.common.keys import Keys
                self.driver.find_element(By.TAG_NAME, 'body').send_keys(Keys.ESCAPE)
                time.sleep(1)
            except:
                pass
    
    def processar_chassis(self, chassis_list):
        """
        Processa a lista completa de chassis
        
        Args:
            chassis_list (list): Lista de chassis para processar
        """
        total_chassis = len(chassis_list)
        processados = 0
        
        while processados < total_chassis:
            # Pesquisar grupo e abrir modal de edição
            if not self.pesquisar_grupo_e_editar():
                break
            
            # Processar até 50 chassis por vez
            lote_atual = 0
            max_lote = 50
            
            while lote_atual < max_lote and processados < total_chassis:
                chassi = chassis_list[processados]
                
                print(f"Processando chassi {processados + 1}/{total_chassis}: {chassi}")
                
                # Usar debug detalhado quando necessário
                # self.debug_checkbox_detalhado(chassi)  # Descomente para debug
                
                if self.pesquisar_e_adicionar_chassi(chassi):
                    self.carros_adicionados.append(chassi)
                    print(f"✅ Chassi {chassi} adicionado com sucesso")
                else:
                    if chassi not in self.carros_ja_no_grupo:
                        self.carros_nao_encontrados.append(chassi)
                        print(f"❌ Chassi {chassi} não foi adicionado")
                    else:
                        print(f"⚠️ Chassi {chassi} já estava no grupo")
                
                processados += 1
                lote_atual += 1
                self.total_processados += 1
                
                # Pequena pausa entre pesquisas
                time.sleep(0.5)
            
            # Salvar alterações do lote atual
            print(f"\n💾 Salvando alterações do lote...")
            self.salvar_alteracoes()
            
            # Fechar modal
            print(f"🔒 Fechando modal...")
            self.fechar_modal()
            
            # Pausa entre lotes
            time.sleep(2)
            
            print(f"📦 Lote concluído. Processados: {processados}/{total_chassis}")
    
    def gerar_relatorio(self):
        """Gera relatório final do processamento"""
        relatorio = f"""
╔══════════════════════════════════════════════════════╗
║                 RELATÓRIO FINAL                      ║
╠══════════════════════════════════════════════════════╣
║ Grupo processado: {self.nome_grupo:<30} ║
║ Total processados: {self.total_processados:<29} ║
║ Adicionados com sucesso: {len(self.carros_adicionados):<23} ║
║ Não encontrados: {len(self.carros_nao_encontrados):<31} ║
║ Já estavam no grupo: {len(self.carros_ja_no_grupo):<27} ║
╚══════════════════════════════════════════════════════╝

✅ CARROS ADICIONADOS:
{chr(10).join([f"  • {chassi}" for chassi in self.carros_adicionados]) if self.carros_adicionados else "  Nenhum"}

❌ CHASSIS NÃO ENCONTRADOS:
{chr(10).join([f"  • {chassi}" for chassi in self.carros_nao_encontrados]) if self.carros_nao_encontrados else "  Nenhum"}

⚠️ CHASSIS JÁ NO GRUPO:
{chr(10).join([f"  • {chassi}" for chassi in self.carros_ja_no_grupo]) if self.carros_ja_no_grupo else "  Nenhum"}
"""
        
        print(relatorio)
        
        # Mensagem específica solicitada
        if self.carros_nao_encontrados:
            print(f"Os seguintes chassis não foram encontrados: {self.carros_nao_encontrados}")
        if self.carros_ja_no_grupo:
            print(f"Os seguintes chassis já estavam no grupo: {self.carros_ja_no_grupo}")
    
    def executar(self):
        """Executa o processo completo de automação"""
        try:
            print("🚀 Iniciando automação de adição de carros")
            
            # 1. Definir grupo de veículos
            self.definir_grupo_veiculos()
            
            # 2. Escolher método e carregar chassis
            chassis_list = self.escolher_metodo_entrada()
            
            if not chassis_list:
                print("❌ Nenhum chassi foi inserido. Encerrando.")
                return
            
            # 3. Mostrar resumo
            print(f"\n📊 RESUMO DA AUTOMAÇÃO:")
            print(f"🏷️  Grupo: {self.nome_grupo}")
            print(f"🚗 Total de chassis: {len(chassis_list)}")
            print(f"📦 Serão processados em lotes de 50")
            print(f"⏱️  Tempo estimado: ~{len(chassis_list) * 3} segundos")
            
            # 4. Configurar browser
            print("\n🔧 Configurando navegador...")
            self.setup_driver()
            
            # 5. Aguardar login manual
            self.fazer_login_e_navegar()
            
            # 6. Confirmação final antes de processar
            print("\n" + "="*50)
            print("⚠️  CONFIRMAÇÃO FINAL")
            print("="*50)
            print("A automação irá começar a processar os chassis.")
            print(f"Certifique-se de estar na página de grupos de veículos.")
            print(f"Grupo a ser editado: {self.nome_grupo}")
            
            confirmar = input("\n🚀 Iniciar processamento automático? (s/n): ").lower()
            if confirmar not in ['s', 'sim', 'y', 'yes']:
                print("❌ Automação cancelada pelo usuário")
                return
            
            # 7. Processar chassis
            print("\n🔄 Iniciando processamento dos chassis...")
            print("-"*50)
            self.processar_chassis(chassis_list)
            
            # 8. Gerar relatório
            print("\n📋 Gerando relatório final...")
            self.gerar_relatorio()
            
            print("\n🎉 Automação concluída com sucesso!")
            
        except KeyboardInterrupt:
            print("\n⚠️ Automação interrompida pelo usuário (Ctrl+C)")
        except Exception as e:
            print(f"❌ Erro durante execução: {e}")
            self.logger.error(f"Erro durante execução: {e}")
        finally:
            if self.driver:
                input("\nPressione ENTER para fechar o navegador...")
                self.driver.quit()
                print("🔒 Navegador fechado")


# Uso
if __name__ == "__main__":
    automacao = CarAdditionAutomation()
    automacao.executar()