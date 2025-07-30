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
        self.driver.get("https://live.mzoneweb.net/mzonex/maintenance/vehiclegroups")
    
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
    
    def fazer_login_inicial(self):
        """
        Aguarda o usuário fazer login manualmente na primeira vez
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
    
    def recarregar_pagina(self):
        """Recarrega a página e aguarda o carregamento"""
        try:
            print("🔄 Recarregando página...")
            self.driver.refresh()
            
            # Aguardar página carregar completamente
            time.sleep(9)
            
            # Aguardar overlay desaparecer (se existir)
            try:
                WebDriverWait(self.driver, 15).until(
                    EC.invisibility_of_element_located((By.CSS_SELECTOR, "div.form-overlay.ng-star-inserted.visible"))
                )
            except TimeoutException:
                pass
            
            print("✅ Página recarregada com sucesso")
            return True
            
        except Exception as e:
            self.logger.error(f"Erro ao recarregar página: {e}")
            return False

    def pesquisar_grupo(self):
        """Pesquisa o grupo específico na lista"""
        try:
            print(f"🔍 Pesquisando grupo '{self.nome_grupo}'...")
            
            # Campo de pesquisa de grupos
            campo_pesquisa_grupo = self.wait.until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='search'][placeholder='Pesquisar']"))
            )
            
            # Limpar e pesquisar o grupo
            campo_pesquisa_grupo.clear()
            time.sleep(0.5)
            campo_pesquisa_grupo.send_keys(self.nome_grupo)
            
            # Aguardar carregamento automático da pesquisa
            time.sleep(2)
            
            # Verificar se o grupo foi encontrado
            grupo_encontrado = self.wait.until(
                EC.presence_of_element_located((By.XPATH, f"//div[contains(@class, 'wj-cell') and contains(text(), '{self.nome_grupo}')]"))
            )
            self.logger.info(f"Grupo '{self.nome_grupo}' encontrado")
            return True
            
        except TimeoutException:
            self.logger.error(f"Grupo '{self.nome_grupo}' não foi encontrado na pesquisa")
            return False
        except Exception as e:
            self.logger.error(f"Erro ao pesquisar grupo: {e}")
            return False

    def clicar_editar_grupo(self):
        """Clica no botão de editar do grupo"""
        try:
            print("✏️ Clicando em editar grupo...")
            
            # Botão editar grupo (botão com ícone lápis)
            botao_editar = self.wait.until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "button.button.btn-mz-grid[title='Editar grupo de veículos'] i.mz7-pencil"))
            )
            
            time.sleep(1)
            
            # Verificar se não há overlay bloqueando
            try:
                overlay_presente = self.driver.find_element(By.CSS_SELECTOR, "div.form-overlay.ng-star-inserted.visible")
                if overlay_presente.is_displayed():
                    print("⚠️ Overlay ainda presente, aguardando...")
                    WebDriverWait(self.driver, 10).until(
                        EC.invisibility_of_element_located((By.CSS_SELECTOR, "div.form-overlay.ng-star-inserted.visible"))
                    )
            except NoSuchElementException:
                pass
            
            # Clicar no botão
            try:
                botao_editar.click()
            except:
                botao_pai = botao_editar.find_element(By.XPATH, "..")
                botao_pai.click()
            
            # Modal de edição carregado
            print("⏱️ Aguardando modal de edição carregar...")
            self.wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "div.editor-section")))
            time.sleep(1)
            
            self.logger.info(f"Modal de edição do grupo '{self.nome_grupo}' aberto com sucesso")
            return True
            
        except TimeoutException:
            self.logger.error("Não foi possível clicar no botão editar ou abrir o modal")
            return False
        except Exception as e:
            self.logger.error(f"Erro ao clicar em editar: {e}")
            return False

    def encontrar_checkbox_por_contexto(self, chassi):
        """
        Tenta encontrar a checkbox específica através da estrutura HTML próxima ao chassi
        """
        try:
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
                except (NoSuchElementException, StaleElementReferenceException):
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
                # Para ADIÇÃO: Verificar se está DESMARCADA (precisa marcar para adicionar)
                is_checked_js = self.driver.execute_script("return arguments[0].checked;", checkbox)
                
                if not is_checked_js:  # Se está desmarcada, pode adicionar
                    # Tentar clicar na checkbox para MARCAR (adicionar ao grupo)
                    self.driver.execute_script("arguments[0].click();", checkbox)
                    time.sleep(0.8)
                    
                    # Verificar se chassi foi adicionado (checkbox marcada)
                    time.sleep(1)
                    try:
                        nova_verificacao = self.driver.execute_script("return arguments[0].checked;", checkbox)
                        if nova_verificacao:
                            self.logger.info(f"Chassi {chassi} adicionado ao grupo com sucesso")
                            return True
                        else:
                            self.logger.warning(f"Clique não adicionou o chassi {chassi}")
                            return False
                    except StaleElementReferenceException:
                        self.logger.info(f"Chassi {chassi} adicionado com sucesso (checkbox stale)")
                        return True
                else:
                    self.logger.warning(f"Chassi {chassi} encontrado mas já estava no grupo")
                    return False
                    
            except StaleElementReferenceException:
                self.logger.warning(f"StaleElementReferenceException na tentativa {tentativa + 1} para {chassi}")
                if tentativa < max_tentativas - 1:
                    time.sleep(1)
                    checkbox = self.encontrar_checkbox_por_contexto(chassi)
                    if not checkbox:
                        checkbox = self.encontrar_checkbox_por_posicao(chassi)
                    if not checkbox:
                        break
                else:
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
            
            # Pegar todas as checkboxes DESMARCADAS (para adicionar)
            checkboxes_desmarcadas = self.driver.find_elements(By.CSS_SELECTOR, "input[type='checkbox']:not(:checked)")
            checkboxes_visiveis = [cb for cb in checkboxes_desmarcadas if cb.is_displayed()]
            
            for i, checkbox in enumerate(checkboxes_visiveis):
                try:
                    # Marcar a checkbox (adicionar ao grupo)
                    self.driver.execute_script("arguments[0].click();", checkbox)
                    time.sleep(1)
                    
                    # Verificar se a checkbox foi marcada (indicando sucesso)
                    try:
                        checkbox_marcada = self.driver.execute_script("return arguments[0].checked;", checkbox)
                        if checkbox_marcada:
                            self.logger.info(f"Chassi {chassi} adicionado com sucesso (método força bruta, checkbox {i})")
                            return True
                        else:
                            continue
                    except StaleElementReferenceException:
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
        """
        try:
            # Campo de pesquisa de chassi dentro do modal
            campo_pesquisa_chassi = self.wait.until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "input.form-input[placeholder='Buscar']"))
            )
            
            # Limpar campo e digitar chassi
            campo_pesquisa_chassi.clear()
            campo_pesquisa_chassi.send_keys(chassi)
            
            # Aguardar carregamento automático da pesquisa
            time.sleep(2)
            
            # Verificar se o chassi foi encontrado
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
        """Salva as alterações no grupo usando o seletor correto"""
        try:
            print("💾 Salvando alterações...")
            
            # Primeiro, tentar pelo XPath mais específico com o texto exato
            try:
                botao_salvar = self.wait.until(
                    EC.element_to_be_clickable((By.XPATH, "//button[@type='submit' and contains(@class, 'button') and contains(@class, 'btn-mz') and contains(@class, 'success') and normalize-space(text())='Salvar']"))
                )
                self.logger.info("Botão 'Salvar' encontrado pelo XPath com texto")
            except TimeoutException:
                # Se não encontrar pelo texto, tentar pelo seletor CSS corrigido
                try:
                    botao_salvar = self.wait.until(
                        EC.element_to_be_clickable((By.CSS_SELECTOR, "button[type='submit'].button.btn-mz.success.ng-star-inserted"))
                    )
                    # Verificar se é realmente o botão "Salvar" pelo texto
                    texto_botao = botao_salvar.text.strip()
                    if "Salvar" not in texto_botao or "criar" in texto_botao.lower():
                        self.logger.warning(f"Botão encontrado não é o correto. Texto: '{texto_botao}'")
                        raise TimeoutException("Botão incorreto encontrado")
                    self.logger.info("Botão 'Salvar' encontrado pelo CSS selector")
                except TimeoutException:
                    # Última tentativa: buscar todos os botões submit e filtrar
                    botoes_submit = self.driver.find_elements(By.CSS_SELECTOR, "button[type='submit']")
                    botao_salvar = None
                    
                    for botao in botoes_submit:
                        try:
                            texto = botao.text.strip()
                            if texto == "Salvar" and "success" in botao.get_attribute("class"):
                                botao_salvar = botao
                                self.logger.info(f"Botão 'Salvar' encontrado por busca manual: '{texto}'")
                                break
                        except:
                            continue
                    
                    if not botao_salvar:
                        raise TimeoutException("Botão 'Salvar' não encontrado em nenhuma tentativa")
            
            # Verificar se o botão está realmente visível e clicável
            if not botao_salvar.is_displayed():
                self.logger.error("Botão 'Salvar' encontrado mas não está visível")
                return False
            
            # Log do botão que será clicado
            texto_final = botao_salvar.text.strip()
            classes_final = botao_salvar.get_attribute("class")
            self.logger.info(f"Clicando no botão: '{texto_final}' com classes: '{classes_final}'")
            
            # Clicar no botão
            try:
                botao_salvar.click()
            except Exception as e:
                # Se o clique normal falhar, tentar com JavaScript
                self.logger.warning(f"Clique normal falhou, tentando com JavaScript: {e}")
                self.driver.execute_script("arguments[0].click();", botao_salvar)
            
            # Aguardar confirmação de salvamento e fechamento do modal
            print("⏱️ Aguardando confirmação do salvamento...")
            time.sleep(3)
            
            # Verificar se o modal foi fechado (isso indica que salvou com sucesso)
            try:
                WebDriverWait(self.driver, 8).until(
                    EC.invisibility_of_element_located((By.CSS_SELECTOR, "div.editor-section"))
                )
                self.logger.info("Modal fechado - alterações salvas com sucesso")
                return True
            except TimeoutException:
                # Se o modal ainda estiver aberto, pode ser que ainda esteja processando
                print("⚠️ Modal ainda aberto, aguardando mais tempo...")
                time.sleep(5)
                try:
                    WebDriverWait(self.driver, 5).until(
                        EC.invisibility_of_element_located((By.CSS_SELECTOR, "div.editor-section"))
                    )
                    self.logger.info("Modal fechado após tempo adicional - alterações salvas")
                    return True
                except TimeoutException:
                    self.logger.warning("Modal ainda aberto após salvamento - pode haver erro")
                    return False
            
        except TimeoutException:
            self.logger.error("Botão salvar não encontrado ou não clicável")
            # Listar todos os botões disponíveis para debug
            try:
                botoes = self.driver.find_elements(By.TAG_NAME, "button")
                self.logger.info("Botões disponíveis no modal:")
                for i, botao in enumerate(botoes):
                    try:
                        texto = botao.text.strip()
                        classes = botao.get_attribute("class")
                        tipo = botao.get_attribute("type")
                        if texto or "btn" in classes:
                            self.logger.info(f"  Botão {i}: '{texto}' | type='{tipo}' | class='{classes}'")
                    except:
                        pass
            except:
                pass
            return False
        except Exception as e:
            self.logger.error(f"Erro ao salvar alterações: {e}")
            return False
    
    def processar_lote(self, chassis_lote, numero_lote):
        """
        Processa um lote de até 60 chassis
        
        Args:
            chassis_lote (list): Lista de chassis para processar neste lote
            numero_lote (int): Número do lote atual
            
        Returns:
            bool: True se o lote foi processado com sucesso
        """
        print(f"\n🔄 PROCESSANDO LOTE {numero_lote}")
        print("="*50)
        print(f"📦 Chassis neste lote: {len(chassis_lote)}")
        
        try:
            # 1. Pesquisar o grupo
            if not self.pesquisar_grupo():
                print("❌ Não foi possível encontrar o grupo")
                return False
            
            # 2. Clicar em editar para abrir o modal
            if not self.clicar_editar_grupo():
                print("❌ Não foi possível abrir o modal de edição")
                return False
            
            print("✅ Modal aberto, processando chassis...")
            
            # 3. Processar cada chassi do lote
            for i, chassi in enumerate(chassis_lote, 1):
                print(f"  Processando {i}/{len(chassis_lote)}: {chassi}")
                
                resultado = self.pesquisar_e_adicionar_chassi(chassi)
                
                if resultado:
                    self.carros_adicionados.append(chassi)
                    print(f"    ✅ Adicionado")
                elif chassi in self.carros_ja_no_grupo:
                    print(f"    ⚠️ Já estava no grupo")
                else:
                    self.carros_nao_encontrados.append(chassi)
                    print(f"    ❌ Não adicionado")
                
                self.total_processados += 1
                time.sleep(0.5)  # Pequena pausa entre pesquisas
            
            # 4. Salvar alterações do lote
            print(f"\n💾 Salvando alterações do lote {numero_lote}...")
            if self.salvar_alteracoes():
                print(f"✅ Lote {numero_lote} salvo com sucesso")
                time.sleep(2)  # Aguardar processamento do servidor
                return True
            else:
                print(f"⚠️ Problema ao salvar lote {numero_lote}")
                return False
                
        except Exception as e:
            self.logger.error(f"Erro ao processar lote {numero_lote}: {e}")
            return False
    
    def processar_todos_chassis(self, chassis_list):
        """
        Processa todos os chassis dividindo em lotes de 60
        """
        total_chassis = len(chassis_list)
        tamanho_lote = 50
        numero_lote = 1
        
        # Dividir chassis em lotes
        for i in range(0, total_chassis, tamanho_lote):
            chassis_lote = chassis_list[i:i + tamanho_lote]
            
            print(f"\n📊 Progresso: {min(i + tamanho_lote, total_chassis)}/{total_chassis} chassis")
            
            # Processar o lote atual
            sucesso_lote = self.processar_lote(chassis_lote, numero_lote)
            
            if not sucesso_lote:
                print(f"❌ Erro no lote {numero_lote}. Continuando para próximo lote...")
            
            # Se não é o último lote, recarregar a página
            if i + tamanho_lote < total_chassis:
                print(f"⏱️ Aguardando 5 segundos antes de recarregar...")
                time.sleep(5)
                
                if not self.recarregar_pagina():
                    print("❌ Erro ao recarregar página. Tentando continuar...")
                    time.sleep(2)
            
            numero_lote += 1
        
        print(f"\n🎉 Todos os lotes processados!")
    
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
            total_lotes = (len(chassis_list) + 59) // 60  # Arredonda para cima
            print(f"\n📊 RESUMO DA AUTOMAÇÃO:")
            print(f"🏷️  Grupo: {self.nome_grupo}")
            print(f"🚗 Total de chassis: {len(chassis_list)}")
            print(f"📦 Serão processados em {total_lotes} lotes de até 60 chassis")
            print(f"🔄 A página será recarregada entre cada lote")
            
            # 4. Configurar browser
            print("\n🔧 Configurando navegador...")
            self.setup_driver()
            
            # 5. Aguardar login manual
            self.fazer_login_inicial()
            
            # 6. Confirmação final
            print("\n" + "="*50)
            print("⚠️  CONFIRMAÇÃO FINAL")
            print("="*50)
            print("A automação irá processar os chassis em lotes.")
            print(f"Grupo a ser editado: {self.nome_grupo}")
            print(f"Total de lotes: {total_lotes}")
            
            confirmar = input("\n🚀 Iniciar processamento? (s/n): ").lower()
            if confirmar not in ['s', 'sim', 'y', 'yes']:
                print("❌ Automação cancelada pelo usuário")
                return
            
            # 7. Processar todos os chassis
            print("\n🔄 Iniciando processamento...")
            self.processar_todos_chassis(chassis_list)
            
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


def main():
    """Função principal da automação"""
    automacao = CarAdditionAutomation()
    automacao.executar()

if __name__ == '__main__':
    main()