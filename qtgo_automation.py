import pandas as pd
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import TimeoutException, NoSuchElementException, WebDriverException
import os

class ChassisAutomation:
    def __init__(self):
        self.driver = None
        self.processed_chassis = []
        self.failed_chassis = []
        self.successful_chassis = []
        
    def setup_driver(self):
        """Configura o driver do Chrome"""
        options = webdriver.ChromeOptions()
        options.add_argument("--disable-blink-features=AutomationControlled")
        options.add_experimental_option("excludeSwitches", ["enable-automation"])
        options.add_experimental_option('useAutomationExtension', False)
        
        self.driver = webdriver.Chrome(options=options)
        self.driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
        self.driver.maximize_window()
        
    def load_chassis_list(self):
        """Carrega a lista de chassis do Excel ou input manual"""
        print("=== CARREGAMENTO DE CHASSIS ===")
        print("1. Carregar de arquivo Excel (QTGO_ID.xlsx)")
        print("2. Inserir lista manualmente")
        
        choice = input("\nEscolha uma opção (1 ou 2): ").strip()
        
        chassis_list = []
        
        if choice == "1":
            # Arquivo fixo na raiz do projeto
            file_path = "QTGO_ID.xlsx"
            
            try:
                if not os.path.exists(file_path):
                    print(f"❌ Arquivo não encontrado: {file_path}")
                    print("📁 Certifique-se de que o arquivo QTGO_ID.xlsx está na raiz do projeto")
                    return self.load_chassis_list()
                
                print(f"📊 Carregando arquivo: {file_path}")
                df = pd.read_excel(file_path)
                
                print("📋 Colunas disponíveis no Excel:")
                for i, col in enumerate(df.columns):
                    print(f"   {i+1}. {col}")
                
                col_choice = input("\n🔢 Digite o número da coluna que contém os chassis: ").strip()
                try:
                    col_index = int(col_choice) - 1
                    column_name = df.columns[col_index]
                    chassis_list = df[column_name].dropna().astype(str).tolist()
                    print(f"✅ Coluna selecionada: {column_name}")
                except (ValueError, IndexError):
                    print("❌ Opção inválida!")
                    return self.load_chassis_list()
                    
            except Exception as e:
                print(f"❌ Erro ao ler o arquivo Excel: {e}")
                print("💡 Verifique se o arquivo não está aberto em outro programa")
                return self.load_chassis_list()
                
        elif choice == "2":
            print("📝 Cole a lista de chassis (um por linha). Digite 'FIM' para finalizar:")
            while True:
                chassis = input("   > ").strip()
                if chassis.upper() == 'FIM':
                    break
                if chassis:
                    chassis_list.append(chassis)
        else:
            print("❌ Opção inválida!")
            return self.load_chassis_list()
        
        chassis_list = [chassis.strip() for chassis in chassis_list if chassis.strip()]
        print(f"\n✅ {len(chassis_list)} chassis carregados com sucesso!")
        
        # Mostra preview dos primeiros chassis carregados
        if chassis_list:
            print("📋 Preview dos chassis carregados:")
            for i, chassis in enumerate(chassis_list[:5]):
                print(f"   {i+1}. {chassis}")
            if len(chassis_list) > 5:
                print(f"   ... e mais {len(chassis_list) - 5} chassis")
        
        return chassis_list
    
    def wait_for_manual_login(self):
        """Abre o sistema e aguarda login manual"""
        print(f"\n=== ABRINDO SISTEMA ===")
        self.driver.get("https://quantigo.scopemp.net/app/subscriptions")
        
        print("🌐 Sistema aberto no navegador.")
        print("🔑 Faça login manualmente e aguarde a página de subscriptions carregar.")
        input("⏳ Pressione ENTER quando estiver na página de subscriptions e pronto para iniciar a automação...")
        
    def search_chassis(self, chassis):
        """Pesquisa um chassis específico"""
        try:
            # Localiza o campo de busca (ID pode variar, então usa placeholder como backup)
            search_input = None
            try:
                search_input = WebDriverWait(self.driver, 10).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "input.mat-input-element[placeholder='Search']"))
                )
            except:
                # Tenta com ID específico se não encontrar pelo placeholder
                search_input = self.driver.find_element(By.ID, "mat-input-3")
            
            # Limpa o campo e digita o chassis
            search_input.clear()
            time.sleep(0.5)
            search_input.send_keys(chassis)
            time.sleep(0.5)
            print(f"🔍 Chassis '{chassis}' digitado no campo de busca")
            
            # Localiza e clica no botão de pesquisa ESPECÍFICO
            try:
                # Método 1: Busca o botão que contém especificamente o ícone "search"
                search_button = WebDriverWait(self.driver, 5).until(
                    EC.element_to_be_clickable((By.XPATH, "//button[contains(@class, 'mat-icon-button')]//mat-icon[text()='search']/ancestor::button"))
                )
                search_button.click()
                print(f"🔍 Botão de busca clicado (método 1) para chassis: {chassis}")
            except:
                try:
                    # Método 2: Busca dentro da div quantigo-table-actions o botão com ícone search
                    search_button = self.driver.find_element(By.CSS_SELECTOR, ".quantigo-table-actions button[mat-icon-button] mat-icon[role='img'][aria-hidden='true']")
                    # Clica no botão pai (o button)
                    search_button = search_button.find_element(By.XPATH, "./ancestor::button")
                    search_button.click()
                    print(f"🔍 Botão de busca clicado (método 2) para chassis: {chassis}")
                except:
                    try:
                        # Método 3: JavaScript click no botão específico que contém "search"
                        search_button = self.driver.find_element(By.XPATH, "//div[contains(@class, 'quantigo-table-actions')]//button//mat-icon[contains(text(), 'search')]/parent::*/parent::button")
                        self.driver.execute_script("arguments[0].click();", search_button)
                        print(f"🔍 Botão de busca clicado (método 3 - JavaScript) para chassis: {chassis}")
                    except:
                        # Método 4: Busca o segundo botão mat-icon-button (assumindo que o primeiro é o menu)
                        search_buttons = self.driver.find_elements(By.CSS_SELECTOR, "button.mat-icon-button")
                        if len(search_buttons) >= 2:
                            # Verifica qual contém o ícone search
                            for btn in search_buttons:
                                try:
                                    icon = btn.find_element(By.CSS_SELECTOR, "mat-icon")
                                    if icon.text == "search":
                                        btn.click()
                                        print(f"🔍 Botão de busca clicado (método 4) para chassis: {chassis}")
                                        break
                                except:
                                    continue
                        else:
                            raise Exception("Não foi possível encontrar o botão de search")
            
            # IMPORTANTE: Aguarda o loading terminar antes de continuar
            print(f"⏳ Aguardando sistema carregar resultados para chassis: {chassis}...")
            self.wait_for_loading_to_finish()
            
            print(f"✅ Pesquisa do chassis {chassis} concluída")
            return True
            
        except Exception as e:
            print(f"❌ Erro ao pesquisar chassis {chassis}: {e}")
            # Método alternativo com ENTER
            try:
                print("🔄 Tentando método alternativo com ENTER...")
                if search_input:
                    search_input.send_keys(Keys.ENTER)
                else:
                    search_input = self.driver.find_element(By.CSS_SELECTOR, "input.mat-input-element[placeholder='Search']")
                    search_input.clear()
                    search_input.send_keys(chassis)
                    search_input.send_keys(Keys.ENTER)
                
                # IMPORTANTE: Aguarda o loading terminar mesmo no método alternativo
                print(f"⏳ Aguardando sistema carregar resultados (método alternativo) para chassis: {chassis}...")
                self.wait_for_loading_to_finish()
                
                print(f"✅ Chassis {chassis} pesquisado (método alternativo com ENTER)")
                return True
            except Exception as e2:
                print(f"❌ Falha também no método alternativo: {e2}")
                return False
    
    def wait_for_loading_to_finish(self, timeout=30):
        """Aguarda o loading spinner desaparecer"""
        try:
            print("🔄 Verificando se há loading ativo...")
            
            # Aguarda um pouco para o loading aparecer se necessário
            time.sleep(1)
            
            # Verifica se existe o spinner de loading
            loading_spinner = self.driver.find_elements(By.CSS_SELECTOR, ".quantigo-progress-spinner")
            
            if loading_spinner:
                print("⏳ Loading detectado, aguardando finalizar...")
                # Aguarda o spinner desaparecer
                WebDriverWait(self.driver, timeout).until(
                    EC.invisibility_of_element_located((By.CSS_SELECTOR, ".quantigo-progress-spinner"))
                )
                print("✅ Loading finalizado!")
                
                # Aguarda mais um pouco para garantir que a página está completamente carregada
                time.sleep(2)
            else:
                print("✅ Nenhum loading detectado")
                
            return True
            
        except TimeoutException:
            print(f"⚠️ Timeout aguardando loading finalizar ({timeout}s)")
            return False
        except Exception as e:
            print(f"❌ Erro ao aguardar loading: {e}")
            return False
    
    def find_active_records(self):
        """Encontra registros com status Active na tabela"""
        try:
            # Aguarda a tabela carregar completamente
            WebDriverWait(self.driver, 15).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "table.mat-table"))
            )
            
            # Aguarda um pouco mais para garantir que os dados carregaram
            time.sleep(2)
            
            # Procura por todas as linhas da tabela usando o seletor correto
            table_rows = self.driver.find_elements(By.CSS_SELECTOR, "tr.material-table-row.mat-row")
            
            active_rows = []
            
            for row in table_rows:
                # Procura pela célula de status na linha usando o seletor correto
                try:
                    status_cell = row.find_element(By.CSS_SELECTOR, "td.mat-cell-bold.mat-cell.cdk-column-status")
                    status_text = status_cell.text.strip()
                    print(f"📊 Status encontrado: '{status_text}'")
                    
                    if status_text == "Active":
                        active_rows.append(row)
                        print(f"✅ Registro ativo encontrado!")
                        
                        # Debug: mostra informações da linha
                        try:
                            description_cell = row.find_element(By.CSS_SELECTOR, "td.cdk-column-description")
                            print(f"📝 Descrição: {description_cell.text[:50]}...")
                        except:
                            pass
                            
                except NoSuchElementException:
                    # Tenta seletor alternativo para status
                    try:
                        status_cell = row.find_element(By.CSS_SELECTOR, "td.cdk-column-status")
                        status_text = status_cell.text.strip()
                        print(f"📊 Status encontrado (alternativo): '{status_text}'")
                        
                        if status_text == "Active":
                            active_rows.append(row)
                            print(f"✅ Registro ativo encontrado (seletor alternativo)!")
                    except NoSuchElementException:
                        print("⚠️ Não foi possível encontrar célula de status nesta linha")
                        continue
            
            print(f"📈 Total de registros com status Active: {len(active_rows)}")
            return active_rows
            
        except Exception as e:
            print(f"❌ Erro ao procurar registros ativos: {e}")
            return []
    
    def wait_for_modal_to_appear(self, timeout=15):
        """Aguarda especificamente o modal de desinstalação aparecer"""
        try:
            print("⏳ Aguardando modal de desinstalação aparecer...")
            
            # Aguarda o campo de location aparecer (indicativo de que o modal carregou)
            WebDriverWait(self.driver, timeout).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "input[formcontrolname='location']"))
            )
            
            # Aguarda um pouco mais para garantir que o modal carregou completamente
            time.sleep(1)
            
            # Verifica se realmente é visível
            location_field = self.driver.find_element(By.CSS_SELECTOR, "input[formcontrolname='location']")
            if location_field.is_displayed():
                print("✅ Modal de desinstalação apareceu e está visível")
                return True
            else:
                print("⚠️ Modal encontrado mas não está visível")
                return False
                
        except TimeoutException:
            print(f"⚠️ Timeout: Modal não apareceu em {timeout} segundos")
            return False
        except Exception as e:
            print(f"❌ Erro aguardando modal aparecer: {e}")
            return False
    
    def click_deinstallation_button(self, row):
        """Clica no botão de desinstalação na linha especificada"""
        try:
            # Aguarda qualquer loading finalizar antes de tentar clicar
            self.wait_for_loading_to_finish()
            
            # Procura pela célula de ações na linha específica
            actions_cell = row.find_element(By.CSS_SELECTOR, "td.mat-cell.cdk-column-actions")
            
            # Dentro da célula de ações, procura especificamente pelo botão Deinstallation
            deinstall_buttons = actions_cell.find_elements(By.CSS_SELECTOR, "button.quantigo-table-row-action")
            
            deinstall_button = None
            for button in deinstall_buttons:
                button_text = button.text.strip()
                print(f"🔍 Botão encontrado: '{button_text}'")
                if "Deinstallation" in button_text:
                    deinstall_button = button
                    break
            
            if not deinstall_button:
                print("❌ Botão Deinstallation não encontrado")
                return False
            
            # Scroll até o botão para garantir que está visível
            self.driver.execute_script("arguments[0].scrollIntoView(true);", deinstall_button)
            time.sleep(1)
            
            # Aguarda o botão estar clicável
            WebDriverWait(self.driver, 10).until(
                EC.element_to_be_clickable(deinstall_button)
            )
            
            # Verifica novamente se não há loading
            if not self.driver.find_elements(By.CSS_SELECTOR, ".quantigo-progress-spinner"):
                # Clica no botão
                deinstall_button.click()
                print("✅ Botão de desinstalação clicado")
            else:
                print("⏳ Loading ainda ativo, aguardando...")
                self.wait_for_loading_to_finish()
                deinstall_button.click()
                print("✅ Botão de desinstalação clicado após aguardar loading")
            
            # IMPORTANTE: Aguarda o modal realmente aparecer antes de retornar True
            if self.wait_for_modal_to_appear():
                return True
            else:
                print("❌ Modal não apareceu após clicar no botão")
                return False
            
        except Exception as e:
            print(f"❌ Erro ao clicar no botão de desinstalação: {e}")
            # Tenta uma abordagem alternativa com JavaScript
            try:
                self.wait_for_loading_to_finish()
                # Procura por qualquer botão que contenha "Deinstallation"
                deinstall_button = row.find_element(By.XPATH, ".//button[contains(., 'Deinstallation')]")
                self.driver.execute_script("arguments[0].click();", deinstall_button)
                print("✅ Botão de desinstalação clicado (método JavaScript)")
                
                # Aguarda o modal aparecer também no método alternativo
                if self.wait_for_modal_to_appear():
                    return True
                else:
                    print("❌ Modal não apareceu após clicar (método alternativo)")
                    return False
                    
            except Exception as e2:
                print(f"❌ Falha também no método alternativo: {e2}")
                return False
    
    def fill_modal_and_confirm(self):
        """Preenche o modal de desinstalação e confirma"""
        try:
            print("⏳ Aguardando modal aparecer...")
            
            # Aguarda o modal aparecer e localiza o campo de location
            location_input = WebDriverWait(self.driver, 15).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "input[formcontrolname='location']"))
            )
            
            print("📝 Campo location encontrado, preenchendo...")
            
            # Preenche com "REMOÇÃO"
            location_input.clear()
            time.sleep(0.5)
            location_input.send_keys("REMOÇÃO")
            
            print("🔍 Campo preenchido, procurando botão OK...")
            
            # Múltiplas tentativas para encontrar e clicar no botão OK
            ok_button = None
            selectors = [
                # Seletor baseado no HTML fornecido
                "button.mat-stroked-button.mat-button-base.mat-primary",
                # Seletor alternativo por texto
                "//button[contains(@class, 'mat-stroked-button') and contains(span, 'Ok')]",
                # Seletor mais específico
                "button[color='primary'][mat-stroked-button]",
                # Seletor genérico por texto
                "//button[contains(text(), 'Ok')]",
                # Seletor alternativo
                "button.mat-stroked-button span:contains('Ok')",
                # Seletor XPath mais robusto
                "//button[contains(@class, 'mat-stroked-button') and .//span[text()='Ok']]"
            ]
            
            for i, selector in enumerate(selectors):
                try:
                    print(f"🔄 Tentativa {i+1}: usando seletor '{selector}'")
                    
                    if selector.startswith("//"):
                        # XPath selector
                        ok_button = WebDriverWait(self.driver, 5).until(
                            EC.element_to_be_clickable((By.XPATH, selector))
                        )
                    else:
                        # CSS selector
                        ok_button = WebDriverWait(self.driver, 5).until(
                            EC.element_to_be_clickable((By.CSS_SELECTOR, selector))
                        )
                    
                    print(f"✅ Botão OK encontrado com seletor {i+1}")
                    break
                    
                except TimeoutException:
                    print(f"⚠️ Seletor {i+1} não funcionou, tentando próximo...")
                    continue
                except Exception as e:
                    print(f"❌ Erro com seletor {i+1}: {e}")
                    continue
            
            if not ok_button:
                print("🔍 Tentando encontrar todos os botões do modal...")
                # Última tentativa: encontrar todos os botões e procurar pelo texto
                all_buttons = self.driver.find_elements(By.CSS_SELECTOR, "button")
                for button in all_buttons:
                    try:
                        if "Ok" in button.text or "OK" in button.text:
                            ok_button = button
                            print(f"✅ Botão OK encontrado por busca em todos os botões: '{button.text}'")
                            break
                    except:
                        continue
            
            if not ok_button:
                raise Exception("Não foi possível encontrar o botão OK em nenhuma tentativa")
            
            # Scroll até o botão para garantir visibilidade
            self.driver.execute_script("arguments[0].scrollIntoView(true);", ok_button)
            time.sleep(1)
            
            # Aguarda o botão estar realmente clicável
            WebDriverWait(self.driver, 10).until(
                EC.element_to_be_clickable(ok_button)
            )
            
            print("🔘 Clicando no botão OK...")
            
            # Tenta clicar normalmente primeiro
            try:
                ok_button.click()
                print("✅ Botão OK clicado (método normal)")
            except Exception as e:
                print(f"❌ Erro no clique normal: {e}, tentando JavaScript...")
                # Se falhar, usa JavaScript
                self.driver.execute_script("arguments[0].click();", ok_button)
                print("✅ Botão OK clicado (método JavaScript)")
            
            print("⏳ Modal confirmado, aguardando processamento...")
            
            # Aguarda o modal desaparecer
            try:
                WebDriverWait(self.driver, 10).until(
                    EC.invisibility_of_element_located((By.CSS_SELECTOR, "input[formcontrolname='location']"))
                )
                print("✅ Modal fechado")
            except TimeoutException:
                print("⚠️ Timeout aguardando modal fechar, continuando...")
            
            # Aguarda qualquer loading que possa aparecer
            time.sleep(2)
            self.wait_for_loading_to_finish()
            
            print("✅ Processamento do modal concluído com sucesso")
            return True
            
        except TimeoutException as e:
            print(f"⚠️ Timeout no modal: {e}")
            # Tenta screenshot para debug
            try:
                self.driver.save_screenshot("modal_error.png")
                print("📸 Screenshot salvo como modal_error.png")
            except:
                pass
            return False
            
        except Exception as e:
            print(f"❌ Erro ao preencher modal: {e}")
            # Tenta screenshot para debug
            try:
                self.driver.save_screenshot("modal_error.png")
                print("📸 Screenshot salvo como modal_error.png")
            except:
                pass
            return False
    
    def process_chassis(self, chassis):
        """Processa um chassi completo"""
        print(f"\n--- 🚗 Processando chassis: {chassis} ---")
        
        try:
            # 1. Pesquisa o chassis
            if not self.search_chassis(chassis):
                self.failed_chassis.append(f"{chassis} - Erro na pesquisa")
                return False
            
            # 2. Encontra registros ativos
            active_rows = self.find_active_records()
            
            if not active_rows:
                print(f"⚠️ Nenhum registro ativo encontrado para {chassis}")
                self.failed_chassis.append(f"{chassis} - Nenhum registro ativo")
                return False
            
            # 3. Processa cada registro ativo
            processed_count = 0
            for i, row in enumerate(active_rows):
                print(f"🔄 Processando registro ativo {i+1}/{len(active_rows)}")
                
                if self.click_deinstallation_button(row):
                    if self.fill_modal_and_confirm():
                        processed_count += 1
                        print(f"✅ Registro {i+1} processado com sucesso")
                    else:
                        print(f"❌ Falha ao confirmar desinstalação do registro {i+1}")
                else:
                    print(f"❌ Falha ao clicar em desinstalação do registro {i+1}")
                
                # Pequena pausa entre registros
                time.sleep(1)
            
            if processed_count > 0:
                self.successful_chassis.append(f"{chassis} - {processed_count} registro(s) processado(s)")
                print(f"🎉 Chassis {chassis} processado com sucesso! ({processed_count} registros)")
                return True
            else:
                self.failed_chassis.append(f"{chassis} - Falha no processamento")
                return False
                
        except Exception as e:
            print(f"❌ Erro geral ao processar chassis {chassis}: {e}")
            self.failed_chassis.append(f"{chassis} - Erro: {str(e)}")
            return False
    
    def check_if_logged_out(self):
        """Verifica se o usuário foi deslogado - chamado apenas quando há erro"""
        try:
            # Verifica se foi redirecionado para página de login
            current_url = self.driver.current_url
            if "login" in current_url.lower() or "auth" in current_url.lower():
                return True
            
            # Verifica se ainda está na página de subscriptions
            if "quantigo.scopemp.net/app/subscriptions" not in current_url:
                return True
            
            # Procura por elementos de login que só aparecem quando deslogado
            login_indicators = [
                "input[type='password']",
                ".login-form",
                ".auth-container",
                ".login-container"
            ]
            
            for indicator in login_indicators:
                if self.driver.find_elements(By.CSS_SELECTOR, indicator):
                    return True
            
            return False
            
        except:
            return True  # Em caso de erro, assume que foi deslogado
    
    def run_automation(self):
        """Executa a automação completa"""
        try:
            # 1. Configurar driver
            self.setup_driver()
            
            # 2. Carregar lista de chassis
            chassis_list = self.load_chassis_list()
            if not chassis_list:
                print("⚠️ Nenhum chassis para processar!")
                return
            
            # 3. Abrir sistema e aguardar login
            self.wait_for_manual_login()
            
            # 4. Processar cada chassis
            print(f"\n=== 🚀 INICIANDO PROCESSAMENTO DE {len(chassis_list)} CHASSIS ===")
            
            for i, chassis in enumerate(chassis_list, 1):
                print(f"\n📊 Progresso: {i}/{len(chassis_list)}")
                
                # Tenta processar o chassis
                try:
                    self.process_chassis(chassis)
                except Exception as e:
                    # Se houver erro relacionado a logout, verifica
                    if "login" in str(e).lower() or "auth" in str(e).lower():
                        if self.check_if_logged_out():
                            print("🚨 ATENÇÃO: Usuário deslogado! Parando a automação.")
                            break
                    print(f"❌ Erro ao processar chassis {chassis}: {e}")
                    self.failed_chassis.append(f"{chassis} - Erro: {str(e)}")
                
                # Pausa entre chassis
                time.sleep(2)
            
            # 5. Mostrar resumo
            self.show_summary()
            
        except KeyboardInterrupt:
            print("\n⚠️ Automação interrompida pelo usuário!")
        except Exception as e:
            print(f"❌ Erro geral na automação: {e}")
        finally:
            if self.driver:
                input("\n⏸️ Pressione ENTER para fechar o navegador...")
                self.driver.quit()
    
    def show_summary(self):
        """Mostra o resumo final da execução"""
        print("\n" + "="*50)
        print("📊 RESUMO DA AUTOMAÇÃO")
        print("="*50)
        
        print(f"\n✅ Chassis processados com sucesso: {len(self.successful_chassis)}")
        for chassis in self.successful_chassis:
            print(f"   ✓ {chassis}")
        
        print(f"\n❌ Chassis com falhas: {len(self.failed_chassis)}")
        for chassis in self.failed_chassis:
            print(f"   ✗ {chassis}")
        
        total = len(self.successful_chassis) + len(self.failed_chassis)
        success_rate = (len(self.successful_chassis) / total * 100) if total > 0 else 0
        
        print(f"\n📈 Taxa de sucesso: {success_rate:.1f}%")
        print("="*50)

def main():
    """Função principal - necessária para compatibilidade com o executável"""
    print("=== 🤖 AUTOMAÇÃO DE DESINSTALAÇÃO DE CHASSIS ===\n")
    
    # Verifica se o pandas está instalado
    try:
        import pandas as pd
    except ImportError:
        print("❌ ERRO: pandas não está instalado!")
        print("💡 Execute: pip install pandas openpyxl")
        return
    
    # Verifica se o selenium está instalado
    try:
        from selenium import webdriver
    except ImportError:
        print("❌ ERRO: selenium não está instalado!")
        print("💡 Execute: pip install selenium")
        return
    
    automation = ChassisAutomation()
    automation.run_automation()

if __name__ == "__main__":
    main()