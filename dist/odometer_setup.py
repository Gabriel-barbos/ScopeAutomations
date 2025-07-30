import pandas as pd
import json
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from selenium.webdriver.common.action_chains import ActionChains
import logging

class OdometerUpdateAutomation:
    def __init__(self):
        self.driver = None
        self.wait = None
        self.credentials = {}
        self.vehicles_data = pd.DataFrame()
        self.report = {
            'success': [],
            'errors': [],
            'not_found': [],
            'login_errors': []
        }
        self.setup_logging()
        self.vehicles_page_initialized = False

    def setup_logging(self):
        """Configura o sistema de logs"""
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[logging.StreamHandler()]
        )
        self.logger = logging.getLogger(__name__)

    def setup_driver(self):
        """Configura o driver do Chrome"""
        try:
            chrome_options = Options()
            chrome_options.add_argument("--start-maximized")
            chrome_options.add_argument("--force-device-scale-factor=0.8")
            chrome_options.add_argument("--disable-notifications")
            chrome_options.add_argument("--disable-translate")
            
            prefs = {
                "profile.default_content_setting_values": {"notifications": 2},
                "credentials_enable_service": False,
                "profile.password_manager_enabled": False
            }
            chrome_options.add_experimental_option("prefs", prefs)
            chrome_options.add_experimental_option("excludeSwitches", ["enable-logging"])
            chrome_options.add_experimental_option('useAutomationExtension', False)
            chrome_options.add_argument("--disable-blink-features=AutomationControlled")
            
            self.driver = webdriver.Chrome(options=chrome_options)
            self.wait = WebDriverWait(self.driver, 15)
            
            self.driver.execute_script("document.body.style.zoom='80%'")
            self.driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
            
            print("🌐 Chrome configurado com sucesso")
            return True
        except Exception as e:
            print(f"❌ Erro ao configurar driver: {str(e)}")
            return False

    def load_excel_data(self):
        """Carrega e valida os dados da planilha Excel"""
        try:
            self.vehicles_data = pd.read_excel('veiculos_setup.xlsx')
            
            required_columns = ['ID', 'CHASSI', 'DESCRIÇÃO', 'PLACA', 'GRUPO DE VEICULOS', 'CLIENTE', 'ODOMETRO']
            missing_columns = [col for col in required_columns if col not in self.vehicles_data.columns]
            
            if missing_columns:
                self.logger.error(f"Colunas faltantes na planilha: {missing_columns}")
                return False
            
            # Remove linhas com dados obrigatórios vazios
            self.vehicles_data = self.vehicles_data.dropna(subset=['ID', 'CHASSI', 'CLIENTE'])
            
            # Ordena por cliente para otimizar logins
            self.vehicles_data = self.vehicles_data.sort_values('CLIENTE')
            
            self.logger.info(f"Planilha carregada: {len(self.vehicles_data)} veículos")
            return True
            
        except FileNotFoundError:
            self.logger.error("Arquivo 'veiculos_setup.xlsx' não encontrado")
            return False
        except Exception as e:
            self.logger.error(f"Erro ao carregar planilha: {str(e)}")
            return False

    def load_credentials(self):
        """Carrega as credenciais do arquivo JSON"""
        try:
            with open('credentials.json', 'r', encoding='utf-8') as file:
                credentials_data = json.load(file)
            
            self.credentials = {}
            for cliente in credentials_data['clientes']:
                self.credentials[cliente['nome']] = {
                    'user': cliente['user'],
                    'senha': cliente['senha']
                }
            
            self.logger.info(f"Credenciais carregadas: {len(self.credentials)} clientes")
            return True
            
        except Exception as e:
            self.logger.error(f"Erro ao carregar credenciais: {str(e)}")
            return False

    def wait_for_loading(self):
        """Aguarda loadings do sistema finalizarem"""
        try:
            time.sleep(2)
            loading_selectors = [
                "//div[contains(@class, 'loading')]",
                "//div[contains(@class, 'spinner')]", 
                "//div[contains(@class, 'loader')]"
            ]
            
            for selector in loading_selectors:
                try:
                    WebDriverWait(self.driver, 3).until_not(
                        EC.presence_of_element_located((By.XPATH, selector))
                    )
                except:
                    pass
            time.sleep(1)
        except Exception:
            time.sleep(3)

    def login(self, client):
        """Realiza login para o cliente especificado"""
        if client not in self.credentials:
            print(f"❌ Credenciais não encontradas para cliente: {client}")
            return False
        
        try:
            print(f"🔐 Fazendo login para cliente: {client}")
            
            self.driver.get("https://live.mzoneweb.net/mzonex/")
            self.wait_for_loading()
            
            username_field = self.wait.until(EC.presence_of_element_located((By.ID, "Username")))
            username_field.clear()
            username_field.send_keys(self.credentials[client]['user'])
            
            password_field = self.wait.until(EC.presence_of_element_located((By.ID, "password-input")))
            password_field.clear()
            password_field.send_keys(self.credentials[client]['senha'])
            
            login_button = self.wait.until(EC.element_to_be_clickable((By.ID, "login")))
            login_button.click()
            
            try:
                WebDriverWait(self.driver, 20).until(
                    lambda driver: "workspace/map" in driver.current_url
                )
                self.wait_for_loading()
                print(f"✅ Login bem-sucedido: {client}")
                return True
            except TimeoutException:
                print(f"❌ Login falhou para cliente: {client}")
                return False
            
        except Exception as e:
            print(f"❌ Erro no login: {str(e)}")
            return False

    def logout(self):
        """Realiza logout do sistema"""
        try:
            user_dropdown = self.wait.until(EC.element_to_be_clickable((By.ID, "userDropDownMenuToggle")))
            user_dropdown.click()
            
            logout_button = self.wait.until(EC.element_to_be_clickable((By.ID, "userMenuSignOut")))
            logout_button.click()
            
            time.sleep(2)
            self.vehicles_page_initialized = False
            print("✅ Logout realizado")
            return True
        except Exception as e:
            print(f"❌ Erro no logout: {str(e)}")
            return False

    def navigate_to_vehicles(self):
        """Navega para a seção de veículos e seleciona 'TODOS OS VEICULOS'"""
        try:
            if self.vehicles_page_initialized:
                return True
                
            print("🚗 Navegando para seção de veículos...")
            
            self.wait_for_loading()
            
            # Clica no botão de veículos
            vehicles_button = self.wait.until(
                EC.element_to_be_clickable((By.XPATH, "//span[contains(@class, 'button') and contains(., 'Veículos')]"))
            )
            vehicles_button.click()
            
            self.wait_for_loading()
            
            # Clica no dropdown de grupos de veículos
            dropdown_button = self.wait.until(
                EC.element_to_be_clickable((By.XPATH, "//div[@class='wj-input-group wj-input-btn-visible']//button[@class='wj-btn wj-btn-default']"))
            )
            dropdown_button.click()
            
            time.sleep(0.6)
            
            # Seleciona "TODOS OS VEICULOS"
            all_vehicles_option = self.wait.until(
                EC.element_to_be_clickable((By.XPATH, "//div[@class='wj-listbox-item'][contains(., 'Todos os') and contains(., 'Veículos')]"))
            )
            all_vehicles_option.click()
            
            self.wait_for_loading()
            self.vehicles_page_initialized = True
            print("✅ Navegação para veículos concluída")
            return True
            
        except Exception as e:
            print(f"❌ Erro ao navegar para veículos: {str(e)}")
            return False

    def search_vehicle(self, vehicle_id, chassi):
        """Busca veículo por ID e depois por chassi se necessário"""
        try:
            # Primeiro tenta buscar por ID
            if self._search_vehicle_by_criteria(vehicle_id, "ID"):
                return True
            
            # Se não encontrou por ID, tenta por chassi
            print(f"⚠️ Veículo não encontrado por ID, tentando por chassi: {chassi}")
            if self._search_vehicle_by_criteria(chassi, "CHASSI"):
                return True
            
            print(f"❌ Veículo não encontrado nem por ID nem por chassi")
            return False
            
        except Exception as e:
            print(f"❌ Erro ao buscar veículo: {str(e)}")
            return False

    def _search_vehicle_by_criteria(self, search_term, criteria_type):
        """Busca veículo por um critério específico"""
        try:
            print(f"🔍 Buscando veículo por {criteria_type}: {search_term}")
            
            self.wait_for_loading()
            
            search_field = self.wait.until(
                EC.presence_of_element_located((By.XPATH, "//input[@type='search' and @placeholder='Procurar veículos']"))
            )
            
            search_field.clear()
            time.sleep(0.5)
            search_field.send_keys(str(search_term))
            
            self.wait_for_loading()
            time.sleep(3)
            
            # Verifica se encontrou resultado na tabela
            result_xpath = f"//div[@class='wj-row']//div[contains(text(), '{search_term}')]"
            
            try:
                WebDriverWait(self.driver, 5).until(
                    EC.presence_of_element_located((By.XPATH, result_xpath))
                )
                print(f"✅ Veículo encontrado por {criteria_type}: {search_term}")
                return True
            except TimeoutException:
                print(f"❌ Veículo não encontrado por {criteria_type}: {search_term}")
                return False
            
        except Exception as e:
            print(f"❌ Erro ao buscar por {criteria_type}: {str(e)}")
            return False

    def edit_vehicle(self):
        """Clica no botão de editar veículo (três pontos)"""
        try:
            print("✏️ Clicando em editar veículo...")
            
            # Botão de três pontos para editar
            edit_button = self.wait.until(
                EC.element_to_be_clickable((By.XPATH, "//div[@class='wj-cell wj-col-buttons wj-state-multi-selected']//span[@class='three-vertical-dots ng-star-inserted']"))
            )
            edit_button.click()
            
            self.wait_for_loading()
            print("✅ Botão editar clicado")
            return True
            
        except Exception as e:
            print(f"❌ Erro ao clicar em editar: {str(e)}")
            return False

    def navigate_to_unit_controller(self):
        """Navega para a aba do Controlador de unidade"""
        try:
            print("🎛️ Clicando na aba Controlador de unidade...")
            
            unit_controller_tab = self.wait.until(
                EC.element_to_be_clickable((By.XPATH, "//a[@class='ng-star-inserted']//i[@class='mz7-unit-controller']//parent::a"))
            )
            unit_controller_tab.click()
            
            self.wait_for_loading()
            print("✅ Aba Controlador de unidade aberta")
            return True
            
        except Exception as e:
            print(f"❌ Erro ao navegar para Controlador de unidade: {str(e)}")
            return False

    def navigate_to_odometer_tab(self):
        """Navega para a aba do odômetro"""
        try:
            print("📊 Navegando para aba do odômetro...")
            
            # Clica na aba do odômetro
            odometer_tab = self.wait.until(
                EC.element_to_be_clickable((By.XPATH, "//li[@class='ng-star-inserted']//a[contains(text(), 'Odometer')]"))
            )
            odometer_tab.click()
            
            self.wait_for_loading()
            
            # Aguarda a seção do odômetro carregar
            odometer_section = self.wait.until(
                EC.presence_of_element_located((By.XPATH, "//div[@class='editor-section-heading'][text()='Odometer']"))
            )
            
            print("✅ Aba do odômetro aberta")
            return True
            
        except Exception as e:
            print(f"❌ Erro ao navegar para aba do odômetro: {str(e)}")
            return False

    def add_odometer_adjustment(self):
        """Clica no botão Add adjustment"""
        try:
            print("➕ Clicando em Add adjustment...")
            
            add_adjustment_button = self.wait.until(
                EC.element_to_be_clickable((By.XPATH, "//button[@class='button small btn-mz default float-right m-t-xsm']//i[@class='mz-plus']//parent::button"))
            )
            add_adjustment_button.click()
            
            self.wait_for_loading()
            print("✅ Botão Add adjustment clicado")
            return True
            
        except Exception as e:
            print(f"❌ Erro ao clicar em Add adjustment: {str(e)}")
            return False

    def update_odometer(self, odometer_value):
        """Atualiza o valor do odômetro"""
        try:
            print(f"🔢 Atualizando odômetro para: {odometer_value}")
            
            # Campo de input do odômetro
            odometer_field = self.wait.until(
                EC.presence_of_element_located((By.XPATH, "//input[@type='tel' and @name='decimalOdometer' and @class='wj-form-control wj-numeric']"))
            )
            
            odometer_field.clear()
            time.sleep(0.5)
            odometer_field.send_keys(str(odometer_value))
            
            print("✅ Valor do odômetro atualizado")
            return True
            
        except Exception as e:
            print(f"❌ Erro ao atualizar odômetro: {str(e)}")
            return False

    def edit_adjustment_start_time(self):
        """Clica no botão Edit adjustment start time"""
        try:
            print("🕐 Clicando em Edit adjustment start time...")
            
            edit_time_button = self.wait.until(
                EC.element_to_be_clickable((By.XPATH, "//a[text()='Edit adjustment start time']"))
            )
            edit_time_button.click()
            
            time.sleep(1)
            print("✅ Edit adjustment start time clicado")
            return True
            
        except Exception as e:
            print(f"❌ Erro ao clicar em Edit adjustment start time: {str(e)}")
            return False

    def scroll_down_if_needed(self):
        """Faz um pequeno scroll para baixo se necessário"""
        try:
            self.driver.execute_script("window.scrollBy(0, 200);")
            time.sleep(0.5)
            return True
        except Exception as e:
            print(f"❌ Erro ao fazer scroll: {str(e)}")
            return False

    def save_changes(self):
        """Salva as alterações (botão Definir)"""
        try:
            print("💾 Salvando alterações...")
            
            # Botão Definir
            save_button = self.wait.until(
                EC.element_to_be_clickable((By.XPATH, "//button[@class='button small btn-mz default float-right m-t-xsm'][contains(text(), 'Definir')]"))
            )
            save_button.click()
            
            self.wait_for_loading()
            print("✅ Alterações salvas")
            return True
            
        except Exception as e:
            print(f"❌ Erro ao salvar: {str(e)}")
            return False

    def close_modal(self):
        """Fecha o modal (botão Fechar)"""
        try:
            print("❌ Fechando modal...")
            
            close_button = self.wait.until(
                EC.element_to_be_clickable((By.XPATH, "//button[@class='button btn-mz gray'][contains(text(), 'Fechar')]"))
            )
            close_button.click()
            
            self.wait_for_loading()
            print("✅ Modal fechado")
            return True
            
        except Exception as e:
            print(f"❌ Erro ao fechar modal: {str(e)}")
            return False

    def process_vehicle(self, vehicle_row):
        """Processa um veículo individual"""
        vehicle_id = vehicle_row['ID']
        chassi = vehicle_row['CHASSI']
        odometer = vehicle_row['ODOMETRO']
        
        try:
            print(f"\n📋 Processando veículo ID: {vehicle_id}")
            
            # Busca o veículo
            if not self.search_vehicle(vehicle_id, chassi):
                self.report['not_found'].append({
                    'id': vehicle_id,
                    'chassi': chassi,
                    'error': 'Veículo não encontrado'
                })
                return False
            
            # Edita o veículo (clica nos três pontos)
            if not self.edit_vehicle():
                self.report['errors'].append({
                    'id': vehicle_id,
                    'chassi': chassi,
                    'error': 'Erro ao clicar em editar'
                                })
                return False
            
            # Navega para aba do Controlador de unidade
            if not self.navigate_to_unit_controller():
                self.report['errors'].append({
                    'id': vehicle_id,
                    'chassi': chassi,
                    'error': 'Erro ao navegar para Controlador de unidade'
                })
                self.close_modal()  # Tenta fechar modal em caso de erro
                return False
            
            # Navega para aba do odômetro
            if not self.navigate_to_odometer_tab():
                self.report['errors'].append({
                    'id': vehicle_id,
                    'chassi': chassi,
                    'error': 'Erro ao navegar para aba do odômetro'
                })
                self.close_modal()
                return False
            
            # Clica em Add adjustment
            if not self.add_odometer_adjustment():
                self.report['errors'].append({
                    'id': vehicle_id,
                    'chassi': chassi,
                    'error': 'Erro ao clicar em Add adjustment'
                })
                self.close_modal()
                return False
            
            # Atualiza o valor do odômetro
            if not self.update_odometer(odometer):
                self.report['errors'].append({
                    'id': vehicle_id,
                    'chassi': chassi,
                    'error': 'Erro ao atualizar valor do odômetro'
                })
                self.close_modal()
                return False
            
            # Clica em Edit adjustment start time
            if not self.edit_adjustment_start_time():
                self.report['errors'].append({
                    'id': vehicle_id,
                    'chassi': chassi,
                    'error': 'Erro ao clicar em Edit adjustment start time'
                })
                self.close_modal()
                return False
            
            # Faz scroll para baixo se necessário
            self.scroll_down_if_needed()
            
            # Salva as alterações (botão Definir)
            if not self.save_changes():
                self.report['errors'].append({
                    'id': vehicle_id,
                    'chassi': chassi,
                    'error': 'Erro ao salvar alterações'
                })
                self.close_modal()
                return False
            
            # Fecha o modal
            if not self.close_modal():
                self.report['errors'].append({
                    'id': vehicle_id,
                    'chassi': chassi,
                    'error': 'Erro ao fechar modal'
                })
                return False
            
            # Sucesso!
            self.report['success'].append({
                'id': vehicle_id,
                'chassi': chassi,
                'odometer': odometer
            })
            
            print(f"✅ Veículo processado com sucesso: {vehicle_id}")
            return True
            
        except Exception as e:
            print(f"❌ Erro geral ao processar veículo {vehicle_id}: {str(e)}")
            self.report['errors'].append({
                'id': vehicle_id,
                'chassi': chassi,
                'error': f'Erro geral: {str(e)}'
            })
            # Tenta fechar modal em caso de erro
            try:
                self.close_modal()
            except:
                pass
            return False

    def run(self):
        """Executa a automação completa"""
        try:
            print("🚀 Iniciando automação de atualização de odômetro")
            print("="*60)
            
            # Configurações iniciais
            if not self.setup_driver():
                return False
            
            if not self.load_excel_data():
                return False
            
            if not self.load_credentials():
                return False
            
            print(f"📊 Total de veículos a processar: {len(self.vehicles_data)}")
            
            # Processa veículos agrupados por cliente
            current_client = None
            processed_count = 0
            
            for index, vehicle_row in self.vehicles_data.iterrows():
                client = vehicle_row['CLIENTE']
                
                print(f"\n{'='*40}")
                print(f"Progresso: {processed_count + 1}/{len(self.vehicles_data)}")
                print(f"{'='*40}")
                
                # Faz login se mudou de cliente
                if client != current_client:
                    if current_client is not None:
                        print(f"🔄 Mudando de cliente: {current_client} → {client}")
                        self.logout()
                    
                    if not self.login(client):
                        print(f"❌ Falha no login para cliente: {client}")
                        # Adiciona todos os veículos deste cliente aos erros de login
                        client_vehicles = self.vehicles_data[self.vehicles_data['CLIENTE'] == client]
                        for _, v in client_vehicles.iterrows():
                            self.report['login_errors'].append({
                                'id': v['ID'],
                                'chassi': v['CHASSI'],
                                'client': client,
                                'error': 'Falha no login'
                            })
                        # Pula todos os veículos deste cliente
                        processed_count += len(client_vehicles)
                        continue
                    
                    if not self.navigate_to_vehicles():
                        print(f"❌ Falha ao navegar para veículos do cliente: {client}")
                        continue
                    
                    current_client = client
                    print(f"✅ Logado como cliente: {client}")
                
                # Processa o veículo
                success = self.process_vehicle(vehicle_row)
                processed_count += 1
                
                if success:
                    print(f"🎯 Sucesso! Veículo {vehicle_row['ID']} processado")
                else:
                    print(f"⚠️ Falha ao processar veículo {vehicle_row['ID']}")
                
                # Pequena pausa entre veículos
                time.sleep(2)
                
                # Mostra progresso a cada 5 veículos
                if processed_count % 5 == 0:
                    success_count = len(self.report['success'])
                    error_count = len(self.report['errors']) + len(self.report['not_found']) + len(self.report['login_errors'])
                    print(f"📈 Progresso atual: {success_count} sucessos, {error_count} erros")
            
            # Logout final
            if current_client is not None:
                self.logout()
            
            # Exibe relatório final
            self.show_final_report()
            
        except KeyboardInterrupt:
            print("\n⚠️ Execução interrompida pelo usuário")
            self.show_final_report()
        except Exception as e:
            print(f"❌ Erro geral na automação: {str(e)}")
            self.show_final_report()
        finally:
            if self.driver:
                print("🔄 Fechando navegador...")
                self.driver.quit()
                print("✅ Navegador fechado")

    def show_final_report(self):
        """Exibe o relatório final da execução"""
        print("\n" + "="*80)
        print("🏁 RELATÓRIO FINAL DA EXECUÇÃO")
        print("="*80)
        
        total_vehicles = len(self.vehicles_data)
        success_count = len(self.report['success'])
        error_count = len(self.report['errors'])
        not_found_count = len(self.report['not_found'])
        login_error_count = len(self.report['login_errors'])
        
        print(f"📊 RESUMO GERAL:")
        print(f"   Total de veículos: {total_vehicles}")
        print(f"   ✅ Sucessos: {success_count}")
        print(f"   ❌ Erros de processamento: {error_count}")
        print(f"   🔍 Não encontrados: {not_found_count}")
        print(f"   🔐 Erros de login: {login_error_count}")
        
        # Calcula taxa de sucesso
        if total_vehicles > 0:
            success_rate = (success_count / total_vehicles) * 100
            print(f"   📈 Taxa de sucesso: {success_rate:.1f}%")
        
        # Detalhes dos sucessos
        if self.report['success']:
            print(f"\n✅ VEÍCULOS PROCESSADOS COM SUCESSO ({success_count}):")
            for success in self.report['success']:
                print(f"   • ID: {success['id']} | Chassi: {success['chassi']} | Odômetro: {success['odometer']}")
        
        # Detalhes dos erros
        if self.report['errors']:
            print(f"\n❌ VEÍCULOS COM ERRO DE PROCESSAMENTO ({error_count}):")
            for error in self.report['errors']:
                print(f"   • ID: {error['id']} | Chassi: {error['chassi']} | Erro: {error['error']}")
        
        # Veículos não encontrados
        if self.report['not_found']:
            print(f"\n🔍 VEÍCULOS NÃO ENCONTRADOS ({not_found_count}):")
            for nf in self.report['not_found']:
                print(f"   • ID: {nf['id']} | Chassi: {nf['chassi']}")
        
        # Erros de login
        if self.report['login_errors']:
            print(f"\n🔐 ERROS DE LOGIN ({login_error_count}):")
            login_by_client = {}
            for le in self.report['login_errors']:
                client = le['client']
                if client not in login_by_client:
                    login_by_client[client] = []
                login_by_client[client].append(le)
            
            for client, errors in login_by_client.items():
                print(f"   Cliente: {client} ({len(errors)} veículos)")
                for error in errors[:3]:  # Mostra apenas os primeiros 3
                    print(f"     • ID: {error['id']}")
                if len(errors) > 3:
                    print(f"     ... e mais {len(errors) - 3} veículos")
        
        print("="*80)
        
        # Salva relatório em arquivo
        self.save_report_to_file()

    def save_report_to_file(self):
        """Salva o relatório em um arquivo Excel"""
        try:
            from datetime import datetime
            
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"relatorio_odometro_{timestamp}.xlsx"
            
            # Cria DataFrames dos relatórios
            success_df = pd.DataFrame(self.report['success'])
            errors_df = pd.DataFrame(self.report['errors'])
            not_found_df = pd.DataFrame(self.report['not_found'])
            login_errors_df = pd.DataFrame(self.report['login_errors'])
            
            # Salva em arquivo Excel com múltiplas abas
            with pd.ExcelWriter(filename, engine='openpyxl') as writer:
                if not success_df.empty:
                    success_df.to_excel(writer, sheet_name='Sucessos', index=False)
                if not errors_df.empty:
                    errors_df.to_excel(writer, sheet_name='Erros', index=False)
                if not not_found_df.empty:
                    not_found_df.to_excel(writer, sheet_name='Não Encontrados', index=False)
                if not login_errors_df.empty:
                    login_errors_df.to_excel(writer, sheet_name='Erros de Login', index=False)
            
            print(f"💾 Relatório salvo em: {filename}")
            
        except Exception as e:
            print(f"⚠️ Erro ao salvar relatório: {str(e)}")


def main():
    """Função principal"""
    print("🔧 AUTOMAÇÃO DE ATUALIZAÇÃO DE ODÔMETRO")
    print("="*50)
    
    # Verificações iniciais
    import os
    
    if not os.path.exists('veiculos_setup.xlsx'):
        print("❌ Arquivo 'veiculos_setup.xlsx' não encontrado!")
        print("   Certifique-se de que o arquivo está na pasta do script")
        return
    
    if not os.path.exists('credentials.json'):
        print("❌ Arquivo 'credentials.json' não encontrado!")
        print("   Certifique-se de que o arquivo está na pasta do script")
        return
    
    # Pergunta se quer continuar
    print("\n📋 Arquivos encontrados:")
    print("   ✅ veiculos_setup.xlsx")
    print("   ✅ credentials.json")
    
    print("\n⚠️  ATENÇÃO:")
    print("   • Certifique-se de que o Chrome está fechado")
    print("   • A automação irá abrir uma nova janela do Chrome")
    print("   • Não interfira no navegador durante a execução")
    print("   • Em caso de erro, o script continuará com o próximo veículo")
    
    response = input("\n🚀 Deseja iniciar a automação? (s/n): ").lower().strip()
    
    if response in ['s', 'sim', 'y', 'yes']:
        automation = OdometerUpdateAutomation()
        automation.run()
    else:
        print("❌ Automação cancelada pelo usuário")


if __name__ == "__main__":
    main()