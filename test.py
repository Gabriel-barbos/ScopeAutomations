import pandas as pd
import json
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import TimeoutException
import logging

class VehicleAutomation:
    def __init__(self):
        self.driver = None
        self.wait = None
        self.credentials = {}
        self.vehicles_data = pd.DataFrame()
        self.report = {
            'success': [],
            'errors': [],
            'not_found': [],
            'manual_login': [],
            'save_errors': [],
            'odometer_errors': []
        }
        self.setup_logging()
        self.vehicles_page_initialized = False

    def setup_logging(self):
        """Configura o sistema de logs apenas para console"""
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
            
            # Configurações básicas
            chrome_options.add_argument("--start-maximized")
            chrome_options.add_argument("--force-device-scale-factor=0.9")
            chrome_options.add_argument("--disable-infobars")
            chrome_options.add_argument("--use-fake-ui-for-media-stream")
            chrome_options.add_argument("--autoplay-policy=no-user-gesture-required")
            chrome_options.add_argument("--disable-translate")
            chrome_options.add_argument("--disable-web-security")
            chrome_options.add_argument("--disable-features=VizDisplayCompositor")
            
            # Configurações de prefs
            prefs = {
                "profile.default_content_setting_values": {
                    "notifications": 2,
                    "media_stream_camera": 2,
                    "media_stream_mic": 2,
                    "geolocation": 2,
                    "plugins": 2,
                    "popups": 2,
                    "automatic_downloads": 2
                },
                "profile.default_content_settings": {"popups": 0},
                "profile.managed_default_content_settings": {"images": 1},
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
            
            print("🌐 Chrome configurado com sucesso ")
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
            
            initial_count = len(self.vehicles_data)
            self.vehicles_data = self.vehicles_data.dropna(subset=['ID', 'CHASSI', 'CLIENTE'])
            final_count = len(self.vehicles_data)
            
            if initial_count != final_count:
                self.logger.warning(f"Removidas {initial_count - final_count} linhas com dados obrigatórios vazios")
            
            mask_placa_vazia = self.vehicles_data['PLACA'].isna() | (self.vehicles_data['PLACA'] == '')
            self.vehicles_data.loc[mask_placa_vazia, 'PLACA'] = self.vehicles_data.loc[mask_placa_vazia, 'CHASSI']
            placas_preenchidas = mask_placa_vazia.sum()
            
            if placas_preenchidas > 0:
                self.logger.info(f"Preenchidas {placas_preenchidas} placas vazias com o valor do chassi")
            
            self.vehicles_data = self.vehicles_data.sort_values('CLIENTE')
            
            self.logger.info(f"Planilha carregada com sucesso: {len(self.vehicles_data)} veículos encontrados")
            return True
            
        except FileNotFoundError:
            self.logger.error("Arquivo 'veiculos_setup.xlsx' não encontrado na raiz do projeto")
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
            
            self.logger.info(f"Credenciais carregadas: {len(self.credentials)} clientes encontrados")
            self.logger.info(f"Clientes disponíveis: {list(self.credentials.keys())}")
            return True
            
        except FileNotFoundError:
            self.logger.error("Arquivo 'credentials.json' não encontrado na raiz do projeto")
            return False
        except json.JSONDecodeError:
            self.logger.error("Erro ao decodificar o arquivo credentials.json")
            return False
        except KeyError as e:
            self.logger.error(f"Estrutura do JSON inválida. Chave não encontrada: {e}")
            return False
        except Exception as e:
            self.logger.error(f"Erro ao carregar credenciais: {str(e)}")
            return False

    def validate_clients(self):
        """Valida se todos os clientes da planilha existem nas credenciais"""
        unique_clients = self.vehicles_data['CLIENTE'].unique()
        missing_clients = []
        manual_clients = []
        
        for client in unique_clients:
            if str(client).upper() == 'MANUAL':
                manual_clients.append(client)
            elif client not in self.credentials:
                missing_clients.append(client)
        
        if missing_clients:
            self.logger.warning(f"Clientes não encontrados nas credenciais: {missing_clients}")
            for client in missing_clients:
                self.report['manual_login'].append({
                    'cliente': client,
                    'motivo': 'Cliente não encontrado nas credenciais'
                })
        
        if manual_clients:
            self.logger.info(f"Clientes configurados para login manual: {manual_clients}")
        
        return True

    def ask_login_method(self):
        """Pergunta ao usuário sobre o método de login"""
        while True:
            print("\n" + "="*50)
            print("MÉTODO DE LOGIN")
            print("="*50)
            print("1 - Login Automático ()")
            print("2 - Login Manual (você fará login manualmente)")
            print("="*50)
            
            choice = input("Escolha uma opção (1 ou 2): ").strip()
            
            if choice in ['1', '2']:
                return choice == '2'
            else:
                print("Opção inválida! Digite 1 ou 2.")

    def wait_for_loading(self):
        """Aguarda loadings do sistema .NET finalizarem"""
        try:
            time.sleep(2)
            
            loading_selectors = [
                "//div[contains(@class, 'loading')]",
                "//div[contains(@class, 'spinner')]", 
                "//div[contains(@class, 'loader')]",
                "//*[contains(@class, 'loading')]"
            ]
            
            for selector in loading_selectors:
                try:
                    WebDriverWait(self.driver, 3).until_not(
                        EC.presence_of_element_located((By.XPATH, selector))
                    )
                except:
                    pass
            
            time.sleep(1)
            
        except Exception as e:
            self.logger.warning(f"Aviso no wait_for_loading: {str(e)}")
            time.sleep(3)
    
    def login_automatic(self, client):
        """Realiza login automático usando as credenciais"""
        if client not in self.credentials:
            print(f"❌ Credenciais não encontradas para cliente: {client}")
            return False
        
        try:
            print(f"🔐 Realizando login automático para cliente: {client}")
            
            self.driver.get("https://live.mzoneweb.net/mzonex/")
            self.wait_for_loading()
            
            username_field = self.wait.until(EC.presence_of_element_located((By.ID, "Username")))
            username_field.clear()
            username_field.send_keys(self.credentials[client]['user'])
            print(f"📝 Usuário preenchido: {self.credentials[client]['user']}")
            
            password_field = self.wait.until(EC.presence_of_element_located((By.ID, "password-input")))
            password_field.clear()
            password_field.send_keys(self.credentials[client]['senha'])
            print("🔑 Senha preenchida")
            
            login_button = self.wait.until(EC.element_to_be_clickable((By.ID, "login")))
            login_button.click()
            print("⏳ Aguardando redirecionamento...")
            
            try:
                WebDriverWait(self.driver, 20).until(
                    lambda driver: "workspace/map" in driver.current_url
                )
                self.wait_for_loading()
                print(f"✅ Login automático bem-sucedido para cliente: {client}")
                return True
                
            except TimeoutException:
                current_url = self.driver.current_url
                print(f"❌ Login falhou para cliente {client}. URL atual: {current_url}")
                return False
            
        except TimeoutException:
            print(f"⏰ Timeout ao tentar fazer login para cliente: {client}")
            return False
        except Exception as e:
            print(f"❌ Erro no login automático para {client}: {str(e)}")
            return False

    def login_manual(self, client):
        """Aguarda login manual do usuário"""
        try:
            self.logger.info(f"Preparando login manual para cliente: {client}")
            
            self.driver.get("https://live.mzoneweb.net/mzonex/")
            time.sleep(2)
            
            print(f"\n{'='*60}")
            print(f"LOGIN MANUAL NECESSÁRIO PARA CLIENTE: {client}")
            print(f"{'='*60}")
            print("A página de login foi carregada no navegador.")
            print("Por favor, realize o login manualmente.")
            print("Após completar o login, pressione ENTER para continuar...")
            
            input()
            
            try:
                current_url = self.driver.current_url
                if "workspace/map" in current_url:
                    self.logger.info(f"Login manual bem-sucedido para cliente: {client}")
                    return True
                else:
                    self.logger.warning(f"Login manual pode não ter sido bem-sucedido para cliente {client}. URL atual: {current_url}")
                    while True:
                        resposta = input("O login foi realizado com sucesso? (s/n): ").lower().strip()
                        if resposta in ['s', 'sim', 'y', 'yes']:
                            self.logger.info(f"Login manual confirmado pelo usuário para cliente: {client}")
                            return True
                        elif resposta in ['n', 'não', 'nao', 'no']:
                            self.logger.error(f"Login manual falhou para cliente: {client}")
                            return False
                        else:
                            print("Por favor, responda com 's' para sim ou 'n' para não.")
                            
            except Exception as e:
                self.logger.error(f"Erro ao verificar login manual para {client}: {str(e)}")
                return False
            
        except Exception as e:
            self.logger.error(f"Erro no login manual para {client}: {str(e)}")
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
            
            self.logger.info("Logout realizado com sucesso")
            return True
            
        except Exception as e:
            self.logger.error(f"Erro no logout: {str(e)}")
            return False
    
    def navigate_to_vehicles(self):
        """Navega para a seção de veículos (apenas uma vez por cliente)"""
        try:
            if self.vehicles_page_initialized:
                print("🔄 Página de veículos já inicializada, reutilizando...")
                return True
                
            print("🚗 Navegando para seção de veículos...")
            
            self.wait_for_loading()
            
            vehicles_button = self.wait.until(
                EC.element_to_be_clickable((By.XPATH, "//span[contains(@class, 'button') and contains(., 'Veículos')]"))
            )
            vehicles_button.click()
            print("📋 Botão 'Veículos' clicado")
            
            self.wait_for_loading()
            
            dropdown_button = self.wait.until(
                EC.element_to_be_clickable((By.XPATH, "//div[@class='wj-input-group wj-input-btn-visible']//button[@class='wj-btn wj-btn-default']"))
            )
            dropdown_button.click()
            print("📂 Dropdown de grupos aberto")
            
            time.sleep(0.6)
            
            all_vehicles_option = self.wait.until(
                EC.element_to_be_clickable((By.XPATH, "//div[@class='wj-listbox-item'][contains(., 'Todos os') and contains(., 'Veículos')]"))
            )
            all_vehicles_option.click()
            print("✅ Opção 'Todos os Veículos' selecionada")
            
            self.wait_for_loading()
            print("🎯 Todos os veículos carregados com sucesso")
            
            self.vehicles_page_initialized = True
            return True
            
        except Exception as e:
            print(f"❌ Erro ao navegar para veículos: {str(e)}")
            return False

    def search_vehicle_by_id(self, vehicle_id):
        """Busca veículo pelo ID com melhor detecção de resultados vazios"""
        try:
            print(f"🔍 Buscando veículo com ID: {vehicle_id}")
            
            self.wait_for_loading()
            
            search_field = self.wait.until(
                EC.presence_of_element_located((By.XPATH, "//input[@type='search' and @placeholder='Procurar veículos']"))
            )
            
            search_field.clear()
            time.sleep(0.5)
            search_field.send_keys(str(vehicle_id))
            
            print(f"⌨️  ID {vehicle_id} digitado no campo de busca")
            
            self.wait_for_loading()
            time.sleep(3)
            print("⏳ Aguardando resultados da busca...")
            
            # Verificar se encontrou resultados
            result_xpath = f"//div[@class='wj-row']//div[contains(text(), '{vehicle_id}')]"
            
            try:
                WebDriverWait(self.driver, 5).until(
                    EC.presence_of_element_located((By.XPATH, result_xpath))
                )
                print(f"✅ Veículo encontrado: {vehicle_id}")
                return True
            except TimeoutException:
                print(f"❌ Veículo não encontrado: {vehicle_id}")
                return False
            
        except Exception as e:
            print(f"❌ Erro ao buscar veículo {vehicle_id}: {str(e)}")
            return False

    def click_edit_vehicle(self, vehicle_id):
        """Clica no botão de editar veículo"""
        try:
            self.logger.info(f"Clicando em editar para veículo: {vehicle_id}")
            
            vehicle_row_xpath = f"//div[@class='wj-row'][.//div[contains(text(), '{vehicle_id}')]]"
            vehicle_row = self.wait.until(EC.presence_of_element_located((By.XPATH, vehicle_row_xpath)))
            
            edit_button = vehicle_row.find_element(By.XPATH, ".//i[@class='pointer mz7-pencil']")
            
            self.driver.execute_script("arguments[0].scrollIntoView(true);", edit_button)
            time.sleep(0.5)
            
            edit_button.click()
            
            self.wait_for_loading()
            
            self.logger.info(f"Modal de edição aberto para veículo: {vehicle_id}")
            return True
            
        except Exception as e:
            self.logger.error(f"Erro ao clicar em editar veículo {vehicle_id}: {str(e)}")
            return False

    def fill_vehicle_form(self, vehicle_data):
        """Preenche o formulário de edição do veículo"""
        try:
            self.logger.info(f"Preenchendo formulário para veículo ID: {vehicle_data['ID']}")
            
            self.wait_for_loading()
            time.sleep(2)
            
            if not self.fill_description_field(vehicle_data):
                return False
            
            if not self.fill_plate_field(vehicle_data):
                return False
            
            if not self.fill_chassis_field(vehicle_data):
                return False
            
            if not self.select_vehicle_group(vehicle_data):
                return False
            
            if not self.save_vehicle_form_and_check(vehicle_data):
                return False
            
            self.logger.info(f"Formulário preenchido e salvo com sucesso para veículo ID: {vehicle_data['ID']}")
            return True
            
        except Exception as e:
            self.logger.error(f"Erro ao preencher formulário: {str(e)}")
            return False

    def fill_description_field(self, vehicle_data):
        """Preenche o campo descrição"""
        try:
            description = vehicle_data.get('DESCRIÇÃO', '')
            if pd.isna(description) or description == '':
                print("⚠️  Descrição vazia, pulando...")
                return True
            
            print(f"📝 Preenchendo descrição: {description}")
            
            description_field = self.wait.until(
                EC.presence_of_element_located((By.XPATH, "//input[@name='description' and @placeholder='Descrição']"))
            )
            
            description_field.clear()
            time.sleep(0.5)
            description_field.send_keys(str(description))
            
            print("✅ Campo descrição preenchido")
            return True
            
        except Exception as e:
            print(f"❌ Erro ao preencher descrição: {str(e)}")
            return False

    def fill_plate_field(self, vehicle_data):
        """Preenche o campo placa (usa chassi se placa estiver vazia)"""
        try:
            plate = vehicle_data.get('PLACA', '')
            chassi = vehicle_data.get('CHASSI', '')
            
            if pd.isna(plate) or plate == '':
                plate = chassi
                print(f"⚠️  Placa vazia, usando chassi: {plate}")
            
            if pd.isna(plate) or plate == '':
                print("⚠️  Placa e chassi vazios, pulando...")
                return True
            
            print(f"🏷️  Preenchendo placa: {plate}")
            
            plate_field = self.wait.until(
                EC.presence_of_element_located((By.XPATH, "//input[@name='registration' and @placeholder='Placa']"))
            )
            
            plate_field.clear()
            time.sleep(0.5)
            plate_field.send_keys(str(plate))
            
            print("✅ Campo placa preenchido")
            return True
            
        except Exception as e:
            print(f"❌ Erro ao preencher placa: {str(e)}")
            return False

    def fill_chassis_field(self, vehicle_data):
        """Preenche o campo chassi"""
        try:
            chassi = vehicle_data.get('CHASSI', '')
            if pd.isna(chassi) or chassi == '':
                print("⚠️  Chassi vazio, pulando...")
                return True
            
            print(f"🔧 Preenchendo chassi: {chassi}")
            
            chassis_field = self.wait.until(
                EC.presence_of_element_located((By.XPATH, "//input[@name='vin' and @placeholder='Número de Chassi']"))
            )
            
            chassis_field.clear()
            time.sleep(0.5)
            chassis_field.send_keys(str(chassi))
            
            print("✅ Campo chassi preenchido")
            return True
            
        except Exception as e:
            print(f"❌ Erro ao preencher chassi: {str(e)}")
            return False

    def select_vehicle_group(self, vehicle_data):
        """Seleciona o grupo de veículos"""
        try:
            group_name = vehicle_data.get('GRUPO DE VEICULOS', '')
            if pd.isna(group_name) or group_name == '':
                print("⚠️  Grupo de veículos vazio, pulando...")
                return True
            
            print(f"👥 Selecionando grupo: {group_name}")
            
            group_button = self.wait.until(
                EC.element_to_be_clickable((By.XPATH, "//li/a[contains(text(), 'Grupos de veículos')]"))
            )
            group_button.click()
            print("📂 Seção de grupos aberta")
            
            time.sleep(1)
            
            search_group_field = self.wait.until(
                EC.presence_of_element_located((By.XPATH, "//input[@placeholder='Buscar' and @class='form-input ng-untouched ng-pristine ng-valid']"))
            )
            
            search_group_field.clear()
            time.sleep(0.5)
            search_group_field.send_keys(str(group_name))
            
            print(f"🔍 Pesquisando grupo: {group_name}")
            
            time.sleep(3)
            
            if not self.select_group_checkbox(group_name):
                return False
            
            print(f"✅ Grupo selecionado: {group_name}")
            return True
            
        except Exception as e:
            print(f"❌ Erro ao selecionar grupo: {str(e)}")
            return False

    def select_group_checkbox(self, group_name):
        """Seleciona o checkbox do grupo encontrado"""
        try:
            group_strategies = [
                f"//label[contains(text(), '{group_name}')]/input[@type='checkbox']",
                f"//label[contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), '{group_name.lower()}')]/input[@type='checkbox']"
            ]
            
            if ' ' in group_name:
                group_strategies.append(f"//label[contains(text(), '{group_name.split()[0]}')]/input[@type='checkbox']")
            
            for strategy_xpath in group_strategies:
                try:
                    checkbox = WebDriverWait(self.driver, 5).until(
                        EC.element_to_be_clickable((By.XPATH, f"//div[@class='scrollbar-container checkboxlist-wrapper']{strategy_xpath}"))
                    )
                    
                    self.driver.execute_script("arguments[0].scrollIntoView(true);", checkbox)
                    time.sleep(0.5)
                    
                    if not checkbox.is_selected():
                        checkbox.click()
                        self.logger.info(f"Checkbox do grupo '{group_name}' marcado com sucesso")
                    else:
                        self.logger.info(f"Checkbox do grupo '{group_name}' já estava marcado")
                    
                    return True
                    
                except TimeoutException:
                    continue
                except Exception as e:
                    self.logger.warning(f"Erro na estratégia {strategy_xpath}: {str(e)}")
                    continue
            
            # Busca genérica
            try:
                checkboxes = self.driver.find_elements(By.XPATH, "//div[@class='scrollbar-container checkboxlist-wrapper']//label")
                
                for label in checkboxes:
                    label_text = label.text.strip()
                    if group_name.upper() in label_text.upper() or label_text.upper() in group_name.upper():
                        checkbox = label.find_element(By.XPATH, ".//input[@type='checkbox']")
                        if not checkbox.is_selected():
                            checkbox.click()
                            self.logger.info(f"Grupo encontrado e selecionado: '{label_text}' para busca '{group_name}'")
                        return True
                
                self.logger.error(f"Grupo '{group_name}' não encontrado nos resultados")
                return False
                
            except Exception as e:
                self.logger.error(f"Erro na busca genérica do grupo: {str(e)}")
                return False
            
        except Exception as e:
            self.logger.error(f"Erro ao selecionar checkbox do grupo: {str(e)}")
            return False

    def save_vehicle_form_and_check(self, vehicle_data):
        """Clica no botão salvar do formulário e verifica se modal fechou"""
        try:
            print("💾 Salvando formulário...")
            
            save_button = self.wait.until(
                EC.element_to_be_clickable((By.XPATH, "//button[@type='submit' and contains(@class, 'button') and contains(@class, 'btn-mz') and contains(@class, 'success') and contains(text(), 'Salvar')]"))
            )
            
            self.driver.execute_script("arguments[0].scrollIntoView(true);", save_button)
            time.sleep(0.5)
            
            save_button.click()
            print("✅ Botão salvar clicado")
            
            self.wait_for_loading()
            
            # Verificar se modal fechou (aguardar 1-2 segundos)
            print("⏳ Verificando se modal foi fechado...")
            time.sleep(1)
            
            try:
                # Verifica se o modal ainda está visível
                modal_overlay = self.driver.find_element(By.XPATH, "//div[@class='form-overlay visible']")
                if modal_overlay.is_displayed():
                    print("❌ Modal não fechou - erro ao salvar formulário")
                    self.click_cancel_button()
                    
                    # Registra erro de salvamento
                    self.report['save_errors'].append({
                        'cliente': vehicle_data['CLIENTE'],
                        'id': vehicle_data['ID'],
                        'erro': 'Modal não fechou após salvar'
                    })
                    return False
                else:
                    print("✅ Modal fechado - formulário salvo com sucesso")
                    return True
                    
            except:
                # Se não encontrou o modal, significa que fechou
                print("✅ Modal fechado - formulário salvo com sucesso")
                return True
            
        except Exception as e:
            print(f"❌ Erro ao salvar formulário: {str(e)}")
            return False

    def click_cancel_button(self):
        """Clica no botão cancelar quando modal não fecha"""
        try:
            print("🚫 Clicando no botão Cancelar...")
            
            cancel_button = self.wait.until(
                EC.element_to_be_clickable((By.XPATH, "//button[@type='button' and contains(@class, 'button') and contains(@class, 'btn-mz') and contains(@class, 'gray') and contains(text(), 'Cancelar')]"))
            )
            
            cancel_button.click()
            print("✅ Botão Cancelar clicado")
            
            time.sleep(1)
            return True
            
        except Exception as e:
            print(f"❌ Erro ao clicar no botão Cancelar: {str(e)}")
            return False

    def update_vehicle_odometer(self, vehicle_data):
        """Atualiza o odômetro do veículo seguindo o fluxo completo"""
        try:
            vehicle_id = vehicle_data['ID']
            odometer_value = vehicle_data.get('ODOMETRO', '')
            
            # Verificar se há valor de odômetro
            if pd.isna(odometer_value) or odometer_value == '':
                print("⚠️  Valor de odômetro vazio, pulando atualização...")
                return True
                
            print(f"🔢 Iniciando atualização do odômetro: {odometer_value}")
            
            # 1. Clicar no botão de três pontos
            if not self.click_three_dots_menu(vehicle_id):
                return False
            
            # 2. Clicar em "Controlador de unidade"
            if not self.click_unit_controller():
                return False
            
            # 3. Clicar na aba "Odometer"
            if not self.click_odometer_tab():
                return False
            
            # 4. Clicar no botão "Add adjustment"
            if not self.click_add_adjustment():
                return False
            
            # 5. Preencher o valor do odômetro
            if not self.fill_odometer_value(odometer_value):
                return False
            
            # 6. Clicar em "Edit adjustment start time"
            if not self.click_edit_adjustment_time():
                return False
            
            # 7. Clicar em "Definir"
            if not self.click_define_button():
                return False
            
            # 8. Fechar o modal
            if not self.close_odometer_modal():
                return False
            
            print(f"✅ Odômetro atualizado com sucesso: {odometer_value}")
            return True
            
        except Exception as e:
            print(f"❌ Erro ao atualizar odômetro: {str(e)}")
            self.report['odometer_errors'].append({
                'cliente': vehicle_data['CLIENTE'],
                'id': vehicle_data['ID'],
                'erro': str(e)
            })
            return False

    def click_three_dots_menu(self, vehicle_id):
        """Clica no botão de três pontos (…) do veículo"""
        try:
            print("📋 Clicando no menu de três pontos...")
            
            vehicle_row_xpath = f"//div[@class='wj-row'][.//div[contains(text(), '{vehicle_id}')]]"
            vehicle_row = self.wait.until(EC.presence_of_element_located((By.XPATH, vehicle_row_xpath)))
            
            three_dots_button = vehicle_row.find_element(By.XPATH, ".//div[@class='wj-cell wj-col-buttons wj-state-multi-selected']//span[@class='three-vertical-dots']")
            
            self.driver.execute_script("arguments[0].scrollIntoView(true);", three_dots_button)
            time.sleep(0.5)
            
            three_dots_button.click()
            print("✅ Menu de três pontos clicado")
            
            time.sleep(1)  # Aguardar menu abrir
            return True
            
        except Exception as e:
            print(f"❌ Erro ao clicar no menu de três pontos: {str(e)}")
            return False

    def click_unit_controller(self):
        """Clica na opção 'Controlador de unidade'"""
        try:
            print("🎛️  Clicando em 'Controlador de unidade'...")
            
            unit_controller_option = self.wait.until(
                EC.element_to_be_clickable((By.XPATH, "//a[@class='ng-star-inserted'][.//i[@class='mz7-unit-controller']]"))
            )
            
            unit_controller_option.click()
            print("✅ 'Controlador de unidade' clicado")
            
            self.wait_for_loading()
            return True
            
        except Exception as e:
            print(f"❌ Erro ao clicar em 'Controlador de unidade': {str(e)}")
            return False

    def click_odometer_tab(self):
        """Clica na aba 'Odometer'"""
        try:
            print("📏 Clicando na aba 'Odometer'...")
            
            odometer_tab = self.wait.until(
                EC.element_to_be_clickable((By.XPATH, "//li[@class='ng-star-inserted']//a[contains(text(), 'Odometer')]"))
            )
            
            odometer_tab.click()
            print("✅ Aba 'Odometer' clicada")
            
            time.sleep(1)
            return True
            
        except Exception as e:
            print(f"❌ Erro ao clicar na aba 'Odometer': {str(e)}")
            return False

    def click_add_adjustment(self):
        """Clica no botão 'Add adjustment'"""
        try:
            print("➕ Clicando em 'Add adjustment'...")
            
            add_adjustment_button = self.wait.until(
                EC.element_to_be_clickable((By.XPATH, "//button[@type='button' and contains(@class, 'button') and contains(@class, 'small') and contains(@class, 'btn-mz') and contains(@class, 'default')]//i[@class='mz-plus']"))
            )
            
            add_adjustment_button.click()
            print("✅ Botão 'Add adjustment' clicado")
            
            time.sleep(1)
            return True
            
        except Exception as e:
            print(f"❌ Erro ao clicar em 'Add adjustment': {str(e)}")
            return False

    def fill_odometer_value(self, odometer_value):
        """Limpa e preenche o valor do odômetro"""
        try:
            print(f"📝 Preenchendo valor do odômetro: {odometer_value}")
            
            odometer_input = self.wait.until(
                EC.presence_of_element_located((By.XPATH, "//input[@type='tel' and @name='decimalOdometer' and contains(@class, 'wj-form-control')]"))
            )
            
            # Limpar campo e inserir novo valor
            odometer_input.clear()
            time.sleep(0.5)
            odometer_input.send_keys(str(odometer_value))
            
            print("✅ Valor do odômetro preenchido")
            return True
            
        except Exception as e:
            print(f"❌ Erro ao preencher valor do odômetro: {str(e)}")
            return False

    def click_edit_adjustment_time(self):
        """Clica em 'Edit adjustment start time'"""
        try:
            print("⏰ Clicando em 'Edit adjustment start time'...")
            
            edit_time_link = self.wait.until(
                EC.element_to_be_clickable((By.XPATH, "//div//a[contains(text(), 'Edit adjustment start time')]"))
            )
            
            edit_time_link.click()
            print("✅ 'Edit adjustment start time' clicado")
            
            time.sleep(1)
            return True
            
        except Exception as e:
            print(f"❌ Erro ao clicar em 'Edit adjustment start time': {str(e)}")
            return False

    def click_define_button(self):
        """Clica no botão 'Definir'"""
        try:
            print("✔️  Clicando em 'Definir'...")
            
            define_button = self.wait.until(
                EC.element_to_be_clickable((By.XPATH, "//button[@type='button' and contains(@class, 'button') and contains(@class, 'small') and contains(@class, 'btn-mz') and contains(@class, 'default') and contains(text(), 'Definir')]"))
            )
            
            define_button.click()
            print("✅ Botão 'Definir' clicado")
            
            time.sleep(1)
            return True
            
        except Exception as e:
            print(f"❌ Erro ao clicar em 'Definir': {str(e)}")
            return False

    def close_odometer_modal(self):
        """Fecha o modal do odômetro"""
        try:
            print("🚪 Fechando modal do odômetro...")
            
            close_button = self.wait.until(
                EC.element_to_be_clickable((By.XPATH, "//button[@type='button' and contains(@class, 'button') and contains(@class, 'btn-mz') and contains(@class, 'gray') and contains(text(), 'Fechar')]"))
            )
            
            close_button.click()
            print("✅ Modal do odômetro fechado")
            
            time.sleep(1)
            return True
            
        except Exception as e:
            print(f"❌ Erro ao fechar modal do odômetro: {str(e)}")
            return False

    def run_automation(self):
        """Executa o fluxo principal da automação"""
        try:
            print("🚀" + "="*58 + "🚀")
            print("🤖           INICIANDO AUTOMAÇÃO DE VEÍCULOS           🤖")
            print("🚀" + "="*58 + "🚀")
            
            print("\n📊 FASE 1: VALIDAÇÃO DOS DADOS")
            print("-" * 40)
            
            if not self.load_excel_data():
                print("❌ Falha ao carregar planilha Excel")
                return False
            
            if not self.load_credentials():
                print("❌ Falha ao carregar credenciais")
                return False
            
            self.validate_clients()
            print("✅ Validação concluída")
            
            print("\n🌐 FASE 2: CONFIGURAÇÃO DO NAVEGADOR")
            print("-" * 40)
            if not self.setup_driver():
                print("❌ Falha ao configurar navegador")
                return False
            
            use_manual_login = self.ask_login_method()
            
            print("\n🔄 FASE 3: PROCESSAMENTO DOS VEÍCULOS")
            print("-" * 40)
            
            current_client = None
            total_vehicles = len(self.vehicles_data)
            processed_count = 0
            
            for index, vehicle in self.vehicles_data.iterrows():
                client = vehicle['CLIENTE']
                processed_count += 1
                
                print(f"\n🔢 Progresso: {processed_count}/{total_vehicles} veículos")
                
                if current_client != client:
                    if current_client is not None:
                        print(f"🚪 Fazendo logout do cliente: {current_client}")
                        self.logout()
                    
                    current_client = client
                    self.vehicles_page_initialized = False
                    
                    print(f"\n👤 MUDANDO PARA CLIENTE: {client}")
                    print("-" * 50)
                    
                    if str(client).upper() == 'MANUAL' or use_manual_login or client not in self.credentials:
                        login_success = self.login_manual(client)
                    else:
                        login_success = self.login_automatic(client)
                    
                    if not login_success:
                        print(f"❌ Login falhou para cliente: {client}")
                        client_vehicles = self.vehicles_data[self.vehicles_data['CLIENTE'] == client]
                        for _, v in client_vehicles.iterrows():
                            self.report['errors'].append({
                                'cliente': client,
                                'id': v['ID'],
                                'erro': 'Falha no login'
                            })
                        continue
                
                self.process_vehicle(vehicle)
            
            if current_client is not None:
                print(f"\n🚪 Logout final do cliente: {current_client}")
                self.logout()
            
            print("\n🎯 FASE 4: GERAÇÃO DO RELATÓRIO")
            print("-" * 40)
            
            self.generate_final_report()
            
            return True
            
        except Exception as e:
            print(f"❌ Erro no fluxo principal: {str(e)}")
            return False
        finally:
            if self.driver:
                print("🔒 Fechando navegador...")
                self.driver.quit()

    def process_vehicle(self, vehicle_data):
        """Processa um veículo individual"""
        try:
            vehicle_id = vehicle_data['ID']
            client = vehicle_data['CLIENTE']
            
            print(f"\n{'='*60}")
            print(f"🔧 PROCESSANDO VEÍCULO")
            print(f"{'='*60}")
            print(f"🆔 ID: {vehicle_id}")
            print(f"👤 Cliente: {client}")
            print(f"🚗 Chassi: {vehicle_data.get('CHASSI', 'N/A')}")
            print(f"🏷️  Placa: {vehicle_data.get('PLACA', 'N/A')}")
            print(f"📝 Descrição: {vehicle_data.get('DESCRIÇÃO', 'N/A')}")
            print(f"👥 Grupo: {vehicle_data.get('GRUPO DE VEICULOS', 'N/A')}")
            print(f"🔢 Odômetro: {vehicle_data.get('ODOMETRO', 'N/A')}")
            print(f"{'='*60}")
            
            if not self.navigate_to_vehicles():
                self.report['errors'].append({
                    'cliente': client,
                    'id': vehicle_id,
                    'erro': 'Erro ao navegar para veículos'
                })
                print("❌ Erro ao navegar para veículos")
                return
            
            print(f"🔍 Buscando veículo ID: {vehicle_id}...")
            if not self.search_vehicle_by_id(vehicle_id):
                self.report['not_found'].append({
                    'cliente': client,
                    'id': vehicle_id,
                    'chassi': vehicle_data.get('CHASSI', 'N/A'),
                    'placa': vehicle_data.get('PLACA', 'N/A')
                })
                print(f"❌ Veículo ID {vehicle_id} não encontrado")
                return
            
            print(f"✏️  Abrindo modal de edição...")
            if not self.click_edit_vehicle(vehicle_id):
                self.report['errors'].append({
                    'cliente': client,
                    'id': vehicle_id,
                    'erro': 'Erro ao abrir modal de edição'
                })
                print("❌ Erro ao abrir modal de edição")
                return
            
            print(f"📝 Preenchendo formulário...")
            form_success = self.fill_vehicle_form(vehicle_data)
            
            if not form_success:
                self.report['errors'].append({
                    'cliente': client,
                    'id': vehicle_id,
                    'erro': 'Erro ao preencher formulário'
                })
                print("❌ Erro ao preencher formulário")
                return
            
            # Se formulário foi salvo com sucesso, atualizar odômetro
            print(f"🔢 Atualizando odômetro...")
            odometer_success = self.update_vehicle_odometer(vehicle_data)
            
            if not odometer_success:
                print("⚠️  Erro ao atualizar odômetro, mas veículo foi processado")
            
            self.report['success'].append({
                'cliente': client,
                'id': vehicle_id,
                'chassi': vehicle_data.get('CHASSI', 'N/A'),
                'placa': vehicle_data.get('PLACA', 'N/A'),
                'odometer_updated': odometer_success
            })
            print(f"✅ Veículo ID {vehicle_id} processado com sucesso!")
            
        except Exception as e:
            self.report['errors'].append({
                'cliente': vehicle_data['CLIENTE'],
                'id': vehicle_data['ID'],
                'erro': str(e)
            })
            print(f"❌ Erro inesperado: {str(e)}")

    def generate_final_report(self):
        """Gera o relatório final da automação no terminal"""
        print("\n" + "="*80)
        print("RELATÓRIO FINAL DA AUTOMAÇÃO")
        print("="*80)
        
        total_vehicles = len(self.vehicles_data)
        success_count = len(self.report['success'])
        error_count = len(self.report['errors'])
        not_found_count = len(self.report['not_found'])
        manual_login_count = len(self.report['manual_login'])
        save_errors_count = len(self.report['save_errors'])
        odometer_errors_count = len(self.report['odometer_errors'])
        
        print(f"\n📊 RESUMO GERAL:")
        print(f"   Total de veículos processados: {total_vehicles}")
        print(f"   ✅ Sucessos: {success_count}")
        print(f"   ❌ Erros: {error_count}")
        print(f"   🔍 Não encontrados: {not_found_count}")
        print(f"   👤 Logins manuais: {manual_login_count}")
        print(f"   💾 Erros ao salvar: {save_errors_count}")
        print(f"   🔢 Erros no odômetro: {odometer_errors_count}")
        
        if success_count > 0:
            print(f"\n✅ VEÍCULOS PROCESSADOS COM SUCESSO ({success_count}):")
            for item in self.report['success']:
                odometer_status = "✓" if item.get('odometer_updated', False) else "✗"
                print(f"   Cliente: {item['cliente']} | ID: {item['id']} | Chassi: {item['chassi']} | Placa: {item['placa']} | Odômetro: {odometer_status}")
        
        if error_count > 0:
            print(f"\n❌ VEÍCULOS COM ERRO ({error_count}):")
            for item in self.report['errors']:
                print(f"   Cliente: {item['cliente']} | ID: {item['id']} | Erro: {item['erro']}")
        
        if save_errors_count > 0:
            print(f"\n💾 ERROS AO SALVAR FORMULÁRIO ({save_errors_count}):")
            for item in self.report['save_errors']:
                print(f"   Cliente: {item['cliente']} | ID: {item['id']} | Erro: {item['erro']}")
        
        if odometer_errors_count > 0:
            print(f"\n🔢 ERROS NA ATUALIZAÇÃO DO ODÔMETRO ({odometer_errors_count}):")
            for item in self.report['odometer_errors']:
                print(f"   Cliente: {item['cliente']} | ID: {item['id']} | Erro: {item['erro']}")
        
        if not_found_count > 0:
            print(f"\n🔍 VEÍCULOS NÃO ENCONTRADOS ({not_found_count}):")
            for item in self.report['not_found']:
                print(f"   Cliente: {item['cliente']} | ID: {item['id']} | Chassi: {item['chassi']} | Placa: {item['placa']}")
        
        if manual_login_count > 0:
            print(f"\n👤 CLIENTES COM LOGIN MANUAL ({manual_login_count}):")
            for item in self.report['manual_login']:
                print(f"   Cliente: {item['cliente']} | Motivo: {item['motivo']}")
        
        if total_vehicles > 0:
            print(f"\n📈 ESTATÍSTICAS POR CLIENTE:")
            clients = self.vehicles_data['CLIENTE'].unique()
            for client in clients:
                client_vehicles = self.vehicles_data[self.vehicles_data['CLIENTE'] == client]
                client_total = len(client_vehicles)
                client_success = len([x for x in self.report['success'] if x['cliente'] == client])
                client_errors = len([x for x in self.report['errors'] if x['cliente'] == client])
                client_not_found = len([x for x in self.report['not_found'] if x['cliente'] == client])
                client_save_errors = len([x for x in self.report['save_errors'] if x['cliente'] == client])
                client_odometer_errors = len([x for x in self.report['odometer_errors'] if x['cliente'] == client])
                
                success_rate = (client_success / client_total) * 100 if client_total > 0 else 0
                print(f"   {client}: {client_success}/{client_total} sucessos ({success_rate:.1f}%) | {client_errors} erros | {client_not_found} não encontrados | {client_save_errors} erros salvamento | {client_odometer_errors} erros odômetro")
        
        print("\n" + "="*80)
        print("AUTOMAÇÃO FINALIZADA")
        print("="*80)


if __name__ == "__main__":
    automation = VehicleAutomation()
    automation.run_automation()