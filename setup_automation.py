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
            'manual_login': []
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
            
            # Preenche placas vazias com chassi
            mask_placa_vazia = self.vehicles_data['PLACA'].isna() | (self.vehicles_data['PLACA'] == '')
            self.vehicles_data.loc[mask_placa_vazia, 'PLACA'] = self.vehicles_data.loc[mask_placa_vazia, 'CHASSI']
            
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
        """Carrega as credenciais fixas"""
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

    def ask_login_method(self):
        """Pergunta ao usuário sobre o método de login"""
        while True:
            print("\n" + "="*50)
            print("MÉTODO DE LOGIN")
            print("="*50)
            print("1 - Login Automático (usar credentials.json)")
            print("2 - Login Manual (você fará login manualmente)")
            print("="*50)
            
            choice = input("Escolha uma opção (1 ou 2): ").strip()
            
            if choice in ['1', '2']:
                return choice == '2'
            else:
                print("Opção inválida! Digite 1 ou 2.")

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
    
    def login_automatic(self, client):
        """Realiza login automático"""
        if client not in self.credentials:
            print(f"❌ Credenciais não encontradas para cliente: {client}")
            return False
        
        try:
            print(f"🔐 Login automático para cliente: {client}")
            
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
                print(f"✅ Login automático bem-sucedido: {client}")
                return True
            except TimeoutException:
                print(f"❌ Login falhou para cliente: {client}")
                return False
            
        except Exception as e:
            print(f"❌ Erro no login automático: {str(e)}")
            return False

    def login_manual(self, client):
        """Aguarda login manual do usuário"""
        try:
            self.driver.get("https://live.mzoneweb.net/mzonex/")
            time.sleep(2)
            
            print(f"\n{'='*60}")
            print(f"LOGIN MANUAL NECESSÁRIO PARA CLIENTE: {client}")
            print(f"{'='*60}")
            print("Realize o login manualmente e pressione ENTER...")
            
            input()
            
            current_url = self.driver.current_url
            if "workspace/map" in current_url:
                print(f"✅ Login manual bem-sucedido: {client}")
                return True
            else:
                while True:
                    resposta = input("O login foi realizado com sucesso? (s/n): ").lower().strip()
                    if resposta in ['s', 'sim', 'y', 'yes']:
                        return True
                    elif resposta in ['n', 'não', 'nao', 'no']:
                        return False
                    else:
                        print("Responda com 's' para sim ou 'n' para não.")
            
        except Exception as e:
            print(f"❌ Erro no login manual: {str(e)}")
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
        """Navega para a seção de veículos"""
        try:
            if self.vehicles_page_initialized:
                return True
                
            print("🚗 Navegando para seção de veículos...")
            
            self.wait_for_loading()
            
            vehicles_button = self.wait.until(
                EC.element_to_be_clickable((By.XPATH, "//span[contains(@class, 'button') and contains(., 'Veículos')]"))
            )
            vehicles_button.click()
            
            self.wait_for_loading()
            
            dropdown_button = self.wait.until(
                EC.element_to_be_clickable((By.XPATH, "//div[@class='wj-input-group wj-input-btn-visible']//button[@class='wj-btn wj-btn-default']"))
            )
            dropdown_button.click()
            
            time.sleep(0.6)
            
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

    def search_vehicle_by_id(self, vehicle_id):
        """Busca veículo pelo ID"""
        try:
            print(f"🔍 Buscando veículo ID: {vehicle_id}")
            
            self.wait_for_loading()
            
            search_field = self.wait.until(
                EC.presence_of_element_located((By.XPATH, "//input[@type='search' and @placeholder='Procurar veículos']"))
            )
            
            search_field.clear()
            time.sleep(0.5)
            search_field.send_keys(str(vehicle_id))
            
            self.wait_for_loading()
            time.sleep(3)
            
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
            print(f"❌ Erro ao buscar veículo: {str(e)}")
            return False

    def click_edit_vehicle(self, vehicle_id):
        """Clica no botão de editar veículo"""
        try:
            vehicle_row_xpath = f"//div[@class='wj-row'][.//div[contains(text(), '{vehicle_id}')]]"
            vehicle_row = self.wait.until(EC.presence_of_element_located((By.XPATH, vehicle_row_xpath)))
            
            edit_button = vehicle_row.find_element(By.XPATH, ".//i[@class='pointer mz7-pencil']")
            
            self.driver.execute_script("arguments[0].scrollIntoView(true);", edit_button)
            time.sleep(0.5)
            edit_button.click()
            
            self.wait_for_loading()
            return True
            
        except Exception as e:
            print(f"❌ Erro ao clicar em editar: {str(e)}")
            return False

    def cancel_modal(self):
        """Clica no botão cancelar do modal em caso de erro"""
        try:
            print("🚫 Cancelando modal devido a erro...")
            cancel_button = WebDriverWait(self.driver, 5).until(
                EC.element_to_be_clickable((By.XPATH, "//button[contains(@class, 'btn') and (contains(text(), 'Cancelar') or contains(text(), 'Cancel'))]"))
            )
            cancel_button.click()
            time.sleep(2)
            print("✅ Modal cancelado")
            return True
        except Exception as e:
            print(f"⚠️ Não foi possível cancelar o modal: {str(e)}")
            # Tenta fechar com ESC
            try:
                from selenium.webdriver.common.keys import Keys
                self.driver.find_element(By.TAG_NAME, 'body').send_keys(Keys.ESCAPE)
                time.sleep(2)
                print("✅ Modal fechado com ESC")
                return True
            except:
                print("❌ Falha ao fechar modal")
                return False

    def fill_vehicle_form(self, vehicle_data):
        """Preenche o formulário de edição do veículo com tratamento de erro melhorado"""
        try:
            print(f"📝 Preenchendo formulário para ID: {vehicle_data['ID']}")
            
            self.wait_for_loading()
            time.sleep(2)
            
            # Tenta preencher cada campo individualmente
            try:
                if not self.fill_description_field(vehicle_data):
                    raise Exception("Erro ao preencher descrição")
            except Exception as e:
                print(f"⚠️ Erro na descrição: {str(e)}")
                self.cancel_modal()
                return False
            
            try:
                if not self.fill_plate_field(vehicle_data):
                    raise Exception("Erro ao preencher placa")
            except Exception as e:
                print(f"⚠️ Erro na placa: {str(e)}")
                self.cancel_modal()
                return False
            
            try:
                if not self.fill_chassis_field(vehicle_data):
                    raise Exception("Erro ao preencher chassi")
            except Exception as e:
                print(f"⚠️ Erro no chassi: {str(e)}")
                self.cancel_modal()
                return False
            
            try:
                if not self.select_vehicle_group(vehicle_data):
                    raise Exception("Erro ao selecionar grupo")
            except Exception as e:
                print(f"⚠️ Erro no grupo: {str(e)}")
                self.cancel_modal()
                return False
            
            try:
                if not self.save_vehicle_form():
                    raise Exception("Erro ao salvar formulário")
            except Exception as e:
                print(f"⚠️ Erro ao salvar: {str(e)}")
                self.cancel_modal()
                return False
            
            print(f"✅ Formulário preenchido e salvo: ID {vehicle_data['ID']}")
            return True
            
        except Exception as e:
            print(f"❌ Erro geral no formulário: {str(e)}")
            self.cancel_modal()
            return False

    def fill_description_field(self, vehicle_data):
        """Preenche o campo descrição"""
        try:
            description = vehicle_data.get('DESCRIÇÃO', '')
            if pd.isna(description) or description == '':
                return True
            
            description_field = self.wait.until(
                EC.presence_of_element_located((By.XPATH, "//input[@name='description' and @placeholder='Descrição']"))
            )
            
            description_field.clear()
            time.sleep(0.5)
            description_field.send_keys(str(description))
            return True
            
        except Exception as e:
            print(f"❌ Erro ao preencher descrição: {str(e)}")
            return False

    def fill_plate_field(self, vehicle_data):
        """Preenche o campo placa"""
        try:
            plate = vehicle_data.get('PLACA', '')
            if pd.isna(plate) or plate == '':
                plate = vehicle_data.get('CHASSI', '')
            
            if pd.isna(plate) or plate == '':
                return True
            
            plate_field = self.wait.until(
                EC.presence_of_element_located((By.XPATH, "//input[@name='registration' and @placeholder='Placa']"))
            )
            
            plate_field.clear()
            time.sleep(0.5)
            plate_field.send_keys(str(plate))
            return True
            
        except Exception as e:
            print(f"❌ Erro ao preencher placa: {str(e)}")
            return False

    def fill_chassis_field(self, vehicle_data):
        """Preenche o campo chassi"""
        try:
            chassi = vehicle_data.get('CHASSI', '')
            if pd.isna(chassi) or chassi == '':
                return True
            
            chassis_field = self.wait.until(
                EC.presence_of_element_located((By.XPATH, "//input[@name='vin' and @placeholder='Número de Chassi']"))
            )
            
            chassis_field.clear()
            time.sleep(0.5)
            chassis_field.send_keys(str(chassi))
            return True
            
        except Exception as e:
            print(f"❌ Erro ao preencher chassi: {str(e)}")
            return False

    def select_vehicle_group(self, vehicle_data):
        """Seleciona o grupo de veículos"""
        try:
            group_name = vehicle_data.get('GRUPO DE VEICULOS', '')
            if pd.isna(group_name) or group_name == '':
                return True
            
            group_button = self.wait.until(
                EC.element_to_be_clickable((By.XPATH, "//li/a[contains(text(), 'Grupos de veículos')]"))
            )
            group_button.click()
            
            time.sleep(1)
            
            search_group_field = self.wait.until(
                EC.presence_of_element_located((By.XPATH, "//input[@placeholder='Buscar' and @class='form-input ng-untouched ng-pristine ng-valid']"))
            )
            
            search_group_field.clear()
            time.sleep(0.5)
            search_group_field.send_keys(str(group_name))
            
            time.sleep(3)
            
            return self.select_group_checkbox(group_name)
            
        except Exception as e:
            print(f"❌ Erro ao selecionar grupo: {str(e)}")
            return False

    def select_group_checkbox(self, group_name):
        """Seleciona o checkbox do grupo"""
        try:
            # Busca direta pelo nome
            checkbox_xpath = f"//div[@class='scrollbar-container checkboxlist-wrapper']//label[contains(text(), '{group_name}')]/input[@type='checkbox']"
            
            try:
                checkbox = WebDriverWait(self.driver, 5).until(
                    EC.element_to_be_clickable((By.XPATH, checkbox_xpath))
                )
                
                if not checkbox.is_selected():
                    checkbox.click()
                
                return True
            except TimeoutException:
                # Busca genérica
                checkboxes = self.driver.find_elements(By.XPATH, "//div[@class='scrollbar-container checkboxlist-wrapper']//label")
                
                for label in checkboxes:
                    label_text = label.text.strip()
                    if group_name.upper() in label_text.upper():
                        checkbox = label.find_element(By.XPATH, ".//input[@type='checkbox']")
                        if not checkbox.is_selected():
                            checkbox.click()
                        return True
                
                print(f"❌ Grupo '{group_name}' não encontrado")
                return False
            
        except Exception as e:
            print(f"❌ Erro ao selecionar checkbox: {str(e)}")
            return False

    def save_vehicle_form(self):
        """Salva o formulário"""
        try:
            save_button = self.wait.until(
                EC.element_to_be_clickable((By.XPATH, "//button[@type='submit' and contains(@class, 'success') and contains(text(), 'Salvar')]"))
            )
            
            save_button.click()
            
            self.wait_for_loading()
            time.sleep(2)
            
            # Verifica se o modal foi fechado
            try:
                WebDriverWait(self.driver, 3).until_not(
                    EC.presence_of_element_located((By.XPATH, "//button[@type='submit' and contains(text(), 'Salvar')]"))
                )
            except:
                pass
            
            return True
            
        except Exception as e:
            print(f"❌ Erro ao salvar: {str(e)}")
            return False

    def process_vehicle(self, vehicle_data):
        """Processa um veículo individual"""
        try:
            vehicle_id = vehicle_data['ID']
            client = vehicle_data['CLIENTE']
            
            print(f"\n🔧 PROCESSANDO VEÍCULO ID: {vehicle_id} | Cliente: {client}")
            
            if not self.navigate_to_vehicles():
                self.report['errors'].append({
                    'cliente': client,
                    'id': vehicle_id,
                    'erro': 'Erro ao navegar para veículos'
                })
                return
            
            if not self.search_vehicle_by_id(vehicle_id):
                self.report['not_found'].append({
                    'cliente': client,
                    'id': vehicle_id,
                    'chassi': vehicle_data.get('CHASSI', 'N/A'),
                    'placa': vehicle_data.get('PLACA', 'N/A'),
                    'odometro': vehicle_data.get('ODOMETRO', 'N/A')
                })
                return
            
            if not self.click_edit_vehicle(vehicle_id):
                self.report['errors'].append({
                    'cliente': client,
                    'id': vehicle_id,
                    'erro': 'Erro ao abrir modal de edição'
                })
                return
            
            if not self.fill_vehicle_form(vehicle_data):
                self.report['errors'].append({
                    'cliente': client,
                    'id': vehicle_id,
                    'erro': 'Erro ao preencher formulário'
                })
                return
            
            self.report['success'].append({
                'cliente': client,
                'id': vehicle_id,
                'chassi': vehicle_data.get('CHASSI', 'N/A'),
                'placa': vehicle_data.get('PLACA', 'N/A'),
                'odometro': vehicle_data.get('ODOMETRO', 'N/A')
            })
            print(f"✅ Veículo processado com sucesso: {vehicle_id}")
            
        except Exception as e:
            self.report['errors'].append({
                'cliente': vehicle_data['CLIENTE'],
                'id': vehicle_data['ID'],
                'erro': str(e)
            })
            print(f"❌ Erro inesperado: {str(e)}")

    def generate_final_report(self):
        """Gera o relatório final e prepara próxima automação"""
        print("\n" + "="*80)
        print("RELATÓRIO FINAL DA AUTOMAÇÃO")
        print("="*80)
        
        success_count = len(self.report['success'])
        error_count = len(self.report['errors'])
        not_found_count = len(self.report['not_found'])
        
        print(f"\n📊 RESUMO:")
        print(f"   ✅ Sucessos: {success_count}")
        print(f"   ❌ Erros: {error_count}")
        print(f"   🔍 Não encontrados: {not_found_count}")
        
        if success_count > 0:
            print(f"\n✅ VEÍCULOS PROCESSADOS COM SUCESSO:")
            for item in self.report['success']:
                print(f"   Cliente: {item['cliente']} | ID: {item['id']} | Chassi: {item['chassi']}")
        
        if error_count > 0:
            print(f"\n❌ VEÍCULOS COM ERRO:")
            for item in self.report['errors']:
                print(f"   Cliente: {item['cliente']} | ID: {item['id']} | Erro: {item['erro']}")
        
        if not_found_count > 0:
            print(f"\n🔍 VEÍCULOS NÃO ENCONTRADOS:")
            for item in self.report['not_found']:
                print(f"   Cliente: {item['cliente']} | ID: {item['id']}")
        
        print("="*80)
        
        # Prepara lista para próxima automação
        if success_count > 0:
            print("\n🚗 CARROS PARA AJUSTAR ODÔMETRO:")
            odometer_list = []
            for item in self.report['success']:
                odometer_info = f"Cliente: {item['cliente']} | ID: {item['id']} | Odômetro: {item['odometro']}"
                odometer_list.append(odometer_info)
                print(f"   {odometer_info}")
            
            # Salva lista em arquivo para próxima automação
            try:
                with open('carros_para_odometro.txt', 'w', encoding='utf-8') as f:
                    f.write("CARROS PARA AJUSTAR ODÔMETRO:\n")
                    f.write("="*50 + "\n")
                    for item in odometer_list:
                        f.write(f"{item}\n")
                print(f"\n💾 Lista salva em 'carros_para_odometro.txt'")
            except Exception as e:
                print(f"⚠️ Erro ao salvar lista: {str(e)}")
            
            print("\n🔄 INICIANDO PRÓXIMA AUTOMAÇÃO...")
            print("Copie a lista acima para a próxima automação de odômetro.")
        
        print("\n" + "="*80)
        print("AUTOMAÇÃO FINALIZADA")
        print("="*80)

    def run_automation(self):
        """Executa o fluxo principal da automação"""
        try:
            print("🚀" + "="*58 + "🚀")
            print("🤖        AUTOMAÇÃO  SETUP       🤖")
            print("🚀" + "="*58 + "🚀")
            
            if not self.load_excel_data():
                return False
            
            if not self.load_credentials():
                return False
            
            if not self.setup_driver():
                return False
            
            use_manual_login = self.ask_login_method()
            
            current_client = None
            total_vehicles = len(self.vehicles_data)
            processed_count = 0
            
            for index, vehicle in self.vehicles_data.iterrows():
                client = vehicle['CLIENTE']
                processed_count += 1
                
                print(f"\n🔢 Progresso: {processed_count}/{total_vehicles}")
                
                # Mudança de cliente
                if current_client != client:
                    if current_client is not None:
                        self.logout()
                    
                    current_client = client
                    self.vehicles_page_initialized = False
                    
                    print(f"\n👤 CLIENTE: {client}")
                    
                    if str(client).upper() == 'MANUAL' or use_manual_login or client not in self.credentials:
                        login_success = self.login_manual(client)
                    else:
                        login_success = self.login_automatic(client)
                    
                    if not login_success:
                        print(f"❌ Login falhou: {client}")
                        # Marca todos os veículos deste cliente como erro
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
                self.logout()
            
            self.generate_final_report()
            return True
            
        except Exception as e:
            print(f"❌ Erro no fluxo principal: {str(e)}")
            return False
        finally:
            if self.driver:
                print("🔒 Fechando navegador...")
                self.driver.quit()



def main():
    """Função principal da automação"""
    automation = VehicleAutomation()
    automation.run_automation()

if __name__ == '__main__':
    main()