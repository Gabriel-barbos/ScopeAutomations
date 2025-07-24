import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from datetime import datetime
import time
import logging
import locale

# Configurar logging para acompanhar o progresso
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# Configurar locale para português (para formatação de data)
try:
    locale.setlocale(locale.LC_TIME, 'pt_BR.UTF-8')
except:
    try:
        locale.setlocale(locale.LC_TIME, 'Portuguese_Brazil.1252')
    except:
        logger.warning("Não foi possível configurar locale português")

class ScopeBillingAutomation:
    def __init__(self, excel_file_path):
        self.excel_file_path = excel_file_path
        self.base_url = "https://billing.scopemp.net/Scope.Billing.Web/"
        self.contracts_url = "https://billing.scopemp.net/Scope.Billing.Web/ContractMaintenance.aspx"
        self.driver = None
        self.wait = None
        self.processed_count = 0
        self.error_count = 0
        self.termination_date = None
        self.total_contracts_terminated = 0
        
    def setup_driver_existing_session(self, debug_port=9222):
        """Conecta a uma sessão existente do Chrome"""
        try:
            chrome_options = Options()
            chrome_options.add_experimental_option("debuggerAddress", f"127.0.0.1:{debug_port}")
            
            self.driver = webdriver.Chrome(options=chrome_options)
            self.wait = WebDriverWait(self.driver, 15)
            logger.info("Conectado à sessão existente do Chrome")
            
        except Exception as e:
            logger.error(f"Erro ao conectar à sessão existente: {e}")
            logger.info("INSTRUÇÕES:")
            logger.info("1. Feche todas as instâncias do Chrome")
            logger.info("2. Abra o Chrome com: chrome --remote-debugging-port=9222")
            logger.info("3. Faça login normalmente no sistema")
            logger.info("4. Execute este script novamente")
            raise
    
    def setup_driver_new_session(self):
        """Cria uma nova sessão do Chrome (para caso não queira usar sessão existente)"""
        chrome_options = Options()
        chrome_options.add_argument("--disable-web-security")
        chrome_options.add_argument("--allow-running-insecure-content")
        chrome_options.add_argument("--window-size=1920,1080")
        
        self.driver = webdriver.Chrome(options=chrome_options)
        self.wait = WebDriverWait(self.driver, 15)
        
        logger.info("Nova sessão do Chrome criada")
        logger.info("Por favor, faça login manualmente no sistema e pressione Enter para continuar...")
        input()
        
    def read_excel_ids(self):
        """Lê os IDs do arquivo Excel"""
        try:
            df = pd.read_excel(self.excel_file_path)
            # Tenta diferentes nomes de coluna possíveis
            possible_columns = ['ID', 'id', 'Id', 'Equipment_ID', 'EquipmentID', 'Equipamento']
            
            id_column = None
            for col in possible_columns:
                if col in df.columns:
                    id_column = col
                    break
            
            if id_column is None:
                logger.info(f"Colunas disponíveis: {list(df.columns)}")
                id_column = input("Digite o nome da coluna que contém os IDs: ")
            
            ids = df[id_column].dropna().astype(str).tolist()
            logger.info(f"Lidos {len(ids)} IDs do arquivo Excel (coluna: {id_column})")
            return ids
            
        except Exception as e:
            logger.error(f"Erro ao ler arquivo Excel: {e}")
            return []
    
    def navigate_to_contracts(self):
        """Navega para a página de contratos"""
        try:
            self.driver.get(self.contracts_url)
            time.sleep(2)
            logger.info("Navegado para página de contratos")
        except Exception as e:
            logger.error(f"Erro ao navegar para contratos: {e}")
            raise
    
    def search_equipment(self, equipment_id):
        """Pesquisa um equipamento pelo ID"""
        try:
            # Limpar campo de pesquisa e inserir novo ID
            search_field = self.wait.until(
                EC.presence_of_element_located((By.ID, "ctl00_ContentPlaceHolder1_txt_BillableEntityDescription"))
            )
            
            search_field.clear()
            time.sleep(0.5)
            search_field.send_keys(str(equipment_id))
            search_field.send_keys(Keys.ENTER)
            
            logger.info(f"Pesquisa enviada para ID: {equipment_id}")
            
            # Aguardar a tabela de resultados carregar - método mais robusto
            try:
                # Aguardar a tabela aparecer
                self.wait.until(
                    EC.presence_of_element_located((By.ID, "ctl00_ContentPlaceHolder1_gv_SearchResults"))
                )
                
                # Aguardar um pouco mais para garantir que os dados carregaram
                time.sleep(2)
                
                # Verificar se há linhas de dados (não apenas cabeçalho)
                rows = self.driver.find_elements(By.XPATH, "//table[@id='ctl00_ContentPlaceHolder1_gv_SearchResults']//tr[position()>1]")
                
                if len(rows) > 0:
                    logger.info(f"✅ Tabela carregada com {len(rows)} linha(s) de dados")
                else:
                    logger.warning("⚠️ Tabela carregada mas sem dados encontrados")
                    
            except Exception as table_error:
                logger.error(f"Erro ao aguardar tabela: {table_error}")
                # Aguardar um tempo fixo como fallback
                time.sleep(3)
            
        except Exception as e:
            logger.error(f"Erro na pesquisa do equipamento {equipment_id}: {e}")
            raise
    
    def debug_table_structure(self):
        """Função de debug para mostrar a estrutura da tabela com foco no status"""
        try:
            table = self.driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_gv_SearchResults")
            rows = table.find_elements(By.TAG_NAME, "tr")
            
            logger.info("=== 🔍 DEBUG: Estrutura da tabela ===")
            
            # Mostrar cabeçalho
            header_row = rows[0] if rows else None
            if header_row:
                header_cells = header_row.find_elements(By.TAG_NAME, "th")
                header_texts = [cell.text.strip() for cell in header_cells]
                logger.info(f"📋 Cabeçalho: {header_texts}")
            
            # Mostrar dados
            data_rows = rows[1:] if len(rows) > 1 else []
            logger.info(f"📊 Total de linhas de dados: {len(data_rows)}")
            
            for i, row in enumerate(data_rows):
                cells = row.find_elements(By.TAG_NAME, "td")
                if cells:
                    # Focar nas colunas mais importantes
                    cell_texts = [cell.text.strip() for cell in cells[:6]]  # Primeiras 6 colunas
                    
                    # Destacar o status (3ª coluna)
                    status = cell_texts[2] if len(cell_texts) > 2 else "N/A"
                    logger.info(f"📄 Linha {i+1}: Status='{status}' | Dados: {cell_texts}")
                    
                    # Verificar se há links nesta linha
                    links = row.find_elements(By.TAG_NAME, "a")
                    if links:
                        link_info = []
                        for link in links:
                            link_text = link.text.strip()
                            link_id = link.get_attribute('id') or 'sem-id'
                            link_enabled = link.is_enabled()
                            link_info.append(f"{link_text}({link_id})[{'habilitado' if link_enabled else 'desabilitado'}]")
                        logger.info(f"   🔗 Links: {link_info}")
                        
            logger.info("=== 🔍 FIM DEBUG ===")
            
        except Exception as e:
            logger.error(f"❌ Erro no debug da tabela: {e}")

    def get_active_contracts(self):
        """Encontra APENAS os contratos com status 'Active' na página atual"""
        try:
            # Aguardar a tabela de resultados carregar
            time.sleep(3)
            
            # Aguardar a tabela aparecer
            search_results_table = self.wait.until(
                EC.presence_of_element_located((By.ID, "ctl00_ContentPlaceHolder1_gv_SearchResults"))
            )
            
            logger.info("Tabela de resultados encontrada, analisando contratos...")
            
            # Encontrar todas as linhas da tabela (excluindo cabeçalho)
            contract_rows = search_results_table.find_elements(By.XPATH, ".//tr[position()>1]")
            
            active_termination_links = []
            total_contracts = 0
            active_contracts = 0
            inactive_contracts = 0
            
            for i, row in enumerate(contract_rows):
                try:
                    # Buscar todas as células da linha atual
                    cells = row.find_elements(By.TAG_NAME, "td")
                    
                    if len(cells) < 3:  # Verificar se tem células suficientes
                        continue
                    
                    total_contracts += 1
                    
                    # A terceira célula (índice 2) contém o status "Active" ou "Inactive"
                    status_cell = cells[2]
                    status_text = status_cell.text.strip().upper()
                    
                    logger.info(f"📋 Linha {i+1}: Status = '{status_text}'")
                    
                    # APENAS processar se o status for exatamente "ACTIVE"
                    if status_text == "ACTIVE":
                        active_contracts += 1
                        
                        # Procurar pelo link "Termination" nesta linha
                        try:
                            termination_link = row.find_element(
                                By.XPATH, 
                                ".//a[contains(@id, 'lb_ContractTermination') and contains(text(), 'Termination')]"
                            )
                            
                            if termination_link.is_displayed() and termination_link.is_enabled():
                                active_termination_links.append(termination_link)
                                logger.info(f"✅ Contrato ATIVO encontrado na linha {i+1} - SERÁ PROCESSADO")
                            else:
                                logger.warning(f"⚠️ Contrato ativo na linha {i+1} mas link 'Termination' não está disponível")
                                
                        except Exception as link_error:
                            logger.warning(f"⚠️ Contrato ativo na linha {i+1} mas erro ao encontrar link 'Termination': {link_error}")
                            
                    elif "INACTIVE" in status_text or "CANCELLED" in status_text:
                        inactive_contracts += 1
                        logger.info(f"⏸️ Contrato na linha {i+1} está INATIVO - IGNORADO")
                        
                    else:
                        logger.info(f"❓ Status desconhecido na linha {i+1}: '{status_text}' - IGNORADO")
                        
                except Exception as row_error:
                    logger.error(f"❌ Erro ao processar linha {i+1}: {row_error}")
                    continue
            
            # Resumo dos contratos encontrados
            logger.info("="*50)
            logger.info("📊 RESUMO DOS CONTRATOS:")
            logger.info(f"📋 Total de contratos: {total_contracts}")
            logger.info(f"✅ Contratos ativos (para processar): {active_contracts}")
            logger.info(f"⏸️ Contratos inativos (ignorados): {inactive_contracts}")
            logger.info(f"🎯 Links válidos para cancelamento: {len(active_termination_links)}")
            logger.info("="*50)
            
            return active_termination_links
            
        except Exception as e:
            logger.error(f"❌ Erro ao buscar contratos ativos: {e}")
            return []
    
    def get_termination_date(self):
        """Solicita a data de terminação do usuário"""
        print("\n" + "="*50)
        print("CONFIGURAÇÃO DA DATA DE TERMINAÇÃO")
        print("="*50)
        
        today = datetime.now()
        months = {
            1: 'jan', 2: 'fev', 3: 'mar', 4: 'abr', 5: 'mai', 6: 'jun',
            7: 'jul', 8: 'ago', 9: 'set', 10: 'out', 11: 'nov', 12: 'dez'
        }
        suggested_date = f"{today.day} {months[today.month]} {today.year}"
        
        print(f"Data sugerida (hoje): {suggested_date}")
        print("\nOpções:")
        print("1. Usar data de hoje")
        print("2. Inserir data personalizada")
        
        while True:
            choice = input("\nEscolha uma opção (1 ou 2): ").strip()
            
            if choice == '1':
                self.termination_date = suggested_date
                print(f"✅ Data selecionada: {self.termination_date}")
                break
                
            elif choice == '2':
                print("\nFormato esperado: DD mmm AAAA (exemplo: 23 jul 2025)")
                print("Meses válidos: jan, fev, mar, abr, mai, jun, jul, ago, set, out, nov, dez")
                
                custom_date = input("Digite a data: ").strip()
                
                # Validar formato básico
                if self.validate_date_format(custom_date):
                    self.termination_date = custom_date
                    print(f"✅ Data selecionada: {self.termination_date}")
                    break
                else:
                    print("❌ Formato inválido. Tente novamente.")
                    
            else:
                print("❌ Opção inválida. Digite 1 ou 2.")
    
    def validate_date_format(self, date_str):
        """Valida se a data está no formato correto"""
        try:
            parts = date_str.split()
            if len(parts) != 3:
                return False
                
            day, month, year = parts
            
            # Verificar se dia é número
            day_num = int(day)
            if not (1 <= day_num <= 31):
                return False
                
            # Verificar se mês é válido
            valid_months = ['jan', 'fev', 'mar', 'abr', 'mai', 'jun', 
                           'jul', 'ago', 'set', 'out', 'nov', 'dez']
            if month.lower() not in valid_months:
                return False
                
            # Verificar se ano é número
            year_num = int(year)
            if not (2020 <= year_num <= 2030):  # Range razoável
                return False
                
            return True
            
        except:
            return False
    
    def terminate_contract(self, termination_link, equipment_id, contract_index):
        """Cancela um contrato específico"""
        try:
            logger.info(f"Cancelando contrato {contract_index} do equipamento {equipment_id}")
            
            # Clicar no link de Termination
            self.driver.execute_script("arguments[0].click();", termination_link)
            time.sleep(2)
            
            # Aguardar o campo de data aparecer
            date_field = self.wait.until(
                EC.presence_of_element_located((By.ID, "ctl00_ContentPlaceHolder1_txt_TerminationDate"))
            )
            
            # Inserir a data selecionada pelo usuário
            date_field.clear()
            time.sleep(0.5)
            date_field.send_keys(self.termination_date)
            
            logger.info(f"Data inserida: {self.termination_date}")
            
            # Aguardar e clicar no botão "Terminate"
            terminate_button = self.wait.until(
                EC.element_to_be_clickable((By.ID, "ctl00_ContentPlaceHolder1_cmd_Terminate"))
            )
            
            logger.info("Clicando no botão 'Terminate'...")
            terminate_button.click()
            
            # Aguardar o processamento
            time.sleep(3)
            
            logger.info(f"✅ Contrato {contract_index} cancelado com sucesso!")
            
        except Exception as e:
            logger.error(f"❌ Erro ao cancelar contrato {contract_index}: {e}")
            raise
    
    def process_equipment(self, equipment_id):
        """Processa um equipamento: pesquisa e cancela TODOS os contratos ativos"""
        try:
            logger.info(f"========== Processando equipamento: {equipment_id} ==========")
            
            contracts_terminated = 0
            attempt = 1
            max_attempts = 10  # Limite de segurança para evitar loop infinito
            
            while attempt <= max_attempts:
                logger.info(f"--- Tentativa {attempt}: Buscando contratos ativos para {equipment_id} ---")
                
                # Navegar para página de contratos
                self.navigate_to_contracts()
                
                # Pesquisar o equipamento
                self.search_equipment(equipment_id)
                
                # Encontrar contratos ativos
                active_contracts = self.get_active_contracts()
                
                if not active_contracts:
                    if contracts_terminated == 0:
                        logger.warning(f"❌ Nenhum contrato ATIVO encontrado para {equipment_id}")
                        logger.info("💡 Isso pode significar que:")
                        logger.info("   • Todos os contratos já estão inativos/cancelados")
                        logger.info("   • O ID do equipamento não existe")
                        logger.info("   • Não há contratos para este equipamento")
                        
                        # Ativar debug para ver a estrutura da tabela
                        logger.info("Executando debug da tabela para análise...")
                        self.debug_table_structure()
                        
                        # Perguntar se deve continuar ou pular
                        response = input(f"\nNenhum contrato ativo encontrado para {equipment_id}. Continuar para próximo? (s/n): ")
                        if response.lower() != 's':
                            return False
                        return True
                    else:
                        # Todos os contratos ativos foram processados
                        logger.info(f"🎉 SUCESSO! Todos os contratos ativos do equipamento {equipment_id} foram cancelados!")
                        logger.info(f"📊 Total de contratos cancelados: {contracts_terminated}")
                        break
                
                # Processar apenas o primeiro contrato ativo encontrado
                # (após cancelar, a página recarrega e precisamos buscar novamente)
                contract_link = active_contracts[0]
                contract_number = contracts_terminated + 1
                
                try:
                    logger.info(f"🎯 Cancelando contrato {contract_number} de {equipment_id}")
                    self.terminate_contract(contract_link, equipment_id, contract_number)
                    contracts_terminated += 1
                    self.total_contracts_terminated += 1
                    
                    # Pequena pausa entre cancelamentos
                    time.sleep(2)
                    
                except Exception as e:
                    logger.error(f"❌ Erro ao cancelar contrato {contract_number} do equipamento {equipment_id}: {e}")
                    
                    # Perguntar se deve tentar novamente ou pular este contrato
                    response = input(f"Erro ao cancelar contrato. Tentar novamente? (s/n): ")
                    if response.lower() != 's':
                        break
                
                attempt += 1
            
            if attempt > max_attempts:
                logger.warning(f"⚠️ Atingido limite máximo de tentativas para {equipment_id}")
            
            if contracts_terminated > 0:
                self.processed_count += 1
                logger.info(f"🎉 Equipamento {equipment_id} processado com sucesso! ({contracts_terminated} contratos cancelados)")
                return True
            else:
                logger.warning(f"⚠️ Nenhum contrato foi cancelado para {equipment_id}")
                return False
            
        except Exception as e:
            logger.error(f"❌ Erro crítico ao processar equipamento {equipment_id}: {e}")
            self.error_count += 1
            return False
    
    def run_automation(self, use_existing_session=True):
        """Executa todo o processo de automação"""
        try:
            # Configurar driver
            if use_existing_session:
                self.setup_driver_existing_session()
            else:
                self.setup_driver_new_session()
            
            # Solicitar data de terminação
            self.get_termination_date()
            
            # Ler IDs do Excel
            equipment_ids = self.read_excel_ids()
            
            if not equipment_ids:
                logger.error("Nenhum ID encontrado no arquivo Excel")
                return
            
            logger.info(f"Iniciando processamento de {len(equipment_ids)} equipamentos")
            
            # Processar cada equipamento
            for i, equipment_id in enumerate(equipment_ids, 1):
                try:
                    logger.info(f"\n--- Progresso: {i}/{len(equipment_ids)} ---")
                    success = self.process_equipment(equipment_id)
                    
                    if not success:
                        # Perguntar se deve continuar em caso de erro
                        response = input(f"Erro ao processar {equipment_id}. Continuar? (s/n): ")
                        if response.lower() != 's':
                            break
                    
                    # Pausa entre equipamentos
                    time.sleep(1)
                    
                except KeyboardInterrupt:
                    logger.info("Automação interrompida pelo usuário")
                    break
                except Exception as e:
                    logger.error(f"Erro crítico ao processar {equipment_id}: {e}")
                    self.error_count += 1
                    continue
            
            # Relatório final
            logger.info("\n" + "="*60)
            logger.info("🎉 RELATÓRIO FINAL DA AUTOMAÇÃO:")
            logger.info("="*60)
            logger.info(f"📋 Total de equipamentos processados: {len(equipment_ids)}")
            logger.info(f"✅ Equipamentos com sucesso: {self.processed_count}")
            logger.info(f"❌ Equipamentos com erro: {self.error_count}")
            logger.info(f"🎯 Total de contratos cancelados: {self.total_contracts_terminated}")
            logger.info(f"📅 Data de terminação usada: {self.termination_date}")
            
            success_rate = (self.processed_count / len(equipment_ids)) * 100 if equipment_ids else 0
            logger.info(f"📊 Taxa de sucesso: {success_rate:.1f}%")
            logger.info("="*60)
            
        except Exception as e:
            logger.error(f"Erro geral na automação: {e}")
        finally:
            if self.driver:
                input("Pressione Enter para fechar o navegador...")
                self.driver.quit()

# Exemplo de uso
if __name__ == "__main__":
    # Configurações
    EXCEL_FILE = "ID_billing.xlsx"  # Ajuste o caminho do seu arquivo
    
    print("=== AUTOMAÇÃO DE CANCELAMENTO DE CONTRATOS ===")
    print("URL: https://billing.scopemp.net/Scope.Billing.Web/")
    print()
    
    # Instruções para uso com sessão existente
    print("INSTRUÇÕES PARA USO:")
    print("1. Feche todas as instâncias do Chrome")
    print("2. Abra o Chrome com: chrome --remote-debugging-port=9222")
    print("3. Faça login no sistema Scope Billing")
    print("4. Execute este script")
    print()
    
    use_existing = input("Usar sessão existente do Chrome? (s/n): ").lower() == 's'
    
    # Criar instância da automação
    automation = ScopeBillingAutomation(EXCEL_FILE)
    
    # Executar automação
    automation.run_automation(use_existing_session=use_existing)