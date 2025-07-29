import os
import sys
import subprocess
from pathlib import Path
import importlib.util

def limpar_tela():
    """Limpa a tela do terminal"""
    os.system('cls' if os.name == 'nt' else 'clear')

def verificar_arquivos():
    """Verifica se os arquivos necessários existem"""
    arquivos_necessarios = [
        'add_automation.py',
        'remove_automation.py', 
        'billing_automation.py',
        'qtgo_automation.py',  # Novo arquivo adicionado
        'AdicionarGrupo.xlsx',
        'RemoverGrupo.xlsx',
        'ID_billing.xlsx',
        'QTGO_ID.xlsx'  # Nova planilha para o QTGO
    ]
    
    arquivos_faltando = []
    for arquivo in arquivos_necessarios:
        if not Path(arquivo).exists():
            arquivos_faltando.append(arquivo)
    
    return arquivos_faltando

def mostrar_banner():
    """Exibe o banner do sistema"""
    print("=" * 60)
    print("           SCOPE AUTOMATIONS - MENU PRINCIPAL")
    print("=" * 60)
    print()

def mostrar_menu():
    """Exibe as opções do menu"""
    print("Escolha uma opção:")
    print()
    print("1 - Adicionar Carros ao Grupo de Veículos")
    print("2 - Remover Carros do Grupo de Veículos") 
    print("3 - Remover Unidades do Billing")
    print("4 - Remover Carros do QTGO")  # Nova opção
    print("5 - Verificar Arquivos do Sistema")  # Opção renumerada
    print("0 - Sair")
    print()
    print("-" * 60)

def executar_script(nome_script):
    """Executa um script Python"""
    try:
        print(f"\n🚀 Iniciando {nome_script}...")
        print("-" * 40)
        
        # Verifica se está executando como executável (PyInstaller)
        if getattr(sys, 'frozen', False):
            # Executando como executável - importa e executa diretamente
            nome_modulo = nome_script.replace('.py', '')
            
            try:
                # Tenta importar o módulo
                if nome_modulo == 'add_automation':
                    import add_automation
                    if hasattr(add_automation, 'main'):
                        add_automation.main()
                    else:
                        # Se não tem função main, executa o código principal
                        exec(compile(open(nome_script, 'rb').read(), nome_script, 'exec'))
                        
                elif nome_modulo == 'remove_automation':
                    import remove_automation
                    if hasattr(remove_automation, 'main'):
                        remove_automation.main()
                    else:
                        exec(compile(open(nome_script, 'rb').read(), nome_script, 'exec'))
                        
                elif nome_modulo == 'billing_automation':
                    import billing_automation
                    if hasattr(billing_automation, 'main'):
                        billing_automation.main()
                    else:
                        exec(compile(open(nome_script, 'rb').read(), nome_script, 'exec'))

                elif nome_modulo == 'qtgo_automation':  # Novo módulo
                    try:
                        import qtgo_automation
                        if hasattr(qtgo_automation, 'main'):
                            qtgo_automation.main()
                        else:
                            # Força reimportação
                            import importlib
                            importlib.reload(qtgo_automation)
                    except Exception as e:
                        print(f"⚠️  Erro na importação: {e}")
                        # Fallback para execução de arquivo
                        script_path = nome_script
                        if hasattr(sys, '_MEIPASS'):
                            script_path = os.path.join(sys._MEIPASS, nome_script)
                        
                        # Tenta modo binário primeiro (mais confiável)
                        try:
                            with open(script_path, 'rb') as f:
                                code = compile(f.read(), script_path, 'exec')
                                exec(code)
                        except Exception:
                            # Último recurso: UTF-8 com ignore
                            with open(script_path, 'r', encoding='utf-8', errors='ignore') as f:
                                exec(f.read())
                        
                print(f"✅ {nome_script} executado com sucesso!")
                
            except ImportError:
                # Se não conseguir importar, tenta executar como arquivo
                print(f"⚠️  Importação falhou, tentando execução direta...")
                try:
                    # Determina o caminho correto do arquivo
                    if hasattr(sys, '_MEIPASS'):
                        script_path = os.path.join(sys._MEIPASS, nome_script)
                    else:
                        script_path = nome_script
                    
                    # Tenta diferentes codificações
                    for encoding in ['utf-8', 'latin-1', 'cp1252']:
                        try:
                            with open(script_path, 'r', encoding=encoding) as f:
                                exec(f.read())
                            break
                        except UnicodeDecodeError:
                            continue
                    else:
                        # Último recurso: modo binário
                        exec(compile(open(script_path, 'rb').read(), script_path, 'exec'))
                        
                except Exception as e:
                    print(f"❌ Erro na execução direta: {str(e)}")
                    raise
                    
                print(f"✅ {nome_script} executado com sucesso!")
                
        else:
            # Executando como script Python normal - usa subprocess
            resultado = subprocess.run([sys.executable, nome_script], 
                                     capture_output=False, 
                                     text=True)
            
            if resultado.returncode == 0:
                print(f"✅ {nome_script} executado com sucesso!")
            else:
                print(f"❌ Erro ao executar {nome_script}")
        
        print("-" * 40)
            
    except FileNotFoundError:
        print(f"❌ Arquivo {nome_script} não encontrado!")
        print("💡 Certifique-se de que o arquivo está na mesma pasta")
    except Exception as e:
        print(f"❌ Erro inesperado: {str(e)}")
        print(f"💡 Detalhes do erro: {type(e).__name__}")
    
    input("\nPressione ENTER para continuar...")

def verificar_planilhas():
    """Verifica e exibe informações sobre as planilhas"""
    planilhas = {
        'AdicionarGrupo.xlsx': 'Planilha para adicionar carros ao grupo',
        'RemoverGrupo.xlsx': 'Planilha para remover carros do grupo',
        'ID_billing.xlsx': 'Planilha com IDs para billing',
        'QTGO_ID.xlsx': 'Planilha para remover carros do QTGO'  # Nova planilha
    }
    
    print("\n📊 STATUS DAS PLANILHAS:")
    print("-" * 40)
    
    for planilha, descricao in planilhas.items():
        if Path(planilha).exists():
            tamanho = Path(planilha).stat().st_size
            print(f"✅ {planilha} - {descricao}")
            print(f"   Tamanho: {tamanho} bytes")
        else:
            print(f"❌ {planilha} - NÃO ENCONTRADA")
        print()

def main():
    """Função principal do menu"""
    
    while True:
        limpar_tela()
        mostrar_banner()
        
        # Verifica arquivos faltando
        arquivos_faltando = verificar_arquivos()
        if arquivos_faltando:
            print("⚠️  ATENÇÃO: Arquivos faltando:")
            for arquivo in arquivos_faltando:
                print(f"   - {arquivo}")
            print()
        
        mostrar_menu()
        
        try:
            opcao = input("Digite sua opção: ").strip()
            
            if opcao == '1':
                limpar_tela()
                print("🚗 ADICIONAR CARROS AO GRUPO")
                print("Certifique-se de que a planilha 'AdicionarGrupo.xlsx' está preenchida")
                input("Pressione ENTER para continuar ou CTRL+C para cancelar...")
                executar_script('add_automation.py')
                
            elif opcao == '2':
                limpar_tela()
                print("🗑️  REMOVER CARROS DO GRUPO")
                print("Certifique-se de que a planilha 'RemoverGrupo.xlsx' está preenchida")
                input("Pressione ENTER para continuar ou CTRL+C para cancelar...")
                executar_script('remove_automation.py')
                
            elif opcao == '3':
                limpar_tela()
                print("💰 REMOVER UNIDADES DO BILLING")
                print("Certifique-se de que a planilha 'ID_billing.xlsx' está preenchida")
                input("Pressione ENTER para continuar ou CTRL+C para cancelar...")
                executar_script('billing_automation.py')
                
            elif opcao == '4':  # Nova opção
                limpar_tela()
                print("🚙 REMOVER CARROS DO QTGO")
                print("Certifique-se de que a planilha 'QTGO_ID.xlsx' está preenchida")
                input("Pressione ENTER para continuar ou CTRL+C para cancelar...")
                executar_script('qtgo_automation.py')
                
            elif opcao == '5':  # Opção renumerada
                limpar_tela()
                print("🔍 VERIFICAÇÃO DO SISTEMA")
                verificar_planilhas()
                
                print("\n📜 SCRIPTS DISPONÍVEIS:")
                print("-" * 40)
                scripts = ['add_automation.py', 'remove_automation.py', 'billing_automation.py', 'qtgo_automation.py']  # Lista atualizada
                for script in scripts:
                    if Path(script).exists():
                        print(f"✅ {script}")
                    else:
                        print(f"❌ {script} - NÃO ENCONTRADO")
                
                input("\nPressione ENTER para voltar ao menu...")
                
            elif opcao == '0':
                limpar_tela()
                print("👋 Obrigado por usar o Scope Automations!")
                print("Sistema encerrado.")
                break
                
            else:
                print("\n❌ Opção inválida! Digite um número de 0 a 5.")  # Faixa atualizada
                input("Pressione ENTER para tentar novamente...")
                
        except KeyboardInterrupt:
            print("\n\n⚠️  Operação cancelada pelo usuário.")
            input("Pressione ENTER para voltar ao menu...")
        except Exception as e:
            print(f"\n❌ Erro inesperado: {str(e)}")
            input("Pressione ENTER para continuar...")

if __name__ == "__main__":
    main()