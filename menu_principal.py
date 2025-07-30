import os
import sys
import subprocess
from pathlib import Path
import importlib.util
import time
import random

def limpar_tela():
    """Limpa a tela do terminal"""
    os.system('cls' if os.name == 'nt' else 'clear')

def animacao_loading(texto="Carregando", duracao=0.5):
    """Exibe uma animação de loading"""
    frames = ["⠋", "⠙", "⠹", "⠸", "⠼", "⠴", "⠦", "⠧", "⠇", "⠏"]
    end_time = time.time() + duracao
    i = 0
    while time.time() < end_time:
        print(f"\r{frames[i % len(frames)]} {texto}...", end='', flush=True)
        time.sleep(0.1)
        i += 1
    print("\r" + " " * (len(texto) + 10), end='\r')

def verificar_arquivos():
    """Verifica se os arquivos necessários existem"""
    arquivos_necessarios = [
        'add_automation.py',
        'remove_automation.py', 
        'billing_automation.py',
        'qtgo_automation.py',
        'setup_automation.py',      
        'odometer_setup.py',        
        'AdicionarGrupo.xlsx',
        'RemoverGrupo.xlsx',
        'ID_billing.xlsx',
        'QTGO_ID.xlsx'
    ]
    
    arquivos_faltando = []
    for arquivo in arquivos_necessarios:
        if not Path(arquivo).exists():
            arquivos_faltando.append(arquivo)
    
    return arquivos_faltando

def mostrar_banner():
    """Exibe o banner do sistema com ASCII art"""
    cores = {
        'azul': '\033[94m',
        'verde': '\033[92m',
        'amarelo': '\033[93m',
        'vermelho': '\033[91m',
        'roxo': '\033[95m',
        'ciano': '\033[96m',
        'reset': '\033[0m',
        'negrito': '\033[1m'
    }
    
    # ASCII art do JARVIS
    jarvis_art = f"""
{cores['ciano']}
    ╔══════════════════════════════════════════════════════════════╗
    ║                                                              ║
    ║     ██╗ █████╗ ██████╗ ██╗   ██╗██╗███████╗                ║
    ║     ██║██╔══██╗██╔══██╗██║   ██║██║██╔════╝                ║
    ║     ██║███████║██████╔╝██║   ██║██║███████╗                ║
    ║██   ██║██╔══██║██╔══██╗╚██╗ ██╔╝██║╚════██║                ║
    ║╚█████╔╝██║  ██║██║  ██║ ╚████╔╝ ██║███████║                ║
    ║ ╚════╝ ╚═╝  ╚═╝╚═╝  ╚═╝  ╚═══╝  ╚═╝╚══════╝                ║
    ║                                                              ║
    ║{cores['amarelo']}           🤖  robô  para Scope 🤖{cores['ciano']}              ║
    ╚══════════════════════════════════════════════════════════════╝
{cores['reset']}"""
    
    print(jarvis_art)
    
    # Mensagem aleatória do JARVIS
    mensagens = [
        "Boa tarde",
        "Jarvis Online !",
        "Bom dia",
        "Bem-vindo !",
        "JARVIS ativo"
    ]
    
    print(f"{cores['verde']}    💬 {random.choice(mensagens)}{cores['reset']}")
    print()

def mostrar_menu():
    """Exibe as opções do menu com visual melhorado"""
    cores = {
        'titulo': '\033[96m',
        'opcao': '\033[93m',
        'descricao': '\033[37m',
        'reset': '\033[0m',
        'verde': '\033[92m',
        'vermelho': '\033[91m'
    }
    
    print(f"{cores['titulo']}┌─────────────────────────────────────────────────────────────┐")
    print(f"│                    🎯 MENU PRINCIPAL 🎯                     │")
    print(f"└─────────────────────────────────────────────────────────────┘{cores['reset']}")
    print()
    
    opcoes = [
        ("1", "🚗  Adicionar Carros ao Grupo", "Adiciono Veiculos em grupo especifico"),
        ("2", "🗑️   Remover Carros do Grupo", "Removo veículos do grupo selecionado"),
        ("3", "💰  Remover Unidades do Billing", "Removo unidades do Billing"),
        ("4", "🚙  Remover Carros do QTGO", "Solicitação desinstalação no QTGO"),
        ("5", "⚙️   Setup Automático", "Faço Setup de veiculos"),
        ("6", "📏  Ajuste de Odômetro", "Ajusto odômetro automaticamente"),
        ("7", "🔍  Verificar Sistema", "Verifica status dos arquivos e planilhas"),
        ("0", "🚪  Sair", "Encerra o sistema")
    ]
    
    for opcao, titulo, descricao in opcoes:
        if opcao == "0":
            print(f"\n{cores['vermelho']}  [{opcao}] {titulo}{cores['reset']}")
        else:
            print(f"{cores['opcao']}  [{opcao}] {titulo}{cores['reset']}")
        print(f"{cores['descricao']}       └─ {descricao}{cores['reset']}")
    
    print(f"\n{cores['titulo']}━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━{cores['reset']}")

def executar_script(nome_script, titulo_operacao):
    """Executa um script Python com visual melhorado"""
    try:
        print(f"\n🚀 {titulo_operacao}")
        print("─" * 50)
        animacao_loading(f"Iniciando {nome_script}", 1)
        
        # Verifica se está executando como executável (PyInstaller)
        if getattr(sys, 'frozen', False):
            # Executando como executável - importa e executa diretamente
            nome_modulo = nome_script.replace('.py', '')
            
            try:
                # Importação dinâmica baseada no nome do módulo
                if nome_modulo in ['add_automation', 'remove_automation', 'billing_automation', 
                                  'qtgo_automation', 'setup_automation', 'odometer_setup']:
                    module = importlib.import_module(nome_modulo)
                    if hasattr(module, 'main'):
                        module.main()
                    else:
                        # Se não tem função main, executa o código principal
                        exec(compile(open(nome_script, 'rb').read(), nome_script, 'exec'))
                        
                print(f"\n✅ {titulo_operacao} concluída com sucesso!")
                
            except ImportError:
                # Se não conseguir importar, tenta executar como arquivo
                print(f"⚠️  Tentando execução direta...")
                try:
                    if hasattr(sys, '_MEIPASS'):
                        script_path = os.path.join(sys._MEIPASS, nome_script)
                    else:
                        script_path = nome_script
                    
                    with open(script_path, 'rb') as f:
                        code = compile(f.read(), script_path, 'exec')
                        exec(code)
                        
                except Exception as e:
                    print(f"❌ Erro na execução: {str(e)}")
                    raise
                    
                print(f"\n✅ {titulo_operacao} concluída com sucesso!")
                
        else:
            # Executando como script Python normal - usa subprocess
            resultado = subprocess.run([sys.executable, nome_script], 
                                     capture_output=False, 
                                     text=True)
            
            if resultado.returncode == 0:
                print(f"\n✅ {titulo_operacao} concluída com sucesso!")
            else:
                print(f"\n❌ Erro ao executar {titulo_operacao}")
        
        print("─" * 50)
            
    except FileNotFoundError:
        print(f"\n❌ Arquivo {nome_script} não encontrado!")
        print("💡 Certifique-se de que o arquivo está na mesma pasta")
    except Exception as e:
        print(f"\n❌ Erro inesperado: {str(e)}")
        print(f"💡 Detalhes do erro: {type(e).__name__}")
    
    input("\n🔙 Pressione ENTER para voltar ao menu...")

def verificar_planilhas():
    """Verifica e exibe informações sobre as planilhas com visual melhorado"""
    planilhas = {
        'AdicionarGrupo.xlsx': ('Planilha para adicionar carros ao grupo', '🚗'),
        'RemoverGrupo.xlsx': ('Planilha para remover carros do grupo', '🗑️'),
        'ID_billing.xlsx': ('Planilha com IDs para billing', '💰'),
        'QTGO_ID.xlsx': ('Planilha para remover carros do QTGO', '🚙')
    }
    
    print("\n📊 STATUS DAS PLANILHAS:")
    print("┌" + "─" * 58 + "┐")
    
    todas_ok = True
    for planilha, (descricao, emoji) in planilhas.items():
        if Path(planilha).exists():
            tamanho = Path(planilha).stat().st_size
            print(f"│ ✅ {emoji} {planilha:<25} │ {tamanho:>10} bytes │")
        else:
            print(f"│ ❌ {emoji} {planilha:<25} │ {'NÃO ENCONTRADA':>16} │")
            todas_ok = False
    
    print("└" + "─" * 58 + "┘")
    
    if todas_ok:
        print("\n✨ Todas as planilhas estão disponíveis!")
    else:
        print("\n⚠️  Algumas planilhas estão faltando!")

def tela_saida():
    """Exibe uma tela de saída animada"""
    limpar_tela()
    print("\n" * 5)
    print("    👋 Obrigado por usar o JARVIS!")
    print("    " + "─" * 30)
    
    mensagens_saida = [
        "Desligando sistemas...",
        "Salvando configurações...",
        "Encerrando protocolos...",
        "Até a próxima!"
    ]
    
    for msg in mensagens_saida:
        animacao_loading(msg, 0.5)
        print(f"    ✓ {msg}")
        time.sleep(0.3)
    
    print("\n    🤖 JARVIS desconectado.")
    print("\n" * 2)

def main():
    """Função principal do menu"""
    
    # Tela de boas-vindas
    limpar_tela()
    animacao_loading("Inicializando JARVIS", 1.5)
    
    while True:
        limpar_tela()
        mostrar_banner()
        
        # Verifica arquivos faltando
        arquivos_faltando = verificar_arquivos()
        if arquivos_faltando:
            print("⚠️  ATENÇÃO: Alguns arquivos estão faltando:")
            for arquivo in arquivos_faltando[:3]:  # Mostra apenas os 3 primeiros
                print(f"   • {arquivo}")
            if len(arquivos_faltando) > 3:
                print(f"   • ... e mais {len(arquivos_faltando) - 3} arquivo(s)")
            print()
        
        mostrar_menu()
        
        try:
            opcao = input("\n💡 Digite sua opção: ").strip()
            
            if opcao == '1':
                limpar_tela()
                executar_script('add_automation.py', 'ADICIONAR CARROS AO GRUPO')
                
            elif opcao == '2':
                limpar_tela()
                executar_script('remove_automation.py', 'REMOVER CARROS DO GRUPO')
                
            elif opcao == '3':
                limpar_tela()
                executar_script('billing_automation.py', 'REMOVER UNIDADES DO BILLING')
                
            elif opcao == '4':
                limpar_tela()
                executar_script('qtgo_automation.py', 'REMOVER CARROS DO QTGO')
                
            elif opcao == '5':
                limpar_tela()
                executar_script('setup_automation.py', 'SETUP AUTOMÁTICO DO SISTEMA')
                
            elif opcao == '6':
                limpar_tela()
                executar_script('odometer_setup.py', 'AJUSTE AUTOMÁTICO DE ODÔMETRO')
                
            elif opcao == '7':
                limpar_tela()
                print("\n🔍 VERIFICAÇÃO COMPLETA DO SISTEMA")
                print("═" * 60)
                
                # Verifica planilhas
                verificar_planilhas()
                
                # Verifica scripts
                print("\n📜 STATUS DOS SCRIPTS:")
                print("┌" + "─" * 58 + "┐")
                
                scripts = [
                    ('add_automation.py', 'Adicionar ao grupo', '🚗'),
                    ('remove_automation.py', 'Remover do grupo', '🗑️'),
                    ('billing_automation.py', 'Automação billing', '💰'),
                    ('qtgo_automation.py', 'Automação QTGO', '🚙'),
                    ('setup_automation.py', 'Setup automático', '⚙️'),
                    ('odometer_setup.py', 'Ajuste odômetro', '📏')
                ]
                
                todos_ok = True
                for script, descricao, emoji in scripts:
                    if Path(script).exists():
                        print(f"│ ✅ {emoji} {script:<25} │ {'DISPONÍVEL':>16} │")
                    else:
                        print(f"│ ❌ {emoji} {script:<25} │ {'NÃO ENCONTRADO':>16} │")
                        todos_ok = False
                
                print("└" + "─" * 58 + "┘")
                
                if todos_ok:
                    print("\n✨ Todos os scripts estão disponíveis!")
                else:
                    print("\n⚠️  Alguns scripts estão faltando!")
                
                # Informações do sistema
                print("\n💻 INFORMAÇÕES DO SISTEMA:")
                print("┌" + "─" * 58 + "┐")
                print(f"│ 🐍 Python: {sys.version.split()[0]:<45} │")
                print(f"│ 💾 Diretório: {os.getcwd()[:42]:<42} │")
                print(f"│ 🖥️  Sistema: {sys.platform:<44} │")
                print("└" + "─" * 58 + "┘")
                
                input("\n🔙 Pressione ENTER para voltar ao menu...")
                
            elif opcao == '0':
                tela_saida()
                break
                
            else:
                print("\n❌ Opção inválida!")
                print("💡 Por favor, digite um número entre 0 e 7.")
                animacao_loading("Voltando ao menu", 1)
                
        except KeyboardInterrupt:
            print("\n\n⚠️  Operação cancelada pelo usuário.")
            resposta = input("Deseja sair do sistema? (S/N): ").strip().upper()
            if resposta == 'S':
                tela_saida()
                break
            else:
                animacao_loading("Voltando ao menu", 1)
                
        except Exception as e:
            print(f"\n❌ Erro inesperado: {str(e)}")
            print("🔧 Tentando recuperar...")
            animacao_loading("Reiniciando menu", 2)

# Adiciona algumas funções extras para melhorar a experiência

def exibir_dica_aleatoria():
    """Exibe uma dica aleatória sobre o sistema"""
    dicas = [
        "💡 Dica: Mantenha suas planilhas sempre atualizadas!",
        "💡 Dica: Use Ctrl+C para cancelar uma operação em andamento.",
        "💡 Dica: Verifique o sistema regularmente para garantir que tudo está OK.",
        "💡 Dica: O JARVIS pode processar múltiplas planilhas sequencialmente.",
        "💡 Dica: Faça backup das suas planilhas antes de executar as automações."
    ]
    return random.choice(dicas)

def criar_arquivo_log():
    """Cria um arquivo de log para registrar as operações"""
    from datetime import datetime
    
    log_dir = Path("logs")
    if not log_dir.exists():
        log_dir.mkdir()
    
    log_file = log_dir / f"jarvis_{datetime.now().strftime('%Y%m%d')}.log"
    return log_file

# Modifica a função mostrar_banner para incluir dicas
def mostrar_banner_com_dica():
    """Exibe o banner do sistema com uma dica aleatória"""
    mostrar_banner()
    print(f"\n{exibir_dica_aleatoria()}\n")

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"\n❌ Erro crítico no sistema: {str(e)}")
        print("📧 Por favor, contate o suporte técnico.")
        input("\nPressione ENTER para encerrar...")