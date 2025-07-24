"""
Script para preparar os arquivos de automação para o build
Execute antes de gerar o executável
"""

import os
from pathlib import Path

def adicionar_funcao_main(nome_arquivo):
    """Adiciona função main() ao arquivo se não existir"""
    
    if not Path(nome_arquivo).exists():
        print(f"❌ {nome_arquivo} não encontrado!")
        return False
    
    # Lê o conteúdo atual
    with open(nome_arquivo, 'r', encoding='utf-8') as f:
        conteudo = f.read()
    
    # Verifica se já tem função main
    if 'def main():' in conteudo:
        print(f"✅ {nome_arquivo} já tem função main()")
        return True
    
    # Verifica se tem if __name__ == "__main__":
    if 'if __name__ == "__main__":' in conteudo:
        # Substitui por estrutura com função main
        linhas = conteudo.split('\n')
        novas_linhas = []
        encontrou_main_block = False
        indentacao_main = ""
        
        for linha in linhas:
            if 'if __name__ == "__main__":' in linha:
                encontrou_main_block = True
                # Adiciona função main antes
                novas_linhas.append("")
                novas_linhas.append("def main():")
                novas_linhas.append("    \"\"\"Função principal da automação\"\"\"")
                continue
            elif encontrou_main_block and linha.strip():
                # Pega a indentação da primeira linha após o if
                if not indentacao_main:
                    indentacao_main = len(linha) - len(linha.lstrip())
                
                # Remove uma indentação e adiciona dentro da função main
                linha_limpa = linha[indentacao_main:] if len(linha) > indentacao_main else linha
                novas_linhas.append("    " + linha_limpa)
            else:
                novas_linhas.append(linha)
        
        # Adiciona chamada da função main
        novas_linhas.append("")
        novas_linhas.append("if __name__ == '__main__':")
        novas_linhas.append("    main()")
        
        # Salva o arquivo modificado
        conteudo_novo = '\n'.join(novas_linhas)
        
    else:
        # Não tem if __name__, adiciona função main envolvendo todo o código
        linhas = conteudo.split('\n')
        
        # Encontra onde começam as importações
        imports_fim = 0
        for i, linha in enumerate(linhas):
            if linha.strip() and not (linha.strip().startswith('import ') or 
                                    linha.strip().startswith('from ') or
                                    linha.strip().startswith('#') or
                                    linha.strip().startswith('"""') or
                                    linha.strip().startswith("'''")):
                imports_fim = i
                break
        
        # Divide em imports e código principal
        imports = linhas[:imports_fim]
        codigo_principal = linhas[imports_fim:]
        
        # Cria nova estrutura
        novas_linhas = imports + [
            "",
            "def main():",
            "    \"\"\"Função principal da automação\"\"\""
        ]
        
        # Adiciona código principal com indentação
        for linha in codigo_principal:
            if linha.strip():
                novas_linhas.append("    " + linha)
            else:
                novas_linhas.append("")
        
        # Adiciona chamada da função
        novas_linhas.extend([
            "",
            "if __name__ == '__main__':",
            "    main()"
        ])
        
        conteudo_novo = '\n'.join(novas_linhas)
    
    # Cria backup do arquivo original
    backup_nome = nome_arquivo + '.backup'
    with open(backup_nome, 'w', encoding='utf-8') as f:
        f.write(conteudo)
    print(f"💾 Backup criado: {backup_nome}")
    
    # Salva arquivo modificado
    with open(nome_arquivo, 'w', encoding='utf-8') as f:
        f.write(conteudo_novo)
    
    print(f"✅ {nome_arquivo} modificado com função main()")
    return True

def verificar_estrutura_arquivos():
    """Verifica a estrutura atual dos arquivos"""
    arquivos = ['add_automation.py', 'remove_automation.py', 'billing_automation.py']
    
    print("🔍 VERIFICANDO ESTRUTURA DOS ARQUIVOS:")
    print("-" * 50)
    
    for arquivo in arquivos:
        if Path(arquivo).exists():
            with open(arquivo, 'r', encoding='utf-8') as f:
                conteudo = f.read()
            
            tem_main = 'def main():' in conteudo
            tem_if_name = 'if __name__ == "__main__":' in conteudo
            
            print(f"\n📄 {arquivo}:")
            print(f"   ✅ Existe" if Path(arquivo).exists() else "   ❌ Não existe")
            print(f"   ✅ Tem função main()" if tem_main else "   ❌ Sem função main()")
            print(f"   ✅ Tem if __name__" if tem_if_name else "   ❌ Sem if __name__")
            
            if not tem_main:
                print(f"   ⚠️  Precisa ser modificado")
        else:
            print(f"\n📄 {arquivo}:")
            print(f"   ❌ Arquivo não encontrado!")

def main():
    print("=" * 60)
    print("        PREPARAR SCRIPTS PARA BUILD - SCOPE AUTOMATIONS")
    print("=" * 60)
    print()
    
    print("Este script irá:")
    print("- Verificar a estrutura dos seus scripts")
    print("- Adicionar função main() se necessário")
    print("- Criar backups dos arquivos originais")
    print("- Tornar os scripts compatíveis com o executável")
    print()
    
    # Verificar estrutura atual
    verificar_estrutura_arquivos()
    
    print("\n" + "=" * 60)
    resposta = input("Deseja modificar os arquivos para compatibilidade? (s/n): ").lower()
    
    if resposta != 's':
        print("Operação cancelada.")
        return
    
    print("\n🔧 MODIFICANDO ARQUIVOS:")
    print("-" * 30)
    
    # Modifica cada arquivo
    arquivos = ['add_automation.py', 'remove_automation.py', 'billing_automation.py']
    sucesso = 0
    
    for arquivo in arquivos:
        if adicionar_funcao_main(arquivo):
            sucesso += 1
    
    print(f"\n✅ {sucesso}/{len(arquivos)} arquivos modificados com sucesso!")
    
    if sucesso == len(arquivos):
        print("\n🎉 PRONTO PARA BUILD!")
        print("Agora você pode executar:")
        print("- python build_executable.py")
        print("- ou python rebuild_quick.py")
        print("\n💡 DICA: Se algo der errado, você pode restaurar os backups (.backup)")
    else:
        print("\n⚠️  Alguns arquivos não foram modificados.")
        print("Verifique os erros acima antes de fazer o build.")

if __name__ == "__main__":
    main()