"""
Script para gerar executável do Scope Automations
Instale primeiro: pip install pyinstaller
Depois execute: python build_executable.py
"""

import os
import sys
import subprocess
from pathlib import Path

def verificar_pyinstaller():
    """Verifica se PyInstaller está instalado"""
    try:
        import PyInstaller
        print("✅ PyInstaller já está instalado")
        return True
    except ImportError:
        print("❌ PyInstaller não encontrado")
        return False

def instalar_pyinstaller():
    """Instala PyInstaller"""
    print("📦 Instalando PyInstaller...")
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])
        print("✅ PyInstaller instalado com sucesso!")
        return True
    except subprocess.CalledProcessError:
        print("❌ Erro ao instalar PyInstaller")
        return False

def criar_spec_file():
    """Cria arquivo .spec personalizado para o build"""
    spec_content = '''# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['menu_principal.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('AdicionarGrupo.xlsx', '.'),
        ('RemoverGrupo.xlsx', '.'),
        ('ID_billing.xlsx', '.'),
        ('QTGO_ID.xlsx', '.'),
        ('qtgo_automation.py', '.'),
    ],
    hiddenimports=[
        'selenium',
        'selenium.webdriver',
        'selenium.webdriver.chrome',
        'selenium.webdriver.common',
        'selenium.webdriver.support',
        'selenium.webdriver.support.ui',
        'selenium.webdriver.support.wait',
        'selenium.common.exceptions',
        'openpyxl',
        'pandas',
        'xlrd',
        'xlsxwriter',
        'add_automation',
        'remove_automation',
        'billing_automation',
        'qtgo_automation'
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='ScopeAutomations',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None
)
'''
    
    with open('ScopeAutomations.spec', 'w', encoding='utf-8') as f:
        f.write(spec_content)
    
    print("✅ Arquivo .spec criado!")

def verificar_arquivos():
    """Verifica se todos os arquivos necessários existem"""
    arquivos_necessarios = [
        'menu_principal.py',
        'add_automation.py',
        'remove_automation.py',
        'billing_automation.py',
        'qtgo_automation.py',  # Novo arquivo adicionado
        'AdicionarGrupo.xlsx',
        'RemoverGrupo.xlsx', 
        'ID_billing.xlsx',
        'QTGO_ID.xlsx'  # Nova planilha
    ]
    
    arquivos_faltando = []
    for arquivo in arquivos_necessarios:
        if not Path(arquivo).exists():
            arquivos_faltando.append(arquivo)
    
    return arquivos_faltando

def gerar_executavel():
    """Gera o arquivo executável"""
    print("🔨 Gerando executável...")
    print("Isso pode demorar alguns minutos...")
    
    try:
        # Usa o arquivo .spec personalizado
        cmd = [sys.executable, "-m", "PyInstaller", "ScopeAutomations.spec", "--clean"]
        
        resultado = subprocess.run(cmd, capture_output=True, text=True)
        
        if resultado.returncode == 0:
            print("✅ Executável gerado com sucesso!")
            print("📁 Localizado em: dist/ScopeAutomations.exe")
            return True
        else:
            print("❌ Erro ao gerar executável:")
            print(resultado.stderr)
            return False
            
    except Exception as e:
        print(f"❌ Erro inesperado: {str(e)}")
        return False

def criar_bat_file():
    """Cria arquivo .bat para facilitar execução"""
    bat_content = '''@echo off
title Scope Automations
cd /d "%~dp0"
ScopeAutomations.exe
pause
'''
    
    with open('dist/Executar_ScopeAutomations.bat', 'w') as f:
        f.write(bat_content)
    
    print("✅ Arquivo .bat criado para facilitar execução!")

def main():
    print("=" * 60)
    print("        GERADOR DE EXECUTÁVEL - SCOPE AUTOMATIONS")
    print("=" * 60)
    print()
    
    # Verificar arquivos
    print("🔍 Verificando arquivos...")
    arquivos_faltando = verificar_arquivos()
    
    if arquivos_faltando:
        print("❌ Arquivos faltando:")
        for arquivo in arquivos_faltando:
            print(f"   - {arquivo}")
        print("\nCertifique-se de que todos os arquivos estão na pasta!")
        return
    
    print("✅ Todos os arquivos encontrados!")
    print()
    
    # Verificar/instalar PyInstaller
    if not verificar_pyinstaller():
        resposta = input("Deseja instalar PyInstaller? (s/n): ").lower()
        if resposta == 's':
            if not instalar_pyinstaller():
                return
        else:
            print("PyInstaller é necessário para gerar o executável.")
            return
    
    print()
    
    # Criar arquivo .spec
    print("📝 Criando configuração de build...")
    criar_spec_file()
    print()
    
    # Confirmar build
    print("🚀 Pronto para gerar o executável!")
    print("Isso irá:")
    print("- Incluir todos os scripts Python (4 automações)")
    print("- Incluir todas as planilhas Excel")
    print("- Criar um executável único")
    print("- Gerar arquivo .bat para execução")
    print()
    
    resposta = input("Continuar? (s/n): ").lower()
    if resposta != 's':
        print("Build cancelado.")
        return
    
    print()
    
    # Gerar executável
    if gerar_executavel():
        criar_bat_file()
        print()
        print("=" * 60)
        print("✅ BUILD CONCLUÍDO COM SUCESSO!")
        print("=" * 60)
        print()
        print("📁 Arquivos gerados:")
        print("   - dist/ScopeAutomations.exe")
        print("   - dist/Executar_ScopeAutomations.bat")
        print()
        print("📋 Para distribuir:")
        print("1. Copie TODA a pasta 'dist' para o computador destino")
        print("2. Execute 'Executar_ScopeAutomations.bat'")
        print("3. Ou execute diretamente 'ScopeAutomations.exe'")
        print()
        print("⚠️  IMPORTANTE:")
        print("- Chrome/Edge deve estar instalado no computador destino")
        print("- ChromeDriver será baixado automaticamente pelo Selenium")
        print()
        print("📋 Automações incluídas:")
        print("1. Adicionar Carros ao Grupo de Veículos")
        print("2. Remover Carros do Grupo de Veículos")
        print("3. Remover Unidades do Billing")
        print("4. Remover Carros do QTGO")
        
    else:
        print("❌ Falha no build. Verifique os erros acima.")

if __name__ == "__main__":
    main()