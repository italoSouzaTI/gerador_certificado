#!/bin/bash

echo "=========================================="
echo "  COMPILADOR PARA EXECUTÁVEL WINDOWS"
echo "=========================================="
echo ""

# Instalar PyInstaller se necessário
echo "Verificando PyInstaller..."
pip3 install pyinstaller --quiet --user

echo ""
echo "Compilando para Windows..."
echo ""

# Compilar
pyinstaller --clean --onefile --console \
    --name "GeradorCertificados" \
    --add-data "CERTIFICADO BASE.jpg:." \
    gerar_certificados_windows.py

echo ""
echo "=========================================="
echo "  COMPILAÇÃO CONCLUÍDA!"
echo "=========================================="
echo ""
echo "O executável está em: dist/GeradorCertificados"
echo ""
echo "ARQUIVOS NECESSÁRIOS PARA DISTRIBUIÇÃO:"
echo "  1. GeradorCertificados.exe (ou GeradorCertificados no macOS)"
echo "  2. CERTIFICADO BASE.jpg"
echo "  3. Fontes Myriad Pro (.ttf ou .otf)"
echo "  4. Pasta 'gerar certificados_gerados' com arquivo 'Controle de Diplomas.xlsx'"
echo ""
echo "ESTRUTURA DE PASTAS PARA O USUÁRIO:"
echo "  MinhaPasta/"
echo "    ├── GeradorCertificados.exe"
echo "    ├── CERTIFICADO BASE.jpg"
echo "    ├── Myriad Pro Bold Italic.ttf"
echo "    ├── myriad-pro-semibold-italic.otf"
echo "    └── gerar certificados_gerados/"
echo "        └── Controle de Diplomas.xlsx"
echo ""

