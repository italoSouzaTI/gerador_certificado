# Gerador de Certificados - Executável Windows

## 📋 Como usar

### 1️⃣ Estrutura de Arquivos

Organize os arquivos da seguinte forma:

```
MinhaPasta/
├── GeradorCertificados.exe          # Executável
├── CERTIFICADO BASE.jpg              # Imagem do certificado base
├── Myriad Pro Bold Italic.ttf        # Fonte Bold
├── myriad-pro-semibold-italic.otf    # Fonte Semibold
└── gerar certificados_gerados/       # Pasta com dados
    └── Controle de Diplomas.xlsx     # Planilha Excel com os dados
```

### 2️⃣ Formato do Excel

A planilha `Controle de Diplomas.xlsx` deve ter a seguinte estrutura:

**Aba: 2025**

| Coluna A (Nome) | Coluna B (Registro) | Coluna C (Graduação) | Coluna F (Key User) |
| --------------- | ------------------- | -------------------- | ------------------- |
| Nome Completo   | CPF: 123.456.789-00 | PRETA                | 001/2025            |
| Outra Pessoa    | RG: 12.345.678-9    | ROXA                 | 002/2025            |

**Colunas:**

-   **A:** Nome completo da pessoa
-   **B:** Registro (CPF, RG, etc)
-   **C:** Graduação (PRETA, ROXA, AZUL, VERDE, LARANJA, etc)
-   **D:** Data (não usado)
-   **E:** (vazio)
-   **F:** Key User (numeração, ex: 001/2025)

> **Importante:** A primeira linha é o cabeçalho. Os dados começam na linha 2.

### 3️⃣ Executar o Programa

1. Coloque todos os arquivos nas pastas corretas (veja estrutura acima)
2. Execute o arquivo `GeradorCertificados.exe`
3. O programa vai:
    - Verificar se todos os arquivos necessários existem
    - Criar a pasta `certificado_final` (se não existir)
    - Ler os dados do Excel
    - Gerar os certificados em PDF
    - Salvar na pasta `certificado_final`

### 4️⃣ Resultado

Os certificados serão gerados em PDF na pasta:

```
MinhaPasta/certificado_final/
├── 1-Nome Completo da Pessoa.pdf
├── 2-Outra Pessoa.pdf
└── ...
```

## 🛠 Como compilar o executável

### No macOS (para gerar executável Windows):

```bash
# Instalar PyInstaller
pip3 install pyinstaller openpyxl pillow

# Compilar
chmod +x build_windows.sh
./build_windows.sh
```

### No Windows:

```cmd
# Instalar dependências
pip install pyinstaller openpyxl pillow

# Compilar
pyinstaller --clean --onefile --console --name "GeradorCertificados" gerar_certificados_windows.py
```

## ⚙️ Requisitos

### Para compilar:

-   Python 3.7+
-   PyInstaller
-   openpyxl (para ler Excel)
-   Pillow (PIL)

### Para usar o executável:

-   Windows 7 ou superior
-   Nenhuma instalação necessária (executável standalone)

## 📝 Notas

-   O executável é standalone (não precisa instalar Python)
-   As fontes Myriad Pro devem estar na mesma pasta do executável
-   A planilha Excel deve estar em `gerar certificados_gerados/Controle de Diplomas.xlsx`
-   A imagem base deve ser `CERTIFICADO BASE.jpg`
-   Os certificados são gerados em alta qualidade (PDF)
-   Os dados devem estar na aba "2025" da planilha

## ❗ Solução de Problemas

### "Arquivo não encontrado"

-   Verifique se todos os arquivos estão nas pastas corretas
-   Confira os nomes dos arquivos (com espaços e extensões)
-   O Excel deve se chamar exatamente "Controle de Diplomas.xlsx"

### "Fontes não encontradas"

-   Coloque os arquivos .ttf ou .otf na mesma pasta do executável
-   Nomes aceitos:
    -   `Myriad Pro Bold Italic.ttf`
    -   `myriad-pro-semibold-italic.otf`
    -   `Myriad Pro Semibold Italic.ttf`

### "Nenhum dado encontrado no Excel"

-   Verifique se há dados na aba "2025"
-   Certifique-se de que a coluna A (Nome) está preenchida
-   Os dados devem começar na linha 2 (linha 1 é cabeçalho)
-   Não deixe linhas vazias entre os registros

### "Aba não encontrada"

-   O programa procura pela aba "2025" primeiro
-   Se não encontrar, usa a primeira aba disponível
-   Certifique-se de que a aba existe e tem o nome correto

## 📧 Suporte

Para dúvidas ou problemas, verifique:

1. Estrutura de pastas correta
2. Nomes dos arquivos corretos (incluindo extensões)
3. Planilha Excel com aba "2025" e dados na estrutura correta
4. Fontes Myriad Pro na pasta do executável
5. Imagem "CERTIFICADO BASE.jpg" presente

## 🔄 Diferença entre versões

| Característica  | create_pdf.py | gerar_certificados_windows.py |
| --------------- | ------------- | ----------------------------- |
| Formato entrada | Excel (.xlsx) | Excel (.xlsx)                 |
| Precisa Python? | ✅ Sim        | ❌ Não (executável)           |
| Portável?       | ❌ Não        | ✅ Sim                        |
| Saída           | PDF           | PDF                           |
| Fontes          | Myriad Pro    | Myriad Pro                    |
