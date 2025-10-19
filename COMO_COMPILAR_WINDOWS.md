# Como Compilar Executável para Windows

## 📦 O que foi criado

Foi gerado um pacote completo pronto para distribuição:

-   **Arquivo:** `GeradorCertificados_Pacote.zip`
-   **Tamanho:** ~11 MB
-   **Conteúdo:** Executável + Fontes + Imagem Base + Excel de exemplo + Documentação

## 🖥️ Versões Disponíveis

### ✅ **macOS (ARM64)** - PRONTO!

-   Já foi compilado: `dist/GeradorCertificados`
-   Funciona em: macOS 10.13+ (ARM64)
-   Tamanho: 8.2 MB

### 🪟 Para compilar versão **Windows**

Você precisa compilar EM UM COMPUTADOR WINDOWS. Siga os passos:

#### Opção 1: Em uma máquina Windows

1. **Instalar Python 3.7+**

    ```cmd
    # Baixar de: https://www.python.org/downloads/
    ```

2. **Instalar dependências**

    ```cmd
    pip install pillow pyinstaller
    ```

3. **Copiar os arquivos**

    - `gerar_certificados_windows.py`
    - `CERTIFICADO BASE.jpg`
    - As duas fontes Myriad Pro

4. **Compilar**

    ```cmd
    pyinstaller --clean --onefile --console --name "GeradorCertificados" gerar_certificados_windows.py
    ```

5. **O executável estará em:** `dist\GeradorCertificados.exe`

#### Opção 2: Usando máquina virtual Windows

1. Instale VirtualBox ou Parallels
2. Crie uma VM com Windows 10/11
3. Siga os passos da Opção 1

#### Opção 3: Usando Wine (macOS/Linux) - NÃO RECOMENDADO

-   Wine pode ter problemas com PyInstaller
-   Melhor usar uma VM real

## 📁 Estrutura do Projeto

```
gerador_certificado/
├── create_pdf.py                      # Script original (Excel)
├── gerar_certificados_windows.py      # Script standalone (Excel)
├── gerar_certificados.spec            # Config PyInstaller
├── build_windows.sh                   # Script de compilação
├── GUIA_INSTALACAO.txt               # Guia para usuário final
├── README_EXECUTAVEL.md              # Documentação técnica
├── COMO_COMPILAR_WINDOWS.md          # Este arquivo
│
├── dist/
│   └── GeradorCertificados           # Executável macOS (8.2 MB)
│
├── pacote_distribuicao/              # Pacote completo
│   ├── GeradorCertificados           # Executável
│   ├── CERTIFICADO BASE.jpg          # Imagem base
│   ├── Myriad Pro Bold Italic.ttf    # Fonte Bold
│   ├── myriad-pro-semibold-italic.otf # Fonte Semibold
│   ├── GUIA_INSTALACAO.txt          # Guia do usuário
│   ├── README_EXECUTAVEL.md         # Documentação
│   └── gerar certificados_gerados/
│       └── Controle de Diplomas.xlsx # Planilha Excel
│
└── GeradorCertificados_Pacote.zip    # Tudo compactado (~11 MB)
```

## 🚀 Distribuição

### Para usuário final (macOS):

1. Extrair `GeradorCertificados_Pacote.zip`
2. Dar permissão de execução: `chmod +x GeradorCertificados`
3. Executar: `./GeradorCertificados`

### Para usuário final (Windows):

1. Extrair o pacote
2. Duplo clique em `GeradorCertificados.exe`
3. (Pode aparecer aviso do Windows Defender - é normal)

## 📝 Diferenças entre versões

| Recurso    | create_pdf.py | gerar_certificados_windows.py |
| ---------- | ------------- | ----------------------------- |
| Entrada    | Excel (.xlsx) | Excel (.xlsx)                 |
| Saída      | PDF           | PDF                           |
| Fontes     | Myriad Pro    | Myriad Pro                    |
| Standalone | ❌            | ✅                            |
| Executável | ❌            | ✅                            |
| Portável   | ❌            | ✅                            |

## 🔧 Requisitos para compilação

### Python

-   Python 3.7 ou superior
-   PyInstaller 6.0+
-   Pillow (PIL)

### Fontes

-   Myriad Pro Bold Italic (.ttf)
-   Myriad Pro Semibold Italic (.otf ou .ttf)

### Arquivos

-   gerar_certificados_windows.py
-   CERTIFICADO BASE.jpg

## ⚠️ Observações Importantes

1. **Antivírus**: Executáveis PyInstaller podem ser flagados como falso positivo
2. **Tamanho**: O executável tem ~8 MB porque inclui o Python embutido
3. **Cross-compile**: Não é possível compilar .exe no macOS diretamente
4. **Fontes**: Devem estar na mesma pasta do executável
5. **Excel**: Deve estar em `gerar certificados_gerados/Controle de Diplomas.xlsx`

## 🐛 Solução de Problemas

### "Comando não encontrado: pyinstaller"

```bash
pip install --upgrade pyinstaller
```

### "ModuleNotFoundError: No module named 'PIL'"

```bash
pip install pillow
```

### Windows Defender bloqueia o executável

-   É um falso positivo comum
-   Clique em "Mais informações" → "Executar assim mesmo"
-   Ou adicione exceção no antivírus

### Executável não abre

-   Verifique se todas as DLLs estão presentes
-   Execute via CMD para ver mensagens de erro
-   Recompile com `--debug all` para mais informações

## 📧 Suporte

Para gerar o executável Windows, você precisará:

1. Acesso a um computador Windows, OU
2. Uma máquina virtual Windows, OU
3. Contratar um serviço de compilação online

**Executável macOS já está pronto em:** `dist/GeradorCertificados`
