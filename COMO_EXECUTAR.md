# 📖 Guia de Execução em Outros Dispositivos

Este guia explica detalhadamente como executar a aplicação de **Automação de Consulta de Débitos Veiculares (Detran-CE & Sefaz-CE)** em qualquer computador, cobrindo dois cenários:
1. **Dispositivo COM Python instalado**
2. **Dispositivo SEM Python instalado (via arquivo executável `.exe`)**

---

## 📋 Pré-requisito Comum a Ambos os Cenários
Independentemente de ter Python ou não, o computador que executará a automação precisa de:
- **Google Chrome** instalado e atualizado (o robô automatiza o navegador Chrome).
- **Acesso à internet** (para consultar os portais governamentais do Detran-CE e Sefaz-CE).
- **Uma planilha Excel (.xlsx)** contendo obrigatoriamente as colunas:
  - `Placa` (ex: `ABC1D23`)
  - `RENAVAM` (ex: `01234567890`)

---

## 🟢 Cenário 1: O dispositivo possui Python instalado

Se o computador de destino já possui o Python 3.8 ou superior:

### Passo 1: Obter os arquivos do projeto
Baixe ou clone este repositório no computador:
```bash
git clone https://github.com/JohnathaDev/automacao-de-frota.git
cd automacao-de-frota
```
*(Ou simplesmente copie o arquivo `automacao_frota_gui.py` e o `requirements.txt` para uma pasta).*

### Passo 2: Instalar as bibliotecas necessárias
Abra o Prompt de Comando (CMD) ou PowerShell dentro da pasta do projeto e execute:
```bash
pip install -r requirements.txt
```
> **Nota:** As bibliotecas `tkinter` e `threading` já vêm instaladas por padrão no Python no Windows.

### Passo 3: Executar o programa
Para abrir a **Interface Gráfica (GUI) sem janela de console ou terminal ao fundo**:
- **Opção 1 (Mais fácil):** Dê um duplo clique no atalho `executar_gui.bat`.
- **Opção 2:** Dê um duplo clique no arquivo `automacao_frota_gui.pyw`.
- **Opção 3 (Terminal):**
```bash
pythonw automacao_frota_gui.py
```
*(Se executar com `python automacao_frota_gui.py`, o script oculta a janela preta de console automaticamente).*

Caso deseje rodar a versão em **Linha de Comando (CLI)** sem interface gráfica:
```bash
python automacao_frota_ce.py
```

### Passo 4: Utilizar a Interface
1. Clique no botão **📁 Selecionar Planilha (.xlsx)**. A planilha pode ter colunas em qualquer ordem e com campos extras.
2. (Opcional) Clique em **📂 Pasta de Destino...** caso deseje salvar os relatórios em uma pasta diferente da planilha original.
3. Em **Consultar Débitos**, marque quais débitos deseja pesquisar: `IPVA`, `Licenciamento` e/ou `Multas` (ou use o botão `Marcar/Desmarcar Todos`).
4. Em **Exportação**, selecione se deseja salvar como `Excel (.xlsx)`, `PDF Paisagem (.pdf)` ou `Ambos`.
5. Escolha se deseja manter o navegador em segundo plano (invisível).
6. Clique em **▶️ Iniciar Consulta**.
7. Acompanhe o progresso e o dashboard em tempo real. Se preferir trabalhar com tema escuro, clique em **🌙 Modo Noturno** no topo superior direito.
8. Ao concluir, utilize os botões rápidos para **Abrir Planilha (Excel)**, **Abrir Relatório (PDF)** ou **Abrir Pasta de Destino**.
9. Caso algum veículo tenha apresentado lentidão transitória dos portais governamentais, clique em **⚠️ Reconsultar Falhas** ou clique com o **botão direito do mouse** sobre o veículo na tabela para reprocessá-lo individualmente.

---

## 🔵 Cenário 2: O dispositivo NÃO possui Python instalado (Gerando o `.exe`)

Para enviar o programa a um usuário final ou computador corporativo que **não possui Python instalado**, transformamos o código em um executável autônomo (`.exe`).

### Passo 1: Gerar o executável `.exe` (faça isso no seu computador com Python)
1. Certifique-se de que o `pyinstaller` está instalado:
   ```bash
   pip install pyinstaller
   ```
2. No terminal, na pasta do projeto, execute o comando de compilação:
   ```bash
   pyinstaller --noconsole --onefile automacao_frota_gui.py
   ```
   - `--onefile`: Compacta tudo (Python + bibliotecas + código) em um **único arquivo `.exe`**.
   - `--noconsole`: Oculta a janela preta do prompt de comando, exibindo apenas a janela da interface gráfica.

### Passo 2: Localizar o arquivo executável gerado
Após a finalização do comando, será criada uma pasta chamada `dist/`. O executável estará pronto dentro dela:
```text
automacao-de-frota/
├── dist/
│   └── automacao_frota_gui.exe   <-- ESTE É O ARQUIVO FINAL
```

### Passo 3: Distribuir e Executar no outro computador
1. Copie apenas o arquivo `automacao_frota_gui.exe` (por pendrive, rede local, email ou nuvem) para o computador de destino.
2. No computador de destino:
   - **NÃO precisa instalar Python.**
   - **NÃO precisa abrir terminal nem digitar comandos.**
   - Basta dar um **duplo clique** em `automacao_frota_gui.exe`.
3. A interface abrirá normalmente, permitindo selecionar qualquer planilha `.xlsx` e consultar os débitos.

---

## ⚠️ Dicas e Soluções de Problemas Comuns

| Situação | Causa | Solução |
| :--- | :--- | :--- |
| **Aviso do Windows Defender ("Windows protegeu o computador")** | O executável é novo e não possui assinatura digital paga da Microsoft. | Clique em **"Mais informações"** e depois no botão **"Executar assim mesmo"**. |
| **Erro ao abrir: "Chrome não encontrado"** | O Google Chrome não está instalado no computador de destino. | Instale o navegador Google Chrome oficial. |
| **Erro: "Planilha deve conter as colunas Placa e RENAVAM"** | Cabeçalhos diferentes na planilha de entrada. | Verifique se a primeira linha da planilha contém exatamente `Placa` e `RENAVAM` (respeitando maiúsculas/minúsculas). |
| **Zeros à esquerda sumindo do RENAVAM** | Excel tratando o número como numérico e não texto. | O script já faz a leitura forçada como texto (`dtype=str`), preservando zeros à esquerda (ex: `01234567890`). |
