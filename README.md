# 🚗 Automação de Frota - Detran/Sefaz CE (Versão 2.1)

Sistema corporativo de web scraping e automação veicular com Interface Gráfica nativa (GUI) para consulta em lote de débitos de **Licenciamento (Detran-CE)**, **IPVA (Sefaz-CE)** e **Multas (Detran-CE)**.

---

## 🌟 Funcionalidades e Melhorias da Versão 2.1

- **Prevenção de Falsos Negativos:** Diferenciação estrita entre veículos confirmados regulares (`R$ 0,00`) e falhas ou timeouts de comunicação (`Falha Detran` / `Falha Sefaz` / `Falha na Consulta`), evitando diagnósticos incorretos de frotas regulares.
- **Proteção contra Arquivos Abertos (PermissionError):** Se a planilha Excel ou relatório PDF estiver aberto no Excel ou leitor no momento do salvamento, o sistema grava automaticamente com sufixo de data/hora (`_YYYYMMDD_HHMMSS`), impedindo a perda de dados.
- **Garantia de 11 Dígitos no RENAVAM (`zfill`):** Sanitização automática de zeros à esquerda para planilhas que salvaram o campo como número.
- **Validação Prévia e Limpeza de Dados:** Limpeza de linhas em branco e validação de placas (Mercosul e Tradicional) via expressão regular antes do envio aos portais governamentais.
- **Salvamento Incremental (Checkpoint):** Gravação periódica de pontos de recuperação (`~checkpoint_frota_ce.xlsx`) a cada 5 veículos e salvamento garantido em caso de interrupção ou cancelamento prematuro.
- **Pasta de Destino Personalizada:** Controle visual na interface gráfica para escolher a pasta onde os arquivos de relatório serão gerados.
- **Reconsulta Inteligente (Falhas e Selecionados):** Botão rápido e menu de contexto (clique com botão direito) para reconsultar apenas os veículos que apresentaram falhas temporárias de conexão.
- **Sessão Unificada no Detran-CE:** Consulta combinada de Licenciamento e Multas na mesma navegação por veículo.
- **Exportação Multiformato:** Planilha Excel executiva e Relatório PDF em layout Paisagem (A4 Landscape) com cards de indicadores e paginação dinâmica.
- **Modo Noturno (Dark Mode) Integrado:** Alternância visual em tempo real entre temas claro e escuro.

---

## 🛠️ Tecnologias Utilizadas

- **Python 3.10+** (Compatível com 3.12 e 3.15)
- **Tkinter & ttk** (Interface gráfica nativa com tema moderno)
- **Selenium WebDriver** & **webdriver-manager** (Automação de navegador Google Chrome)
- **Pandas** & **Openpyxl** (Manipulação e formatação executiva de planilhas Excel)
- **ReportLab** (Geração de relatórios PDF executivos em Paisagem)
- **Threading** (Processamento concorrente não bloqueante)
- **PyInstaller** (Compilação para executável autônomo `.exe`)

---

## 🚀 Como Executar

Consulte o guia completo:
👉 **[Guia Completo de Execução (COM e SEM Python)](COMO_EXECUTAR.md)**

### Execução Rápida:
- **Interface Gráfica:** Dê um duplo clique em `executar_gui.bat` ou `automacao_frota_gui.pyw`.
- **Linha de Comando:** Dê um duplo clique em `executar_cli.bat` ou execute:
  ```bash
  python automacao_frota_ce.py
  ```

---

## 📄 Estrutura da Planilha de Entrada

A planilha `.xlsx` selecionada pode conter qualquer número de colunas adicionais. Ela precisa apenas possuir colunas que façam referência à placa e ao RENAVAM:

| Proprietário | Placa | RENAVAM | Chassi |
| :---: | :---: | :---: | :---: |
| Empresa Frota Ltda | TII5F89 | 01458899249 | 9BWZZZ377VT... |
| Empresa Frota Ltda | THU5D09 | 01458895766 | 9BWZZZ377AB... |

---

## 📜 Licença

Distribuído sob a licença MIT.
