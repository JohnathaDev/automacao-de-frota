# 🚗 Automação de Frota - Detran/Sefaz CE (Versão 3.0)

Sistema corporativo de web scraping e automação veicular com Interface Gráfica nativa (GUI) e CLI para consulta em lote de débitos de **Licenciamento (Detran-CE)**, **IPVA (Sefaz-CE)** e **Multas (Detran-CE)**, além de **emissão e download automático de boletos/DAEs**.

---

## 🌟 Novidades e Funcionalidades da Versão 3.0 Corporativa

- **Arquitetura Modular em Camadas (`core/`):** Estrutura limpa e escalável com módulos desacoplados para raspagem, processamento de dados, emissão de boletos, exportação de relatórios, persistência de preferências e logging estruturado.
- **Download e Emissão Automática de Boletos (PDF):** Módulo dedicado para emissão e download direto dos boletos/DAEs de Licenciamento e IPVA, com organização automática por pastas e nomenclatura padronizada por veículo e data de vencimento.
- **Suporte a Múltiplos Workers (Concorrência):** Processamento paralelo configurável com múltiplos navegadores em simultâneo para acelerar a consulta de grandes frotas.
- **Mecanismo de Auto-Heal (Recuperação Automática):** Detecção proativa de congelamento ou perda de conexão do Selenium WebDriver, realizando a reinicialização automática da instância do navegador sem interromper o lote.
- **Gerenciador de Preferências e Histórico (`config.json`):** Persistência do estado de configurações (tema, pasta padrão de destino, quantidade de workers, débitos selecionados) e histórico de arquivos recentes consultados.
- **Logging Estruturado e Rotativo (`logs/`):** Rastreabilidade total das operações com logs diários formatados e rotação automática.
- **Suíte de Testes Automatizados (`tests/`):** Cobertura com testes unitários para a camada core (`test_core.py`) e para a interface gráfica (`test_gui.py`).
- **Prevenção de Falsos Negativos:** Diferenciação estrita entre veículos regulares (`R$ 0,00`) e falhas ou timeouts de comunicação (`Falha Detran` / `Falha Sefaz` / `Falha na Consulta`), evitando diagnósticos incorretos de frotas regulares.
- **Proteção contra Arquivos Abertos (PermissionError):** Se a planilha Excel ou relatório PDF estiver aberto no Excel ou leitor no momento do salvamento, o sistema grava automaticamente com sufixo de data/hora (`_YYYYMMDD_HHMMSS`), impedindo a perda de dados.
- **Garantia de 11 Dígitos no RENAVAM (`zfill`):** Sanitização automática de zeros à esquerda para planilhas que salvaram o campo como número.
- **Validação Prévia e Limpeza de Dados:** Limpeza de linhas em branco e validação de placas (Mercosul e Tradicional) via expressão regular antes do envio aos portais governamentais.
- **Salvamento Incremental (Checkpoint):** Gravação periódica de pontos de recuperação (`~checkpoint_frota_ce.xlsx`) a cada 5 veículos e salvamento garantido em caso de interrupção ou cancelamento prematuro.
- **Pasta de Destino Personalizada:** Controle visual na interface gráfica para escolher a pasta onde os arquivos de relatório e boletos serão gerados.
- **Reconsulta Inteligente (Falhas e Selecionados):** Botão rápido e menu de contexto (clique com botão direito) para reconsultar apenas os veículos que apresentaram falhas temporárias de conexão.
- **Sessão Unificada no Detran-CE:** Consulta combinada de Licenciamento e Multas na mesma navegação por veículo.
- **Exportação Multiformato:** Planilha Excel executiva e Relatório PDF em layout Paisagem (A4 Landscape) com cards de indicadores e paginação dinâmica.
- **Modo Noturno (Dark Mode) Integrado & Notificação Sonora:** Alternância visual em tempo real entre temas claro e escuro, com alerta sonoro configurável ao término do processamento.

---

## 🏗️ Estrutura do Projeto

```text
automação-de-frota/
├── automacao_frota_gui.py    # Interface Gráfica Principal (Tkinter/ttk)
├── automacao_frota_gui.pyw   # Lançador sem console no Windows
├── automacao_frota_ce.py     # Execução via Linha de Comando (CLI)
├── core/                     # Pacote central de lógica de negócio
│   ├── __init__.py           # Exportação da API do core
│   ├── boleto_downloader.py  # Download e emissão de boletos/DAEs
│   ├── config_manager.py     # Gerenciamento de configurações e histórico
│   ├── data_processor.py     # Validação, limpeza e tratamento financeiro
│   ├── logger_config.py      # Configuração de logs rotativos
│   ├── relatorio_exporter.py # Exportação de relatórios (Excel e PDF Paisagem)
│   └── scraper_engine.py     # Motores de automação Selenium e auto-heal
├── tests/                    # Suíte de testes automatizados
│   ├── test_core.py          # Testes unitários do core
│   └── test_gui.py           # Testes de integração da interface gráfica
├── executar_gui.bat          # Atalho para iniciar a GUI no Windows
├── executar_cli.bat          # Atalho para iniciar o CLI no Windows
├── config.example.json       # Exemplo de configurações padrão
├── requirements.txt          # Dependências do projeto
├── COMO_EXECUTAR.md          # Guia passo a passo de execução (com e sem Python)
└── README.md                 # Documentação geral do sistema
```

---

## 🛠️ Tecnologias Utilizadas

- **Python 3.10+** (Compatível com 3.12 e 3.15)
- **Tkinter & ttk** (Interface gráfica nativa com tema moderno)
- **Selenium WebDriver** & **webdriver-manager** (Automação de navegador Google Chrome)
- **Pandas** & **Openpyxl** (Manipulação e formatação executiva de planilhas Excel)
- **ReportLab** (Geração de relatórios PDF executivos em Paisagem)
- **Threading & Concorrência** (Processamento não bloqueante com múltiplos workers)
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

### Executar Testes Unitários:
```bash
python -m unittest discover tests
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
