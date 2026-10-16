# 🚗 Automação de Frota - Detran/Sefaz CE (Versão 4.0)

Sistema corporativo de web scraping e automação veicular com Interface Gráfica nativa (GUI Comercial — *Fleet Command Pro*) e CLI para consulta em lote de débitos de **Licenciamento (Detran-CE)**, **IPVA (Sefaz-CE)** e **Multas (Detran-CE)**, além de **emissão automática de boletos/DAEs** e **Business Intelligence (BI) integrado**.

---

## 🌟 Funcionalidades da Versão 4.0

- **Arquitetura Modular em Camadas (`core/`):** Estrutura limpa e escalável com módulos desacoplados para raspagem, processamento de dados, emissão de boletos, exportação de relatórios, persistência de preferências e logging estruturado.
- **Painel Analítico de BI Integrado (Top 10):** Aba dedicada na interface gráfica com **Gráfico de Rosca (Donut Chart)** de composição tributária, **Saúde e Regularidade da Frota** e **Ranking Top 10 Veículos de Maior Impacto Financeiro**, com layout anti-colisão em tempo real.
- **Duplo Relatório PDF Executivo (Relatório Geral + Relatório de BI):** Além da planilha Excel (`.xlsx`) e do relatório consolidado (`.pdf`), gera automaticamente um **2º PDF Executivo de BI (`*_BI.pdf`)** com gráficos vetoriais, KPIs, Top 10 veículos mais críticos e consolidação por Proprietário/Centro de Custo.
- **Raio-X Executivo do Veículo:** Duplo clique em qualquer veículo na tabela abre o detalhamento individual completo com botão para copiar resumo formatado para WhatsApp ou E-mail.
- **Download e Emissão Automática de Boletos (PDF):** Módulo dedicado para emissão e download direto dos boletos/DAEs de Licenciamento e IPVA, com organização automática por pastas e nomenclatura padronizada por veículo e data de vencimento.
- **Suporte a Múltiplos Workers (Concorrência):** Processamento paralelo configurável com múltiplos navegadores em simultâneo para acelerar a consulta de grandes frotas.
- **Mecanismo de Auto-Heal (Recuperação Automática):** Detecção proativa de congelamento ou perda de conexão do Selenium WebDriver, realizando a reinicialização automática da instância do navegador sem interromper o lote.
- **Gerenciador de Preferências e Histórico (`config.json`):** Persistência do estado de configurações (tema, pasta padrão de destino, quantidade de workers, débitos selecionados) e menu rápido de arquivos recentes.
- **Prevenção de Falsos Negativos:** Diferenciação estrita entre veículos regulares (`R$ 0,00`) e falhas ou timeouts de comunicação (`Falha Detran` / `Falha Sefaz` / `Não Localizado`), evitando diagnósticos incorretos de frotas regulares.
- **Proteção contra Arquivos Abertos (PermissionError):** Se a planilha Excel ou relatório PDF estiver aberto no momento do salvamento, o sistema grava automaticamente com sufixo de data/hora (`_YYYYMMDD_HHMMSS`), impedindo a perda de dados.
- **Garantia de 11 Dígitos no RENAVAM (`zfill`):** Sanitização automática de zeros à esquerda para planilhas que salvaram o campo como número.
- **Nomenclatura Dinâmica por Frota:** Analisa automaticamente o nome da planilha consultada (ex: `frota_final_8.xlsx`) e gera os arquivos de saída correspondentes (`resultados_final_8.xlsx`, `resultados_final_8.pdf` e `resultados_final_8_BI.pdf`).
- **Salvamento Incremental (Checkpoint Dinâmico):** Gravação periódica de pontos de recuperação específicos por frota (`~checkpoint_resultados_<frota>.xlsx`) a cada 5 veículos e salvamento garantido em caso de interrupção ou cancelamento.

---

## 🏗️ Estrutura do Projeto

```text
automação-de-frota/
├── automacao_frota_gui.py    # Interface Gráfica Principal v4.0 (Fleet Command Pro)
├── automacao_frota_gui.pyw   # Lançador sem console no Windows
├── automacao_frota_ce.py     # Execução via Linha de Comando (CLI)
├── core/                     # Pacote central de lógica de negócio
│   ├── __init__.py           # Exportação da API do core
│   ├── boleto_downloader.py  # Download e emissão de boletos/DAEs
│   ├── config_manager.py     # Gerenciamento de configurações e histórico
│   ├── data_processor.py     # Validação, limpeza e tratamento financeiro
│   ├── logger_config.py      # Configuração de logs rotativos
│   ├── relatorio_exporter.py # Exportação de relatórios (Excel, PDF Geral e PDF BI)
│   └── scraper_engine.py     # Motores de automação Selenium e auto-heal
├── executar_gui.bat          # Atalho para iniciar a GUI no Windows
├── executar_cli.bat          # Atalho para iniciar o CLI no Windows
├── config.example.json       # Exemplo de configurações padrão
├── requirements.txt          # Dependências do projeto
├── COMO_EXECUTAR.md          # Guia passo a passo de execução (com e sem Python)
└── README.md                 # Documentação geral do sistema
```

---

## 🛠️ Tecnologias Utilizadas

- **Python 3.10+** (Compatível com 3.12+)
- **Tkinter & ttk + DWM API** (Interface gráfica nativa com Modo Claro / Noturno e gráficos em Canvas)
- **Selenium WebDriver** & **webdriver-manager** (Automação de navegador Google Chrome)
- **Pandas** & **Openpyxl** (Manipulação e formatação executiva de planilhas Excel)
- **ReportLab** (Geração de relatórios PDF executivos e gráficos vetoriais de BI em Paisagem)
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

---

## 📜 Licença

Distribuído sob a licença MIT.
