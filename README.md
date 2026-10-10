# 🚗 Automação de Frota - Detran/Sefaz CE

Sistema de web scraping e automação veicular com Interface Gráfica nativa (GUI) para consulta em lote de débitos de **Licenciamento (Detran-CE)** e **IPVA (Sefaz-CE)**.

---

## 🌟 Funcionalidades

- **Interface Gráfica Amigável (`Tkinter`):** Desenvolvida para usuários finais, com seleção de arquivos, caixas de status, barra de progresso e logs em tempo real.
- **Execução em Segundo Plano (`Threading`):** A interface não congela durante as requisições web.
- **Portais Atendidos:**
  - **Detran-CE:** Consulta automatizada de Licenciamento via Single Page Application (SPA).
  - **Sefaz-CE:** Consulta automatizada de débitos de IPVA em aberto via Angular SPA.
- **Espera Explícita Estrita:** Zero `time.sleep` no carregamento de páginas; utiliza exclusivamente `WebDriverWait` (15s) e `expected_conditions`.
- **Tratamento Resiliente:** Captura de `TimeoutException` que assume automaticamente `R$ 0,00` para veículos sem pendências financeiras.
- **Integridade de Dados:** Leitura de planilhas garantindo que o campo `RENAVAM` não perca os zeros à esquerda.
- **Exportação Formatada:** Geração automática da planilha `resultados_frota_ce.xlsx` com cabeçalhos corporativos e bordas estilizadas via `openpyxl`.

---

## 🛠️ Tecnologias Utilizadas

- **Python 3.12+**
- **Tkinter & ttk** (Interface gráfica nativa e temas)
- **Selenium WebDriver** & **webdriver-manager** (Automação de navegador Chrome)
- **Pandas** & **Openpyxl** (Manipulação e formatação executiva de planilhas Excel)
- **Threading** (Processamento concorrente não bloqueante)
- **PyInstaller** (Compilação para arquivo executável `.exe`)

---

## 🚀 Como Executar em Outro Dispositivo

Consulte o nosso guia passo a passo completo:
👉 **[Guia Completo de Execução (COM e SEM Python)](COMO_EXECUTAR.md)**

### Resumo Rápido (Dispositivo com Python):
```bash
# 1. Clone o repositório
git clone https://github.com/JohnathaDev/automacao-de-frota.git
cd automacao-de-frota

# 2. Instale as dependências
pip install -r requirements.txt

# 3. Inicie o aplicativo com interface gráfica
python automacao_frota_gui.py
```

### Resumo Rápido (Dispositivo sem Python - Gerar `.exe`):
```bash
# Compile em um computador com Python:
pip install pyinstaller
pyinstaller --noconsole --onefile automacao_frota_gui.py

# O executável autônomo estará pronto em dist/automacao_frota_gui.exe
# Basta copiá-lo para qualquer máquina Windows e dar duplo clique!
```

---

## 📄 Estrutura da Planilha de Entrada

O arquivo `.xlsx` selecionado pelo usuário deve conter as seguintes colunas:

| Placa | RENAVAM |
| :---: | :---: |
| TII5F89 | 01458899249 |
| THT5F69 | 01458894204 |
| THV0A49 | 01434835143 |

---

## 📜 Licença

Distribuído sob a licença MIT. Veja `LICENSE` para mais informações.
