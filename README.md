# BOT-SERVICO
> Automação do lançamento de notas de serviço no TOTVS usando Python e Selenium, com interface gráfica e análise de XML.

## 💡 Objetivo

Este projeto automatiza o processo de análise e registro de notas fiscais de serviço emitidas. Ele atua diretamente na interface do TOTVS, realizando o preenchimento de campos com base nos dados extraídos dos arquivos XMLs das notas.

## 🚀 Funcionalidades

- Extração de informações de XMLs.
- Análise e categorização das notas.
- Formatação de Ordens de Serviço e número das notas.
- Seleção do ambiente de produção correto.
- Busca automática de arquivos em pastas padronizadas.
- Interação com sistema TOTVS para preenchimento automatizado.
- Banco de dados baseado em SQLite para controle e rastreabilidade.
- Interface gráfica simples para controle da automação.
- Exportável como `.exe` para execução sem dependências externas.

## 📁 Estrutura de Pastas
```bash
BOT-SERVICO/
│
├── app/ # Interface gráfica (controle da automação)
├── assets/ # Arquivos XML organizados por mês/ano
├── build/ # Pasta gerada pelo PyInstaller
├── dist/ # Executável e bancos de dados locais
├── env/ # Ambiente virtual (excluído pelo .gitignore)
├── icons/ # Ícones usados na aplicação
├── path/ # JSONs e configurações de caminhos
├── processos/ # Scripts principais de automação (análise, extração, interação web)
├── utils/ # Funções auxiliares (ex: serviços.py)
├── web/ # Módulos relacionados à automação web
├── .gitignore # Arquivos/pastas ignoradas pelo Git
├── KADRIX S.spec # Configuração do PyInstaller
├── README.md # Este arquivo
├── requirements.txt # Dependências do projeto
```

## ⚙️ Tecnologias Utilizadas

- Python 3.x
- OpenPyXL
- PyAutoGUI / Pyperclip
- Selenium
- Tkinter (CustomTkinter)
- SQLite

## 🧪 Como Executar

1. Crie e ative um ambiente virtual:
   ```bash
   python -m venv env
   source env/bin/activate  # ou .\env\Scripts\activate no Windows

## Instale as dependências:
> pip install -r requirements.txt

## Execute o script principal:
> python -m app.app

## Para criar o .exe:
> pyinstaller --onefile -w --icon=icons/"ICONE CRIADO".ico --name="NOME DA SUA ESCOLHA" app/app.py

## ⚠️ Avisos
- Este projeto é uma versão genérica, sem qualquer vínculo com dados sensíveis ou proprietários. Adaptado exclusivamente para fins educacionais e de portfólio.
