<div align="center">

# ⚖️ Sentinela Jurídico

### Automação Inteligente de Pautas Judiciais & Agenda Outlook

![Python](https://img.shields.io/badge/Python-3.11+-3776AB?style=for-the-badge&logo=python&logoColor=white)
![CustomTkinter](https://img.shields.io/badge/CustomTkinter-5.2+-1F6FEB?style=for-the-badge)
![Outlook](https://img.shields.io/badge/Microsoft_Outlook-COM_Automation-0078D4?style=for-the-badge&logo=microsoftoutlook&logoColor=white)
![Pandas](https://img.shields.io/badge/Pandas-Data_Engine-150458?style=for-the-badge&logo=pandas&logoColor=white)
![License](https://img.shields.io/badge/License-Proprietary-E8721C?style=for-the-badge)

<br>

**Sistema desktop para gestão automatizada de audiências judiciais.**
Lê planilhas Excel, valida dados com regras estritas, agenda compromissos no Outlook,
envia convites e gera relatórios — tudo com interface gráfica moderna.

<br>

[Funcionalidades](#-funcionalidades) •
[Arquitetura](#-arquitetura) •
[Stack Técnica](#-stack-técnica) •
[Screenshots](#-screenshots) •
[Autor](#-autor)

</div>

---

## Problema

Escritórios de advocacia gerenciam dezenas de audiências judiciais por semana. O processo manual envolve:

- Ler planilha Excel com dados das audiências
- Criar cada compromisso individualmente no Outlook
- Enviar convites para advogados e prepostos
- Monitorar alterações e cancelamentos
- Manter registro de tudo que foi feito

**Resultado:** horas de trabalho repetitivo, erros humanos e audiências perdidas.

## 💡 Solução

O **Sentinela Jurídico** automatiza 100% desse fluxo:

```
📊 Planilha Excel → 🔍 Validação Estrita → 📅 Outlook Agenda → 📧 Convites → 📋 Relatório
```

O operador carrega a planilha, clica "Iniciar" e o sistema faz o resto — incluindo monitoramento contínuo de alterações.

---

## Funcionalidades

### Motor de Análise
- **Validação estrita de dados** — datas (DD/MM/AAAA, máx. 60 dias futuro), horários (06h-22h), formatos
- **Detecção de anomalias** — CNJ inválido, horários sobrepostos, links suspeitos, nomes sem cadastro
- **Busca inteligente de e-mails** — normalização Unicode, matching parcial, case-insensitive
- **Classificação por severidade** — erros bloqueantes vs. avisos informativos

### Integração Outlook (COM Automation)
- **Agendamento automático** com Subject, Body, Duration, Reminder (15 min)
- **Evento-sombra** de alerta "Entrar na Sala" (5 min antes)
- **Envio de convites** (.ics) apenas para destinatários externos
- **Limpeza automática** — compromissos passados são removidos da agenda automaticamente
- **Cancelamento inteligente** — detecta remoções na planilha e apaga do Outlook

### Monitoramento em Tempo Real
- **File watcher** por hash MD5 — detecta alterações na planilha
- **Threading** — UI responsiva durante processamento
- **Event-driven** — pause/resume sem perda de estado

### Interface Gráfica (CustomTkinter)
- Dark mode nativo com paleta corporativa customizada
- Dashboard com cards de status: Total, Erros, Avisos, Outlook
- Log de atividades com color-coding por severidade
- Relatório de entregas com histórico persistente

### Relatório de Entregas
- Exportação automática em Excel (.xlsx)
- Todas as colunas relevantes: Processo, Cliente, Advogado, Preposto, Vara, Comarca, UF
- IDs internos de controle separados (não expostos no arquivo)
- Histórico completo: SUCESSO, FALHA, IGNORADO, REMOVIDO

---

## Arquitetura

```
┌─────────────────────────────────────────────────┐
│                    App (CTk)                     │
│  ┌──────────┐  ┌─────────────────────────────┐  │
│  │ Sidebar  │  │     Main Panel              │  │
│  │ ────────── │  │  ┌─────┬─────┬─────┬─────┐ │  │
│  │ Logo     │  │  │ Tot │ Err │Warn │ OL  │ │  │
│  │ Status   │  │  └─────┴─────┴─────┴─────┘ │  │
│  │ Carregar │  │  ┌─────────────────────────┐ │  │
│  │ Iniciar  │  │  │  TabView               │ │  │
│  │ Pausar   │  │  │  ├─ Log de Atividades  │ │  │
│  │ Parar    │  │  │  └─ Relatório/Backup   │ │  │
│  └──────────┘  │  └─────────────────────────┘ │  │
│                └─────────────────────────────┘  │
└──────────────────────┬──────────────────────────┘
                       │
        ┌──────────────┼──────────────┐
        ▼              ▼              ▼
  ┌───────────┐  ┌───────────┐  ┌──────────┐
  │  Motor    │  │ Executor  │  │Historico │
  │ Análise   │  │ Outlook   │  │ Manager  │
  │           │  │           │  │          │
  │ • Carga   │  │ • COM API │  │ • CRUD   │
  │ • Valid.  │  │ • CRUD    │  │ • Excel  │
  │ • Extract │  │ • Convite │  │ • IDs    │
  └─────┬─────┘  └─────┬─────┘  └────┬─────┘
        │              │              │
        ▼              ▼              ▼
   ┌─────────┐   ┌──────────┐   ┌─────────┐
   │  Excel  │   │ Outlook  │   │ .xlsx   │
   │ (input) │   │ Calendar │   │ Report  │
   └─────────┘   └──────────┘   └─────────┘
```

### Padrões Aplicados
- **MVC implícito** — Motor (Model), App (View+Controller), Historico (Persistence)
- **Observer** — File watcher com hash comparison
- **Strategy** — Validação modular por tipo de campo
- **Thread-safe UI** — `self.after()` para atualizações cross-thread

---

## Stack Técnica

| Camada | Tecnologia | Motivo |
|--------|-----------|--------|
| **Linguagem** | Python 3.11+ | Ecossistema maduro, pywin32 para COM |
| **UI** | CustomTkinter | Desktop nativo, dark mode, sem Electron overhead |
| **Dados** | Pandas + OpenPyXL | ETL robusto para Excel |
| **Automação** | pywin32 (COM) | Integração nativa com Microsoft Outlook |
| **Imagens** | Pillow | Manipulação de logos e ícones |

### Decisões Técnicas Relevantes

- **`CTkTextbox` em vez de `CTkScrollableFrame`** para tabelas — contorna bug de compatibilidade com Tcl/Tk 8.6.13 no Python 3.13+
- **`dayfirst=True`** no parsing de datas — essencial para formato brasileiro DD/MM/AAAA
- **`datetime.now(timezone.utc) - timedelta(hours=3)`** em vez de `datetime.utcnow()` — deprecated no Python 3.12+
- **Hash MD5 para file watching** — mais confiável que mtime para detectar alterações reais
- **IDs internos com prefixo `_`** — separação clara entre dados visíveis e de controle no relatório

---

## Screenshots

<div align="center">

| Tela Principal | Validação com Erros |
|:-:|:-:|
| Dashboard com cards de status e log em tempo real | Sistema rejeita datas inválidas e horários fora de faixa |

</div>

> *Screenshots omitidos por confidencialidade do cliente. Disponíveis sob NDA.*

---

## Requisitos da Planilha

O sistema aceita planilhas `.xlsx` com estrutura padrão:

**Aba principal (INFORMAÇÕES):**
| Coluna | Validação |
|--------|-----------|
| Data Audiência | DD/MM/AAAA, máx. 60 dias futuro |
| Hora | HH:MM, entre 06:00 e 22:00 |
| Advogado Resp. Audiência | Nome completo, match com aba de e-mails |
| Preposto | Nome completo, match com aba de e-mails |
| Link ou Redesignação | URL, "PRESENCIAL" ou status |

**Aba obrigatória (NOME + E-MAIL):**
| Nome | E-MAIL |
|------|--------|
| Case-insensitive | Domínio obrigatório: @empresa.com.br |

---

## Execução

```bash
# Instalar dependências
pip install customtkinter Pillow pandas openpyxl pywin32

# Executar
python sentinela_rgf.py
```

> ⚠️ Requer Windows 10+ com Microsoft Outlook instalado e configurado.

---

## 📁 Estrutura do Projeto

```
sentinela-juridico/
├── sentinela_rgf.py         # Aplicação principal (~1400 linhas)
├── assets/
│   ├── app.ico              # Ícone da aplicação
│   ├── logo_transparente.png
│   └── ...
├── excel/                   # Pasta para planilhas (input)
├── RELATORIO_ENTREGAS.xlsx  # Gerado automaticamente (output)
├── requirements.txt
├── .gitignore
├── LICENSE
└── README.md
```

---

## Métricas

- **~1.400 linhas** de código Python
- **5 classes** principais (MotorAnalise, ExecutorOutlook, HistoricoManager, LogPanel, App)
- **15+ validações** de dados por linha da planilha
- **Tempo médio de processamento:** < 2s para 50 audiências
- **Zero dependências externas** além do ecossistema Python

---

## 🔒 Licença

Software proprietário. Desenvolvido sob encomenda para uso exclusivo do cliente.

O código-fonte neste repositório é disponibilizado apenas como **portfólio profissional**. Não é permitida a cópia, redistribuição ou uso comercial sem autorização expressa do autor.

---

## 👤 Autor

**Daniel Ferreira da Silva**

Desenvolvedor de software especializado em automação corporativa, integração com Microsoft Office (COM/Win32) e aplicações desktop com Python.

[![Linkedln](https://img.shields.io/badge/LinkedIn-0077B5?style=for-the-badge&logo=linkedin&logoColor=white)](https://www.linkedin.com/in/daniel-silva-8475071b0?)
[![GitHub](https://img.shields.io/badge/GitHub-181717?style=for-the-badge&logo=github&logoColor=white)](https://github.com/dAnniel-11)

---

<div align="center">
<sub>Desenvolvido com ☕ e Python — São Paulo, Brasil — 2026</sub>
</div>
