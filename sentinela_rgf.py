#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Rosenthal | Guarita Advogados
Automacao Inteligente de Pautas Judiciais & Agenda Outlook
Versao: 1.4.0
Desenvolvido por: Daniel Ferreira da Silva
Copyright 2026 Daniel Ferreira da Silva. Todos os direitos reservados.
Licenciado exclusivamente para Rosenthal | Guarita Advogados.
"""

__version__ = "1.4.0"
__author__ = "Daniel Ferreira da Silva"
__copyright__ = "Copyright 2026 Daniel Ferreira da Silva. Todos os direitos reservados."
__product__ = "Rosenthal | Guarita Advogados"
__license__ = "Proprietario - Uso licenciado para Rosenthal | Guarita Advogados"

import os, sys, time, threading, hashlib
from datetime import datetime, timedelta, date, timezone
import tkinter as tk
from tkinter import filedialog, messagebox

import customtkinter as ctk
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

from PIL import Image, ImageDraw, ImageFont, ImageTk
import pandas as pd

try:
    import ctypes
    ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("RosenthalGuarita.PautaJudicial.1")
except Exception:
    pass

# ══════════════════ CORES ══════════════════
CHARCOAL="#2D2D2D"; CHARCOAL_DARK="#1E1E1E"; CHARCOAL_MID="#383838"
ORANGE="#E8721C"; ORANGE_HOVER="#FF8C3A"; ORANGE_DARK="#C45E12"
WHITE="#FFFFFF"; MEDIUM_GRAY="#B0B0B0"; DARK_TEXT="#E0E0E0"
SUCCESS_GREEN="#4CAF50"; ERROR_RED="#EF5350"; WARN_YELLOW="#FFB74D"; INFO_BLUE="#42A5F5"

# ══════════════════ CONFIGURACAO ══════════════════
HORAS_CORRECAO = 3
ARQUIVO_RELATORIO = "RELATORIO_ENTREGAS.xlsx"
DOMINIO_EMAIL = "rgfadv.com.br"
MAX_DIAS_FUTURO = 60  # Datas com mais de 60 dias = ERRO

FORMATOS_ACEITOS = [
    ("Planilhas Excel / CSV", "*.xlsx *.xls *.csv"),
    ("Excel (xlsx)", "*.xlsx"), ("Excel (xls)", "*.xls"), ("CSV", "*.csv"),
]

COLUNAS_OBRIGATORIAS = [
    "Data Audien", "Hora", "Advogado Responsavel Audiencia",
    "Preposto", "Link ou Redesignacao ou Cancelada",
]

# Horarios validos: 06:00 a 22:00
HORA_MIN = 6
HORA_MAX = 22

PALAVRAS_IGNORAR = [
    "CANCELADA","CANCELADO","SUSPENSO","SUSPENSA","PREPOSTO EXTERNO",
    "EXTERNO","VERIFICAR","SEM PREPOSTO","EQUIPE EXTERNA","EQUIPE INTERNA",
    "EQUIPE MISTA","REDESIGNADA","ADVOGADO EXTERNO",
]

# Colunas do relatório visível (sem IDs internos)
RELATORIO_COLS_VISIVEIS = [
    'Data Execucao','Status','Processo','Cliente','Advogado','Preposto',
    'Destinatarios','Data Audiencia','Hora Audiencia','Vara','Comarca',
    'UF','Link','Observacoes','Detalhes'
]
# Colunas internas (ocultas no Excel final)
RELATORIO_COLS_INTERNAS = ['_ID_CONTROLE','_EntryID_Main','_EntryID_Shadow']
RELATORIO_ALL_COLS = RELATORIO_COLS_VISIVEIS + RELATORIO_COLS_INTERNAS


def resource_path(rel):
    base = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, rel)

def gerar_hash(cam):
    h = hashlib.md5()
    try:
        with open(cam, 'rb') as f:
            for c in iter(lambda: f.read(8192), b''): h.update(c)
        return h.hexdigest()
    except: return None

def agora_brasilia():
    return datetime.now(timezone.utc).replace(tzinfo=None) - timedelta(hours=3)

def fmt_br(dt):
    if isinstance(dt, datetime): return dt.strftime("%d/%m/%Y %H:%M")
    if isinstance(dt, date): return dt.strftime("%d/%m/%Y")
    return str(dt)

def norm_nome(n):
    if pd.isna(n): return ""
    return str(n).strip().upper()

def _col_match(colunas, *palavras):
    def sem_acento(s):
        rep = {'Á':'A','É':'E','Í':'I','Ó':'O','Ú':'U','Ã':'A','Õ':'O',
               'Â':'A','Ê':'E','Ô':'O','Ç':'C','á':'a','é':'e','í':'i',
               'ó':'o','ú':'u','ã':'a','õ':'o','â':'a','ê':'e','ô':'o','ç':'c'}
        for k, v in rep.items(): s = s.replace(k, v)
        return s.upper()
    for c in colunas:
        cu = sem_acento(c)
        if all(sem_acento(p) in cu for p in palavras): return c
    return None


# ══════════════════ HISTORICO ══════════════════

class HistoricoManager:
    def __init__(self, pasta):
        self.arquivo = os.path.join(pasta, ARQUIVO_RELATORIO)
        self.df = pd.DataFrame(columns=RELATORIO_ALL_COLS)
        self._load()

    def _load(self):
        if os.path.exists(self.arquivo):
            try:
                self.df = pd.read_excel(self.arquivo, engine='openpyxl')
                for c in RELATORIO_ALL_COLS:
                    if c not in self.df.columns: self.df[c] = None
            except: self.df = pd.DataFrame(columns=RELATORIO_ALL_COLS)

    def ids_ativos(self):
        if self.df.empty: return set()
        return set(self.df[self.df['Status']=='SUCESSO']['_ID_CONTROLE'].astype(str).tolist())

    def dados_por_id(self, uid):
        if self.df.empty: return None
        r = self.df[self.df['_ID_CONTROLE']==str(uid)]
        return r.iloc[0].to_dict() if not r.empty else None

    def ja_existe(self, uid):
        if self.df.empty: return False
        return str(uid) in self.df['_ID_CONTROLE'].astype(str).values

    def registrar(self, d):
        self.df = pd.concat([self.df, pd.DataFrame([d])], ignore_index=True)
        self._save()

    def marcar_removido(self, uid):
        if not self.df.empty:
            self.df.loc[self.df['_ID_CONTROLE']==str(uid), 'Status'] = 'REMOVIDO'
            self._save()

    def recarregar(self): self._load()

    def _save(self):
        try:
            # Salvar APENAS colunas visíveis no Excel (sem IDs internos)
            df_save = self.df.copy()
            cols_save = [c for c in RELATORIO_COLS_VISIVEIS if c in df_save.columns]
            df_save[cols_save].to_excel(self.arquivo, index=False, engine='openpyxl')
        except: pass

    def log_visivel(self):
        if self.df.empty: return []
        cols = [c for c in ['Status','Processo','Advogado','Preposto',
                'Data Execucao','Data Audiencia','Detalhes'] if c in self.df.columns]
        return self.df[cols].to_dict('records')


# ══════════════════ MOTOR DE ANALISE ══════════════════

class MotorAnalise:
    def __init__(self, cam, hist, log_cb=None):
        self.cam = cam
        self.hist = hist
        self.log = log_cb or (lambda *a, **k: None)
        self.df_dados = None
        self.df_emails = None
        self.erros = []
        self.avisos = []

    def carregar(self):
        ext = os.path.splitext(self.cam)[1].lower()
        if ext == '.csv':
            self.erros.append("CSV nao possui aba 'NOME + E-MAIL'. Use XLSX.")
            return False
        if ext not in ('.xlsx','.xls'):
            self.erros.append("Formato nao suportado. Use XLSX ou XLS.")
            return False
        try:
            engine = 'openpyxl'
            xl = pd.ExcelFile(self.cam, engine=engine)
            abas = xl.sheet_names

            # Aba principal
            aba_princ = None
            for a in abas:
                n = a.upper().replace("Ç","C").replace("Õ","O")
                if "INFORMA" in n and "OES" in n:
                    aba_princ = a; break
            if not aba_princ:
                aba_princ = abas[0]
                self.avisos.append("Aba 'INFORMACOES' nao encontrada. Usando: '%s'" % aba_princ)

            self.df_dados = pd.read_excel(self.cam, sheet_name=aba_princ, engine=engine)
            self.df_dados.columns = [str(c).strip() for c in self.df_dados.columns]

            # Aba NOME + E-MAIL (obrigatoria)
            aba_email = None
            for a in abas:
                au = a.upper().strip()
                if ("NOME" in au and "MAIL" in au) or "E-MAIL" in au:
                    aba_email = a; break
            if not aba_email:
                self.erros.append("Aba 'NOME + E-MAIL' nao encontrada!")
                return False

            df_e = pd.read_excel(self.cam, sheet_name=aba_email, engine=engine)
            df_e.columns = [str(c).strip().upper() for c in df_e.columns]
            col_n = next((c for c in df_e.columns if 'NOME' in c), None)
            col_e = next((c for c in df_e.columns if 'MAIL' in c), None)
            if not col_n:
                self.erros.append("Coluna 'NOME' ausente na aba de e-mails.")
                return False
            if not col_e:
                self.erros.append("Coluna 'E-MAIL' ausente na aba de e-mails.")
                return False

            self.df_emails = df_e.rename(columns={col_n:'NOME', col_e:'E-MAIL'})
            self.df_emails['NOME_NORM'] = self.df_emails['NOME'].apply(norm_nome)
            self.df_emails['E-MAIL'] = self.df_emails['E-MAIL'].astype(str).str.strip()

            if self.df_emails.empty:
                self.erros.append("Aba 'NOME + E-MAIL' vazia.")
                return False

            # Validar dominio
            for _, row in self.df_emails.iterrows():
                email = str(row['E-MAIL']).strip().lower()
                nome = str(row['NOME']).strip()
                if not email or email == 'nan':
                    self.erros.append("Contato '%s' sem e-mail." % nome)
                elif not email.endswith("@" + DOMINIO_EMAIL):
                    self.erros.append("E-mail '%s' (%s) deve ser @%s." % (nome, email, DOMINIO_EMAIL))

            self.log("info", "%d contato(s) carregado(s)." % len(self.df_emails))
            return True
        except Exception as e:
            self.erros.append("Falha ao ler: %s" % str(e))
            return False

    def validar_estrutura(self):
        if self.df_dados is None or self.df_dados.empty:
            self.erros.append("Planilha vazia.")
            return False
        cols = [str(c).strip() for c in self.df_dados.columns]
        falta = []
        for obr in COLUNAS_OBRIGATORIAS:
            if not _col_match(cols, obr):
                palavras = obr.split()
                encontrou = False
                for c in cols:
                    cu = c.upper()
                    if all(p.upper()[:4] in cu for p in palavras[:2]):
                        encontrou = True; break
                if not encontrou: falta.append(obr)
        if falta:
            self.erros.append("Colunas ausentes: %s" % ', '.join(falta))
            return False
        return True

    def validar_dados(self):
        """Validacao ESTRITA de cada linha."""
        if self.df_dados is None: return

        col_cnj  = _col_match(self.df_dados.columns, "CNJ")
        col_data = _col_match(self.df_dados.columns, "Data", "Audi")
        col_hora = _col_match(self.df_dados.columns, "Hora")
        col_adv  = _col_match(self.df_dados.columns, "Advogado", "Audi")
        col_prep = _col_match(self.df_dados.columns, "Preposto")
        col_link = _col_match(self.df_dados.columns, "Link")

        agora = agora_brasilia()
        hoje = agora.date()
        limite_futuro = hoje + timedelta(days=MAX_DIAS_FUTURO)
        horarios = {}

        for idx, row in self.df_dados.iterrows():
            ln = idx + 2
            cnj = str(row.get(col_cnj, '')).strip() if col_cnj else ''
            if not cnj or cnj == 'nan': continue

            if col_cnj and not any(c.isdigit() for c in cnj):
                self.erros.append("Linha %d: CNJ '%s' invalido." % (ln, cnj))

            # ── DATA (ESTRITA) ────────────────────────────────────
            data_ok = False
            d_parsed = None
            if col_data:
                v = row.get(col_data)
                if pd.isna(v):
                    self.erros.append("Linha %d: Data da audiencia vazia." % ln)
                    continue
                try:
                    d_parsed = pd.to_datetime(v, dayfirst=True).date()
                    data_ok = True
                except Exception:
                    self.erros.append("Linha %d: Data invalida '%s'. Use DD/MM/AAAA." % (ln, v))
                    continue

                # Data no passado = ERRO
                if d_parsed < hoje:
                    self.erros.append(
                        "Linha %d: Data %s ja passou. Corrija na planilha." % (ln, fmt_br(d_parsed))
                    )
                    continue

                # Data muito no futuro = ERRO
                if d_parsed > limite_futuro:
                    self.erros.append(
                        "Linha %d: Data %s esta a mais de %d dias no futuro. Verifique o ano."
                        % (ln, fmt_br(d_parsed), MAX_DIAS_FUTURO)
                    )
                    continue

            # ── HORA (ESTRITA) ────────────────────────────────────
            hora_ok = False
            if col_hora:
                hv = row.get(col_hora)
                if pd.isna(hv):
                    self.erros.append("Linha %d: Hora vazia." % ln)
                    continue

                hora_val = None
                try:
                    if isinstance(hv, str):
                        hv_clean = hv.strip()
                        # Tentar HH:MM:SS ou HH:MM
                        for fmt in ("%H:%M:%S", "%H:%M"):
                            try:
                                hora_val = datetime.strptime(hv_clean, fmt).time()
                                break
                            except: continue
                        if hora_val is None:
                            raise ValueError("Formato invalido: %s" % hv_clean)
                    elif hasattr(hv, 'hour'):
                        hora_val = hv
                    else:
                        raise ValueError("Tipo inesperado: %s" % type(hv).__name__)
                except Exception:
                    self.erros.append(
                        "Linha %d: Hora invalida '%s'. Use formato HH:MM (ex: 08:30)." % (ln, hv)
                    )
                    continue

                # Validar faixa horaria (06:00 a 22:00)
                if hora_val.hour < HORA_MIN or hora_val.hour >= HORA_MAX:
                    self.erros.append(
                        "Linha %d: Hora %02d:%02d fora do horario comercial (%02d:00 a %02d:00)."
                        % (ln, hora_val.hour, hora_val.minute, HORA_MIN, HORA_MAX)
                    )
                    continue

                hora_ok = True

            # ── SOBREPOSICAO ──────────────────────────────────────
            if data_ok and hora_ok and d_parsed and hora_val:
                k = "%s_%02d:%02d" % (d_parsed, hora_val.hour, hora_val.minute)
                if k in horarios:
                    self.avisos.append(
                        "Linha %d: Horario sobreposto com linha %d." % (ln, horarios[k])
                    )
                else:
                    horarios[k] = ln

            # ── ADVOGADO E PREPOSTO ───────────────────────────────
            for col_nm, label in [(col_adv, "advogado"), (col_prep, "preposto")]:
                if col_nm:
                    nm = str(row.get(col_nm, '')).strip()
                    if nm and nm.lower() != 'nan':
                        if any(c.isdigit() for c in nm):
                            self.avisos.append("Linha %d: %s contem numeros: '%s'." % (ln, label.title(), nm))
                        if self.df_emails is not None and not self.df_emails.empty:
                            nn = norm_nome(nm)
                            if nn and nn != 'NAN' and not self._buscar_email(nm):
                                if not any(p in nn for p in PALAVRAS_IGNORAR):
                                    self.avisos.append("Linha %d: %s '%s' sem e-mail cadastrado." % (ln, label.title(), nm))

            # ── LINK ──────────────────────────────────────────────
            if col_link:
                lv = str(row.get(col_link, '')).strip()
                if not lv or lv.lower() in ('nan','none',''):
                    self.avisos.append("Linha %d: Link/Local vazio." % ln)
                elif lv.upper() != 'PRESENCIAL' and not lv.lower().startswith('http'):
                    if len(lv) > 3:
                        self.avisos.append("Linha %d: Link suspeito: '%s'." % (ln, lv[:50]))

    def extrair_tarefas(self):
        if self.df_dados is None: return [], []

        col_cnj  = _col_match(self.df_dados.columns, "CNJ")
        col_data = _col_match(self.df_dados.columns, "Data", "Audi")
        col_hora = _col_match(self.df_dados.columns, "Hora")
        col_adv  = _col_match(self.df_dados.columns, "Advogado", "Audi")
        col_prep = _col_match(self.df_dados.columns, "Preposto")
        col_link = _col_match(self.df_dados.columns, "Link")
        col_cli  = _col_match(self.df_dados.columns, "Cliente")
        col_vara = _col_match(self.df_dados.columns, "Vara", "turma")
        col_nvara= _col_match(self.df_dados.columns, "Nome", "vara")
        col_obs  = _col_match(self.df_dados.columns, "Compromisso")
        col_pasta= _col_match(self.df_dados.columns, "Pasta")
        col_parte= _col_match(self.df_dados.columns, "Parte", "Contr")
        col_comarca= _col_match(self.df_dados.columns, "Comarca")
        col_uf   = _col_match(self.df_dados.columns, "UF")
        col_valor= _col_match(self.df_dados.columns, "Valor")
        col_adv_proc = _col_match(self.df_dados.columns, "Advogado", "Processo")

        if not (col_cnj and col_data and col_hora): return [], []

        agora = agora_brasilia()
        hoje = agora.date()
        limite = hoje + timedelta(days=MAX_DIAS_FUTURO)
        ids_xl = set()
        novos = []

        for idx, row in self.df_dados.iterrows():
            cnj = str(row.get(col_cnj, '')).strip()
            if not cnj or cnj == 'nan': continue
            try:
                d = pd.to_datetime(row[col_data], dayfirst=True).date()
                hv = row.get(col_hora)
                hora_val = None
                if isinstance(hv, str):
                    for fmt in ("%H:%M:%S", "%H:%M"):
                        try: hora_val = datetime.strptime(hv.strip(), fmt).time(); break
                        except: continue
                elif hasattr(hv, 'hour'):
                    hora_val = hv

                if hora_val is None: continue

                # Revalidar regras estritas
                if d < hoje: continue
                if d > limite: continue
                if hora_val.hour < HORA_MIN or hora_val.hour >= HORA_MAX: continue

                dt = datetime.combine(d, hora_val)
                uid = "%s_%s" % (cnj, dt.strftime('%Y%m%d_%H%M'))

                na  = self._val_nome(row.get(col_adv)) if col_adv else None
                np_ = self._val_nome(row.get(col_prep)) if col_prep else None
                dest = list(set(filter(None, [na, np_])))
                if not dest: continue

                emails = []
                for n in dest:
                    e = self._buscar_email(n)
                    if e: emails.append(e)

                ids_xl.add(uid)

                if uid not in self.hist.ids_ativos() and dt > agora:
                    def _g(col, default=''):
                        v = row.get(col, default) if col else default
                        s = str(v).strip()
                        return s if s and s.lower() != 'nan' else default

                    novos.append({
                        'cnj': cnj, 'datetime_final': dt, 'id_unico': uid,
                        'advogado': na or "", 'preposto': np_ or "",
                        'nomes_destinatarios': ", ".join(dest),
                        'lista_emails': emails,
                        'cliente': _g(col_cli, 'CLIENTE').upper(),
                        'vara': _g(col_nvara or col_vara, 'N/A'),
                        'n_vara': _g(col_vara, ''),
                        'comarca': _g(col_comarca, ''),
                        'uf': _g(col_uf, ''),
                        'link': _g(col_link, ''),
                        'obs': _g(col_obs, ''),
                        'pasta': _g(col_pasta, ''),
                        'parte_contraria': _g(col_parte, ''),
                        'valor_causa': _g(col_valor, ''),
                        'adv_processo': _g(col_adv_proc, ''),
                    })
            except Exception:
                continue

        # Itens a remover
        remover = []
        for uid_db in self.hist.ids_ativos():
            if uid_db not in ids_xl:
                dd = self.hist.dados_por_id(uid_db)
                if dd:
                    try:
                        dt_db = datetime.strptime(str(dd.get('Data Audiencia','')), '%d/%m/%Y %H:%M')
                        if dt_db > agora: remover.append(dd)
                    except: pass

        return novos, remover

    def _val_nome(self, v):
        if pd.isna(v): return None
        s = str(v).strip()
        if not s or s.lower() == 'nan': return None
        for p in PALAVRAS_IGNORAR:
            if p in s.upper(): return None
        return s

    def _buscar_email(self, nome):
        if not nome or self.df_emails is None or self.df_emails.empty: return None
        try:
            b = norm_nome(nome)
            if not b: return None
            r = self.df_emails.loc[self.df_emails['NOME_NORM']==b, 'E-MAIL']
            if not r.empty: return r.values[0]
            for _, row in self.df_emails.iterrows():
                nd = str(row.get('NOME_NORM',''))
                if b in nd or nd in b: return row['E-MAIL']
        except: pass
        return None


# ══════════════════ EXECUTOR OUTLOOK ══════════════════

class ExecutorOutlook:
    def __init__(self, hist, log_cb=None):
        self.hist = hist
        self.log = log_cb or (lambda *a, **k: None)
        self.outlook = None
        self.meu_email = "indefinido"
        self._ok = False

    def conectar(self):
        try:
            import win32com.client as w32
            self.outlook = w32.Dispatch('outlook.application')
            try:
                self.meu_email = (
                    self.outlook.Session.CurrentUser.AddressEntry
                    .GetExchangeUser().PrimarySmtpAddress.lower()
                )
            except Exception:
                try: self.meu_email = self.outlook.Session.CurrentUser.Address.lower()
                except: pass
            self._ok = True
            self.log("info", "Outlook conectado (%s)" % self.meu_email)
            return True
        except ImportError:
            self.log("error", "pywin32 nao instalado.")
            return False
        except Exception as e:
            self.log("error", "Outlook falhou: %s" % str(e))
            return False

    @staticmethod
    def verificar_outlook():
        """Verifica se Outlook esta disponivel (sem iniciar sessao)."""
        try:
            import win32com.client as w32
            ol = w32.Dispatch('outlook.application')
            nome = ol.Session.CurrentUser.Name
            return True, nome
        except ImportError:
            return False, "pywin32 nao instalado"
        except Exception as e:
            return False, str(e)

    @property
    def disponivel(self): return self._ok

    def limpar(self, lista):
        if not self._ok: return
        for it in lista:
            try:
                eid = it.get('_EntryID_Main')
                if eid and str(eid) != 'nan':
                    self.outlook.GetItemFromID(eid).Delete()
                sid = it.get('_EntryID_Shadow')
                if sid and str(sid) != 'nan':
                    self.outlook.GetItemFromID(sid).Delete()
            except: pass
            self.hist.marcar_removido(it.get('_ID_CONTROLE'))
            self.log("warn", "Removido: %s" % it.get('Processo', '?'))

    def limpar_passados(self):
        """Remove automaticamente compromissos ja realizados (data passou)."""
        if not self._ok: return 0
        agora = agora_brasilia()
        removidos = 0
        if self.hist.df.empty: return 0

        passados = self.hist.df[self.hist.df['Status'] == 'SUCESSO'].copy()
        for _, row in passados.iterrows():
            try:
                da = str(row.get('Data Audiencia', ''))
                ha = str(row.get('Hora Audiencia', ''))
                if ha and ha != 'nan':
                    dt = datetime.strptime("%s %s" % (da, ha), '%d/%m/%Y %H:%M')
                else:
                    dt = datetime.strptime(da, '%d/%m/%Y')
                # Só apaga se já passou (com margem de 2 horas após o fim)
                if dt + timedelta(hours=2) < agora:
                    uid = str(row.get('_ID_CONTROLE', ''))
                    # Deletar do Outlook
                    for key in ('_EntryID_Main', '_EntryID_Shadow'):
                        eid = row.get(key)
                        if eid and str(eid) != 'nan':
                            try: self.outlook.GetItemFromID(eid).Delete()
                            except: pass
                    self.hist.marcar_removido(uid)
                    removidos += 1
                    self.log("info", "Limpeza: %s (passado)" % row.get('Processo', '?'))
            except: continue
        return removidos

    def agendar(self, lista):
        if not self._ok:
            self.log("error", "Outlook indisponivel.")
            return

        for i, d in enumerate(lista):
            cnj = d['cnj']
            try:
                dt_r = d['datetime_final']
                dt_a = dt_r - timedelta(hours=HORAS_CORRECAO)
                lnk = d.get('link', '')
                if len(lnk) < 5 or lnk.lower() in ('nan','none',''): lnk = "LINK NAO IDENTIFICADO"

                cli  = d.get('cliente','CLIENTE')
                vara = d.get('vara','N/A')
                nms  = d.get('nomes_destinatarios','')
                obs  = d.get('obs','')

                corpo = (
                    "="*60+"\nAUDIENCIA - %s\n"%cli+"="*60+"\n"
                    "LINK/LOCAL: %s\n"%lnk+"-"*60+"\n"
                    "Processo: %s\n"%cnj+
                    "DATA: %s\n"%dt_r.strftime('%d/%m/%Y as %H:%M')+
                    "Vara: %s\n"%vara+"Responsaveis: %s\n"%nms+"OBS: %s\n"%obs
                )

                ap = self.outlook.CreateItem(1)
                ap.Subject = "AUDIENCIA: %s (CNJ: %s)" % (cli, cnj)
                ap.Start = dt_a; ap.Duration = 60; ap.Body = corpo
                ap.ReminderSet = True; ap.ReminderMinutesBeforeStart = 15
                ap.MeetingStatus = 0; ap.Save()
                em = ap.EntryID

                ec = [e for e in d.get('lista_emails',[]) if e.lower() != self.meu_email]
                if ec:
                    ap.MeetingStatus = 1
                    for e in ec: ap.Recipients.Add(e)
                    ap.Recipients.ResolveAll(); ap.Send()

                es = None
                if (dt_r - agora_brasilia()).total_seconds() > 600:
                    s = self.outlook.CreateItem(1)
                    s.Subject = "ENTRAR NA SALA - %s" % cli
                    s.Start = dt_a; s.Duration = 0
                    s.ReminderSet = True; s.ReminderMinutesBeforeStart = 5
                    s.Body = "LINK:\n%s\nResp: %s" % (lnk, nms)
                    s.MeetingStatus = 0; s.Save()
                    es = s.EntryID

                self.hist.registrar({
                    "Data Execucao": agora_brasilia().strftime("%d/%m/%Y %H:%M:%S"),
                    "_ID_CONTROLE": d['id_unico'],
                    "Processo": cnj,
                    "Cliente": d.get('cliente',''),
                    "Advogado": d.get('advogado',''),
                    "Preposto": d.get('preposto',''),
                    "Destinatarios": nms,
                    "Data Audiencia": dt_r.strftime('%d/%m/%Y'),
                    "Hora Audiencia": dt_r.strftime('%H:%M'),
                    "Vara": d.get('vara',''),
                    "Comarca": d.get('comarca',''),
                    "UF": d.get('uf',''),
                    "Link": lnk,
                    "Observacoes": obs,
                    "Status": "SUCESSO",
                    "_EntryID_Main": em,
                    "_EntryID_Shadow": es,
                    "Detalhes": "Convites: %d" % len(ec),
                })

                self.log("success", "[%d/%d] %s - %s" % (i+1, len(lista), cnj, nms))

            except Exception as e:
                self.log("error", "Falha %s: %s" % (cnj, str(e)))
                self.hist.registrar({
                    "Data Execucao": agora_brasilia().strftime("%d/%m/%Y %H:%M:%S"),
                    "_ID_CONTROLE": d.get('id_unico',''),
                    "Processo": cnj, "Advogado": d.get('advogado',''),
                    "Preposto": d.get('preposto',''),
                    "Status": "FALHA", "Detalhes": str(e)[:200],
                })


# ══════════════════ COMPONENTES UI ══════════════════

class LogPanel(ctk.CTkFrame):
    def __init__(self, master, **kw):
        super().__init__(master, fg_color=CHARCOAL_DARK, corner_radius=12, **kw)
        ctk.CTkLabel(self, text="  LOG DE ATIVIDADES",
            font=ctk.CTkFont(size=14, weight="bold"),
            text_color=MEDIUM_GRAY, anchor="w").pack(fill="x", padx=16, pady=(12,4))
        self.tb = ctk.CTkTextbox(self, fg_color=CHARCOAL_MID, text_color=DARK_TEXT,
            font=ctk.CTkFont(family="Consolas", size=12),
            corner_radius=8, border_width=0, wrap="word", state="disabled")
        self.tb.pack(fill="both", expand=True, padx=12, pady=(4,12))
        for tag, cor in [("success",SUCCESS_GREEN),("error",ERROR_RED),
            ("warn",WARN_YELLOW),("info",INFO_BLUE),("dim",MEDIUM_GRAY)]:
            self.tb._textbox.tag_config(tag, foreground=cor)

    def add(self, tipo, msg):
        def _do():
            self.tb.configure(state="normal")
            ts = agora_brasilia().strftime("%H:%M:%S")
            ic = {"success":"[OK]","error":"[X]","warn":"[!]","info":"[i]"}.get(tipo,"[*]")
            self.tb._textbox.insert("end", "[%s] " % ts, "dim")
            self.tb._textbox.insert("end", "%s %s\n" % (ic, msg), tipo)
            self.tb._textbox.see("end")
            self.tb.configure(state="disabled")
        self.after(0, _do)

    def clear(self):
        self.tb.configure(state="normal")
        self.tb._textbox.delete("1.0","end")
        self.tb.configure(state="disabled")


class BackupTable(ctk.CTkFrame):
    def __init__(self, master, **kw):
        super().__init__(master, fg_color=CHARCOAL_DARK, corner_radius=12, **kw)
        ctk.CTkLabel(self, text="  RELATORIO DE ENTREGAS",
            font=ctk.CTkFont(size=14, weight="bold"),
            text_color=MEDIUM_GRAY, anchor="w").pack(fill="x", padx=16, pady=(12,4))
        self.tb = ctk.CTkTextbox(self, fg_color=CHARCOAL_MID, text_color=DARK_TEXT,
            font=ctk.CTkFont(family="Consolas", size=11),
            corner_radius=8, border_width=0, wrap="none", state="disabled")
        self.tb.pack(fill="both", expand=True, padx=12, pady=(4,12))
        for tag, cor in [("header",ORANGE),("success",SUCCESS_GREEN),
            ("error",ERROR_RED),("warn",WARN_YELLOW),("dim",MEDIUM_GRAY)]:
            self.tb._textbox.tag_config(tag, foreground=cor)

    def update_data(self, data):
        self.tb.configure(state="normal")
        self.tb._textbox.delete("1.0","end")
        hdr = "%-12s %-28s %-18s %-18s %-16s %-16s" % (
            "Status","Processo","Advogado","Preposto","Execucao","Audiencia")
        self.tb._textbox.insert("end", hdr+"\n", "header")
        self.tb._textbox.insert("end", "-"*110+"\n", "dim")
        for it in reversed(data[-50:]):
            st = str(it.get('Status',''))
            tag = "success" if st=="SUCESSO" else "error" if "FALHA" in st else "warn" if "IGNORADO" in st else "dim"
            linha = "%-12s %-28s %-18s %-18s %-16s %-16s" % (
                st[:12], str(it.get('Processo',''))[:28],
                str(it.get('Advogado',''))[:18], str(it.get('Preposto',''))[:18],
                str(it.get('Data Execucao',''))[:16], str(it.get('Data Audiencia',''))[:16])
            self.tb._textbox.insert("end", linha+"\n", tag)
        self.tb.configure(state="disabled")


class SobreDialog(ctk.CTkToplevel):
    def __init__(self, master, logo_ref=None):
        super().__init__(master)
        self.title("Sobre"); self.geometry("500x420"); self.resizable(False,False)
        self.configure(fg_color=CHARCOAL); self.transient(master); self.grab_set()
        if logo_ref:
            ctk.CTkLabel(self, image=logo_ref, text="").pack(pady=(28,8))
        else:
            ctk.CTkLabel(self, text="ROSENTHAL | GUARITA\nADVOGADOS",
                font=ctk.CTkFont(size=18, weight="bold"), text_color=WHITE,
                justify="center").pack(pady=(28,8))
        ctk.CTkLabel(self, text="Gestao de Pautas Judiciais",
            font=ctk.CTkFont(size=14), text_color=ORANGE).pack(pady=(0,4))
        ctk.CTkLabel(self, text="Versao %s" % __version__,
            font=ctk.CTkFont(size=12), text_color=MEDIUM_GRAY).pack(pady=(0,16))
        ctk.CTkFrame(self, fg_color=CHARCOAL_MID, height=1).pack(fill="x", padx=40, pady=4)
        info = ("%s\n\nDesenvolvido por %s\n%s\n\n"
            "Este software e proprietario. A reproducao, distribuicao\n"
            "ou engenharia reversa sem autorizacao expressa do autor\n"
            "constitui violacao (Lei 9.609/98 e Lei 9.610/98)."
        ) % (__copyright__, __author__, __license__)
        ctk.CTkLabel(self, text=info, font=ctk.CTkFont(size=11),
            text_color=DARK_TEXT, justify="center", wraplength=420).pack(pady=12)
        ctk.CTkButton(self, text="Fechar", command=self.destroy,
            fg_color=ORANGE, hover_color=ORANGE_HOVER,
            text_color=WHITE, corner_radius=8, width=120).pack(pady=(4,20))


# ══════════════════ JANELA PRINCIPAL ══════════════════

class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Rosenthal | Guarita Advogados")
        self.geometry("1300x800"); self.minsize(980,620)
        self.configure(fg_color=CHARCOAL_DARK)

        self.arq = None; self.hash_arq = None; self.estado = "parado"
        self.thread_mon = None; self.ev_parar = threading.Event()
        self.ev_pausar = threading.Event()
        self.pasta = os.path.dirname(os.path.abspath(__file__))
        self.hist = HistoricoManager(self.pasta)
        self._logo_ref = None; self._icon_refs = []

        self._set_icon()
        self._build_ui()
        self._refresh_table()
        self._check_outlook_status()
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    def _set_icon(self):
        ico = resource_path("assets/app.ico")
        if os.path.exists(ico):
            try: self.after(50, lambda: self._try_set_icon(ico))
            except: pass

    def _try_set_icon(self, ico_path):
        try: self.iconbitmap(default=ico_path)
        except:
            try: self.iconbitmap(ico_path)
            except: pass

    def _load_logo(self):
        for n in ['logo_transparente.png','logo_clean.png','logo_light.png']:
            p = resource_path("assets/"+n)
            if os.path.exists(p):
                try:
                    img = Image.open(p); tw = 200
                    ratio = tw/img.width; th = int(img.height*ratio)
                    img = img.resize((tw,th), Image.LANCZOS)
                    return ctk.CTkImage(light_image=img, dark_image=img, size=(tw,th))
                except: pass
        return None

    def _check_outlook_status(self):
        """Verifica Outlook ao abrir o app."""
        def _check():
            ok, info = ExecutorOutlook.verificar_outlook()
            if ok:
                self.after(0, lambda: self.c_outl.configure(text="OK"))
                self.after(0, lambda: self.logp.add("info", "Outlook: %s" % info))
            else:
                self.after(0, lambda: self.c_outl.configure(text="OFFLINE"))
                self.after(0, lambda: self.logp.add("warn", "Outlook nao detectado: %s" % info))
        threading.Thread(target=_check, daemon=True).start()

    def _build_ui(self):
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # SIDEBAR
        sb = ctk.CTkFrame(self, fg_color=CHARCOAL, width=250, corner_radius=0)
        sb.grid(row=0, column=0, sticky="nsew"); sb.grid_propagate(False)

        self._logo_ref = self._load_logo()
        if self._logo_ref:
            ctk.CTkLabel(sb, image=self._logo_ref, text="").pack(pady=(24,12))
        else:
            ctk.CTkLabel(sb, text="ROSENTHAL | GUARITA\nADVOGADOS",
                font=ctk.CTkFont(size=14, weight="bold"),
                text_color=WHITE, justify="center").pack(pady=(24,12))

        ctk.CTkFrame(sb, fg_color=CHARCOAL_MID, height=1).pack(fill="x", padx=20, pady=4)

        sf = ctk.CTkFrame(sb, fg_color=CHARCOAL_MID, corner_radius=10)
        sf.pack(fill="x", padx=16, pady=8)
        self.st_dot = ctk.CTkLabel(sf, text="*", font=ctk.CTkFont(size=18, weight="bold"), text_color=MEDIUM_GRAY)
        self.st_dot.pack(side="left", padx=(12,6), pady=8)
        self.st_lbl = ctk.CTkLabel(sf, text="Parado", font=ctk.CTkFont(size=13, weight="bold"), text_color=DARK_TEXT)
        self.st_lbl.pack(side="left", pady=8)

        self.arq_lbl = ctk.CTkLabel(sb, text="Nenhum arquivo carregado",
            font=ctk.CTkFont(size=11), text_color=MEDIUM_GRAY, wraplength=210, justify="center")
        self.arq_lbl.pack(padx=16, pady=(8,4))

        self.btn_load = ctk.CTkButton(sb, text="Carregar Planilha", command=self._load_excel,
            fg_color=ORANGE, hover_color=ORANGE_HOVER, text_color=WHITE,
            corner_radius=8, font=ctk.CTkFont(size=13, weight="bold"), height=40)
        self.btn_load.pack(fill="x", padx=20, pady=(12,6))

        ctk.CTkFrame(sb, fg_color=CHARCOAL_MID, height=1).pack(fill="x", padx=20, pady=12)
        ctk.CTkLabel(sb, text="CONTROLES", font=ctk.CTkFont(size=10, weight="bold"),
            text_color=MEDIUM_GRAY).pack(padx=20, anchor="w")

        self.btn_start = ctk.CTkButton(sb, text="Iniciar", command=self._start,
            fg_color=SUCCESS_GREEN, hover_color="#66BB6A", text_color=WHITE,
            corner_radius=8, font=ctk.CTkFont(size=13, weight="bold"), height=38)
        self.btn_start.pack(fill="x", padx=20, pady=(8,4))

        self.btn_pause = ctk.CTkButton(sb, text="Pausar", command=self._pause,
            fg_color=WARN_YELLOW, hover_color="#FFA726", text_color=CHARCOAL,
            corner_radius=8, font=ctk.CTkFont(size=13, weight="bold"), height=38, state="disabled")
        self.btn_pause.pack(fill="x", padx=20, pady=4)

        self.btn_stop = ctk.CTkButton(sb, text="Parar", command=self._stop,
            fg_color=ERROR_RED, hover_color="#EF5350", text_color=WHITE,
            corner_radius=8, font=ctk.CTkFont(size=13, weight="bold"), height=38, state="disabled")
        self.btn_stop.pack(fill="x", padx=20, pady=4)

        ctk.CTkLabel(sb, text="").pack(expand=True)
        ctk.CTkButton(sb, text="Sobre", command=lambda: SobreDialog(self, self._logo_ref),
            fg_color="transparent", hover_color=CHARCOAL_MID, text_color=MEDIUM_GRAY,
            corner_radius=8, font=ctk.CTkFont(size=12), height=32,
            border_width=1, border_color=CHARCOAL_MID).pack(fill="x", padx=20, pady=(4,16))

        # MAIN
        ma = ctk.CTkFrame(self, fg_color=CHARCOAL_DARK, corner_radius=0)
        ma.grid(row=0, column=1, sticky="nsew")
        ma.grid_columnconfigure(0, weight=1); ma.grid_rowconfigure(1, weight=1)

        tb = ctk.CTkFrame(ma, fg_color=CHARCOAL, height=56, corner_radius=0)
        tb.grid(row=0, column=0, sticky="ew"); tb.grid_propagate(False)
        ctk.CTkLabel(tb, text="Painel de Controle - Audiencias Judiciais",
            font=ctk.CTkFont(size=16, weight="bold"), text_color=WHITE).pack(side="left", padx=20, pady=12)
        self.clock_lbl = ctk.CTkLabel(tb, text="", font=ctk.CTkFont(size=12), text_color=MEDIUM_GRAY)
        self.clock_lbl.pack(side="right", padx=20, pady=12)
        self._tick()

        cf = ctk.CTkFrame(ma, fg_color="transparent")
        cf.grid(row=0, column=0, sticky="ew", padx=16, pady=(64,0))
        cf.grid_columnconfigure((0,1,2,3), weight=1)

        self.c_tot  = self._card(cf, "Total Agendados", "0", SUCCESS_GREEN, 0)
        self.c_err  = self._card(cf, "Erros", "0", ERROR_RED, 1)
        self.c_warn = self._card(cf, "Avisos", "0", WARN_YELLOW, 2)
        self.c_outl = self._card(cf, "Outlook", "--", INFO_BLUE, 3)

        tv = ctk.CTkTabview(ma, fg_color=CHARCOAL_DARK,
            segmented_button_fg_color=CHARCOAL, segmented_button_selected_color=ORANGE,
            segmented_button_selected_hover_color=ORANGE_HOVER,
            segmented_button_unselected_color=CHARCOAL_MID,
            segmented_button_unselected_hover_color=CHARCOAL_MID,
            text_color=WHITE, corner_radius=12)
        tv.grid(row=1, column=0, sticky="nsew", padx=16, pady=(8,16))

        t1 = tv.add("Log de Atividades")
        t1.grid_columnconfigure(0, weight=1); t1.grid_rowconfigure(0, weight=1)
        self.logp = LogPanel(t1)
        self.logp.grid(row=0, column=0, sticky="nsew", padx=4, pady=4)

        t2 = tv.add("Relatorio / Backup")
        t2.grid_columnconfigure(0, weight=1); t2.grid_rowconfigure(0, weight=1)
        self.bktbl = BackupTable(t2)
        self.bktbl.grid(row=0, column=0, sticky="nsew", padx=4, pady=4)

    def _card(self, parent, title, val, color, col):
        f = ctk.CTkFrame(parent, fg_color=CHARCOAL, corner_radius=12)
        f.grid(row=0, column=col, sticky="ew", padx=6, pady=6)
        ctk.CTkLabel(f, text=title, font=ctk.CTkFont(size=11), text_color=MEDIUM_GRAY).pack(padx=12, pady=(10,0), anchor="w")
        l = ctk.CTkLabel(f, text=val, font=ctk.CTkFont(size=28, weight="bold"), text_color=color)
        l.pack(padx=12, pady=(0,10), anchor="w")
        return l

    def _tick(self):
        a = agora_brasilia()
        self.clock_lbl.configure(text="Brasilia: %s" % a.strftime('%d/%m/%Y  %H:%M:%S'))
        self.after(1000, self._tick)

    def _set_status(self, st, txt=None):
        self.estado = st
        mapa = {"parado":(MEDIUM_GRAY,"Parado"), "rodando":(SUCCESS_GREEN,"Em execucao"),
            "pausado":(WARN_YELLOW,"Pausado"), "analisando":(INFO_BLUE,"Analisando..."),
            "erro":(ERROR_RED,"Erro")}
        cor, td = mapa.get(st, (MEDIUM_GRAY,"--"))
        self.st_dot.configure(text_color=cor); self.st_lbl.configure(text=txt or td)

        if st == "rodando":
            self.btn_start.configure(state="disabled"); self.btn_pause.configure(state="normal")
            self.btn_stop.configure(state="normal"); self.btn_load.configure(state="disabled")
        elif st == "pausado":
            self.btn_start.configure(state="normal"); self.btn_pause.configure(state="disabled")
            self.btn_stop.configure(state="normal")
        else:
            self.btn_start.configure(state="normal" if self.arq else "disabled")
            self.btn_pause.configure(state="disabled"); self.btn_stop.configure(state="disabled")
            self.btn_load.configure(state="normal")

    def _load_excel(self):
        cam = filedialog.askopenfilename(title="Selecionar Planilha",
            filetypes=FORMATOS_ACEITOS, initialdir=os.path.join(self.pasta,"excel"))
        if not cam: return
        ext = os.path.splitext(cam)[1].lower()
        if ext not in ('.xlsx','.xls','.csv'):
            messagebox.showerror("Formato Invalido", "Use XLSX, XLS ou CSV.")
            return
        self.arq = cam; self.hash_arq = gerar_hash(cam); self._first = False
        self.arq_lbl.configure(text=os.path.basename(cam), text_color=WHITE)
        self.logp.clear()
        self.logp.add("info", "Arquivo: %s" % os.path.basename(cam))
        self._set_status("parado")
        self._pre_analyze()

    def _pre_analyze(self):
        if not self.arq: return
        self._set_status("analisando")
        self.logp.add("info", "Analisando planilha...")

        m = MotorAnalise(self.arq, self.hist, lambda t, msg: self.logp.add(t, msg))
        if not m.carregar():
            for e in m.erros: self.logp.add("error", e)
            self._set_status("erro", "Erro na leitura"); return
        if not m.validar_estrutura():
            for e in m.erros: self.logp.add("error", e)
            self._set_status("erro", "Estrutura invalida"); return

        m.validar_dados()
        for e in m.erros: self.logp.add("error", e)
        for a in m.avisos: self.logp.add("warn", a)

        novos, rem = m.extrair_tarefas()
        self.c_tot.configure(text=str(len(novos)))
        self.c_err.configure(text=str(len(m.erros)))
        self.c_warn.configure(text=str(len(m.avisos)))

        if m.erros:
            self.logp.add("error", "%d erro(s) encontrado(s). Corrija a planilha." % len(m.erros))
            self._set_status("erro", "%d erro(s)" % len(m.erros))
        elif m.avisos:
            self.logp.add("warn", "%d aviso(s)." % len(m.avisos))
            self.logp.add("success", "%d audiencia(s) prontas." % len(novos))
            self._set_status("parado")
        else:
            self.logp.add("success", "Planilha OK! %d audiencia(s)." % len(novos))
            self._set_status("parado")

        if rem: self.logp.add("warn", "%d obsoleto(s) a remover." % len(rem))

    def _start(self):
        if not self.arq:
            messagebox.showwarning("Atencao","Carregue uma planilha."); return
        if self.estado == "pausado":
            self.ev_pausar.clear(); self._set_status("rodando","Retomando...")
            self.logp.add("info","Retomado."); return
        self.ev_parar.clear(); self.ev_pausar.clear(); self._first = False
        self._set_status("rodando"); self.logp.add("info","Iniciando...")
        self.thread_mon = threading.Thread(target=self._loop, daemon=True)
        self.thread_mon.start()

    def _pause(self):
        self.ev_pausar.set(); self._set_status("pausado"); self.logp.add("warn","Pausado.")

    def _stop(self):
        self.ev_parar.set(); self.ev_pausar.clear()
        self._set_status("parado"); self.logp.add("info","Parado.")

    def _loop(self):
        last_h = self.hash_arq
        while not self.ev_parar.is_set():
            while self.ev_pausar.is_set() and not self.ev_parar.is_set(): time.sleep(0.5)
            if self.ev_parar.is_set(): break
            try:
                nh = gerar_hash(self.arq)
                if nh and nh != last_h:
                    self.after(0, lambda: self.logp.add("info","Alteracao detectada..."))
                    last_h = nh; self.hash_arq = nh; self._cycle()
                elif not getattr(self, '_first', False):
                    self._first = True; self._cycle()
            except Exception as e:
                self.after(0, lambda err=e: self.logp.add("error","Erro: %s" % str(err)))
            for _ in range(50):
                if self.ev_parar.is_set(): break
                time.sleep(0.1)
        self.after(0, lambda: self._set_status("parado"))

    def _cycle(self):
        def _log(t, msg): self.after(0, lambda: self.logp.add(t, msg))

        m = MotorAnalise(self.arq, self.hist, _log)
        if not m.carregar():
            for e in m.erros: _log("error", e); return
        if not m.validar_estrutura():
            for e in m.erros: _log("error", e); return

        m.validar_dados()
        for e in m.erros: _log("error", e)
        for a in m.avisos: _log("warn", a)

        if m.erros:
            _log("error","Erros criticos. Agendamento cancelado.")
            self.after(0, lambda n=len(m.erros): self.c_err.configure(text=str(n)))
            return

        novos, rem = m.extrair_tarefas()
        self.after(0, lambda n=len(novos): self.c_tot.configure(text=str(n)))
        self.after(0, lambda n=len(m.erros): self.c_err.configure(text=str(n)))
        self.after(0, lambda n=len(m.avisos): self.c_warn.configure(text=str(n)))

        if not novos and not rem:
            _log("success","Sincronizado. Aguardando..."); return

        ex = ExecutorOutlook(self.hist, _log)
        if ex.conectar():
            self.after(0, lambda: self.c_outl.configure(text="OK"))
            # Limpeza automática de compromissos passados
            n_limp = ex.limpar_passados()
            if n_limp > 0:
                _log("info", "%d compromisso(s) passado(s) removido(s) da agenda." % n_limp)
            if rem: ex.limpar(rem)
            if novos: ex.agendar(novos)
        else:
            self.after(0, lambda: self.c_outl.configure(text="FALHA"))

        self.hist.recarregar()
        self.after(0, self._refresh_table)

    def _refresh_table(self): self.bktbl.update_data(self.hist.log_visivel())

    def _on_close(self):
        if self.estado in ("rodando","pausado"):
            if not messagebox.askyesno("Rosenthal | Guarita","Monitoramento ativo. Sair?"): return
        self.ev_parar.set(); self.destroy()


def main():
    try:
        app = App(); app.mainloop()
    except Exception as e:
        try:
            root = tk.Tk(); root.withdraw()
            messagebox.showerror("Erro","Falha ao iniciar:\n%s\n\nExecute DIAGNOSTICO.bat" % str(e))
            root.destroy()
        except: print("[ERRO] %s" % str(e)); input("Enter...")

if __name__ == "__main__":
    main()
