import json
import os
import re
import time
import threading
import datetime
from pathlib import Path
from zoneinfo import ZoneInfo

import requests
from dotenv import load_dotenv
from requests.auth import HTTPBasicAuth

from kivy.app import App
from kivy.clock import Clock
from kivy.core.window import Window
from kivy.graphics import Color, Rectangle
from kivy.metrics import dp
from kivy.properties import BooleanProperty, StringProperty
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.button import Button
from kivy.uix.gridlayout import GridLayout
from kivy.uix.label import Label
from kivy.uix.screenmanager import ScreenManager, Screen
from kivy.uix.scrollview import ScrollView
from kivy.uix.spinner import Spinner
from kivy.uix.textinput import TextInput


# =========================
# CONFIG / ENV
# =========================
BASE_DIR = Path(__file__).resolve().parent
load_dotenv(BASE_DIR / "teste.env")

SUPABASE_URL = os.getenv("SUPABASE_URL", "").strip().rstrip("/")
SUPABASE_KEY = os.getenv("SUPABASE_KEY", os.getenv("SUPABASE_ANON_KEY", "")).strip()

TOTVS_API_BASE = os.getenv(
    "TOTVS_API_BASE",
    "http://200.201.240.47:8383/rest01/PY_APONTAMEN"
).strip().rstrip("/")

TOTVS_TIMEOUT = int(os.getenv("TOTVS_TIMEOUT", "100"))
TOTVS_TENANT_ID = os.getenv("TOTVS_TENANT_ID", "").strip()

INACTIVITY_TIMEOUT_SEC = 180

TZ = ZoneInfo("America/Sao_Paulo")
Window.clearcolor = (0.06, 0.08, 0.11, 1)


# =========================
# HELPERS
# =========================
def normalizar_texto(valor) -> str:
    if valor is None:
        return ""
    return str(valor).strip()


def _normaliza_codigo(v) -> str:
    if v is None:
        return ""
    s = str(v).strip()
    if s.endswith(".0") and s[:-2].isdigit():
        s = s[:-2]
    return s


def agora_iso():
    return datetime.datetime.now(datetime.timezone.utc).isoformat()


def hora_agora():
    return datetime.datetime.now(TZ).strftime("%H:%M:%S")


def headers_totvs():
    headers = {
        "Content-Type": "application/json",
        "Accept": "application/json",
    }
    if TOTVS_TENANT_ID:
        headers["tenantId"] = TOTVS_TENANT_ID
    return headers


def supabase_headers():
    return {
        "apikey": SUPABASE_KEY,
        "Authorization": f"Bearer {SUPABASE_KEY}",
        "Content-Type": "application/json",
        "Accept": "application/json",
    }


def supabase_url(table_name: str) -> str:
    return f"{SUPABASE_URL}/rest/v1/{table_name}"


def supabase_get(table_name: str, params: dict):
    resp = requests.get(
        supabase_url(table_name),
        headers=supabase_headers(),
        params=params,
        timeout=20,
    )
    if not (200 <= resp.status_code < 300):
        raise RuntimeError(f"GET {table_name} -> HTTP {resp.status_code} | {resp.text}")
    return resp.json()


def supabase_post(table_name: str, payload):
    headers = supabase_headers().copy()
    headers["Prefer"] = "return=representation"
    resp = requests.post(
        supabase_url(table_name),
        headers=headers,
        json=payload,
        timeout=20,
    )
    if not (200 <= resp.status_code < 300):
        raise RuntimeError(f"POST {table_name} -> HTTP {resp.status_code} | {resp.text}")
    try:
        return resp.json()
    except Exception:
        return []


def supabase_patch(table_name: str, params: dict, payload: dict):
    headers = supabase_headers().copy()
    headers["Prefer"] = "return=representation"
    resp = requests.patch(
        supabase_url(table_name),
        headers=headers,
        params=params,
        json=payload,
        timeout=20,
    )
    if not (200 <= resp.status_code < 300):
        raise RuntimeError(f"PATCH {table_name} -> HTTP {resp.status_code} | {resp.text}")
    try:
        return resp.json()
    except Exception:
        return []


# =========================
# TRATAMENTO TOTVS
# =========================
def corpo_resposta_http(resp):
    try:
        return resp.json()
    except Exception:
        txt = (resp.text or "").strip()
        return txt if txt else "Sem conteúdo"


def normalizar_quebras(msg: str) -> str:
    return str(msg).replace("\r\n", "\n").replace("\r", "\n")


def remover_cabecalho_ajuda(msg: str) -> str:
    msg = normalizar_quebras(msg)
    linhas = msg.split("\n")
    if linhas and linhas[0].strip().upper().startswith("AJUDA:"):
        linhas = linhas[1:]
    return "\n".join(linhas).strip()


def limpar_mensagem_totvs(msg) -> str:
    if msg is None:
        return "Sem conteúdo"
    msg = remover_cabecalho_ajuda(str(msg))
    linhas = [re.sub(r"\s+", " ", linha).strip() for linha in msg.split("\n") if linha.strip()]
    return " ".join(linhas).strip()


def converter_saldo_para_float(valor):
    if valor is None:
        return None
    txt = str(valor).strip()
    if not txt:
        return None
    txt = txt.replace(".", "").replace(",", ".") if txt.count(",") == 1 and txt.count(".") > 1 else txt.replace(",", ".")
    try:
        return float(txt)
    except Exception:
        return None


def _fmt_qtd_br(valor):
    num = converter_saldo_para_float(valor)
    if num is None:
        return str(valor).strip()
    txt = f"{num:,.2f}"
    txt = txt.replace(",", "X").replace(".", ",").replace("X", ".")
    txt = txt.replace("-", "- ")
    return txt


def _fmt_status_item(valor):
    txt = re.sub(r"\s+", " ", str(valor or "")).strip()
    txt_upper = txt.upper()

    mapa = {
        "SEM SALDO EM ESTOQUE": "Falta de saldo",
        "SALDO INDISPONIVEL": "Saldo indisponível",
        "SALDO INDISPONÍVEL": "Saldo indisponível",
    }

    if txt_upper in mapa:
        return mapa[txt_upper]

    if not txt:
        return "-"

    return txt[:1].upper() + txt[1:]


def extrair_itens_erro_estoque(msg):
    if msg is None:
        return []

    texto = remover_cabecalho_ajuda(str(msg))
    texto = re.sub(r"\s+", " ", texto).strip()

    padrao_item = re.compile(
        r'(?P<produto>[A-Z0-9]{2,}(?:\.[A-Z0-9]{1,})+)'
        r'\s+'
        r'(?P<armazem>[A-Z0-9]{1,10})'
        r'\s+'
        r'(?P<saldo>-?\d+(?:[.,]\d+)?)'
        r'\s+'
        r'(?P<ocorrencia>.*?)(?=(?:[A-Z0-9]{2,}(?:\.[A-Z0-9]{1,})+\s+[A-Z0-9]{1,10}\s+-?\d)|$)',
        flags=re.IGNORECASE
    )

    itens = []
    for m in padrao_item.finditer(texto):
        produto = m.group("produto").strip()
        armazem = m.group("armazem").strip()
        saldo = m.group("saldo").strip()
        ocorrencia = re.sub(r"\s+", " ", m.group("ocorrencia")).strip()

        if produto.upper() == "PRODUTO":
            continue

        itens.append({
            "produto": produto,
            "armazem": armazem,
            "saldo": saldo,
            "saldo_num": converter_saldo_para_float(saldo),
            "ocorrencia": ocorrencia,
            "motivo": ocorrencia,
        })

    return itens


def formatar_mensagem_tela_totvs(msg):
    if msg is None:
        return "Sem conteúdo", []

    texto = remover_cabecalho_ajuda(str(msg))
    texto_unico = re.sub(r"\s+", " ", texto).strip()
    texto_upper = texto_unico.upper()

    itens = extrair_itens_erro_estoque(texto)

    if (
        "NÃO EXISTE QUANTIDADE SUFICIENTE EM ESTOQUE" in texto_upper
        or "NAO EXISTE QUANTIDADE SUFICIENTE EM ESTOQUE" in texto_upper
    ):
        linhas = [
            "Não existe quantidade suficiente em estoque para atender esta requisição.",
            "",
            "CÓDIGO | ARMAZEM | QUANTIDADE | STATUS"
        ]

        if itens:
            for item in itens:
                codigo = item.get("produto", "").strip()
                armazem = item.get("armazem", "").strip()
                quantidade = _fmt_qtd_br(item.get("saldo"))
                status = _fmt_status_item(item.get("ocorrencia"))
                linhas.append(f"{codigo} | {armazem} | {quantidade} | {status}")
        else:
            linhas.append("Sem itens identificados.")

        return "\n".join(linhas).strip(), itens

    linhas = [
        re.sub(r"\s+", " ", linha).strip()
        for linha in texto.split("\n")
        if linha.strip()
    ]
    return "\n".join(linhas).strip(), itens


def interpretar_retorno_totvs(response):
    corpo = corpo_resposta_http(response)

    if isinstance(corpo, dict):
        body_raw = json.dumps(corpo, ensure_ascii=False, indent=2)

        note = corpo.get("note")
        message = corpo.get("message")
        error = corpo.get("error")
        error_id = corpo.get("errorId")

        mensagem_base = note or message or error or body_raw

        mensagem_amigavel = limpar_mensagem_totvs(mensagem_base)
        mensagem_tela, itens_estoque = formatar_mensagem_tela_totvs(mensagem_base)

        texto_upper = mensagem_amigavel.upper()
        erro_negocio = bool(error_id or error)
        if "OP NÃO EXISTE" in texto_upper or "OP NAO EXISTE" in texto_upper:
            erro_negocio = True

        sucesso = (200 <= response.status_code < 300) and not erro_negocio

        return {
            "sucesso": sucesso,
            "status_code": response.status_code,
            "body_raw": body_raw,
            "body_json": corpo,
            "mensagem_amigavel": mensagem_amigavel,
            "mensagem_tela": mensagem_tela,
            "itens_estoque": itens_estoque,
            "headers": dict(response.headers),
        }

    body_raw = str(corpo)
    mensagem_amigavel = limpar_mensagem_totvs(body_raw)
    mensagem_tela, itens_estoque = formatar_mensagem_tela_totvs(body_raw)
    texto_upper = mensagem_amigavel.upper()

    erro_negocio = "OP NÃO EXISTE" in texto_upper or "OP NAO EXISTE" in texto_upper
    sucesso = (200 <= response.status_code < 300) and not erro_negocio

    return {
        "sucesso": sucesso,
        "status_code": response.status_code,
        "body_raw": body_raw,
        "body_json": None,
        "mensagem_amigavel": mensagem_amigavel,
        "mensagem_tela": mensagem_tela,
        "itens_estoque": itens_estoque,
        "headers": dict(response.headers),
    }


# =========================
# SUPABASE / FILA
# =========================
def buscar_apontamento_por_serie(numero_serie):
    dados = supabase_get(
        "apontamentos_manga_pnm",
        {
            "select": "*",
            "numero_serie": f"eq.{_normaliza_codigo(numero_serie)}",
            "order": "id.desc",
            "limit": "1",
        },
    )
    return dados[0] if dados else None


def inserir_apontamento(numero_serie, op, tipo_producao, usuario):
    payload = {
        "numero_serie": _normaliza_codigo(numero_serie),
        "op": _normaliza_codigo(op),
        "tipo_producao": _normaliza_codigo(tipo_producao),
        "usuario": _normaliza_codigo(usuario),
        "data_hora": agora_iso(),
    }
    dados = supabase_post("apontamentos_manga_pnm", payload)
    return dados[0] if dados else None


def buscar_item_fila_por_apontamento(apontamento_id):
    dados = supabase_get(
        "fila_apontamento_totvs",
        {
            "select": "*",
            "apontamento_id": f"eq.{int(apontamento_id)}",
            "order": "id.desc",
            "limit": "1",
        },
    )
    return dados[0] if dados else None


def criar_item_fila_totvs(apontamento_id, numero_serie, op, tipo_producao, usuario):
    payload = {
        "apontamento_id": int(apontamento_id),
        "numero_serie": _normaliza_codigo(numero_serie),
        "op": _normaliza_codigo(op),
        "tipo_producao": _normaliza_codigo(tipo_producao),
        "usuario": _normaliza_codigo(usuario),
        "data_hora": agora_iso(),
        "status": "pendente",
        "tentativas": 0,
        "ultimo_erro": None,
        "resposta_api": None,
    }
    dados = supabase_post("fila_apontamento_totvs", payload)
    return dados[0] if dados else None


def atualizar_item_fila(fila_id, payload):
    supabase_patch(
        "fila_apontamento_totvs",
        {"id": f"eq.{int(fila_id)}"},
        payload,
    )


def carregar_fila_pendente(limit=10):
    return supabase_get(
        "fila_apontamento_totvs",
        {
            "select": "*",
            "status": "in.(pendente,erro)",
            "order": "id.desc",
            "limit": str(limit),
        },
    )


def garantir_apontamento_e_fila(numero_serie, op, tipo_producao, usuario):
    existente = buscar_apontamento_por_serie(numero_serie)

    if existente:
        fila = buscar_item_fila_por_apontamento(existente["id"])
        if fila:
            return existente, fila, False, "Apontamento já existia no Supabase."
        fila = criar_item_fila_totvs(existente["id"], numero_serie, op, tipo_producao, usuario)
        return existente, fila, False, "Apontamento já existia no Supabase. Fila recriada."

    novo = inserir_apontamento(numero_serie, op, tipo_producao, usuario)
    if not novo:
        raise RuntimeError("Não foi possível criar o apontamento no Supabase.")

    fila = criar_item_fila_totvs(novo["id"], numero_serie, op, tipo_producao, usuario)
    return novo, fila, True, "Apontamento salvo no Supabase."


# =========================
# OCORRÊNCIAS
# =========================
def salvar_ocorrencias_apontamento_totvs(numero_serie, op, tipo_producao, usuario, itens_estoque, resposta_api, tentativas):
    numero_serie = _normaliza_codigo(numero_serie)
    op = _normaliza_codigo(op)
    tipo_producao = _normaliza_codigo(tipo_producao)
    usuario = _normaliza_codigo(usuario)

    if not itens_estoque:
        return

    for item in itens_estoque:
        codigo = _normaliza_codigo(item.get("produto"))
        ocorrencia = normalizar_texto(item.get("ocorrencia") or item.get("motivo"))

        if not codigo or not ocorrencia:
            continue

        existente = supabase_get(
            "ocorrencias_apontamento_totvs",
            {
                "select": "id",
                "numero_serie": f"eq.{numero_serie}",
                "op": f"eq.{op}",
                "codigo": f"eq.{codigo}",
                "ocorrencia": f"eq.{ocorrencia}",
                "limit": "1",
            },
        )

        if existente:
            continue

        payload = {
            "numero_serie": numero_serie,
            "op": op,
            "codigo": codigo,
            "quantidade": item.get("saldo_num"),
            "tipo_producao": tipo_producao,
            "usuario": usuario,
            "data_hora": agora_iso(),
            "armazem": item.get("armazem"),
            "ocorrencia": ocorrencia,
            "tentativas": int(tentativas),
            "resposta_api": resposta_api,
        }

        try:
            supabase_post("ocorrencias_apontamento_totvs", payload)
        except Exception:
            pass


# =========================
# TOTVS
# =========================
def apontar_op_totvs(op: str, lotectl: str, username: str, password: str):
    if not username or not password:
        return False, {"erro": "Informe usuário e senha do TOTVS."}

    payload = {
        "quant": 1,
        "lotectl": lotectl if lotectl else " ",
        "op": op,
    }

    url = f"{TOTVS_API_BASE}/NEW"

    try:
        response = requests.post(
            url,
            json=payload,
            headers=headers_totvs(),
            auth=HTTPBasicAuth(username, password),
            timeout=TOTVS_TIMEOUT
        )

        interpretado = interpretar_retorno_totvs(response)

        retorno = {
            "url": url,
            "payload": payload,
            "status_code": interpretado["status_code"],
            "headers": interpretado["headers"],
            "body": interpretado["body_raw"],
            "body_json": interpretado["body_json"],
            "mensagem_amigavel": interpretado["mensagem_amigavel"],
            "mensagem_tela": interpretado["mensagem_tela"],
            "itens_estoque": interpretado["itens_estoque"],
        }

        return interpretado["sucesso"], retorno

    except Exception as e:
        return False, {
            "url": url,
            "payload": payload,
            "erro": f"Falha ao chamar API TOTVS: {e}"
        }


def executar_reenvio_totvs(fila_item, username, password):
    fila_id = int(fila_item["id"])
    numero_serie = _normaliza_codigo(fila_item.get("numero_serie"))
    op = _normaliza_codigo(fila_item.get("op"))
    tipo_producao = _normaliza_codigo(fila_item.get("tipo_producao"))
    usuario = _normaliza_codigo(fila_item.get("usuario")) or _normaliza_codigo(username)
    tentativas = int(fila_item.get("tentativas") or 0) + 1

    sucesso, retorno = apontar_op_totvs(op, numero_serie, username, password)

    if retorno.get("erro"):
        atualizar_item_fila(
            fila_id,
            {
                "status": "erro",
                "tentativas": tentativas,
                "ultimo_erro": retorno["erro"],
                "resposta_api": retorno["erro"],
            },
        )
        return {
            "supabase_status": "OK",
            "totvs_status": "ERRO",
            "texto": f"SUPABASE: OK\nTOTVS: ERRO\n\n{retorno['erro']}",
            "numero_serie": numero_serie,
            "op": op,
            "tipo": tipo_producao,
        }

    resposta_json = json.dumps(
        {
            "url": retorno.get("url"),
            "payload": retorno.get("payload"),
            "status_code": retorno.get("status_code"),
            "headers": retorno.get("headers"),
            "body": retorno.get("body"),
            "body_json": retorno.get("body_json"),
            "mensagem_amigavel": retorno.get("mensagem_amigavel"),
            "mensagem_tela": retorno.get("mensagem_tela"),
            "itens_estoque": retorno.get("itens_estoque", []),
        },
        ensure_ascii=False,
        indent=2,
    )

    if sucesso:
        atualizar_item_fila(
            fila_id,
            {
                "status": "enviado",
                "tentativas": tentativas,
                "ultimo_erro": None,
                "resposta_api": resposta_json,
            },
        )
        return {
            "supabase_status": "OK",
            "totvs_status": "OK",
            "texto": (
                "SUPABASE: OK\n"
                "TOTVS: OK\n"
                f"HTTP TOTVS: {retorno.get('status_code')}\n\n"
                f"Retorno TOTVS:\n{retorno.get('mensagem_tela') or retorno.get('mensagem_amigavel') or 'OK'}"
            ),
            "numero_serie": numero_serie,
            "op": op,
            "tipo": tipo_producao,
        }

    atualizar_item_fila(
        fila_id,
        {
            "status": "erro",
            "tentativas": tentativas,
            "ultimo_erro": retorno.get("mensagem_amigavel") or "Erro no TOTVS",
            "resposta_api": resposta_json,
        },
    )

    itens = retorno.get("itens_estoque", [])
    if itens:
        salvar_ocorrencias_apontamento_totvs(
            numero_serie=numero_serie,
            op=op,
            tipo_producao=tipo_producao,
            usuario=usuario,
            itens_estoque=itens,
            resposta_api=resposta_json,
            tentativas=tentativas,
        )

    return {
        "supabase_status": "OK",
        "totvs_status": "ERRO",
        "texto": (
            "SUPABASE: OK\n"
            "TOTVS: ERRO\n"
            f"HTTP TOTVS: {retorno.get('status_code')}\n\n"
            f"Retorno TOTVS:\n{retorno.get('mensagem_tela') or retorno.get('mensagem_amigavel') or 'Erro'}"
        ),
        "numero_serie": numero_serie,
        "op": op,
        "tipo": tipo_producao,
    }


# =========================
# UI BASE
# =========================
class BaseScreen(Screen):
    def on_touch_down(self, touch):
        app = App.get_running_app()
        if app and hasattr(app, "register_activity"):
            app.register_activity()
        return super().on_touch_down(touch)


class Card(BoxLayout):
    def __init__(self, bg=(0.10, 0.12, 0.16, 1), **kwargs):
        super().__init__(**kwargs)
        with self.canvas.before:
            self._color = Color(*bg)
            self._rect = Rectangle(pos=self.pos, size=self.size)
        self.bind(pos=self._update_rect, size=self._update_rect)

    def _update_rect(self, *_):
        self._rect.pos = self.pos
        self._rect.size = self.size


class StyledInput(TextInput):
    def __init__(self, hint="", password=False, **kwargs):
        super().__init__(
            multiline=False,
            hint_text=hint,
            password=password,
            size_hint_y=None,
            height=dp(42),
            padding=[dp(10), dp(10), dp(10), dp(10)],
            foreground_color=(1, 1, 1, 1),
            hint_text_color=(0.70, 0.74, 0.80, 1),
            background_color=(0.16, 0.19, 0.24, 1),
            cursor_color=(1, 1, 1, 1),
            background_normal="",
            background_active="",
            **kwargs
        )


class StyledButton(Button):
    def __init__(self, text="", primary=True, **kwargs):
        bg = (0.08, 0.55, 0.78, 1) if primary else (0.22, 0.26, 0.32, 1)
        super().__init__(
            text=text,
            size_hint_y=None,
            height=dp(44),
            background_normal="",
            background_down="",
            background_color=bg,
            color=(1, 1, 1, 1),
            **kwargs
        )


class SmallButton(Button):
    def __init__(self, text="", **kwargs):
        super().__init__(
            text=text,
            size_hint_x=None,
            width=dp(92),
            size_hint_y=None,
            height=dp(34),
            background_normal="",
            background_down="",
            background_color=(0.08, 0.55, 0.78, 1),
            color=(1, 1, 1, 1),
            **kwargs
        )


class FieldRow(BoxLayout):
    def __init__(self, titulo, campo, **kwargs):
        super().__init__(orientation="horizontal", size_hint_y=None, height=dp(42), spacing=dp(8), **kwargs)
        lbl = Label(text=titulo, size_hint_x=0.26, color=(1, 1, 1, 1), halign="left", valign="middle")
        lbl.bind(size=lambda inst, val: setattr(inst, "text_size", val))
        self.add_widget(lbl)
        self.add_widget(campo)


# =========================
# TELAS
# =========================
class LoginScreen(BaseScreen):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)

        outer = BoxLayout(orientation="vertical", padding=dp(14), spacing=dp(12))
        self.add_widget(outer)

        header = Card(orientation="vertical", size_hint_y=None, height=dp(100), padding=dp(16), spacing=dp(4), bg=(0.08, 0.11, 0.15, 1))
        outer.add_widget(header)
        header.add_widget(Label(text="Apontamento TOTVS", font_size="24sp", color=(1, 1, 1, 1)))
        header.add_widget(Label(text="Login TOTVS", font_size="15sp", color=(0.80, 0.86, 0.92, 1)))

        form = Card(orientation="vertical", padding=dp(16), spacing=dp(10), bg=(0.10, 0.12, 0.16, 1))
        outer.add_widget(form)

        self.totvs_user = StyledInput("Usuário TOTVS")
        self.totvs_pass = StyledInput("Senha TOTVS", password=True)

        form.add_widget(FieldRow("Usuário", self.totvs_user))
        form.add_widget(FieldRow("Senha", self.totvs_pass))

        self.status = Label(text="", size_hint_y=None, height=dp(24), color=(1, 0.75, 0.65, 1))
        form.add_widget(self.status)

        btn = StyledButton("Entrar", primary=True)
        btn.bind(on_release=lambda *_: self.entrar())
        form.add_widget(btn)

        outer.add_widget(Label())

    def show_status(self, msg):
        self.status.text = msg

    def entrar(self):
        user = normalizar_texto(self.totvs_user.text)
        password = normalizar_texto(self.totvs_pass.text)

        if not user or not password:
            self.status.text = "Informe usuário e senha do TOTVS."
            return

        app = App.get_running_app()
        app.totvs_user = user
        app.totvs_pass = password
        app.register_activity()
        self.status.text = ""
        self.manager.current = "apontamento"


class ApontamentoScreen(BaseScreen):
    busy = BooleanProperty(False)

    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.history_cells = []
        self.pending_box = None

        outer = BoxLayout(orientation="vertical", padding=dp(10), spacing=dp(10))
        self.add_widget(outer)

        topo = Card(orientation="horizontal", size_hint_y=None, height=dp(64), padding=dp(12), spacing=dp(8), bg=(0.08, 0.11, 0.15, 1))
        outer.add_widget(topo)

        left = BoxLayout(orientation="vertical")
        self.lbl_title = Label(text="Apontamento TOTVS", font_size="22sp", color=(1, 1, 1, 1), halign="left", valign="middle")
        self.lbl_title.bind(size=lambda inst, val: setattr(inst, "text_size", val))
        self.lbl_user = Label(text="", font_size="13sp", color=(0.80, 0.86, 0.92, 1), halign="left", valign="middle")
        self.lbl_user.bind(size=lambda inst, val: setattr(inst, "text_size", val))
        left.add_widget(self.lbl_title)
        left.add_widget(self.lbl_user)

        btn_sair = StyledButton("Sair", primary=False, size_hint_x=None, width=dp(86))
        btn_sair.bind(on_release=lambda *_: self.logout())

        topo.add_widget(left)
        topo.add_widget(btn_sair)

        scroll = ScrollView(do_scroll_x=False)
        outer.add_widget(scroll)

        content = BoxLayout(orientation="vertical", size_hint_y=None, spacing=dp(10), padding=[0, 0, 0, dp(10)])
        content.bind(minimum_height=content.setter("height"))
        scroll.add_widget(content)

        form_card = Card(orientation="vertical", padding=dp(14), spacing=dp(10), size_hint_y=None, bg=(0.10, 0.12, 0.16, 1))
        form_card.bind(minimum_height=form_card.setter("height"))
        content.add_widget(form_card)

        self.tipo = Spinner(
            text="MANGA",
            values=["MANGA", "PNM"],
            size_hint_y=None,
            height=dp(42),
            background_normal="",
            background_color=(0.16, 0.19, 0.24, 1),
            color=(1, 1, 1, 1),
        )
        self.op = StyledInput("Digite a OP")
        self.numero_serie = StyledInput("Digite o número de série")

        form_card.add_widget(FieldRow("Tipo", self.tipo))
        form_card.add_widget(FieldRow("OP", self.op))
        form_card.add_widget(FieldRow("Número de série", self.numero_serie))

        botoes = BoxLayout(size_hint_y=None, height=dp(44), spacing=dp(8))
        self.btn_apontar = StyledButton("Apontar", primary=True)
        btn_limpar = StyledButton("Limpar", primary=False)
        self.btn_apontar.bind(on_release=lambda *_: self.on_apontar())
        btn_limpar.bind(on_release=lambda *_: self.on_limpar())
        botoes.add_widget(self.btn_apontar)
        botoes.add_widget(btn_limpar)
        form_card.add_widget(botoes)

        status_card = Card(orientation="vertical", padding=dp(14), spacing=dp(8), size_hint_y=None, height=dp(190), bg=(0.10, 0.12, 0.16, 1))
        content.add_widget(status_card)

        self.lbl_status_title = Label(text="Concluído", size_hint_y=None, height=dp(22), color=(0.85, 0.90, 1, 1), halign="left", valign="middle")
        self.lbl_status_title.bind(size=lambda inst, val: setattr(inst, "text_size", val))
        status_card.add_widget(self.lbl_status_title)

        self.resposta = TextInput(
            readonly=True,
            multiline=True,
            text="Pronto para apontar.",
            foreground_color=(1, 1, 1, 1),
            background_color=(0.16, 0.19, 0.24, 1),
            background_normal="",
            cursor_color=(1, 1, 1, 1),
        )
        status_card.add_widget(self.resposta)

        hist_card = Card(orientation="vertical", padding=dp(14), spacing=dp(8), size_hint_y=None, bg=(0.10, 0.12, 0.16, 1))
        hist_card.bind(minimum_height=hist_card.setter("height"))
        content.add_widget(hist_card)

        hist_title = Label(text="Últimos 5 apontamentos", size_hint_y=None, height=dp(24), color=(1, 1, 1, 1), halign="left", valign="middle")
        hist_title.bind(size=lambda inst, val: setattr(inst, "text_size", val))
        hist_card.add_widget(hist_title)

        grid = GridLayout(cols=6, size_hint_y=None, spacing=dp(4))
        grid.bind(minimum_height=grid.setter("height"))
        hist_card.add_widget(grid)

        headers = ["Hora", "Tipo", "OP", "Série", "SUPABASE", "TOTVS"]
        for h in headers:
            lbl = Label(text=h, size_hint_y=None, height=dp(24), color=(0.82, 0.90, 1, 1))
            grid.add_widget(lbl)

        for _ in range(5):
            row = []
            for __ in range(6):
                lbl = Label(text="-", size_hint_y=None, height=dp(24), color=(1, 1, 1, 1))
                grid.add_widget(lbl)
                row.append(lbl)
            self.history_cells.append(row)

        pend_card = Card(orientation="vertical", padding=dp(14), spacing=dp(8), size_hint_y=None, bg=(0.10, 0.12, 0.16, 1))
        pend_card.bind(minimum_height=pend_card.setter("height"))
        content.add_widget(pend_card)

        pend_title = Label(text="Pendentes / erro para reenviar TOTVS", size_hint_y=None, height=dp(24), color=(1, 1, 1, 1), halign="left", valign="middle")
        pend_title.bind(size=lambda inst, val: setattr(inst, "text_size", val))
        pend_card.add_widget(pend_title)

        self.pending_box = BoxLayout(orientation="vertical", size_hint_y=None, spacing=dp(6))
        self.pending_box.bind(minimum_height=self.pending_box.setter("height"))
        pend_card.add_widget(self.pending_box)

    def on_pre_enter(self, *args):
        app = App.get_running_app()
        self.lbl_user.text = f"Usuário TOTVS: {app.totvs_user or '-'}"
        self.refresh_history()
        self.refresh_pending()

    def logout(self):
        app = App.get_running_app()
        app.force_logout("Sessão encerrada.")

    def set_status(self, texto):
        self.resposta.text = texto

    def on_limpar(self):
        self.tipo.text = "MANGA"
        self.op.text = ""
        self.numero_serie.text = ""
        self.set_status("Campos limpos.")

    def on_apontar(self):
        if self.busy:
            return

        app = App.get_running_app()

        tipo = normalizar_texto(self.tipo.text)
        op = _normaliza_codigo(self.op.text)
        numero_serie = _normaliza_codigo(self.numero_serie.text)
        user = normalizar_texto(app.totvs_user)
        password = normalizar_texto(app.totvs_pass)

        if not user or not password:
            self.set_status("Faça login novamente.")
            self.manager.current = "login"
            return

        if not op:
            self.set_status("Informe a OP.")
            return

        if not numero_serie:
            self.set_status("Informe o número de série.")
            return

        app.register_activity()
        self.busy = True
        self.btn_apontar.disabled = True
        self.set_status("Salvando no Supabase / fila e enviando ao TOTVS...")

        dados = {
            "tipo": tipo,
            "op": op,
            "numero_serie": numero_serie,
            "usuario_totvs": user,
            "senha_totvs": password,
        }

        threading.Thread(target=self._worker_apontar, args=(dados,), daemon=True).start()

    def _worker_apontar(self, dados: dict):
        try:
            _, fila_item, _, msg_supa = garantir_apontamento_e_fila(
                numero_serie=dados["numero_serie"],
                op=dados["op"],
                tipo_producao=dados["tipo"],
                usuario=dados["usuario_totvs"],
            )

            if not fila_item:
                texto = "SUPABASE: ERRO\nTOTVS: -\n\nFalha ao criar ou localizar fila TOTVS."
                historico = {
                    "hora": hora_agora(),
                    "tipo": dados["tipo"],
                    "op": dados["op"],
                    "serie": dados["numero_serie"],
                    "supabase": "ERRO",
                    "totvs": "-",
                }
                Clock.schedule_once(lambda dt: self._finalizar(texto, historico), 0)
                return

            if str(fila_item.get("status", "")).lower() == "enviado":
                texto = "SUPABASE: OK\nTOTVS: OK\n\nEste apontamento já foi enviado ao TOTVS anteriormente."
                historico = {
                    "hora": hora_agora(),
                    "tipo": dados["tipo"],
                    "op": dados["op"],
                    "serie": dados["numero_serie"],
                    "supabase": "OK",
                    "totvs": "OK",
                }
                Clock.schedule_once(lambda dt: self._finalizar(texto, historico), 0)
                return

            resultado = executar_reenvio_totvs(
                fila_item=fila_item,
                username=dados["usuario_totvs"],
                password=dados["senha_totvs"],
            )

            texto = f"{resultado['texto']}\n\nSupabase: {msg_supa}"
            historico = {
                "hora": hora_agora(),
                "tipo": resultado["tipo"],
                "op": resultado["op"],
                "serie": resultado["numero_serie"],
                "supabase": resultado["supabase_status"],
                "totvs": resultado["totvs_status"],
            }
            Clock.schedule_once(lambda dt: self._finalizar(texto, historico), 0)

        except Exception as e:
            texto = f"SUPABASE: ERRO\nTOTVS: -\n\nFalha geral: {e}"
            historico = {
                "hora": hora_agora(),
                "tipo": dados["tipo"],
                "op": dados["op"],
                "serie": dados["numero_serie"],
                "supabase": "ERRO",
                "totvs": "-",
            }
            Clock.schedule_once(lambda dt: self._finalizar(texto, historico), 0)

    def on_reenviar_click(self, fila_item):
        if self.busy:
            return

        app = App.get_running_app()
        user = normalizar_texto(app.totvs_user)
        password = normalizar_texto(app.totvs_pass)

        if not user or not password:
            self.set_status("Faça login novamente.")
            self.manager.current = "login"
            return

        app.register_activity()
        self.busy = True
        self.btn_apontar.disabled = True
        self.set_status(f"Reenviando item {fila_item.get('id')} para o TOTVS...")

        threading.Thread(
            target=self._worker_reenviar,
            args=(fila_item, user, password),
            daemon=True
        ).start()

    def _worker_reenviar(self, fila_item, user, password):
        try:
            resultado = executar_reenvio_totvs(fila_item, user, password)

            historico = {
                "hora": hora_agora(),
                "tipo": resultado["tipo"],
                "op": resultado["op"],
                "serie": resultado["numero_serie"],
                "supabase": resultado["supabase_status"],
                "totvs": resultado["totvs_status"],
            }

            Clock.schedule_once(lambda dt: self._finalizar(resultado["texto"], historico), 0)

        except Exception as e:
            texto = f"SUPABASE: OK\nTOTVS: ERRO\n\nFalha ao reenviar: {e}"
            historico = {
                "hora": hora_agora(),
                "tipo": _normaliza_codigo(fila_item.get("tipo_producao")),
                "op": _normaliza_codigo(fila_item.get("op")),
                "serie": _normaliza_codigo(fila_item.get("numero_serie")),
                "supabase": "OK",
                "totvs": "ERRO",
            }
            Clock.schedule_once(lambda dt: self._finalizar(texto, historico), 0)

    def _finalizar(self, texto_final: str, historico: dict):
        app = App.get_running_app()
        self.busy = False
        self.btn_apontar.disabled = False
        self.set_status(texto_final)

        app.history.insert(0, historico)
        app.history = app.history[:5]
        self.refresh_history()
        self.refresh_pending()

    def refresh_history(self):
        app = App.get_running_app()

        for row in self.history_cells:
            for cell in row:
                cell.text = "-"
                cell.color = (1, 1, 1, 1)

        for idx, item in enumerate(app.history[:5]):
            row = self.history_cells[idx]
            row[0].text = item["hora"]
            row[1].text = item["tipo"]
            row[2].text = item["op"]
            row[3].text = item["serie"]
            row[4].text = item["supabase"]
            row[5].text = item["totvs"]

            row[4].color = (0.30, 1, 0.40, 1) if item["supabase"] == "OK" else ((1, 0.40, 0.40, 1) if item["supabase"] == "ERRO" else (1, 1, 1, 1))
            row[5].color = (0.30, 1, 0.40, 1) if item["totvs"] == "OK" else ((1, 0.40, 0.40, 1) if item["totvs"] == "ERRO" else (1, 1, 1, 1))

    def refresh_pending(self):
        self.pending_box.clear_widgets()

        try:
            pendentes = carregar_fila_pendente(limit=10)
        except Exception as e:
            self.pending_box.add_widget(Label(text=f"Erro ao carregar pendentes: {e}", size_hint_y=None, height=dp(24), color=(1, 0.45, 0.45, 1)))
            return

        if not pendentes:
            self.pending_box.add_widget(Label(text="Nenhum item pendente ou com erro.", size_hint_y=None, height=dp(24), color=(0.75, 0.82, 0.90, 1)))
            return

        for item in pendentes:
            status = str(item.get("status", "")).upper()
            numero_serie = _normaliza_codigo(item.get("numero_serie"))
            op = _normaliza_codigo(item.get("op"))
            tipo = _normaliza_codigo(item.get("tipo_producao"))
            tentativas = int(item.get("tentativas") or 0)

            row = BoxLayout(size_hint_y=None, height=dp(38), spacing=dp(6))
            info = Label(
                text=f"[{status}] {tipo} | OP {op} | Série {numero_serie} | Tent. {tentativas}",
                halign="left",
                valign="middle",
                color=(1, 1, 1, 1),
            )
            info.bind(size=lambda inst, val: setattr(inst, "text_size", val))

            btn = SmallButton(text="Reenviar")
            btn.bind(on_release=lambda _, x=item: self.on_reenviar_click(x))

            row.add_widget(info)
            row.add_widget(btn)
            self.pending_box.add_widget(row)


# =========================
# APP
# =========================
class KivyTotvsApp(App):
    totvs_user = StringProperty("")
    totvs_pass = StringProperty("")

    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.history = []
        self.last_activity = time.monotonic()
        self.sm = None

    def build(self):
        self.title = "Apontamento TOTVS"

        Window.bind(on_key_down=self._on_window_key_down)

        self.sm = ScreenManager()
        self.sm.add_widget(LoginScreen(name="login"))
        self.sm.add_widget(ApontamentoScreen(name="apontamento"))

        Clock.schedule_interval(self._check_inactivity, 1)

        return self.sm

    def _on_window_key_down(self, *args):
        self.register_activity()
        return False

    def register_activity(self):
        self.last_activity = time.monotonic()

    def _check_inactivity(self, dt):
        if not self.sm:
            return

        if self.sm.current == "login":
            return

        if not self.totvs_user or not self.totvs_pass:
            return

        elapsed = time.monotonic() - self.last_activity
        if elapsed >= INACTIVITY_TIMEOUT_SEC:
            self.force_logout("Sessão expirada após 3 minutos de inatividade.")

    def force_logout(self, msg="Sessão encerrada."):
        self.totvs_user = ""
        self.totvs_pass = ""
        self.register_activity()

        if self.sm:
            login = self.sm.get_screen("login")
            login.totvs_pass.text = ""
            login.show_status(msg)
            self.sm.current = "login"


if __name__ == "__main__":
    KivyTotvsApp().run()
