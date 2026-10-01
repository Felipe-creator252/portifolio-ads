"""
Calculadora de Precificação de Impressão 3D — v2
=====================================================
Aplicativo desktop (CustomTkinter, modo escuro) para calcular custo e preço
de venda de peças impressas em 3D, com suporte a:

    - Interface organizada em 3 abas: Calculadora/Produtos, Parâmetros
      Globais e Insumos & Filamentos.
    - Filamentos cadastrados por peso de carretel (250g, 1kg, 2kg, 3kg,
      5kg ou qualquer valor) + cor, com cálculo automático do R$/kg real.
    - Produtos multicoloridos (multi-filamento / CFS / AMS / MMU), com
      peso de purga e custo extra de mão de obra em trocas manuais.
    - Produtos compostos/modulares (BOM), com múltiplos módulos/sub-peças,
      cada um com seu próprio peso, tempo e lista de filamentos.
    - Insumos extras (parafusos, ímãs, embalagem etc.) e preparação de
      mesa por produto.
    - Modelos de margem flexíveis: personalizada, por categoria ou
      markup fixo (multiplicador).
    - Simulador de desconto em lote / atacado.
    - Múltiplas moedas (BRL, USD, EUR) para exibição.
    - Importador de G-code / 3MF (peso e tempo de impressão).
    - Gerador de orçamento em PDF (requer reportlab).
    - Persistência automática em JSON local (dados_calculadora_v2.json),
      com backup de segurança antes de cada gravação.

Dependências:
    pip install customtkinter reportlab

Autor: Gerado com Claude
"""

import customtkinter as ctk
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import csv
import json
import os
import re
import shutil
import sys
import zipfile
import datetime

try:
    from reportlab.lib.pagesizes import A4
    from reportlab.lib import colors
    from reportlab.lib.units import cm
    from reportlab.platypus import (
        SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, HRFlowable
    )
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.enums import TA_RIGHT
    REPORTLAB_DISPONIVEL = True
except ImportError:
    REPORTLAB_DISPONIVEL = False


# ----------------------------------------------------------------------------
# Configurações gerais de aparência
# ----------------------------------------------------------------------------
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

COR_FUNDO = "#1a1a1a"
COR_PAINEL = "#242424"
COR_DESTAQUE = "#2fa572"
COR_TEXTO_SECUNDARIO = "#a0a0a0"
COR_ERRO = "#8b2e2e"
COR_AZUL = "#2f5f8f"

# ----------------------------------------------------------------------------
# Local dos arquivos de dados
# ----------------------------------------------------------------------------
# IMPORTANTE: não usamos os.path.dirname(__file__) para isso. Quando o app é
# empacotado com PyInstaller (--onefile), __file__ aponta para uma pasta
# temporária (_MEIPASS) que é extraída a cada execução e apagada ao fechar —
# ou seja, os dados pareceriam "não salvar" mesmo salvando de fato. Além
# disso, se o .exe estiver em uma pasta protegida (ex: Program Files), a
# escrita pode ser silenciosamente bloqueada ou virtualizada pelo Windows.
#
# Por isso, gravamos sempre numa pasta de dados do usuário, estável entre
# execuções e com permissão de escrita garantida, tanto rodando o .py quanto
# o .exe gerado.
def _obter_pasta_dados():
    nome_pasta_app = "CalculadoraImpressao3D"
    if os.name == "nt":
        base = os.environ.get("APPDATA") or os.path.expanduser("~")
    else:
        base = os.environ.get("XDG_DATA_HOME") or os.path.join(os.path.expanduser("~"), ".local", "share")
    pasta = os.path.join(base, nome_pasta_app)
    os.makedirs(pasta, exist_ok=True)
    return pasta


def _obter_pasta_script():
    """Pasta onde o .py (ou o .exe, quando congelado) realmente está —
    usada apenas para localizar o arquivo de dados da versão 1 (migração),
    que era salvo ao lado do script."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


PASTA_DADOS = _obter_pasta_dados()
ARQUIVO_DADOS = os.path.join(PASTA_DADOS, "dados_calculadora_v2.json")
ARQUIVO_BACKUP = ARQUIVO_DADOS + ".bak"
ARQUIVO_DADOS_V1 = os.path.join(_obter_pasta_script(), "dados_calculadora.json")

CURRENCY_SYMBOLS = {"BRL": "R$", "USD": "US$", "EUR": "€"}
MOEDA_DISPLAY_PARA_CODIGO = {
    "BRL (R$)": "BRL", "USD (US$)": "USD", "EUR (€)": "EUR",
}
MOEDA_CODIGO_PARA_DISPLAY = {v: k for k, v in MOEDA_DISPLAY_PARA_CODIGO.items()}

CATEGORIAS_PADRAO = {"Peça Técnica": 200.0, "Action Figure": 150.0, "Protótipo": 100.0}

# Colunas da tabela principal de produtos
COLUNAS_PRODUTOS = (
    "id", "nome", "categoria", "modulos", "peso", "tempo_imp", "tempo_mo",
    "custo_mat", "custo_energ", "custo_manut", "custo_deprec", "custo_mo",
    "custo_prep", "custo_insumos", "custo_perdas", "custo_total",
    "preco_venda", "lucro",
)
COLUNAS_NUMERICAS = set(COLUNAS_PRODUTOS) - {"nome", "categoria"}
TITULOS_COLUNAS = {
    "id": "ID", "nome": "Nome", "categoria": "Categoria", "modulos": "Módulos",
    "peso": "Peso (g)", "tempo_imp": "T. Impr. (h)", "tempo_mo": "T. M.Obra (h)",
    "custo_mat": "Custo Mat.", "custo_energ": "Custo Energ.",
    "custo_manut": "Custo Manut.", "custo_deprec": "Depreciação",
    "custo_mo": "Mão de Obra", "custo_prep": "Preparo", "custo_insumos": "Insumos",
    "custo_perdas": "Perdas", "custo_total": "Custo Total",
    "preco_venda": "Preço Venda", "lucro": "Lucro",
}
LARGURAS_COLUNAS = {
    "id": 45, "nome": 150, "categoria": 100, "modulos": 70, "peso": 75,
    "tempo_imp": 85, "tempo_mo": 90, "custo_mat": 95, "custo_energ": 95,
    "custo_manut": 95, "custo_deprec": 95, "custo_mo": 95, "custo_prep": 85,
    "custo_insumos": 85, "custo_perdas": 85, "custo_total": 100,
    "preco_venda": 100, "lucro": 90,
}


# ----------------------------------------------------------------------------
# Utilidades numéricas, de moeda e de tempo (funções de módulo)
# ----------------------------------------------------------------------------
def texto_para_float(texto):
    """Converte texto em float aceitando vírgula ou ponto decimal."""
    if texto is None:
        raise ValueError("valor não informado")
    texto_limpo = str(texto).strip()
    if not texto_limpo:
        raise ValueError("campo vazio")
    try:
        if "," in texto_limpo:
            texto_limpo = texto_limpo.replace(".", "").replace(",", ".")
        return float(texto_limpo)
    except (TypeError, ValueError):
        raise ValueError(f"'{texto}' não é um número válido")


def formatar_numero(valor, casas=2):
    """Formata um número simples (peso, tempo, %) no padrão 1.234,56."""
    try:
        texto = f"{float(valor):,.{casas}f}"
    except (TypeError, ValueError):
        return "0,00"
    return texto.replace(",", "X").replace(".", ",").replace("X", ".")


def formatar_moeda(valor, moeda="BRL"):
    """Formata um valor monetário de acordo com a moeda selecionada."""
    simbolo = CURRENCY_SYMBOLS.get(moeda, "R$")
    try:
        valor = float(valor)
    except (TypeError, ValueError):
        valor = 0.0
    if moeda == "USD":
        texto = f"{valor:,.2f}"
    else:
        texto = f"{valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"{simbolo} {texto}"


def valor_numerico_de_celula(texto):
    """Extrai um float aproximado de uma célula formatada (moeda, g, h)
    para permitir a ordenação numérica da tabela."""
    if texto is None:
        return 0.0
    limpo = re.sub(r"[^0-9,.\-]", "", str(texto))
    if not limpo:
        return 0.0
    if "," in limpo and "." in limpo:
        if limpo.rfind(",") > limpo.rfind("."):
            limpo = limpo.replace(".", "").replace(",", ".")
        else:
            limpo = limpo.replace(",", "")
    elif "," in limpo:
        limpo = limpo.replace(".", "").replace(",", ".")
    try:
        return float(limpo)
    except ValueError:
        return 0.0


def _tempo_texto_para_horas(texto):
    """Converte '1h 23m 45s' ou um número puro de segundos em horas."""
    texto = texto.strip()
    if re.fullmatch(r"\d+", texto):
        return float(texto) / 3600.0
    total_seg = 0
    for valor, unidade in re.findall(r"(\d+)\s*([dhms])", texto, re.IGNORECASE):
        valor = int(valor)
        unidade = unidade.lower()
        if unidade == "d":
            total_seg += valor * 86400
        elif unidade == "h":
            total_seg += valor * 3600
        elif unidade == "m":
            total_seg += valor * 60
        elif unidade == "s":
            total_seg += valor
    if total_seg == 0:
        raise ValueError("tempo não reconhecido")
    return total_seg / 3600.0


PADROES_PESO = [
    r"filament\s*used\s*\[g\]\s*=\s*([0-9.,\s]+)",
    r";\s*total\s*filament\s*used\s*\[g\]\s*:?=?\s*([0-9.,]+)",
    r";\s*filament_used_g\s*=\s*([0-9.,]+)",
    r";\s*Filament\s*used\s*:\s*([0-9.,]+)\s*g",
    r"total\s*filament\s*weight\s*\[g\]\s*:?=?\s*([0-9.,]+)",
]
PADROES_TEMPO = [
    r"estimated\s*printing\s*time[^=]*=\s*([0-9dhms\s]+)",
    r";\s*model\s*printing\s*time\s*:\s*([0-9dhms\s]+)",
    r";TIME:(\d+)",
    r";\s*Print\s*time\s*:?=?\s*([0-9dhms\s]+)",
]


def extrair_dados_arquivo_fatiado(caminho):
    """Tenta extrair peso (g) e tempo de impressão (h) de um arquivo
    .gcode ou .3mf exportado por fatiadores comuns (Creality Print,
    OrcaSlicer, PrusaSlicer, Cura/Bambu Studio). Retorna (dados, erro),
    onde dados é um dict {"peso_g":, "tempo_h":} (chaves podem ser None
    se não encontradas) ou None se a leitura falhou completamente."""
    ext = os.path.splitext(caminho)[1].lower()
    conteudo_total = ""
    try:
        if ext == ".3mf":
            with zipfile.ZipFile(caminho) as z:
                for nome in z.namelist():
                    if nome.lower().endswith((".gcode", ".config", ".json", ".xml", ".md5")):
                        try:
                            conteudo_total += z.read(nome).decode("utf-8", errors="ignore") + "\n"
                        except Exception:
                            continue
        else:
            with open(caminho, "r", encoding="utf-8", errors="ignore") as arquivo:
                # Metadados de tempo/peso costumam ficar no início ou fim do
                # arquivo; para não estourar memória em gcodes gigantes,
                # lemos o arquivo inteiro apenas se ele for razoavelmente
                # pequeno, senão lemos início + fim.
                conteudo_total = arquivo.read(4_000_000)
    except Exception as e:
        return None, str(e)

    if not conteudo_total:
        return None, "arquivo vazio ou ilegível"

    peso = None
    for padrao in PADROES_PESO:
        m = re.search(padrao, conteudo_total, re.IGNORECASE)
        if m:
            try:
                partes = [float(x.replace(",", ".")) for x in re.split(r"[,\s]+", m.group(1).strip()) if x]
                if partes:
                    peso = sum(partes)
                    break
            except ValueError:
                continue

    tempo = None
    for padrao in PADROES_TEMPO:
        m = re.search(padrao, conteudo_total, re.IGNORECASE)
        if m:
            try:
                tempo = _tempo_texto_para_horas(m.group(1))
                break
            except ValueError:
                continue

    if peso is None and tempo is None:
        return None, "não foi possível localizar peso ou tempo no arquivo"
    return {"peso_g": peso, "tempo_h": tempo}, None


class CampoInvalidoError(Exception):
    """Erro de validação que carrega o nome do campo com problema."""
    def __init__(self, nome_campo, motivo):
        self.nome_campo = nome_campo
        self.motivo = motivo
        super().__init__(f"{nome_campo}: {motivo}")


# ----------------------------------------------------------------------------
# Painel reutilizável para editar um "módulo" (peça única de um produto
# simples OU uma sub-peça de um produto composto/modular)
# ----------------------------------------------------------------------------
class EditorModuloFrame(ctk.CTkFrame):
    def __init__(self, master, app_ref, mostrar_nome=False, **kwargs):
        super().__init__(master, fg_color="transparent", **kwargs)
        self.app_ref = app_ref
        self.mostrar_nome = mostrar_nome
        self._itens_filamento = []  # [{"filamento_id":, "peso_g":}, ...]
        self._construir()

    def _construir(self):
        if self.mostrar_nome:
            ctk.CTkLabel(self, text="Nome do Módulo", text_color=COR_TEXTO_SECUNDARIO,
                         font=ctk.CTkFont(size=12)).pack(anchor="w", pady=(0, 4))
            self.entrada_nome_modulo = ctk.CTkEntry(self, placeholder_text="Ex: Base")
            self.entrada_nome_modulo.pack(fill="x", pady=(0, 10))

        linha1 = ctk.CTkFrame(self, fg_color="transparent")
        linha1.pack(fill="x", pady=(0, 10))
        linha1.grid_columnconfigure((0, 1), weight=1)

        col_a = ctk.CTkFrame(linha1, fg_color="transparent")
        col_a.grid(row=0, column=0, sticky="ew", padx=(0, 6))
        ctk.CTkLabel(col_a, text="Tempo de Impressão (h)", text_color=COR_TEXTO_SECUNDARIO,
                     font=ctk.CTkFont(size=12)).pack(anchor="w")
        self.entrada_tempo_imp = ctk.CTkEntry(col_a, placeholder_text="Ex: 3,5")
        self.entrada_tempo_imp.pack(fill="x", pady=(4, 0))

        col_b = ctk.CTkFrame(linha1, fg_color="transparent")
        col_b.grid(row=0, column=1, sticky="ew", padx=(6, 0))
        ctk.CTkLabel(col_b, text="Tempo de Mão de Obra (h)", text_color=COR_TEXTO_SECUNDARIO,
                     font=ctk.CTkFont(size=12)).pack(anchor="w")
        self.entrada_tempo_mo = ctk.CTkEntry(col_b, placeholder_text="Ex: 0,5")
        self.entrada_tempo_mo.pack(fill="x", pady=(4, 0))

        linha2 = ctk.CTkFrame(self, fg_color="transparent")
        linha2.pack(fill="x", pady=(0, 6))
        linha2.grid_columnconfigure((0, 1), weight=1)

        col_c = ctk.CTkFrame(linha2, fg_color="transparent")
        col_c.grid(row=0, column=0, sticky="ew", padx=(0, 6))
        ctk.CTkLabel(col_c, text="Peso de Purga / Torre (g)", text_color=COR_TEXTO_SECUNDARIO,
                     font=ctk.CTkFont(size=12)).pack(anchor="w")
        self.entrada_peso_purga = ctk.CTkEntry(col_c, placeholder_text="Ex: 8")
        self.entrada_peso_purga.pack(fill="x", pady=(4, 0))

        col_d = ctk.CTkFrame(linha2, fg_color="transparent")
        col_d.grid(row=0, column=1, sticky="ew", padx=(6, 0))
        ctk.CTkLabel(col_d, text="Troca de Cor", text_color=COR_TEXTO_SECUNDARIO,
                     font=ctk.CTkFont(size=12)).pack(anchor="w")
        self.combo_tipo_troca = ctk.CTkComboBox(
            col_d, values=["Automática (CFS/AMS)", "Manual (Pausa no G-code)"],
            state="readonly", command=self._ao_mudar_tipo_troca,
        )
        self.combo_tipo_troca.pack(fill="x", pady=(4, 0))

        self.frame_tempo_troca = ctk.CTkFrame(self, fg_color="transparent")
        ctk.CTkLabel(self.frame_tempo_troca, text="Tempo por Troca Manual (h)",
                     text_color=COR_TEXTO_SECUNDARIO, font=ctk.CTkFont(size=12)).pack(anchor="w")
        self.entrada_tempo_troca = ctk.CTkEntry(self.frame_tempo_troca, placeholder_text="Ex: 0,05 (3 min)")
        self.entrada_tempo_troca.pack(fill="x", pady=(4, 0))
        self.frame_tempo_troca.pack(fill="x", pady=(6, 6))

        btn_gcode = ctk.CTkButton(
            self, text="📥  Importar Peso/Tempo de G-code ou 3MF",
            command=self.importar_gcode, fg_color=COR_AZUL, hover_color="#3a72ab",
        )
        btn_gcode.pack(fill="x", pady=(2, 10))

        ctk.CTkLabel(self, text="Filamentos / Cores Utilizados", font=ctk.CTkFont(size=13, weight="bold")
                     ).pack(anchor="w", pady=(0, 6))

        self.mini_tabela_filamentos = ttk.Treeview(
            self, columns=("filamento", "peso"), show="headings",
            style="Custom.Treeview", selectmode="extended", height=4,
        )
        self.mini_tabela_filamentos.heading("filamento", text="Filamento (Cor)")
        self.mini_tabela_filamentos.heading("peso", text="Peso Usado")
        self.mini_tabela_filamentos.column("filamento", width=220, anchor="w")
        self.mini_tabela_filamentos.column("peso", width=100, anchor="center")
        self.mini_tabela_filamentos.pack(fill="x", pady=(0, 6))

        linha_add = ctk.CTkFrame(self, fg_color="transparent")
        linha_add.pack(fill="x", pady=(0, 4))
        self.combo_filamento_uso = ctk.CTkComboBox(linha_add, values=[], state="readonly")
        self.combo_filamento_uso.pack(side="left", fill="x", expand=True, padx=(0, 6))
        self.entrada_peso_filamento_uso = ctk.CTkEntry(linha_add, placeholder_text="Peso (g)", width=90)
        self.entrada_peso_filamento_uso.pack(side="left", padx=(0, 6))
        ctk.CTkButton(linha_add, text="+", width=32, command=self._adicionar_item_filamento,
                      fg_color=COR_DESTAQUE, hover_color="#268a5e").pack(side="left")

        ctk.CTkButton(self, text="🗑️  Remover Filamento Selecionado do Módulo",
                      command=self._remover_item_filamento, fg_color=COR_ERRO,
                      hover_color="#a83a3a").pack(fill="x", pady=(2, 4))

        self._ao_mudar_tipo_troca()

    # -- filamentos usados no módulo -------------------------------------
    def _adicionar_item_filamento(self):
        display = self.combo_filamento_uso.get()
        mapa = self.app_ref.mapa_display_filamento()
        if display not in mapa:
            messagebox.showwarning("Filamento", "Cadastre e selecione um filamento válido.")
            return
        try:
            peso = self.app_ref._ler_campo_numerico(
                self.entrada_peso_filamento_uso.get(), "Peso do filamento no módulo",
                permitir_zero=False, minimo=0.001,
            )
        except CampoInvalidoError as e:
            messagebox.showerror("Valor inválido", str(e))
            return
        self._itens_filamento.append({"filamento_id": mapa[display], "peso_g": peso})
        self.entrada_peso_filamento_uso.delete(0, "end")
        self._atualizar_mini_tabela_filamentos()

    def _remover_item_filamento(self):
        selecao = self.mini_tabela_filamentos.selection()
        if not selecao:
            return
        indices = sorted((int(i) for i in selecao), reverse=True)
        for i in indices:
            if 0 <= i < len(self._itens_filamento):
                self._itens_filamento.pop(i)
        self._atualizar_mini_tabela_filamentos()

    def _atualizar_mini_tabela_filamentos(self):
        for item in self.mini_tabela_filamentos.get_children(""):
            self.mini_tabela_filamentos.delete(item)
        for idx, item in enumerate(self._itens_filamento):
            fil = self.app_ref.filamentos.get(item["filamento_id"])
            if fil:
                nome_cor = f"{fil['nome']} ({fil.get('cor', '-')})" if fil.get("cor") else fil["nome"]
            else:
                nome_cor = "Filamento removido"
            self.mini_tabela_filamentos.insert(
                "", "end", iid=str(idx), values=(nome_cor, f"{formatar_numero(item['peso_g'])} g")
            )

    def atualizar_opcoes_filamento(self):
        mapa = self.app_ref.mapa_display_filamento()
        valores = list(mapa.keys())
        self.combo_filamento_uso.configure(values=valores)
        if valores and self.combo_filamento_uso.get() not in valores:
            self.combo_filamento_uso.set(valores[0])
        elif not valores:
            self.combo_filamento_uso.set("")

    def _ao_mudar_tipo_troca(self, valor=None):
        if self.combo_tipo_troca.get().startswith("Manual"):
            self.frame_tempo_troca.pack(fill="x", pady=(6, 6))
        else:
            self.frame_tempo_troca.pack_forget()

    def importar_gcode(self):
        caminho = filedialog.askopenfilename(
            title="Selecionar arquivo do fatiador",
            filetypes=[("G-code / 3MF", "*.gcode *.g *.3mf"), ("Todos os arquivos", "*.*")],
        )
        if not caminho:
            return
        dados, erro = extrair_dados_arquivo_fatiado(caminho)
        if dados is None:
            messagebox.showwarning(
                "Importar", f"Não foi possível extrair os dados automaticamente ({erro}).\n"
                            "Preencha manualmente."
            )
            return
        mensagens = []
        if dados.get("tempo_h") is not None:
            self.entrada_tempo_imp.delete(0, "end")
            self.entrada_tempo_imp.insert(0, formatar_numero(dados["tempo_h"]))
            mensagens.append(f"Tempo de impressão: {formatar_numero(dados['tempo_h'])} h")
        if dados.get("peso_g") is not None:
            self.entrada_peso_filamento_uso.delete(0, "end")
            self.entrada_peso_filamento_uso.insert(0, formatar_numero(dados["peso_g"]))
            mensagens.append(
                f"Peso identificado: {formatar_numero(dados['peso_g'])} g "
                "(selecione o filamento e clique em '+' para adicioná-lo)"
            )
        if mensagens:
            messagebox.showinfo("Importado", "\n".join(mensagens))
        else:
            messagebox.showwarning("Importar", "Nenhum dado reconhecido no arquivo. Preencha manualmente.")

    # -- leitura / preenchimento / limpeza --------------------------------
    def obter_dados(self):
        rotulo_grupo = "Módulo" if self.mostrar_nome else "Produto"
        nome = self.entrada_nome_modulo.get().strip() if self.mostrar_nome else "Peça Única"
        if self.mostrar_nome and not nome:
            raise CampoInvalidoError("Nome do Módulo", "é obrigatório")

        tempo_imp = self.app_ref._ler_campo_numerico(
            self.entrada_tempo_imp.get(), f"Tempo de Impressão ({rotulo_grupo})", permitir_zero=True, minimo=0)
        tempo_mo = self.app_ref._ler_campo_numerico(
            self.entrada_tempo_mo.get(), f"Tempo de Mão de Obra ({rotulo_grupo})", permitir_zero=True, minimo=0)
        peso_purga = self.app_ref._ler_campo_numerico(
            self.entrada_peso_purga.get(), f"Peso de Purga ({rotulo_grupo})", permitir_zero=True, minimo=0)

        if not self._itens_filamento:
            raise CampoInvalidoError(f"Filamentos do {rotulo_grupo}", "adicione ao menos um filamento/cor utilizado")

        tipo_troca = "manual" if self.combo_tipo_troca.get().startswith("Manual") else "automatica"
        tempo_por_troca = 0.0
        if tipo_troca == "manual":
            tempo_por_troca = self.app_ref._ler_campo_numerico(
                self.entrada_tempo_troca.get(), "Tempo por Troca Manual", permitir_zero=True, minimo=0)

        return {
            "nome": nome, "tempo_imp_h": tempo_imp, "tempo_mo_h": tempo_mo,
            "peso_purga_g": peso_purga, "tipo_troca_cor": tipo_troca,
            "tempo_por_troca_h": tempo_por_troca,
            "filamentos": [dict(it) for it in self._itens_filamento],
        }

    def preencher(self, modulo):
        if self.mostrar_nome:
            self.entrada_nome_modulo.delete(0, "end")
            self.entrada_nome_modulo.insert(0, modulo.get("nome", ""))
        self.entrada_tempo_imp.delete(0, "end")
        self.entrada_tempo_imp.insert(0, formatar_numero(modulo.get("tempo_imp_h", 0)))
        self.entrada_tempo_mo.delete(0, "end")
        self.entrada_tempo_mo.insert(0, formatar_numero(modulo.get("tempo_mo_h", 0)))
        self.entrada_peso_purga.delete(0, "end")
        self.entrada_peso_purga.insert(0, formatar_numero(modulo.get("peso_purga_g", 0)))
        self.combo_tipo_troca.set(
            "Manual (Pausa no G-code)" if modulo.get("tipo_troca_cor") == "manual" else "Automática (CFS/AMS)"
        )
        self._ao_mudar_tipo_troca()
        self.entrada_tempo_troca.delete(0, "end")
        self.entrada_tempo_troca.insert(0, formatar_numero(modulo.get("tempo_por_troca_h", 0)))
        self._itens_filamento = [dict(it) for it in modulo.get("filamentos", [])]
        self.atualizar_opcoes_filamento()
        self._atualizar_mini_tabela_filamentos()

    def limpar(self):
        if self.mostrar_nome:
            self.entrada_nome_modulo.delete(0, "end")
        self.entrada_tempo_imp.delete(0, "end")
        self.entrada_tempo_mo.delete(0, "end")
        self.entrada_tempo_mo.insert(0, "0")
        self.entrada_peso_purga.delete(0, "end")
        self.entrada_peso_purga.insert(0, "0")
        self.combo_tipo_troca.set("Automática (CFS/AMS)")
        self._ao_mudar_tipo_troca()
        self.entrada_tempo_troca.delete(0, "end")
        self.entrada_tempo_troca.insert(0, "0")
        self._itens_filamento = []
        self.atualizar_opcoes_filamento()
        self._atualizar_mini_tabela_filamentos()


class ModuloDialog(ctk.CTkToplevel):
    """Janela pop-up para criar/editar um módulo (sub-peça) de um produto
    composto/modular."""
    def __init__(self, master, app_ref, modulo_existente=None):
        super().__init__(master)
        self.title("Módulo / Sub-peça do Produto")
        self.geometry("560x680")
        self.configure(fg_color=COR_FUNDO)
        self.resultado = None

        scroll = ctk.CTkScrollableFrame(self, fg_color="transparent")
        scroll.pack(fill="both", expand=True, padx=20, pady=(20, 0))

        self.editor = EditorModuloFrame(scroll, app_ref, mostrar_nome=True)
        self.editor.pack(fill="both", expand=True)
        self.editor.atualizar_opcoes_filamento()
        if modulo_existente:
            self.editor.preencher(modulo_existente)
        else:
            self.editor.limpar()

        botoes = ctk.CTkFrame(self, fg_color="transparent")
        botoes.pack(fill="x", padx=20, pady=20)
        ctk.CTkButton(botoes, text="💾  Salvar Módulo", command=self._salvar,
                      fg_color=COR_DESTAQUE, hover_color="#268a5e").pack(side="right")
        ctk.CTkButton(botoes, text="✖  Cancelar", command=self._cancelar,
                      fg_color="#3a3a3a", hover_color="#4a4a4a").pack(side="right", padx=(0, 10))

        self.protocol("WM_DELETE_WINDOW", self._cancelar)
        self.transient(master)
        self.after(50, self.grab_set)

    def _salvar(self):
        try:
            self.resultado = self.editor.obter_dados()
        except CampoInvalidoError as e:
            messagebox.showerror("Campo inválido", str(e))
            return
        self.destroy()

    def _cancelar(self):
        self.resultado = None
        self.destroy()


# ----------------------------------------------------------------------------
# Aplicativo principal
# ----------------------------------------------------------------------------
class CalculadoraImpressao3D(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Calculadora de Precificação de Impressão 3D")
        self.geometry("1500x880")
        self.minsize(1200, 700)
        self.configure(fg_color=COR_FUNDO)

        # Contadores de ID e armazenamento em memória
        self.proximo_id = 1
        self.proximo_id_filamento = 1
        self.proximo_id_insumo = 1
        self.produtos = {}
        self.filamentos = {}     # id -> {nome, cor, peso_carretel_g, preco_carretel, preco_kg}
        self.insumos = {}        # id -> {nome, preco_unit}
        self.categorias_margem = dict(CATEGORIAS_PADRAO)
        self.moeda_atual = "BRL"

        self._modulos_produto_atual = []
        self._insumos_produto_atual = []
        self._editando_id = None
        self._editando_id_filamento = None
        self._editando_id_insumo = None
        self._editando_categoria_nome = None
        self._ordenacao_atual = (None, False)
        self._pronto_para_salvar = False

        self._configurar_estilo_treeview()

        self.grid_rowconfigure(1, weight=1)
        self.grid_columnconfigure(0, weight=1)

        self._criar_cabecalho()
        self._criar_tabview()

        self.carregar_dados()
        self._pronto_para_salvar = True

        self.protocol("WM_DELETE_WINDOW", self._ao_fechar_janela)

    def _ao_fechar_janela(self):
        self.salvar_dados()
        self.destroy()

    # ------------------------------------------------------------------
    # Estilo compartilhado do ttk.Treeview
    # ------------------------------------------------------------------
    def _configurar_estilo_treeview(self):
        estilo = ttk.Style()
        estilo.theme_use("default")
        estilo.configure(
            "Custom.Treeview", background="#1e1e1e", fieldbackground="#1e1e1e",
            foreground="#e6e6e6", rowheight=28, bordercolor="#1e1e1e",
            borderwidth=0, font=("Segoe UI", 10),
        )
        estilo.configure(
            "Custom.Treeview.Heading", background="#2f2f2f", foreground="#ffffff",
            font=("Segoe UI", 10, "bold"), relief="flat",
        )
        estilo.map("Custom.Treeview", background=[("selected", COR_DESTAQUE)],
                   foreground=[("selected", "#ffffff")])
        estilo.map("Custom.Treeview.Heading", background=[("active", "#3a3a3a")])

    def _criar_cabecalho(self):
        header = ctk.CTkFrame(self, fg_color="transparent")
        header.grid(row=0, column=0, sticky="ew", padx=20, pady=(20, 10))
        ctk.CTkLabel(header, text="🖨️  Calculadora de Precificação de Impressão 3D",
                     font=ctk.CTkFont(size=22, weight="bold")).pack(side="left")

    # ------------------------------------------------------------------
    # Abas principais
    # ------------------------------------------------------------------
    def _criar_tabview(self):
        self.tabview = ctk.CTkTabview(self, fg_color=COR_PAINEL, segmented_button_selected_color=COR_DESTAQUE)
        self.tabview.grid(row=1, column=0, sticky="nsew", padx=20, pady=(0, 20))

        aba_produtos = self.tabview.add("🖨️  Calculadora / Produtos")
        aba_parametros = self.tabview.add("⚙️  Parâmetros Globais")
        aba_insumos = self.tabview.add("🧵  Insumos & Filamentos")

        self._criar_aba_produtos(aba_produtos)
        self._criar_aba_parametros(aba_parametros)
        self._criar_aba_insumos(aba_insumos)

    # ==================================================================
    # ABA: Parâmetros Globais
    # ==================================================================
    def _criar_aba_parametros(self, aba):
        aba.grid_columnconfigure(0, weight=1)
        aba.grid_columnconfigure(1, weight=1)
        aba.grid_rowconfigure(0, weight=1)

        # ---- coluna esquerda: parâmetros numéricos, moeda, empresa ----
        esquerda = ctk.CTkScrollableFrame(aba, fg_color="transparent")
        esquerda.grid(row=0, column=0, sticky="nsew", padx=(15, 8), pady=15)

        ctk.CTkLabel(esquerda, text="Parâmetros Financeiros", font=ctk.CTkFont(size=16, weight="bold")
                     ).pack(anchor="w", pady=(0, 12))

        self.entradas_parametros = {}
        campos = [
            ("energia", "Tarifa de Energia (R$/kWh)", "0,85"),
            ("consumo", "Consumo da Impressora (kW)", "0,15"),
            ("manutencao", "Taxa de Manutenção (R$/h)", "2,00"),
            ("valor_impressora", "Valor de Compra da Impressora (R$)", "2500,00"),
            ("vida_util", "Vida Útil Estimada (horas)", "3000"),
            ("mao_obra", "Custo de Mão de Obra (R$/h)", "20,00"),
            ("taxa_falha", "Taxa de Falhas / Perdas (%)", "5"),
            ("margem", "Margem de Lucro Padrão (%)", "150"),
        ]
        for chave, rotulo, padrao in campos:
            self._criar_campo_parametro(esquerda, chave, rotulo, padrao)

        ctk.CTkButton(esquerda, text="🔄  Recalcular Todos os Produtos",
                      command=self.recalcular_todos_produtos, fg_color="#3a3a3a",
                      hover_color="#4a4a4a").pack(fill="x", pady=(6, 15))

        sep1 = ctk.CTkFrame(esquerda, fg_color="#3a3a3a", height=1)
        sep1.pack(fill="x", pady=(0, 15))

        ctk.CTkLabel(esquerda, text="Moeda de Exibição", font=ctk.CTkFont(size=16, weight="bold")
                     ).pack(anchor="w", pady=(0, 8))
        self.combo_moeda = ctk.CTkComboBox(
            esquerda, values=list(MOEDA_DISPLAY_PARA_CODIGO.keys()),
            state="readonly", command=self._ao_mudar_moeda,
        )
        self.combo_moeda.set(MOEDA_CODIGO_PARA_DISPLAY["BRL"])
        self.combo_moeda.pack(fill="x", pady=(0, 15))

        sep2 = ctk.CTkFrame(esquerda, fg_color="#3a3a3a", height=1)
        sep2.pack(fill="x", pady=(0, 15))

        ctk.CTkLabel(esquerda, text="Dados da Empresa (usados no PDF)", font=ctk.CTkFont(size=16, weight="bold")
                     ).pack(anchor="w", pady=(0, 8))
        self.entrada_empresa_nome = ctk.CTkEntry(esquerda, placeholder_text="Ex: Explorer Fabricação Digital")
        self.entrada_empresa_nome.pack(fill="x", pady=(0, 8))
        self.entrada_empresa_contato = ctk.CTkEntry(esquerda, placeholder_text="Ex: (51) 99999-0000 · contato@empresa.com")
        self.entrada_empresa_contato.pack(fill="x", pady=(0, 8))
        for entrada in (self.entrada_empresa_nome, self.entrada_empresa_contato):
            entrada.bind("<FocusOut>", lambda e: self.salvar_dados())

        # ---- coluna direita: categorias de margem ----
        direita = ctk.CTkFrame(aba, fg_color=COR_PAINEL, corner_radius=12)
        direita.grid(row=0, column=1, sticky="nsew", padx=(8, 15), pady=15)

        ctk.CTkLabel(direita, text="Categorias e Margens", font=ctk.CTkFont(size=16, weight="bold")
                     ).pack(anchor="w", padx=20, pady=(20, 4))
        ctk.CTkLabel(direita, text="Usadas quando o produto tem o modelo de margem \"Por Categoria\".",
                     text_color=COR_TEXTO_SECUNDARIO, font=ctk.CTkFont(size=11), wraplength=320, justify="left"
                     ).pack(anchor="w", padx=20, pady=(0, 12))

        self.tabela_categorias = ttk.Treeview(
            direita, columns=("categoria", "margem"), show="headings",
            style="Custom.Treeview", selectmode="browse", height=8,
        )
        self.tabela_categorias.heading("categoria", text="Categoria")
        self.tabela_categorias.heading("margem", text="Margem (%)")
        self.tabela_categorias.column("categoria", width=180, anchor="w")
        self.tabela_categorias.column("margem", width=90, anchor="center")
        self.tabela_categorias.pack(fill="x", padx=20)
        self.tabela_categorias.bind("<Double-1>", self._ao_duplo_clique_categoria)

        form_cat = ctk.CTkFrame(direita, fg_color="transparent")
        form_cat.pack(fill="x", padx=20, pady=(10, 6))
        self.entrada_nome_categoria = ctk.CTkEntry(form_cat, placeholder_text="Ex: Peça Técnica")
        self.entrada_nome_categoria.pack(fill="x", pady=(0, 6))
        self.entrada_margem_categoria = ctk.CTkEntry(form_cat, placeholder_text="Margem (%) — ex: 200")
        self.entrada_margem_categoria.pack(fill="x", pady=(0, 6))

        self.texto_botao_categoria = tk.StringVar(value="➕  Adicionar Categoria")
        ctk.CTkButton(direita, textvariable=self.texto_botao_categoria,
                      command=self._salvar_categoria, fg_color=COR_DESTAQUE,
                      hover_color="#268a5e").pack(fill="x", padx=20, pady=(0, 6))
        ctk.CTkButton(direita, text="🗑️  Remover Categoria Selecionada",
                      command=self._remover_categoria, fg_color=COR_ERRO,
                      hover_color="#a83a3a").pack(fill="x", padx=20, pady=(0, 20))

    def _criar_campo_parametro(self, parent, chave, rotulo, valor_padrao):
        frame = ctk.CTkFrame(parent, fg_color="transparent")
        frame.pack(fill="x", pady=6)
        ctk.CTkLabel(frame, text=rotulo, font=ctk.CTkFont(size=12),
                     text_color=COR_TEXTO_SECUNDARIO, anchor="w").pack(anchor="w")
        entrada = ctk.CTkEntry(frame, placeholder_text=valor_padrao)
        entrada.insert(0, valor_padrao)
        entrada.pack(fill="x", pady=(4, 0))
        entrada.bind("<KeyRelease>", lambda e: self.salvar_dados())
        entrada.bind("<FocusOut>", lambda e: self.salvar_dados())
        entrada.bind("<Return>", lambda e: self.salvar_dados())
        self.entradas_parametros[chave] = entrada

    def _ao_mudar_moeda(self, valor_display):
        self.moeda_atual = MOEDA_DISPLAY_PARA_CODIGO.get(valor_display, "BRL")
        self._preencher_tabela()
        self._atualizar_resumo()
        self.salvar_dados()

    # -- categorias de margem ---------------------------------------------
    def _salvar_categoria(self):
        nome = self.entrada_nome_categoria.get().strip()
        if not nome:
            messagebox.showwarning("Campo obrigatório", "Informe o nome da categoria.")
            return
        try:
            margem = self._ler_campo_numerico(self.entrada_margem_categoria.get(), "Margem da Categoria",
                                               permitir_zero=True, minimo=0)
        except CampoInvalidoError as e:
            messagebox.showerror("Valor inválido", str(e))
            return

        if self._editando_categoria_nome and self._editando_categoria_nome != nome:
            self.categorias_margem.pop(self._editando_categoria_nome, None)

        self.categorias_margem[nome] = margem
        self._atualizar_tabela_categorias()
        self._atualizar_opcoes_categoria_produto()
        self.entrada_nome_categoria.delete(0, "end")
        self.entrada_margem_categoria.delete(0, "end")
        self._editando_categoria_nome = None
        self.texto_botao_categoria.set("➕  Adicionar Categoria")
        self.salvar_dados()

    def _remover_categoria(self):
        selecao = self.tabela_categorias.selection()
        if not selecao:
            messagebox.showinfo("Nenhuma seleção", "Selecione uma categoria para remover.")
            return
        for item_id in selecao:
            self.categorias_margem.pop(item_id, None)
        self._atualizar_tabela_categorias()
        self._atualizar_opcoes_categoria_produto()
        self.salvar_dados()

    def _ao_duplo_clique_categoria(self, event):
        item = self.tabela_categorias.identify_row(event.y)
        if not item:
            return
        margem = self.categorias_margem.get(item, 0)
        self.entrada_nome_categoria.delete(0, "end")
        self.entrada_nome_categoria.insert(0, item)
        self.entrada_margem_categoria.delete(0, "end")
        self.entrada_margem_categoria.insert(0, formatar_numero(margem))
        self._editando_categoria_nome = item
        self.texto_botao_categoria.set("💾  Salvar Categoria")

    def _atualizar_tabela_categorias(self):
        for item in self.tabela_categorias.get_children(""):
            self.tabela_categorias.delete(item)
        for nome, margem in self.categorias_margem.items():
            self.tabela_categorias.insert("", "end", iid=nome, values=(nome, formatar_numero(margem)))

    # ==================================================================
    # ABA: Insumos & Filamentos
    # ==================================================================
    def _criar_aba_insumos(self, aba):
        aba.grid_columnconfigure(0, weight=1)
        aba.grid_columnconfigure(1, weight=1)
        aba.grid_rowconfigure(0, weight=1)

        # ---- Filamentos ----
        painel_fil = ctk.CTkScrollableFrame(aba, fg_color=COR_PAINEL, corner_radius=12)
        painel_fil.grid(row=0, column=0, sticky="nsew", padx=(15, 8), pady=15)

        ctk.CTkLabel(painel_fil, text="Filamentos Cadastrados", font=ctk.CTkFont(size=16, weight="bold")
                     ).pack(anchor="w", pady=(0, 12))

        self.tabela_filamentos = ttk.Treeview(
            painel_fil, columns=("nome", "cor", "peso_carretel", "preco_carretel", "preco_kg"),
            show="headings", style="Custom.Treeview", selectmode="browse", height=8,
        )
        titulos_fil = {"nome": "Nome", "cor": "Cor", "peso_carretel": "Peso Carretel",
                       "preco_carretel": "Preço Carretel", "preco_kg": "R$/kg"}
        larguras_fil = {"nome": 90, "cor": 110, "peso_carretel": 95, "preco_carretel": 95, "preco_kg": 85}
        for col, titulo in titulos_fil.items():
            self.tabela_filamentos.heading(col, text=titulo)
            self.tabela_filamentos.column(col, width=larguras_fil[col], anchor="center")
        self.tabela_filamentos.column("nome", anchor="w")
        self.tabela_filamentos.column("cor", anchor="w")
        self.tabela_filamentos.pack(fill="x", pady=(0, 10))
        self.tabela_filamentos.bind("<Double-1>", self._ao_duplo_clique_filamento)

        self.entrada_nome_filamento = ctk.CTkEntry(painel_fil, placeholder_text="Ex: PLA")
        self.entrada_nome_filamento.pack(fill="x", pady=(0, 6))
        self.entrada_cor_filamento = ctk.CTkEntry(painel_fil, placeholder_text="Ex: Preto Cadáver / Slot CFS 1")
        self.entrada_cor_filamento.pack(fill="x", pady=(0, 6))

        linha_peso = ctk.CTkFrame(painel_fil, fg_color="transparent")
        linha_peso.pack(fill="x", pady=(0, 4))
        self.entrada_peso_carretel = ctk.CTkEntry(linha_peso, placeholder_text="Peso do Carretel (g) — ex: 1000")
        self.entrada_peso_carretel.pack(side="left", fill="x", expand=True)

        linha_presets = ctk.CTkFrame(painel_fil, fg_color="transparent")
        linha_presets.pack(fill="x", pady=(0, 6))
        for gramas in (250, 1000, 2000, 3000, 5000):
            ctk.CTkButton(
                linha_presets, text=f"{gramas}g", width=54,
                command=lambda g=gramas: self._definir_peso_carretel(g),
                fg_color="#3a3a3a", hover_color="#4a4a4a",
            ).pack(side="left", padx=(0, 4))

        self.entrada_preco_carretel = ctk.CTkEntry(painel_fil, placeholder_text="Preço do Carretel (R$) — ex: 95,63")
        self.entrada_preco_carretel.pack(fill="x", pady=(0, 4))

        self.label_preview_preco_kg = ctk.CTkLabel(
            painel_fil, text="≈ R$/kg: —", text_color=COR_TEXTO_SECUNDARIO, font=ctk.CTkFont(size=12)
        )
        self.label_preview_preco_kg.pack(anchor="w", pady=(0, 8))
        for entrada in (self.entrada_peso_carretel, self.entrada_preco_carretel):
            entrada.bind("<KeyRelease>", lambda e: self._atualizar_preview_preco_kg())

        self.texto_botao_filamento = tk.StringVar(value="➕  Adicionar Filamento")
        ctk.CTkButton(painel_fil, textvariable=self.texto_botao_filamento,
                      command=self._salvar_filamento, fg_color=COR_DESTAQUE,
                      hover_color="#268a5e").pack(fill="x", pady=(0, 6))
        ctk.CTkButton(painel_fil, text="🗑️  Remover Filamento Selecionado",
                      command=self.remover_filamento, fg_color=COR_ERRO,
                      hover_color="#a83a3a").pack(fill="x")
        ctk.CTkLabel(painel_fil, text="Dica: dê um duplo clique em um filamento na tabela para editá-lo.",
                     text_color=COR_TEXTO_SECUNDARIO, font=ctk.CTkFont(size=11), wraplength=320, justify="left"
                     ).pack(anchor="w", pady=(8, 0))

        # ---- Insumos extras ----
        painel_ins = ctk.CTkScrollableFrame(aba, fg_color=COR_PAINEL, corner_radius=12)
        painel_ins.grid(row=0, column=1, sticky="nsew", padx=(8, 15), pady=15)

        ctk.CTkLabel(painel_ins, text="Insumos Extras Cadastrados", font=ctk.CTkFont(size=16, weight="bold")
                     ).pack(anchor="w", pady=(0, 12))
        ctk.CTkLabel(painel_ins, text="Parafusos, ímãs, insertos, embalagem, caixa etc.",
                     text_color=COR_TEXTO_SECUNDARIO, font=ctk.CTkFont(size=11)).pack(anchor="w", pady=(0, 10))

        self.tabela_insumos = ttk.Treeview(
            painel_ins, columns=("nome", "preco"), show="headings",
            style="Custom.Treeview", selectmode="browse", height=8,
        )
        self.tabela_insumos.heading("nome", text="Insumo")
        self.tabela_insumos.heading("preco", text="Preço Unitário")
        self.tabela_insumos.column("nome", width=180, anchor="w")
        self.tabela_insumos.column("preco", width=110, anchor="center")
        self.tabela_insumos.pack(fill="x", pady=(0, 10))
        self.tabela_insumos.bind("<Double-1>", self._ao_duplo_clique_insumo)

        self.entrada_nome_insumo = ctk.CTkEntry(painel_ins, placeholder_text="Ex: Inserto Rosqueado M3")
        self.entrada_nome_insumo.pack(fill="x", pady=(0, 6))
        self.entrada_preco_insumo = ctk.CTkEntry(painel_ins, placeholder_text="Preço Unitário (R$) — ex: 0,80")
        self.entrada_preco_insumo.pack(fill="x", pady=(0, 6))

        self.texto_botao_insumo = tk.StringVar(value="➕  Adicionar Insumo")
        ctk.CTkButton(painel_ins, textvariable=self.texto_botao_insumo,
                      command=self._salvar_insumo, fg_color=COR_DESTAQUE,
                      hover_color="#268a5e").pack(fill="x", pady=(0, 6))
        ctk.CTkButton(painel_ins, text="🗑️  Remover Insumo Selecionado",
                      command=self.remover_insumo, fg_color=COR_ERRO,
                      hover_color="#a83a3a").pack(fill="x")
        ctk.CTkLabel(painel_ins, text="Dica: dê um duplo clique em um insumo na tabela para editá-lo.",
                     text_color=COR_TEXTO_SECUNDARIO, font=ctk.CTkFont(size=11), wraplength=320, justify="left"
                     ).pack(anchor="w", pady=(8, 0))

    def _definir_peso_carretel(self, gramas):
        self.entrada_peso_carretel.delete(0, "end")
        self.entrada_peso_carretel.insert(0, str(gramas))
        self._atualizar_preview_preco_kg()

    def _atualizar_preview_preco_kg(self):
        try:
            peso_g = texto_para_float(self.entrada_peso_carretel.get())
            preco = texto_para_float(self.entrada_preco_carretel.get())
            if peso_g <= 0:
                raise ValueError
            preco_kg = preco / (peso_g / 1000.0)
            self.label_preview_preco_kg.configure(text=f"≈ R$/kg: {formatar_numero(preco_kg)}")
        except (ValueError, ZeroDivisionError):
            self.label_preview_preco_kg.configure(text="≈ R$/kg: —")

    # -- filamentos: CRUD ---------------------------------------------------
    def _salvar_filamento(self):
        nome = self.entrada_nome_filamento.get().strip()
        cor = self.entrada_cor_filamento.get().strip()
        if not nome:
            messagebox.showwarning("Campo obrigatório", "Informe o nome do filamento.")
            return
        try:
            peso_carretel = self._ler_campo_numerico(self.entrada_peso_carretel.get(), "Peso do Carretel",
                                                       permitir_zero=False, minimo=0.01)
            preco_carretel = self._ler_campo_numerico(self.entrada_preco_carretel.get(), "Preço do Carretel",
                                                        permitir_zero=False, minimo=0)
        except CampoInvalidoError as e:
            messagebox.showerror("Valor inválido", str(e))
            return

        preco_kg = preco_carretel / (peso_carretel / 1000.0)
        dados = {"nome": nome, "cor": cor, "peso_carretel_g": peso_carretel,
                 "preco_carretel": preco_carretel, "preco_kg": preco_kg}

        if self._editando_id_filamento is not None:
            self.filamentos[self._editando_id_filamento] = dados
            self._editando_id_filamento = None
            self.texto_botao_filamento.set("➕  Adicionar Filamento")
        else:
            filamento_id = self.proximo_id_filamento
            self.proximo_id_filamento += 1
            self.filamentos[filamento_id] = dados

        self._atualizar_tabela_filamentos()
        self._atualizar_todas_opcoes_filamento()
        self.entrada_nome_filamento.delete(0, "end")
        self.entrada_cor_filamento.delete(0, "end")
        self.entrada_peso_carretel.delete(0, "end")
        self.entrada_preco_carretel.delete(0, "end")
        self.label_preview_preco_kg.configure(text="≈ R$/kg: —")
        self.salvar_dados()

    def remover_filamento(self):
        selecionado = self.tabela_filamentos.selection()
        if not selecionado:
            messagebox.showinfo("Nenhuma seleção", "Selecione um filamento na lista para remover.")
            return
        if not messagebox.askyesno(
            "Confirmar remoção",
            "Remover o filamento selecionado?\nProdutos já cadastrados manterão a última cor/preço utilizado."
        ):
            return
        for item_id in selecionado:
            self.filamentos.pop(int(item_id), None)
        self._atualizar_tabela_filamentos()
        self._atualizar_todas_opcoes_filamento()
        self.salvar_dados()

    def _ao_duplo_clique_filamento(self, event):
        item = self.tabela_filamentos.identify_row(event.y)
        if not item:
            return
        filamento_id = int(item)
        dados = self.filamentos.get(filamento_id)
        if not dados:
            return
        self.entrada_nome_filamento.delete(0, "end")
        self.entrada_nome_filamento.insert(0, dados["nome"])
        self.entrada_cor_filamento.delete(0, "end")
        self.entrada_cor_filamento.insert(0, dados.get("cor", ""))
        self.entrada_peso_carretel.delete(0, "end")
        self.entrada_peso_carretel.insert(0, formatar_numero(dados["peso_carretel_g"]))
        self.entrada_preco_carretel.delete(0, "end")
        self.entrada_preco_carretel.insert(0, formatar_numero(dados["preco_carretel"]))
        self._atualizar_preview_preco_kg()
        self._editando_id_filamento = filamento_id
        self.texto_botao_filamento.set("💾  Salvar Edição do Filamento")

    def _atualizar_tabela_filamentos(self):
        for item in self.tabela_filamentos.get_children(""):
            self.tabela_filamentos.delete(item)
        for fid, dados in self.filamentos.items():
            self.tabela_filamentos.insert(
                "", "end", iid=str(fid),
                values=(dados["nome"], dados.get("cor", "-"),
                        f"{formatar_numero(dados['peso_carretel_g'], 0)} g",
                        formatar_moeda(dados["preco_carretel"], self.moeda_atual),
                        formatar_moeda(dados["preco_kg"], self.moeda_atual))
            )

    def mapa_display_filamento(self):
        mapa = {}
        for fid, dados in self.filamentos.items():
            cor = (dados.get("cor") or "").strip()
            display = f"{dados['nome']} - {cor}" if cor else dados["nome"]
            if display in mapa:
                display = f"{display} #{fid}"
            mapa[display] = fid
        return mapa

    def _atualizar_todas_opcoes_filamento(self):
        """Atualiza a combobox de filamentos em todo formulário de módulo
        atualmente visível (modo simples do produto)."""
        if hasattr(self, "editor_modulo_simples"):
            self.editor_modulo_simples.atualizar_opcoes_filamento()

    # -- insumos: CRUD --------------------------------------------------
    def _salvar_insumo(self):
        nome = self.entrada_nome_insumo.get().strip()
        if not nome:
            messagebox.showwarning("Campo obrigatório", "Informe o nome do insumo.")
            return
        try:
            preco = self._ler_campo_numerico(self.entrada_preco_insumo.get(), "Preço do Insumo",
                                              permitir_zero=True, minimo=0)
        except CampoInvalidoError as e:
            messagebox.showerror("Valor inválido", str(e))
            return

        dados = {"nome": nome, "preco_unit": preco}
        if self._editando_id_insumo is not None:
            self.insumos[self._editando_id_insumo] = dados
            self._editando_id_insumo = None
            self.texto_botao_insumo.set("➕  Adicionar Insumo")
        else:
            insumo_id = self.proximo_id_insumo
            self.proximo_id_insumo += 1
            self.insumos[insumo_id] = dados

        self._atualizar_tabela_insumos()
        if hasattr(self, "atualizar_opcoes_insumo_uso"):
            self.atualizar_opcoes_insumo_uso()
        self.entrada_nome_insumo.delete(0, "end")
        self.entrada_preco_insumo.delete(0, "end")
        self.salvar_dados()

    def remover_insumo(self):
        selecionado = self.tabela_insumos.selection()
        if not selecionado:
            messagebox.showinfo("Nenhuma seleção", "Selecione um insumo na lista para remover.")
            return
        for item_id in selecionado:
            self.insumos.pop(int(item_id), None)
        self._atualizar_tabela_insumos()
        self.atualizar_opcoes_insumo_uso()
        self.salvar_dados()

    def _ao_duplo_clique_insumo(self, event):
        item = self.tabela_insumos.identify_row(event.y)
        if not item:
            return
        insumo_id = int(item)
        dados = self.insumos.get(insumo_id)
        if not dados:
            return
        self.entrada_nome_insumo.delete(0, "end")
        self.entrada_nome_insumo.insert(0, dados["nome"])
        self.entrada_preco_insumo.delete(0, "end")
        self.entrada_preco_insumo.insert(0, formatar_numero(dados["preco_unit"]))
        self._editando_id_insumo = insumo_id
        self.texto_botao_insumo.set("💾  Salvar Edição do Insumo")

    def _atualizar_tabela_insumos(self):
        for item in self.tabela_insumos.get_children(""):
            self.tabela_insumos.delete(item)
        for iid, dados in self.insumos.items():
            self.tabela_insumos.insert("", "end", iid=str(iid),
                                        values=(dados["nome"], formatar_moeda(dados["preco_unit"], self.moeda_atual)))

    # ==================================================================
    # ABA: Calculadora / Produtos
    # ==================================================================
    def _criar_aba_produtos(self, aba):
        aba.grid_columnconfigure(0, weight=1)
        aba.grid_rowconfigure(0, weight=0)
        aba.grid_rowconfigure(1, weight=1)

        form = ctk.CTkScrollableFrame(aba, fg_color=COR_PAINEL, corner_radius=12, height=430)
        form.grid(row=0, column=0, sticky="ew", padx=15, pady=(15, 8))

        ctk.CTkLabel(form, text="Cadastro do Produto", font=ctk.CTkFont(size=16, weight="bold")
                     ).pack(anchor="w", pady=(0, 12))

        ctk.CTkLabel(form, text="Nome do Produto", text_color=COR_TEXTO_SECUNDARIO,
                     font=ctk.CTkFont(size=12)).pack(anchor="w")
        self.entrada_nome_produto = ctk.CTkEntry(form, placeholder_text="Ex: Suporte de Celular")
        self.entrada_nome_produto.pack(fill="x", pady=(4, 12))

        linha_cat = ctk.CTkFrame(form, fg_color="transparent")
        linha_cat.pack(fill="x", pady=(0, 12))
        linha_cat.grid_columnconfigure((0, 1, 2), weight=1)

        col1 = ctk.CTkFrame(linha_cat, fg_color="transparent")
        col1.grid(row=0, column=0, sticky="ew", padx=(0, 6))
        ctk.CTkLabel(col1, text="Categoria", text_color=COR_TEXTO_SECUNDARIO, font=ctk.CTkFont(size=12)).pack(anchor="w")
        self.combo_categoria_produto = ctk.CTkComboBox(col1, values=list(self.categorias_margem.keys()), state="readonly")
        self.combo_categoria_produto.pack(fill="x", pady=(4, 0))

        col2 = ctk.CTkFrame(linha_cat, fg_color="transparent")
        col2.grid(row=0, column=1, sticky="ew", padx=6)
        ctk.CTkLabel(col2, text="Modelo de Margem", text_color=COR_TEXTO_SECUNDARIO, font=ctk.CTkFont(size=12)).pack(anchor="w")
        self.combo_tipo_margem = ctk.CTkComboBox(
            col2, values=["Personalizada (%)", "Por Categoria", "Markup Fixo (x)"],
            state="readonly", command=self._ao_mudar_tipo_margem_produto,
        )
        self.combo_tipo_margem.set("Personalizada (%)")
        self.combo_tipo_margem.pack(fill="x", pady=(4, 0))

        col3 = ctk.CTkFrame(linha_cat, fg_color="transparent")
        col3.grid(row=0, column=2, sticky="ew", padx=(6, 0))
        self.label_valor_margem = ctk.CTkLabel(col3, text="Margem (%)", text_color=COR_TEXTO_SECUNDARIO, font=ctk.CTkFont(size=12))
        self.label_valor_margem.pack(anchor="w")
        self.entrada_valor_margem = ctk.CTkEntry(col3, placeholder_text="Ex: 150")
        self.entrada_valor_margem.insert(0, "150")
        self.entrada_valor_margem.pack(fill="x", pady=(4, 0))

        self.switch_composto = ctk.CTkSwitch(form, text="Produto Composto / Modular (múltiplas sub-peças)",
                                              command=self._ao_alternar_composto)
        self.switch_composto.pack(anchor="w", pady=(0, 12))

        self.frame_modo_simples = ctk.CTkFrame(form, fg_color="#1f1f1f", corner_radius=10)
        self.editor_modulo_simples = EditorModuloFrame(self.frame_modo_simples, self)
        self.editor_modulo_simples.pack(fill="x", padx=12, pady=12)

        self.frame_modo_composto = ctk.CTkFrame(form, fg_color="#1f1f1f", corner_radius=10)
        ctk.CTkLabel(self.frame_modo_composto, text="Módulos / Sub-peças", font=ctk.CTkFont(size=13, weight="bold")
                     ).pack(anchor="w", padx=12, pady=(12, 6))
        self.tabela_modulos = ttk.Treeview(
            self.frame_modo_composto, columns=("nome", "peso", "tempo_imp", "tempo_mo", "n_filamentos"),
            show="headings", style="Custom.Treeview", selectmode="browse", height=5,
        )
        for col, titulo, largura in (
            ("nome", "Nome", 130), ("peso", "Peso", 80), ("tempo_imp", "T. Impr.", 80),
            ("tempo_mo", "T. M.Obra", 80), ("n_filamentos", "Cores", 60),
        ):
            self.tabela_modulos.heading(col, text=titulo)
            self.tabela_modulos.column(col, width=largura, anchor="center")
        self.tabela_modulos.column("nome", anchor="w")
        self.tabela_modulos.pack(fill="x", padx=12, pady=(0, 8))
        self.tabela_modulos.bind("<Double-1>", lambda e: self._editar_modulo_selecionado())

        linha_botoes_modulo = ctk.CTkFrame(self.frame_modo_composto, fg_color="transparent")
        linha_botoes_modulo.pack(fill="x", padx=12, pady=(0, 12))
        ctk.CTkButton(linha_botoes_modulo, text="➕  Adicionar Módulo", command=self._adicionar_modulo,
                      fg_color=COR_DESTAQUE, hover_color="#268a5e").pack(side="left", padx=(0, 6))
        ctk.CTkButton(linha_botoes_modulo, text="✏️  Editar Selecionado", command=self._editar_modulo_selecionado,
                      fg_color=COR_AZUL, hover_color="#3a72ab").pack(side="left", padx=(0, 6))
        ctk.CTkButton(linha_botoes_modulo, text="🗑️  Remover Selecionado", command=self._remover_modulo_selecionado,
                      fg_color=COR_ERRO, hover_color="#a83a3a").pack(side="left")

        self.frame_modo_simples.pack(fill="x", pady=(0, 12))

        linha_extra = ctk.CTkFrame(form, fg_color="transparent")
        linha_extra.pack(fill="x", pady=(0, 12))
        linha_extra.grid_columnconfigure((0, 1), weight=1)

        col_montagem = ctk.CTkFrame(linha_extra, fg_color="transparent")
        col_montagem.grid(row=0, column=0, sticky="ew", padx=(0, 6))
        ctk.CTkLabel(col_montagem, text="Tempo de Montagem/Acabamento (h)", text_color=COR_TEXTO_SECUNDARIO,
                     font=ctk.CTkFont(size=12)).pack(anchor="w")
        self.entrada_tempo_montagem = ctk.CTkEntry(col_montagem, placeholder_text="Ex: 0,5")
        self.entrada_tempo_montagem.insert(0, "0")
        self.entrada_tempo_montagem.pack(fill="x", pady=(4, 0))

        col_prep = ctk.CTkFrame(linha_extra, fg_color="transparent")
        col_prep.grid(row=0, column=1, sticky="ew", padx=(6, 0))
        ctk.CTkLabel(col_prep, text="Preparação de Mesa (R$ por peça/mesa)", text_color=COR_TEXTO_SECUNDARIO,
                     font=ctk.CTkFont(size=12)).pack(anchor="w")
        self.entrada_custo_preparacao = ctk.CTkEntry(col_prep, placeholder_text="Ex: 1,50")
        self.entrada_custo_preparacao.insert(0, "0")
        self.entrada_custo_preparacao.pack(fill="x", pady=(4, 0))

        ctk.CTkLabel(form, text="Insumos Extras do Produto", font=ctk.CTkFont(size=13, weight="bold")
                     ).pack(anchor="w", pady=(0, 6))
        self.mini_tabela_insumos_uso = ttk.Treeview(
            form, columns=("insumo", "qtd", "subtotal"), show="headings",
            style="Custom.Treeview", selectmode="extended", height=3,
        )
        self.mini_tabela_insumos_uso.heading("insumo", text="Insumo")
        self.mini_tabela_insumos_uso.heading("qtd", text="Qtd")
        self.mini_tabela_insumos_uso.heading("subtotal", text="Subtotal")
        self.mini_tabela_insumos_uso.column("insumo", width=180, anchor="w")
        self.mini_tabela_insumos_uso.column("qtd", width=60, anchor="center")
        self.mini_tabela_insumos_uso.column("subtotal", width=100, anchor="center")
        self.mini_tabela_insumos_uso.pack(fill="x", pady=(0, 6))

        linha_add_insumo = ctk.CTkFrame(form, fg_color="transparent")
        linha_add_insumo.pack(fill="x", pady=(0, 6))
        self.combo_insumo_uso = ctk.CTkComboBox(linha_add_insumo, values=[], state="readonly")
        self.combo_insumo_uso.pack(side="left", fill="x", expand=True, padx=(0, 6))
        self.entrada_qtd_insumo_uso = ctk.CTkEntry(linha_add_insumo, placeholder_text="Qtd", width=70)
        self.entrada_qtd_insumo_uso.insert(0, "1")
        self.entrada_qtd_insumo_uso.pack(side="left", padx=(0, 6))
        ctk.CTkButton(linha_add_insumo, text="+", width=32, command=self._adicionar_insumo_uso,
                      fg_color=COR_DESTAQUE, hover_color="#268a5e").pack(side="left")
        ctk.CTkButton(form, text="🗑️  Remover Insumo Selecionado do Produto",
                      command=self._remover_insumo_uso, fg_color=COR_ERRO,
                      hover_color="#a83a3a").pack(fill="x", pady=(0, 12))

        self.botao_produto_texto = tk.StringVar(value="➕  Adicionar Produto")
        linha_final = ctk.CTkFrame(form, fg_color="transparent")
        linha_final.pack(fill="x", pady=(0, 6))
        self.btn_adicionar_produto = ctk.CTkButton(
            linha_final, textvariable=self.botao_produto_texto, command=self.adicionar_produto,
            fg_color=COR_DESTAQUE, hover_color="#268a5e", height=38, font=ctk.CTkFont(size=13, weight="bold"),
        )
        self.btn_adicionar_produto.pack(side="left", fill="x", expand=True, padx=(0, 10))
        self._btn_cancelar_edicao_produto = ctk.CTkButton(
            linha_final, text="✖  Cancelar Edição", command=self._cancelar_edicao_produto,
            fg_color="#3a3a3a", hover_color="#4a4a4a",
        )

        self._ao_alternar_composto()

        # ---- Tabela de produtos ----
        painel_tabela = ctk.CTkFrame(aba, fg_color=COR_PAINEL, corner_radius=12)
        painel_tabela.grid(row=1, column=0, sticky="nsew", padx=15, pady=(8, 15))

        cabecalho_tabela = ctk.CTkFrame(painel_tabela, fg_color="transparent")
        cabecalho_tabela.pack(fill="x", padx=20, pady=(15, 10))
        ctk.CTkLabel(cabecalho_tabela, text="Produtos Cadastrados", font=ctk.CTkFont(size=16, weight="bold")
                     ).pack(side="left")
        self.entrada_busca = ctk.CTkEntry(cabecalho_tabela, placeholder_text="🔎  Buscar por nome...", width=260)
        self.entrada_busca.pack(side="right")
        self.entrada_busca.bind("<KeyRelease>", lambda e: self._preencher_tabela())

        frame_tree = ctk.CTkFrame(painel_tabela, fg_color="transparent")
        frame_tree.pack(fill="both", expand=True, padx=20, pady=(0, 10))
        self.tabela = ttk.Treeview(frame_tree, columns=COLUNAS_PRODUTOS, show="headings",
                                    style="Custom.Treeview", selectmode="extended")
        for col in COLUNAS_PRODUTOS:
            ancora = "w" if col in ("nome", "categoria") else "center"
            self.tabela.heading(col, text=TITULOS_COLUNAS[col], command=lambda c=col: self._ordenar_coluna(c))
            self.tabela.column(col, width=LARGURAS_COLUNAS[col], anchor=ancora)

        scroll_y = ttk.Scrollbar(frame_tree, orient="vertical", command=self.tabela.yview)
        scroll_x = ttk.Scrollbar(frame_tree, orient="horizontal", command=self.tabela.xview)
        self.tabela.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)
        frame_tree.grid_rowconfigure(0, weight=1)
        frame_tree.grid_columnconfigure(0, weight=1)
        self.tabela.grid(row=0, column=0, sticky="nsew")
        scroll_y.grid(row=0, column=1, sticky="ns")
        scroll_x.grid(row=1, column=0, sticky="ew")
        self.tabela.bind("<Double-1>", self._ao_duplo_clique_tabela)

        self.label_resumo = ctk.CTkLabel(painel_tabela, text="", font=ctk.CTkFont(size=12, weight="bold"),
                                          text_color=COR_TEXTO_SECUNDARIO)
        self.label_resumo.pack(anchor="w", padx=20, pady=(0, 10))

        rodape = ctk.CTkFrame(painel_tabela, fg_color="transparent")
        rodape.pack(fill="x", padx=20, pady=(0, 15))
        ctk.CTkButton(rodape, text="🗑️  Remover Produto", command=self.remover_produto,
                      fg_color=COR_ERRO, hover_color="#a83a3a").pack(side="left")
        ctk.CTkButton(rodape, text="📊  Simulador de Atacado", command=self.abrir_simulador_atacado,
                      fg_color="#3a3a3a", hover_color="#4a4a4a").pack(side="left", padx=(10, 0))
        ctk.CTkButton(rodape, text="📄  Gerar Orçamento PDF", command=self.abrir_dialogo_pdf,
                      fg_color=COR_AZUL, hover_color="#3a72ab").pack(side="right")
        ctk.CTkButton(rodape, text="📑  Exportar Tabela para CSV", command=self.exportar_csv,
                      fg_color="#3a3a3a", hover_color="#4a4a4a").pack(side="right", padx=(0, 10))

    # -- alternância simples / composto -------------------------------------
    def _ao_alternar_composto(self):
        if self.switch_composto.get():
            self.frame_modo_simples.pack_forget()
            self.frame_modo_composto.pack(fill="x", pady=(0, 12), before=self._localizar_widget_apos_switch())
        else:
            self.frame_modo_composto.pack_forget()
            self.frame_modo_simples.pack(fill="x", pady=(0, 12), before=self._localizar_widget_apos_switch())

    def _localizar_widget_apos_switch(self):
        # Mantém a ordem visual correta reposicionando o frame antes do
        # próximo elemento fixo do formulário (linha de tempo/preparo).
        return self.entrada_tempo_montagem.master.master if hasattr(self.entrada_tempo_montagem, "master") else None

    def _ao_mudar_tipo_margem_produto(self, valor=None):
        if valor == "Por Categoria":
            self.label_valor_margem.configure(text="(definida pela categoria)")
            self.entrada_valor_margem.configure(state="disabled")
        elif valor == "Markup Fixo (x)":
            self.label_valor_margem.configure(text="Markup (x) — ex: 2,5")
            self.entrada_valor_margem.configure(state="normal")
        else:
            self.label_valor_margem.configure(text="Margem (%)")
            self.entrada_valor_margem.configure(state="normal")

    def _atualizar_opcoes_categoria_produto(self):
        valores = list(self.categorias_margem.keys()) or ["Sem Categoria"]
        atual = self.combo_categoria_produto.get()
        self.combo_categoria_produto.configure(values=valores)
        if atual not in valores:
            self.combo_categoria_produto.set(valores[0])

    # -- módulos (produto composto) ------------------------------------
    def _adicionar_modulo(self):
        dlg = ModuloDialog(self, self)
        self.wait_window(dlg)
        if dlg.resultado:
            self._modulos_produto_atual.append(dlg.resultado)
            self._atualizar_tabela_modulos_produto()

    def _editar_modulo_selecionado(self):
        selecao = self.tabela_modulos.selection()
        if not selecao:
            messagebox.showinfo("Nenhuma seleção", "Selecione um módulo na lista para editar.")
            return
        idx = int(selecao[0])
        dlg = ModuloDialog(self, self, modulo_existente=self._modulos_produto_atual[idx])
        self.wait_window(dlg)
        if dlg.resultado:
            self._modulos_produto_atual[idx] = dlg.resultado
            self._atualizar_tabela_modulos_produto()

    def _remover_modulo_selecionado(self):
        selecao = self.tabela_modulos.selection()
        if not selecao:
            return
        indices = sorted((int(i) for i in selecao), reverse=True)
        for i in indices:
            if 0 <= i < len(self._modulos_produto_atual):
                self._modulos_produto_atual.pop(i)
        self._atualizar_tabela_modulos_produto()

    def _atualizar_tabela_modulos_produto(self):
        for item in self.tabela_modulos.get_children(""):
            self.tabela_modulos.delete(item)
        for idx, m in enumerate(self._modulos_produto_atual):
            peso = sum(it["peso_g"] for it in m["filamentos"]) + m.get("peso_purga_g", 0)
            self.tabela_modulos.insert(
                "", "end", iid=str(idx),
                values=(m["nome"], f"{formatar_numero(peso)} g", f"{formatar_numero(m['tempo_imp_h'])} h",
                        f"{formatar_numero(m['tempo_mo_h'])} h", str(len(m["filamentos"])))
            )

    # -- insumos usados no produto ----------------------------------------
    def atualizar_opcoes_insumo_uso(self):
        if not hasattr(self, "combo_insumo_uso"):
            return
        valores = [dados["nome"] for dados in self.insumos.values()]
        self.combo_insumo_uso.configure(values=valores)
        if valores and self.combo_insumo_uso.get() not in valores:
            self.combo_insumo_uso.set(valores[0])
        elif not valores:
            self.combo_insumo_uso.set("")

    def _mapa_nome_para_id_insumo(self):
        return {dados["nome"]: iid for iid, dados in self.insumos.items()}

    def _adicionar_insumo_uso(self):
        nome = self.combo_insumo_uso.get()
        mapa = self._mapa_nome_para_id_insumo()
        if nome not in mapa:
            messagebox.showwarning("Insumo", "Cadastre e selecione um insumo válido.")
            return
        try:
            qtd = self._ler_campo_numerico(self.entrada_qtd_insumo_uso.get(), "Quantidade do Insumo",
                                            permitir_zero=False, minimo=0.01)
        except CampoInvalidoError as e:
            messagebox.showerror("Valor inválido", str(e))
            return
        self._insumos_produto_atual.append({"insumo_id": mapa[nome], "quantidade": qtd})
        self._atualizar_mini_tabela_insumos_uso()

    def _remover_insumo_uso(self):
        selecao = self.mini_tabela_insumos_uso.selection()
        if not selecao:
            return
        indices = sorted((int(i) for i in selecao), reverse=True)
        for i in indices:
            if 0 <= i < len(self._insumos_produto_atual):
                self._insumos_produto_atual.pop(i)
        self._atualizar_mini_tabela_insumos_uso()

    def _atualizar_mini_tabela_insumos_uso(self):
        for item in self.mini_tabela_insumos_uso.get_children(""):
            self.mini_tabela_insumos_uso.delete(item)
        for idx, item in enumerate(self._insumos_produto_atual):
            insumo = self.insumos.get(item["insumo_id"])
            nome = insumo["nome"] if insumo else "Insumo removido"
            preco_unit = insumo["preco_unit"] if insumo else 0
            subtotal = preco_unit * item["quantidade"]
            self.mini_tabela_insumos_uso.insert(
                "", "end", iid=str(idx),
                values=(nome, formatar_numero(item["quantidade"], 0), formatar_moeda(subtotal, self.moeda_atual))
            )

    # ------------------------------------------------------------------
    # Utilidades numéricas
    # ------------------------------------------------------------------
    def _ler_campo_numerico(self, texto, nome_campo, permitir_zero=True, minimo=None):
        try:
            valor = texto_para_float(texto)
        except ValueError as e:
            raise CampoInvalidoError(nome_campo, str(e))
        if not permitir_zero and valor == 0:
            raise CampoInvalidoError(nome_campo, "não pode ser zero")
        if minimo is not None and valor < minimo:
            raise CampoInvalidoError(nome_campo, f"deve ser maior ou igual a {minimo}")
        return valor

    def _obter_parametros(self):
        rotulos = {
            "energia": "Tarifa de Energia", "consumo": "Consumo da Impressora",
            "manutencao": "Taxa de Manutenção", "valor_impressora": "Valor de Compra da Impressora",
            "vida_util": "Vida Útil Estimada", "mao_obra": "Custo de Mão de Obra",
            "taxa_falha": "Taxa de Falhas / Perdas", "margem": "Margem de Lucro Padrão",
        }
        valores = {}
        for chave, entrada in self.entradas_parametros.items():
            valores[chave] = self._ler_campo_numerico(entrada.get(), rotulos.get(chave, chave))
        return valores

    # ------------------------------------------------------------------
    # Motor de cálculo
    # ------------------------------------------------------------------
    def _calcular_modulo_custos(self, modulo, parametros):
        peso_filamentos = sum(item["peso_g"] for item in modulo["filamentos"])
        peso_purga = modulo.get("peso_purga_g", 0.0)

        custo_material = 0.0
        for item in modulo["filamentos"]:
            fil = self.filamentos.get(item["filamento_id"])
            preco_kg = fil["preco_kg"] if fil else 0.0
            custo_material += (preco_kg / 1000.0) * item["peso_g"]

        if peso_filamentos > 0:
            preco_medio_kg = sum(
                (self.filamentos.get(it["filamento_id"], {}).get("preco_kg", 0.0)) * it["peso_g"]
                for it in modulo["filamentos"]
            ) / peso_filamentos
        else:
            preco_medio_kg = 0.0
        custo_material += (preco_medio_kg / 1000.0) * peso_purga

        tempo_imp = modulo.get("tempo_imp_h", 0.0)
        tempo_mo = modulo.get("tempo_mo_h", 0.0)

        custo_energia = parametros["consumo"] * tempo_imp * parametros["energia"]
        custo_manutencao = parametros["manutencao"] * tempo_imp
        vida_util = parametros.get("vida_util", 0)
        taxa_dep_hora = (parametros["valor_impressora"] / vida_util) if vida_util > 0 else 0.0
        custo_depreciacao = taxa_dep_hora * tempo_imp

        custo_mao_obra = parametros["mao_obra"] * tempo_mo
        if modulo.get("tipo_troca_cor") == "manual" and len(modulo["filamentos"]) > 1:
            trocas = len(modulo["filamentos"]) - 1
            custo_mao_obra += trocas * modulo.get("tempo_por_troca_h", 0.0) * parametros["mao_obra"]

        return {
            "peso": peso_filamentos + peso_purga, "tempo_imp": tempo_imp, "tempo_mo": tempo_mo,
            "custo_material": custo_material, "custo_energia": custo_energia,
            "custo_manutencao": custo_manutencao, "custo_depreciacao": custo_depreciacao,
            "custo_mao_obra": custo_mao_obra,
        }

    def _calcular_produto_completo(self, dados_produto, parametros):
        total_peso = total_tempo_imp = total_tempo_mo = 0.0
        total_material = total_energia = total_manutencao = total_deprec = total_mao_obra = 0.0

        for modulo in dados_produto["modulos"]:
            r = self._calcular_modulo_custos(modulo, parametros)
            total_peso += r["peso"]
            total_tempo_imp += r["tempo_imp"]
            total_tempo_mo += r["tempo_mo"]
            total_material += r["custo_material"]
            total_energia += r["custo_energia"]
            total_manutencao += r["custo_manutencao"]
            total_deprec += r["custo_depreciacao"]
            total_mao_obra += r["custo_mao_obra"]

        tempo_montagem = dados_produto.get("tempo_montagem_h", 0.0)
        total_mao_obra += parametros["mao_obra"] * tempo_montagem
        total_tempo_mo += tempo_montagem

        num_modulos = max(len(dados_produto["modulos"]), 1)
        custo_preparacao_total = dados_produto.get("custo_preparacao_mesa", 0.0) * num_modulos

        custo_insumos_total = 0.0
        for item in dados_produto.get("insumos_usados", []):
            insumo = self.insumos.get(item["insumo_id"])
            if insumo:
                custo_insumos_total += insumo["preco_unit"] * item["quantidade"]

        custo_perdas = (total_material + total_energia) * (parametros["taxa_falha"] / 100.0)

        custo_total = (total_material + total_energia + total_manutencao + total_deprec
                       + total_mao_obra + custo_perdas + custo_preparacao_total + custo_insumos_total)

        tipo_margem = dados_produto["tipo_margem"]
        if tipo_margem == "markup":
            multiplicador = dados_produto.get("markup_fixo") or 1.0
            preco_venda = custo_total * multiplicador
        elif tipo_margem == "categoria":
            margem_pct = self.categorias_margem.get(dados_produto["categoria"], parametros["margem"])
            preco_venda = custo_total * (1 + margem_pct / 100.0)
        else:
            margem_pct = dados_produto.get("margem_personalizada", parametros["margem"])
            preco_venda = custo_total * (1 + margem_pct / 100.0)

        lucro = preco_venda - custo_total

        return {
            "peso": total_peso, "tempo_imp": total_tempo_imp, "tempo_mo": total_tempo_mo,
            "custo_material": total_material, "custo_energia": total_energia,
            "custo_manutencao": total_manutencao, "custo_depreciacao": total_deprec,
            "custo_mao_obra": total_mao_obra, "custo_preparacao": custo_preparacao_total,
            "custo_insumos": custo_insumos_total, "custo_perdas": custo_perdas,
            "custo_total": custo_total, "preco_venda": preco_venda, "lucro": lucro,
        }

    # ------------------------------------------------------------------
    # Formulário de produto: leitura / preenchimento / limpeza
    # ------------------------------------------------------------------
    def _ler_formulario_produto(self):
        nome = self.entrada_nome_produto.get().strip()
        if not nome:
            raise CampoInvalidoError("Nome do Produto", "é obrigatório")

        categoria = self.combo_categoria_produto.get().strip() or "Sem Categoria"
        tipo_margem_display = self.combo_tipo_margem.get()
        tipo_margem = {"Personalizada (%)": "personalizada", "Por Categoria": "categoria",
                       "Markup Fixo (x)": "markup"}.get(tipo_margem_display, "personalizada")

        valor_margem = 0.0
        if tipo_margem != "categoria":
            valor_margem = self._ler_campo_numerico(self.entrada_valor_margem.get(), "Margem/Markup",
                                                      permitir_zero=True, minimo=0)

        composto = bool(self.switch_composto.get())
        tempo_montagem = self._ler_campo_numerico(self.entrada_tempo_montagem.get(),
                                                    "Tempo de Montagem/Acabamento", permitir_zero=True, minimo=0)
        custo_preparacao = self._ler_campo_numerico(self.entrada_custo_preparacao.get(),
                                                      "Custo de Preparação de Mesa", permitir_zero=True, minimo=0)

        if composto:
            if not self._modulos_produto_atual:
                raise CampoInvalidoError("Módulos", "adicione ao menos um módulo/sub-peça para um produto composto")
            modulos = [dict(m) for m in self._modulos_produto_atual]
        else:
            modulos = [self.editor_modulo_simples.obter_dados()]

        insumos_usados = [dict(it) for it in self._insumos_produto_atual]

        return {
            "nome": nome, "categoria": categoria, "tipo_margem": tipo_margem,
            "margem_personalizada": valor_margem if tipo_margem == "personalizada" else 0.0,
            "markup_fixo": valor_margem if tipo_margem == "markup" else 0.0,
            "composto": composto, "modulos": modulos,
            "tempo_montagem_h": tempo_montagem, "custo_preparacao_mesa": custo_preparacao,
            "insumos_usados": insumos_usados,
        }

    def _limpar_formulario_produto(self):
        self.entrada_nome_produto.delete(0, "end")
        self.combo_tipo_margem.set("Personalizada (%)")
        self._ao_mudar_tipo_margem_produto("Personalizada (%)")
        self.entrada_valor_margem.delete(0, "end")
        self.entrada_valor_margem.insert(0, formatar_numero(self._valor_margem_padrao(), 0))
        self.switch_composto.deselect()
        self._modulos_produto_atual = []
        self._atualizar_tabela_modulos_produto()
        self.editor_modulo_simples.limpar()
        self.entrada_tempo_montagem.delete(0, "end")
        self.entrada_tempo_montagem.insert(0, "0")
        self.entrada_custo_preparacao.delete(0, "end")
        self.entrada_custo_preparacao.insert(0, "0")
        self._insumos_produto_atual = []
        self._atualizar_mini_tabela_insumos_uso()
        self._ao_alternar_composto()
        self.entrada_nome_produto.focus()

    def _valor_margem_padrao(self):
        try:
            return texto_para_float(self.entradas_parametros["margem"].get())
        except (ValueError, KeyError):
            return 150.0

    # ------------------------------------------------------------------
    # Ações de produto: adicionar / editar / remover / recalcular
    # ------------------------------------------------------------------
    def adicionar_produto(self):
        if self._editando_id is not None:
            self._salvar_edicao_produto()
            return
        try:
            dados_produto = self._ler_formulario_produto()
            parametros = self._obter_parametros()
        except CampoInvalidoError as e:
            messagebox.showerror("Campo inválido", str(e))
            return

        resultado = self._calcular_produto_completo(dados_produto, parametros)
        produto_id = self.proximo_id
        self.proximo_id += 1
        self.produtos[produto_id] = {**dados_produto, **resultado}

        self._preencher_tabela()
        self._atualizar_resumo()
        self._limpar_formulario_produto()
        self.salvar_dados()

    def _salvar_edicao_produto(self):
        produto_id = self._editando_id
        try:
            dados_produto = self._ler_formulario_produto()
            parametros = self._obter_parametros()
        except CampoInvalidoError as e:
            messagebox.showerror("Campo inválido", str(e))
            return

        resultado = self._calcular_produto_completo(dados_produto, parametros)
        self.produtos[produto_id] = {**dados_produto, **resultado}

        self._preencher_tabela()
        self._atualizar_resumo()
        self._cancelar_edicao_produto()
        self.salvar_dados()

    def _cancelar_edicao_produto(self):
        self._editando_id = None
        self.botao_produto_texto.set("➕  Adicionar Produto")
        self._btn_cancelar_edicao_produto.pack_forget()
        self._limpar_formulario_produto()

    def remover_produto(self):
        selecionado = self.tabela.selection()
        if not selecionado:
            messagebox.showinfo("Nenhuma seleção", "Selecione um produto na tabela para remover.")
            return
        nomes = [self.produtos[int(iid)]["nome"] for iid in selecionado if int(iid) in self.produtos]
        pergunta = (f"Remover o produto '{nomes[0]}'?" if len(nomes) == 1
                    else f"Remover {len(nomes)} produtos selecionados?")
        if not messagebox.askyesno("Confirmar remoção", pergunta):
            return
        for item_id in selecionado:
            self.produtos.pop(int(item_id), None)
        self._preencher_tabela()
        self._atualizar_resumo()
        self.salvar_dados()

    def recalcular_todos_produtos(self):
        if not self.produtos:
            return
        try:
            parametros = self._obter_parametros()
        except CampoInvalidoError as e:
            messagebox.showerror("Parâmetro inválido", str(e))
            return
        for produto_id, p in self.produtos.items():
            resultado = self._calcular_produto_completo(p, parametros)
            p.update(resultado)
        self._preencher_tabela()
        self._atualizar_resumo()
        messagebox.showinfo("Recalculado", "Todos os produtos foram recalculados com os novos parâmetros.")
        self.salvar_dados()

    # ------------------------------------------------------------------
    # Tabela de produtos
    # ------------------------------------------------------------------
    def _valores_linha(self, produto_id):
        p = self.produtos[produto_id]
        modulos_txt = str(len(p["modulos"])) if p.get("composto") else "-"
        return (
            produto_id, p["nome"], p.get("categoria", "-"), modulos_txt,
            f"{formatar_numero(p['peso'])} g", f"{formatar_numero(p['tempo_imp'])} h",
            f"{formatar_numero(p['tempo_mo'])} h",
            formatar_moeda(p["custo_material"], self.moeda_atual),
            formatar_moeda(p["custo_energia"], self.moeda_atual),
            formatar_moeda(p["custo_manutencao"], self.moeda_atual),
            formatar_moeda(p["custo_depreciacao"], self.moeda_atual),
            formatar_moeda(p["custo_mao_obra"], self.moeda_atual),
            formatar_moeda(p["custo_preparacao"], self.moeda_atual),
            formatar_moeda(p["custo_insumos"], self.moeda_atual),
            formatar_moeda(p["custo_perdas"], self.moeda_atual),
            formatar_moeda(p["custo_total"], self.moeda_atual),
            formatar_moeda(p["preco_venda"], self.moeda_atual),
            formatar_moeda(p["lucro"], self.moeda_atual),
        )

    def _preencher_tabela(self):
        filtro = self.entrada_busca.get().strip().lower() if hasattr(self, "entrada_busca") else ""
        for item in self.tabela.get_children(""):
            self.tabela.delete(item)
        for produto_id, p in self.produtos.items():
            if filtro and filtro not in p["nome"].lower():
                continue
            self.tabela.insert("", "end", iid=str(produto_id), values=self._valores_linha(produto_id))

    def _atualizar_resumo(self):
        total_produtos = len(self.produtos)
        custo_total = sum(p["custo_total"] for p in self.produtos.values())
        lucro_total = sum(p["lucro"] for p in self.produtos.values())
        self.label_resumo.configure(
            text=(f"Total de produtos: {total_produtos}   |   "
                  f"Custo Total: {formatar_moeda(custo_total, self.moeda_atual)}   |   "
                  f"Lucro Total: {formatar_moeda(lucro_total, self.moeda_atual)}")
        )

    def _ordenar_coluna(self, col):
        coluna_atual, reverso_atual = self._ordenacao_atual
        reverso = (not reverso_atual) if coluna_atual == col else False
        itens = [(self.tabela.set(k, col), k) for k in self.tabela.get_children("")]
        if col in COLUNAS_NUMERICAS:
            itens.sort(key=lambda t: valor_numerico_de_celula(t[0]), reverse=reverso)
        else:
            itens.sort(key=lambda t: t[0].lower(), reverse=reverso)
        for indice, (_, k) in enumerate(itens):
            self.tabela.move(k, "", indice)
        self._ordenacao_atual = (col, reverso)

    # ------------------------------------------------------------------
    # Edição de produto (duplo clique na tabela)
    # ------------------------------------------------------------------
    def _ao_duplo_clique_tabela(self, event):
        item = self.tabela.identify_row(event.y)
        if not item:
            return
        produto_id = int(item)
        if produto_id not in self.produtos:
            return
        self._entrar_modo_edicao_produto(produto_id)

    def _entrar_modo_edicao_produto(self, produto_id):
        p = self.produtos[produto_id]
        self._editando_id = produto_id

        self.entrada_nome_produto.delete(0, "end")
        self.entrada_nome_produto.insert(0, p["nome"])
        self._atualizar_opcoes_categoria_produto()
        self.combo_categoria_produto.set(p.get("categoria", "Sem Categoria"))

        tipo_display = {"personalizada": "Personalizada (%)", "categoria": "Por Categoria",
                        "markup": "Markup Fixo (x)"}.get(p.get("tipo_margem", "personalizada"), "Personalizada (%)")
        self.combo_tipo_margem.set(tipo_display)
        self._ao_mudar_tipo_margem_produto(tipo_display)
        valor_margem = p.get("markup_fixo") if p.get("tipo_margem") == "markup" else p.get("margem_personalizada", 0)
        self.entrada_valor_margem.configure(state="normal")
        self.entrada_valor_margem.delete(0, "end")
        self.entrada_valor_margem.insert(0, formatar_numero(valor_margem or 0))
        if tipo_display == "Por Categoria":
            self.entrada_valor_margem.configure(state="disabled")

        if p.get("composto"):
            self.switch_composto.select()
            self._modulos_produto_atual = [dict(m) for m in p["modulos"]]
            self._atualizar_tabela_modulos_produto()
        else:
            self.switch_composto.deselect()
            self._modulos_produto_atual = []
            self.editor_modulo_simples.preencher(p["modulos"][0])
        self._ao_alternar_composto()

        self.entrada_tempo_montagem.delete(0, "end")
        self.entrada_tempo_montagem.insert(0, formatar_numero(p.get("tempo_montagem_h", 0)))
        self.entrada_custo_preparacao.delete(0, "end")
        self.entrada_custo_preparacao.insert(0, formatar_numero(p.get("custo_preparacao_mesa", 0)))

        self._insumos_produto_atual = [dict(it) for it in p.get("insumos_usados", [])]
        self._atualizar_mini_tabela_insumos_uso()

        self.botao_produto_texto.set("💾  Salvar Edição")
        self._btn_cancelar_edicao_produto.pack(side="left", padx=(10, 0))
        self.entrada_nome_produto.focus()

    # ------------------------------------------------------------------
    # Exportação CSV
    # ------------------------------------------------------------------
    def exportar_csv(self):
        if not self.produtos:
            messagebox.showinfo("Sem dados", "Não há produtos para exportar.")
            return
        caminho = filedialog.asksaveasfilename(
            defaultextension=".csv", filetypes=[("Arquivo CSV", "*.csv"), ("Todos os arquivos", "*.*")],
            title="Exportar tabela como CSV", initialfile="precificacao_impressao_3d.csv",
        )
        if not caminho:
            return
        cabecalho = [TITULOS_COLUNAS[c] for c in COLUNAS_PRODUTOS]
        try:
            with open(caminho, mode="w", newline="", encoding="utf-8-sig") as arquivo:
                escritor = csv.writer(arquivo, delimiter=";")
                escritor.writerow(cabecalho)
                for produto_id in self.produtos:
                    escritor.writerow(self._valores_linha(produto_id))
            messagebox.showinfo("Exportado", f"Tabela exportada com sucesso para:\n{caminho}")
        except Exception as e:
            messagebox.showerror("Erro ao exportar", f"Não foi possível exportar o arquivo:\n{e}")

    # ------------------------------------------------------------------
    # Simulador de desconto em lote / atacado
    # ------------------------------------------------------------------
    def abrir_simulador_atacado(self):
        if not self.produtos:
            messagebox.showinfo("Sem produtos", "Cadastre ao menos um produto antes de usar o simulador.")
            return

        dlg = ctk.CTkToplevel(self)
        dlg.title("Simulador de Desconto em Lote / Atacado")
        dlg.geometry("460x460")
        dlg.configure(fg_color=COR_FUNDO)

        mapa_nome_id = {p["nome"]: pid for pid, p in self.produtos.items()}
        nomes = list(mapa_nome_id.keys())

        ctk.CTkLabel(dlg, text="Produto", text_color=COR_TEXTO_SECUNDARIO).pack(anchor="w", padx=20, pady=(20, 4))
        combo_produto = ctk.CTkComboBox(dlg, values=nomes, state="readonly")
        combo_produto.pack(fill="x", padx=20)
        if nomes:
            combo_produto.set(nomes[0])

        ctk.CTkLabel(dlg, text="Quantidade", text_color=COR_TEXTO_SECUNDARIO).pack(anchor="w", padx=20, pady=(15, 4))
        entrada_qtd = ctk.CTkEntry(dlg)
        entrada_qtd.insert(0, "10")
        entrada_qtd.pack(fill="x", padx=20)

        frame_presets = ctk.CTkFrame(dlg, fg_color="transparent")
        frame_presets.pack(fill="x", padx=20, pady=(6, 0))
        for valor in (10, 50, 100):
            ctk.CTkButton(
                frame_presets, text=str(valor), width=50,
                command=lambda v=valor: (entrada_qtd.delete(0, "end"), entrada_qtd.insert(0, str(v))),
                fg_color="#3a3a3a", hover_color="#4a4a4a",
            ).pack(side="left", padx=(0, 6))

        ctk.CTkLabel(dlg, text="Desconto Adicional no Preço (%)", text_color=COR_TEXTO_SECUNDARIO
                     ).pack(anchor="w", padx=20, pady=(15, 4))
        entrada_desconto = ctk.CTkEntry(dlg)
        entrada_desconto.insert(0, "0")
        entrada_desconto.pack(fill="x", padx=20)

        label_resultado = ctk.CTkLabel(dlg, text="", justify="left", font=ctk.CTkFont(size=13))
        label_resultado.pack(fill="x", padx=20, pady=(20, 0))

        def calcular():
            pid = mapa_nome_id.get(combo_produto.get())
            if pid is None:
                return
            p = self.produtos[pid]
            try:
                qtd = self._ler_campo_numerico(entrada_qtd.get(), "Quantidade", permitir_zero=False, minimo=1)
                desconto = self._ler_campo_numerico(entrada_desconto.get(), "Desconto", permitir_zero=True, minimo=0)
            except CampoInvalidoError as e:
                messagebox.showerror("Valor inválido", str(e))
                return
            preco_unit_final = p["preco_venda"] * (1 - desconto / 100.0)
            custo_unit = p["custo_total"]
            total = preco_unit_final * qtd
            custo_lote = custo_unit * qtd
            lucro_lote = total - custo_lote
            margem_efetiva = ((preco_unit_final - custo_unit) / custo_unit * 100.0) if custo_unit > 0 else 0.0
            label_resultado.configure(text=(
                f"Preço unitário com desconto: {formatar_moeda(preco_unit_final, self.moeda_atual)}\n"
                f"Total do lote ({int(qtd)} un.): {formatar_moeda(total, self.moeda_atual)}\n"
                f"Custo do lote: {formatar_moeda(custo_lote, self.moeda_atual)}\n"
                f"Lucro do lote: {formatar_moeda(lucro_lote, self.moeda_atual)}\n"
                f"Margem efetiva: {margem_efetiva:.1f}%"
            ))

        ctk.CTkButton(dlg, text="Calcular", command=calcular, fg_color=COR_DESTAQUE,
                      hover_color="#268a5e").pack(fill="x", padx=20, pady=(15, 20))
        dlg.transient(self)
        dlg.after(50, dlg.grab_set)

    # ------------------------------------------------------------------
    # Geração de orçamento em PDF
    # ------------------------------------------------------------------
    def abrir_dialogo_pdf(self):
        if not REPORTLAB_DISPONIVEL:
            messagebox.showerror("Biblioteca ausente",
                                  "Para gerar PDFs, instale a biblioteca reportlab:\n\npip install reportlab")
            return
        if not self.produtos:
            messagebox.showinfo("Sem produtos", "Cadastre ao menos um produto antes de gerar um orçamento.")
            return

        dlg = ctk.CTkToplevel(self)
        dlg.title("Gerar Orçamento em PDF")
        dlg.geometry("640x720")
        dlg.configure(fg_color=COR_FUNDO)

        ctk.CTkLabel(dlg, text="Dados do Cliente", font=ctk.CTkFont(size=15, weight="bold")
                     ).pack(anchor="w", padx=20, pady=(20, 10))
        entrada_cliente_nome = ctk.CTkEntry(dlg, placeholder_text="Nome do Cliente")
        entrada_cliente_nome.pack(fill="x", padx=20, pady=(0, 8))
        entrada_cliente_tel = ctk.CTkEntry(dlg, placeholder_text="Telefone / Contato")
        entrada_cliente_tel.pack(fill="x", padx=20, pady=(0, 8))
        entrada_validade = ctk.CTkEntry(dlg, placeholder_text="Validade da proposta (dias) — ex: 7")
        entrada_validade.insert(0, "7")
        entrada_validade.pack(fill="x", padx=20, pady=(0, 8))

        ctk.CTkLabel(dlg, text="Itens do Orçamento", font=ctk.CTkFont(size=15, weight="bold")
                     ).pack(anchor="w", padx=20, pady=(10, 10))
        frame_itens = ctk.CTkScrollableFrame(dlg, fg_color=COR_PAINEL, height=180)
        frame_itens.pack(fill="x", padx=20)

        checks, entradas_qtd = {}, {}
        for pid, p in self.produtos.items():
            linha = ctk.CTkFrame(frame_itens, fg_color="transparent")
            linha.pack(fill="x", pady=3)
            var = tk.BooleanVar(value=False)
            ctk.CTkCheckBox(linha, text=p["nome"], variable=var).pack(side="left")
            e_qtd = ctk.CTkEntry(linha, width=60)
            e_qtd.insert(0, "1")
            e_qtd.pack(side="right")
            ctk.CTkLabel(linha, text="Qtd:", text_color=COR_TEXTO_SECUNDARIO).pack(side="right", padx=(0, 6))
            checks[pid] = var
            entradas_qtd[pid] = e_qtd

        ctk.CTkLabel(dlg, text="Condições de Pagamento", text_color=COR_TEXTO_SECUNDARIO
                     ).pack(anchor="w", padx=20, pady=(15, 4))
        texto_condicoes = ctk.CTkTextbox(dlg, height=60)
        texto_condicoes.insert("1.0", "50% de entrada e 50% na entrega. Pagamento via Pix.")
        texto_condicoes.pack(fill="x", padx=20)

        ctk.CTkLabel(dlg, text="Observações", text_color=COR_TEXTO_SECUNDARIO
                     ).pack(anchor="w", padx=20, pady=(10, 4))
        texto_obs = ctk.CTkTextbox(dlg, height=60)
        texto_obs.pack(fill="x", padx=20, pady=(0, 10))

        def gerar():
            itens_selecionados = []
            for pid, var in checks.items():
                if var.get():
                    p = self.produtos[pid]
                    try:
                        qtd = self._ler_campo_numerico(entradas_qtd[pid].get(), f"Quantidade de {p['nome']}",
                                                         permitir_zero=False, minimo=1)
                    except CampoInvalidoError as e:
                        messagebox.showerror("Valor inválido", str(e))
                        return
                    itens_selecionados.append((p, qtd))
            if not itens_selecionados:
                messagebox.showwarning("Nenhum item", "Selecione ao menos um produto para o orçamento.")
                return

            caminho = filedialog.asksaveasfilename(
                defaultextension=".pdf", filetypes=[("Arquivo PDF", "*.pdf")],
                title="Salvar Orçamento em PDF", initialfile="orcamento.pdf",
            )
            if not caminho:
                return

            cliente = {
                "nome": entrada_cliente_nome.get().strip() or "Cliente",
                "telefone": entrada_cliente_tel.get().strip(),
                "validade": entrada_validade.get().strip() or "7",
            }
            try:
                self._gerar_pdf_orcamento(caminho, cliente, itens_selecionados,
                                           texto_condicoes.get("1.0", "end").strip(),
                                           texto_obs.get("1.0", "end").strip())
                messagebox.showinfo("PDF Gerado", f"Orçamento salvo em:\n{caminho}")
                dlg.destroy()
            except Exception as e:
                messagebox.showerror("Erro ao gerar PDF", str(e))

        ctk.CTkButton(dlg, text="📄  Gerar PDF", command=gerar, fg_color=COR_DESTAQUE,
                      hover_color="#268a5e", height=38).pack(fill="x", padx=20, pady=(5, 20))
        dlg.transient(self)
        dlg.after(50, dlg.grab_set)

    def _descrever_produto_para_pdf(self, p):
        partes = []
        for m in p["modulos"]:
            itens = []
            for it in m["filamentos"]:
                fil = self.filamentos.get(it["filamento_id"])
                if fil:
                    nome_cor = f"{fil['nome']} ({fil.get('cor', '-')})" if fil.get("cor") else fil["nome"]
                else:
                    nome_cor = "Filamento removido"
                itens.append(f"{nome_cor} {formatar_numero(it['peso_g'])}g")
            texto_itens = ", ".join(itens)
            if p.get("composto") and m.get("nome"):
                partes.append(f"<b>{m['nome']}:</b> {texto_itens}")
            else:
                partes.append(texto_itens)
        return "<br/>".join(partes) if partes else "-"

    def _gerar_pdf_orcamento(self, caminho, cliente, itens, condicoes, observacoes):
        styles = getSampleStyleSheet()
        estilo_titulo = ParagraphStyle("TituloEmpresa", parent=styles["Heading1"], fontSize=18,
                                        textColor=colors.HexColor("#0A2540"))
        estilo_normal = styles["Normal"]

        doc = SimpleDocTemplate(caminho, pagesize=A4, topMargin=1.5 * cm, bottomMargin=1.5 * cm,
                                 leftMargin=1.5 * cm, rightMargin=1.5 * cm)
        elementos = []

        nome_empresa = self.entrada_empresa_nome.get().strip() or "Minha Empresa"
        contato_empresa = self.entrada_empresa_contato.get().strip()

        elementos.append(Paragraph(nome_empresa, estilo_titulo))
        if contato_empresa:
            elementos.append(Paragraph(contato_empresa, estilo_normal))
        elementos.append(Spacer(1, 10))
        elementos.append(HRFlowable(width="100%", color=colors.HexColor("#14C7C2"), thickness=2))
        elementos.append(Spacer(1, 14))

        data_hoje = datetime.date.today().strftime("%d/%m/%Y")
        elementos.append(Paragraph(f"<b>Orçamento</b> — {data_hoje}", styles["Heading2"]))
        elementos.append(Spacer(1, 6))
        elementos.append(Paragraph(f"<b>Cliente:</b> {cliente['nome']}", estilo_normal))
        if cliente["telefone"]:
            elementos.append(Paragraph(f"<b>Contato:</b> {cliente['telefone']}", estilo_normal))
        elementos.append(Paragraph(f"<b>Validade da proposta:</b> {cliente['validade']} dia(s)", estilo_normal))
        elementos.append(Spacer(1, 14))

        dados_tabela = [["Item", "Detalhes", "Qtd", "Vlr. Unit.", "Subtotal"]]
        total_geral = 0.0
        for p, qtd in itens:
            detalhes = self._descrever_produto_para_pdf(p)
            subtotal = p["preco_venda"] * qtd
            total_geral += subtotal
            dados_tabela.append([
                Paragraph(p["nome"], estilo_normal), Paragraph(detalhes, estilo_normal), str(int(qtd)),
                formatar_moeda(p["preco_venda"], self.moeda_atual), formatar_moeda(subtotal, self.moeda_atual),
            ])

        tabela = Table(dados_tabela, colWidths=[3.2 * cm, 6.3 * cm, 1.3 * cm, 2.6 * cm, 2.6 * cm], repeatRows=1)
        tabela.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#0A2540")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("FONTSIZE", (0, 0), (-1, -1), 9),
            ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#cccccc")),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("ALIGN", (2, 0), (-1, -1), "CENTER"),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#f2f2f2")]),
        ]))
        elementos.append(tabela)
        elementos.append(Spacer(1, 10))
        elementos.append(Paragraph(
            f"<b>Total Geral: {formatar_moeda(total_geral, self.moeda_atual)}</b>",
            ParagraphStyle("Total", parent=styles["Heading2"], alignment=TA_RIGHT)
        ))
        elementos.append(Spacer(1, 20))

        if condicoes:
            elementos.append(Paragraph("<b>Condições de Pagamento</b>", styles["Heading3"]))
            elementos.append(Paragraph(condicoes.replace("\n", "<br/>"), estilo_normal))
            elementos.append(Spacer(1, 10))
        if observacoes:
            elementos.append(Paragraph("<b>Observações</b>", styles["Heading3"]))
            elementos.append(Paragraph(observacoes.replace("\n", "<br/>"), estilo_normal))

        doc.build(elementos)

    # ------------------------------------------------------------------
    # Persistência de dados (JSON local com backup automático)
    # ------------------------------------------------------------------
    def salvar_dados(self):
        if not self._pronto_para_salvar:
            return
        if os.path.exists(ARQUIVO_DADOS):
            try:
                shutil.copyfile(ARQUIVO_DADOS, ARQUIVO_BACKUP)
            except Exception:
                pass

        dados = {
            "proximo_id": self.proximo_id,
            "proximo_id_filamento": self.proximo_id_filamento,
            "proximo_id_insumo": self.proximo_id_insumo,
            "moeda": self.moeda_atual,
            "parametros": {chave: entrada.get() for chave, entrada in self.entradas_parametros.items()},
            "empresa": {
                "nome": self.entrada_empresa_nome.get() if hasattr(self, "entrada_empresa_nome") else "",
                "contato": self.entrada_empresa_contato.get() if hasattr(self, "entrada_empresa_contato") else "",
            },
            "categorias_margem": self.categorias_margem,
            "filamentos": {str(fid): dados_f for fid, dados_f in self.filamentos.items()},
            "insumos": {str(iid): dados_i for iid, dados_i in self.insumos.items()},
            "produtos": {str(pid): produto for pid, produto in self.produtos.items()},
        }
        try:
            with open(ARQUIVO_DADOS, "w", encoding="utf-8") as arquivo:
                json.dump(dados, arquivo, ensure_ascii=False, indent=2)
        except Exception as e:
            messagebox.showwarning("Falha ao salvar", f"Não foi possível salvar os dados automaticamente:\n{e}")

    def carregar_dados(self):
        dados = self._ler_json_com_fallback(ARQUIVO_DADOS, ARQUIVO_BACKUP)
        if dados is None:
            dados = self._importar_dados_v1_se_existir()
        if dados is None:
            self._atualizar_tabela_categorias()
            self._atualizar_opcoes_categoria_produto()
            self._atualizar_resumo()
            return

        parametros_salvos = dados.get("parametros", {})
        for chave, valor in parametros_salvos.items():
            entrada = self.entradas_parametros.get(chave)
            if entrada is not None:
                entrada.delete(0, "end")
                entrada.insert(0, valor)

        empresa = dados.get("empresa", {})
        if empresa.get("nome"):
            self.entrada_empresa_nome.insert(0, empresa["nome"])
        if empresa.get("contato"):
            self.entrada_empresa_contato.insert(0, empresa["contato"])

        self.moeda_atual = dados.get("moeda", "BRL")
        self.combo_moeda.set(MOEDA_CODIGO_PARA_DISPLAY.get(self.moeda_atual, MOEDA_CODIGO_PARA_DISPLAY["BRL"]))

        categorias_salvas = dados.get("categorias_margem")
        if categorias_salvas:
            self.categorias_margem = dict(categorias_salvas)
        self._atualizar_tabela_categorias()
        self._atualizar_opcoes_categoria_produto()

        filamentos_salvos = dados.get("filamentos", {})
        maior_id_fil = 0
        for id_texto, dados_f in filamentos_salvos.items():
            try:
                fid = int(id_texto)
            except (TypeError, ValueError):
                continue
            dados_f.setdefault("cor", "")
            dados_f.setdefault("peso_carretel_g", 1000.0)
            dados_f.setdefault("preco_kg", dados_f.get("preco_carretel", 0.0))
            self.filamentos[fid] = dados_f
            maior_id_fil = max(maior_id_fil, fid)
        self.proximo_id_filamento = dados.get("proximo_id_filamento", maior_id_fil + 1)
        self._atualizar_tabela_filamentos()
        self._atualizar_todas_opcoes_filamento()

        insumos_salvos = dados.get("insumos", {})
        maior_id_ins = 0
        for id_texto, dados_i in insumos_salvos.items():
            try:
                iid = int(id_texto)
            except (TypeError, ValueError):
                continue
            self.insumos[iid] = dados_i
            maior_id_ins = max(maior_id_ins, iid)
        self.proximo_id_insumo = dados.get("proximo_id_insumo", maior_id_ins + 1)
        self._atualizar_tabela_insumos()
        self.atualizar_opcoes_insumo_uso()

        produtos_salvos = dados.get("produtos", {})
        maior_id = 0
        for id_texto, produto in produtos_salvos.items():
            try:
                pid = int(id_texto)
            except (TypeError, ValueError):
                continue
            produto.setdefault("categoria", "Sem Categoria")
            produto.setdefault("tipo_margem", "personalizada")
            produto.setdefault("margem_personalizada", 0.0)
            produto.setdefault("markup_fixo", 0.0)
            produto.setdefault("composto", False)
            produto.setdefault("modulos", [])
            produto.setdefault("tempo_montagem_h", 0.0)
            produto.setdefault("custo_preparacao_mesa", 0.0)
            produto.setdefault("insumos_usados", [])
            produto.setdefault("custo_preparacao", 0.0)
            produto.setdefault("custo_insumos", 0.0)
            self.produtos[pid] = produto
            maior_id = max(maior_id, pid)
        self.proximo_id = dados.get("proximo_id", maior_id + 1)

        self._preencher_tabela()
        self._atualizar_resumo()
        self._limpar_formulario_produto()

    def _importar_dados_v1_se_existir(self):
        """Se não houver dados no novo formato mas existir o arquivo da
        versão anterior, importa parâmetros globais e filamentos (assumindo
        carretel de 1kg) para não obrigar o usuário a recadastrar tudo."""
        if not os.path.exists(ARQUIVO_DADOS_V1):
            return None
        try:
            with open(ARQUIVO_DADOS_V1, "r", encoding="utf-8") as arquivo:
                dados_v1 = json.load(arquivo)
        except Exception:
            return None

        dados_convertidos = {"parametros": dados_v1.get("parametros", {}), "filamentos": {}}
        for id_texto, dados_f in dados_v1.get("filamentos", {}).items():
            preco_kg = dados_f.get("preco_kg", 0.0)
            dados_convertidos["filamentos"][id_texto] = {
                "nome": dados_f.get("nome", "?"), "cor": "",
                "peso_carretel_g": 1000.0, "preco_carretel": preco_kg, "preco_kg": preco_kg,
            }
        messagebox.showinfo(
            "Dados importados",
            "Parâmetros globais e filamentos da versão anterior foram importados.\n"
            "Os produtos precisarão ser recadastrados na nova estrutura "
            "(com suporte a multicoloridos e módulos)."
        )
        return dados_convertidos

    @staticmethod
    def _ler_json_com_fallback(caminho_principal, caminho_backup):
        for caminho, e_backup in ((caminho_principal, False), (caminho_backup, True)):
            if not os.path.exists(caminho):
                continue
            try:
                with open(caminho, "r", encoding="utf-8") as arquivo:
                    dados = json.load(arquivo)
                if e_backup:
                    messagebox.showinfo(
                        "Recuperado do backup",
                        "O arquivo de dados principal estava corrompido.\n"
                        "Os dados foram recuperados a partir do backup automático."
                    )
                return dados
            except Exception:
                continue
        if os.path.exists(caminho_principal) or os.path.exists(caminho_backup):
            messagebox.showwarning(
                "Falha ao carregar",
                "Não foi possível ler os dados salvos (arquivo principal e backup "
                "estão corrompidos ou ilegíveis).\nO aplicativo será iniciado com "
                "os valores padrão."
            )
        return None


if __name__ == "__main__":
    app = CalculadoraImpressao3D()
    app.mainloop()
