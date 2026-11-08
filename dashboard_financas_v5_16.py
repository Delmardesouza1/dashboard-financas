#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Dashboard de finanças pessoais - versão 5.4.

COMO USAR: salve este arquivo como dashboard_financas_v5_16.py (troque a extensão .txt por .py)

Lê, da mesma pasta:
  - faturas do Itaú (.xlsx)                -> lançamentos do cartão
  - extratos do Nubank (.csv)              -> lançamentos do cartão
  - fotos de cupons fiscais (.jpg/.png...) -> itens do supermercado, categorizados por produto

Gera:
  - dashboard_financas.html      -> gráficos e tabelas (abre no navegador)
  - dashboard_financas.xlsx      -> dashboard em Excel com gráficos nativos
  - transacoes_consolidadas.csv  -> todos os lançamentos (abre no Excel em pt-BR)
  - ocr_<nome_da_foto>.txt       -> texto extraído de cada foto de cupom

Fluxo em duas etapas (novidade da versão 5.4):
  1) Extrair o texto das fotos:        python dashboard_financas_v5_16.py --so-texto
     Gera um ocr_<foto>.txt por foto. Você pode abrir esses arquivos e corrigir
     o que o OCR leu errado (ex.: um preço ou um nome de produto).
  2) Interpretar os textos e gerar o dashboard:
        python dashboard_financas_v5_16.py --usar-textos
     Usa os ocr_<foto>.txt já salvos, sem refazer o OCR.
  Sem opções (python dashboard_financas_v5_16.py): faz o OCR e o dashboard de uma vez.

Versão 5.16:
  - Correção: os cupons conferidos entram no dashboard mesmo que a foto não esteja na pasta.
    Antes, o script só usava cupons cuja foto estivesse presente com o mesmo nome, e por isso
    não gerava o dashboard quando as fotos tinham outros nomes.

Versão 5.15:
  - Correção dos avisos do Pylance ("Object of type None is not subscriptable") na verificação de
    duplicatas. Não muda o resultado do script.

Versão 5.14:
  - Verificação de duplicata pela data e hora de emissão do cupom (ex.: 29-10-26 17:14:37).
    Essa informação é mais estável no OCR do que nomes e valores.
  - Nova coluna 'emissao' em cupons_conferidos.csv. Se o arquivo antigo não tiver a coluna,
    o script adiciona, preenchendo os cupons já conhecidos.
  - Se a foto tiver a mesma emissão de um cupom já conferido, ela vira DUPLICATA e não vai para
    para_conferir.csv. Se não for possível ler a data e hora, a foto segue para conferência e o
    aviso aparece.
  - Resultado de cada foto salvo em comparacao_ocr.csv.

Versão 5.12:
  - Correção das palavras cortadas: cada faixa agora é lida com psm 6 (bloco de texto), e não
    mais com psm 7 (uma única linha). A faixa pode conter várias linhas, e o psm 7 lia só uma
    parte, cortando o resto.

Versão 5.11:
  - Quando não há nenhum lançamento confirmado (só fotos pendentes), o script não para mais
    com erro. Ele mostra o que falta fazer e termina normalmente, sem gerar o dashboard.
  - Mensagem orientando a aprovar as fotos em para_conferir.csv (status OK).

Versão 5.10:
  - Fluxo com conferência: fotos novas (sem linha em cupons_conferidos.csv) passam pelo OCR
    e viram linhas em para_conferir.csv, com status PENDENTE. Elas NÃO entram no dashboard.
  - Para aprovar: abra para_conferir.csv, corrija os nomes e troque o status para OK em
    todas as linhas da foto. Na próxima execução, a foto é movida para cupons_conferidos.csv
    e passa a entrar no dashboard e na planilha.
  - Edições feitas no para_conferir.csv são preservadas: o script não refaz a leitura de
    fotos que já estão na lista.
  - O código de barras de cada item é mantido ao longo do fluxo.

Versão 5.9:
  - Cupons conferidos manualmente: o arquivo cupons_conferidos.csv (criado na pasta das fotos)
    traz os itens exatamente como foram lidos na conferência. Para as fotos listadas nele,
    o OCR é pulado e os itens desse arquivo são usados. Assim os dados não dependem da leitura.
  - Para o modelo vir com os dois cupons de Andreazza, renomeie as fotos para
    andreazza_29-10-26.jpg e andreazza_28-10-26.jpg (ou ajuste a coluna 'arquivo' do CSV).
  - Cada linha do CSV é um item: valor = valor bruto da linha; desconto = valor a subtrair.
  - Se uma foto listada no CSV não estiver na pasta, o script avisa.

Versão 5.8:
  - Correção do aviso do Pylance '"image_to_string" is not a known attribute of "None"'
    na função de leitura por linhas (ocr_por_linhas), com a mesma verificação explícita
    que já existia no OCR. É um aviso do editor; o script já funcionava sem ele.

Versão 5.7:
  - Foto inclinada é endireitada antes do OCR (testa ângulos de -6° a +6°).
  - Filtro de ruído leve na foto antes da leitura.
  - A foto é cortada em faixas, uma por linha de texto, e cada faixa é lida separadamente.
    Cada linha do cupom vira uma linha do ocr_<foto>.txt, o que ajuda a separar nome e preço.
  - Se a leitura por linhas achar poucos preços, o script usa a leitura do bloco inteiro
    (os três modos de antes) e fica com a que achar mais preços.

Versão 5.6:
  - Catálogo de produtos (produtos_catalogo.csv, criado na pasta das fotos na primeira
    execução). O nome correto do produto é buscado pelo código de barras/código do item,
    então a descrição não depende da leitura do OCR.
  - Se o código não estiver no catálogo, o script tenta uma correspondência aproximada
    pelo nome. Se nada bater, o item sai como "nome lido (revisar)" e aparece nos avisos.
  - Preços com ponto decimal (ex.: "14.99") são lidos como "14,99".
  - Categoria vinda do catálogo tem prioridade sobre as regras por palavra.
  - Para incluir um produto novo, acrescente uma linha no CSV: código;produto;categoria.

Versão 5.5:
  - "Suvalan" (bebida) entra em Bebidas. A regra fica no topo da lista de produtos,
    antes de Hortifruti, porque "Suvalan Uva" e "Suvalan Laranja" teriam caído em
    Hortifruti pelas palavras "uva" e "laranja".
  - "Doguinho" entra em Carnes e aves.
  - Unidade "k" (além de "kg") é removida do nome do produto (ex.: "Donuts Chocolate k").

Versão 5.4:
  - Separação entre extrair o texto (OCR) e interpretar o texto (parser).
  - Opções --so-texto e --usar-textos.
  - Removida uma função antiga sem uso.
  - "Doguinho" saiu de Hortifruti (categoria incerta; agora cai em "Outros (mercado)").

Versão 5.3:
  - Itens em duas linhas, descontos por item, datas dd-mm-aa e total pela linha "Valor Total".

Versão 5.2:
  - Verificação explícita do pytesseract antes do OCR.

Versão 5.1:
  - Filtro de redimensionamento compatível com todas as versões do Pillow (LANCZOS).

Versão 5:
  - Diagnóstico do Tesseract no início (versão, caminho e idiomas instalados).
  - Se o idioma português não estiver instalado, usa inglês e avisa.
  - Tenta três modos de leitura do OCR e fica com o que achar mais preços.
  - Corrigido o erro KeyError 'parcelas_restantes' quando há só cupons.

Como funciona com o cupom:
  - Se a compra foi no cartão, ela já está na fatura. O cupom é casado com a linha
    da fatura (mesmo valor, data com até 3 dias de diferença e, de preferência,
    nome da loja na descrição) e a linha genérica é trocada pelos itens do cupom.
    Assim nada é contado duas vezes.
  - Se não houver fatura correspondente (ex.: pago em dinheiro ou Pix), os itens do
    cupom entram como lançamentos próprios.

Uso:
  python dashboard_financas_v5_16.py
  python dashboard_financas_v5_16.py --so-texto
  python dashboard_financas_v5_16.py --usar-textos
  python dashboard_financas_v5_16.py "C:\\entrada" "C:\\saida"
  (as opções podem ser combinadas com a pasta, ex.: --usar-textos "C:\\entrada")

Dependências:
  pip install pandas openpyxl plotly pytesseract pillow numpy
  + Tesseract OCR instalado no Windows, com o idioma português:
    https://github.com/UB-Mannheim/tesseract/wiki
    (durante a instalação, marque "Portuguese" em Additional language data)
"""

import csv
import difflib
import re
import shutil
import sys
import unicodedata
import warnings
from datetime import datetime
from pathlib import Path

# Bibliotecas obrigatórias: se faltar alguma, o script mostra o comando de instalação
# para o Python que está rodando e para antes de dar erro confuso.
DEPENDENCIAS_OBRIGATORIAS = {"pandas": "pandas", "openpyxl": "openpyxl",
                             "plotly": "plotly", "PIL": "pillow", "numpy": "numpy"}


def verificar_dependencias():
    faltando = []
    for modulo, pacote in DEPENDENCIAS_OBRIGATORIAS.items():
        try:
            __import__(modulo)
        except ImportError:
            faltando.append(pacote)
    if faltando:
        print("ERRO: bibliotecas não instaladas no Python que está rodando o script:")
        print(f"   {sys.executable}")
        print("Instale com este comando (cole no terminal do VS Code):")
        print(f'   "{sys.executable}" -m pip install {" ".join(faltando)}')
        print("Depois rode o script de novo.")
        sys.exit(1)


verificar_dependencias()

import numpy as np  # noqa: E402
import pandas as pd  # noqa: E402
import openpyxl  # noqa: E402
import plotly.express as px  # noqa: E402  # pyright: ignore[reportMissingImports]
import plotly.graph_objects as go  # noqa: E402  # pyright: ignore[reportMissingImports]
from PIL import Image, ImageFilter, ImageOps  # noqa: E402
from openpyxl import Workbook  # noqa: E402
from openpyxl.styles import Font, PatternFill, Alignment  # noqa: E402
from openpyxl.chart import PieChart, BarChart, Reference  # noqa: E402
from openpyxl.utils import get_column_letter  # noqa: E402

warnings.filterwarnings("ignore", message="Workbook contains no default style")

try:
    import pytesseract  # pyright: ignore[reportMissingImports]
except ImportError:  # o restante do script funciona sem OCR
    pytesseract = None

# Filtro de redimensionamento compatível com todas as versões do Pillow
# (versões novas usam Image.Resampling.LANCZOS; as antigas usam Image.LANCZOS).
FILTRO_LANCZOS = getattr(getattr(Image, "Resampling", Image), "LANCZOS")
FILTRO_BICUBIC = getattr(getattr(Image, "Resampling", Image), "BICUBIC")


# ---------------------------------------------------------------
# Configurações
# ---------------------------------------------------------------
PASTA_PADRAO = Path(r"C:\Users\delma\OneDrive\Desktop\finanças")

# Caminho do executável do Tesseract no Windows (ajuste se instalou em outro lugar).
# Se não existir nesse caminho, o script procura "tesseract" no PATH do sistema.
CAMINHO_TESSERACT = Path(r"C:\Program Files\Tesseract-OCR\tesseract.exe")
IDIOMA_OCR = "por"
IDIOMA_FALLBACK = "eng"
LINK_TESSERACT = "https://github.com/UB-Mannheim/tesseract/wiki"

# Modos de leitura testados em ordem; fica com o que encontrar mais preços.
CONFIGS_OCR = ["--psm 6", "--psm 4", "--psm 11"]
LARGURA_MINIMA_OCR = 1800  # fotos menores são ampliadas antes do OCR

EXTENSOES_IMAGEM = {".jpg", ".jpeg", ".png", ".bmp", ".tif", ".tiff", ".webp"}

# Categoria especial: pagamentos de fatura não são gasto (evita dupla contagem).
CATEGORIA_PAGAMENTO_FATURA = "Pagamento de fatura"

# Colunas que o restante do script espera existir, mesmo com a base vazia.
COLUNAS_OBRIGATORIAS = [
    "data", "descricao", "parcelamento", "valor", "cartao", "tipo_cartao",
    "vencimento", "arquivo", "categoria", "eh_gasto", "origem", "loja",
    "parcela_atual", "parcela_total", "parcelas_restantes", "saldo_restante",
]

# Categorias de lançamentos da fatura/extrato. A primeira regra que bater vence.
# Os padrões são aplicados ao texto sem acentos e em minúsculas.
REGRAS_CATEGORIA_BANCO = [
    (CATEGORIA_PAGAMENTO_FATURA, r"pagamento (de |do )?fatura|pag(to|amento)? (da |de )?fatura|pagamento (do )?cartao"),
    ("Pix no crédito",        r"pix no credito"),
    ("Transporte (apps)",     r"\bridesao\b|\buber\b|\b99\b"),
    ("Assinaturas/streaming", r"\bgoogle claude\b|\banthropic\b|\byoutube\b|\bnetflix\b|\bspotify\b"),
    ("Saúde",                 r"\bbradsaude|\bodonto|\bpanvel\b|\braia\b|\bfarma"),
    ("Telefonia",             r"\b(vivo|claro|tim|oi)\b"),
    ("Restaurantes/lanches",  r"\bchurrasc|\brestaur|\bmcdonald|\bsushi\b|\bbiglanch|\bburger|\bpadaria\b|\bcafe\b|\blanch|\bpizza\b|\bru\b"),  # confirme se \bru\b é o restaurante universitário
    ("Mercado/conveniência",  r"\bandreazza|\bmacromix|\balimen|\bpkr\b|\bgelb|\bmercad|\bsupermerc|\batacad"),
    ("Lazer",                 r"\barena\b|\btrampolim|\bgames?\b|\bcinema|\bjogos\b"),
    ("Compras/eletrônicos",   r"planeta do celular|\bcelular|\boboticario|\bbambin|\bmagazine|\bloja\b"),
]

# Categorias de PRODUTO (usadas nos itens dos cupons). Ordem importa.
REGRAS_PRODUTO = [
    ("Bebidas",             r"\bsuvalan\b"),  # antes de Hortifruti: "uva" e "laranja" estão no nome
    ("Pet",                 r"\bracao\b|\bpet\b|areia gato"),
    ("Congelados",          r"\bcongelad|\bsorvete|\bnugget|\bempanad|\bpizza\b"),
    ("Doces e snacks",      r"\bchocolate|\bbombom|\bsalgadinho|\bbatata palha|\bdoce\b|\bchiclete|\bbala\b|\bpipoca|\brecheado|\bbis\b|\btrident|\bcookie|\bwafer|\bdonut"),
    ("Hortifruti",          r"\bbanana|\bmaca\b|\btomate|\bcebola|\balface|\bbatata\b(?! palha)|\blaranja|\blimao|\bfruta|\blegume|\bverdura|\bcenoura|\bmamao|\bmelancia|\babacaxi|\buva\b|\bmorango|\bpepino|\babobora|\bmandioca"),
    ("Carnes e aves",       r"\bdoguinho|\bcarne|\bpatinho|\bpicanha|\bfrango|\bpeixe|\bbovin|\bsuin|\blinguica|\bsalsicha|\bbacon|\bfile\b|\bcostela|\bacem\b|\bcoxa|\bsobrecoxa|\balmondega|\bhamburguer|\blombo|\bperu\b"),
    ("Laticínios e frios",  r"\bleite|\bqueijo|\biogurte|\bmanteiga|\brequeijao|\bpresunto|\bmussarela|\bnata\b|\bcreme de leite|\bmargarina|\bmortadela|\blacteo|\biog\b|\bcre\b|\bcreme\b"),
    ("Padaria",             r"\bpao\b|\bpaes\b|\bbisnaga|\bcroissant|\bbolo\b|\brosca|\btorrada"),
    ("Bebidas",             r"\brefri|\bcoca\b|\bguarana|\bsuco\b|\bcerveja|\bagua\b|\bvinho|\benergetico|\bcha pronto|\bisot|\brefresco"),
    ("Mercearia",           r"\barroz|\bfeijao|\bmacarrao|\boleo\b|\bacucar|\bsal\b|\bfarinha|\bcafe\b|\bmolho|\btempero|\bextrato de tomate|\bbiscoito|\bbolacha|\bcereal|\baveia|\bsardinha|\batum\b|\bmilho|\bfuba|\bvinagre|\bmaionese|\bcatchup|\bketchup|\bbisc\b"),
    ("Limpeza",             r"\bsabao\b|\bdetergente|\bdesinfetante|\bamaciante|\bsanitaria|\besponja|\bpapel toalha|\balvejante|\blimpador|\bmultiuso|\bpano\b|\bvassoura|\bsaco de lixo|\blixo\b"),
    ("Higiene e beleza",    r"\bshampoo|\bcondicionador|\bsabonete|\bcreme dental|\bpasta de dente|\bescova de dente|\bdesodorante|\bpapel higienico|\babsorvente|\bfralda|\bbarbear|\blenco|\bhidratante"),
]
CATEGORIA_PRODUTO_PADRAO = "Outros (mercado)"

# Categorias que costumam ser cortáveis (aparecem em "onde dá para economizar").
CATEGORIAS_ECONOMIZAVEIS = {
    "Transporte (apps)",
    "Restaurantes/lanches",
    "Mercado/conveniência",
    "Pix no crédito",
    "Assinaturas/streaming",
    "Lazer",
    "Doces e snacks",
    "Bebidas",
    "Congelados",
    "Padaria",
}


# ---------------------------------------------------------------
# Utilitários
# ---------------------------------------------------------------
def normalizar(texto: str) -> str:
    """Remove acentos e põe em minúsculas, para comparar palavras-chave."""
    texto = unicodedata.normalize("NFKD", str(texto))
    return texto.encode("ascii", "ignore").decode("ascii").lower()


def parse_brl(valor) -> float:
    """Converte '1.234,56', '184,90', '- 71,13' ou 12.5 em float."""
    if isinstance(valor, (int, float)):
        return float(valor)
    s = str(valor).strip().replace("R$", "").replace(" ", "")
    s = s.replace(".", "").replace(",", ".")
    return float(s)


def valor_numerico(valor):
    """
    Converte o conteúdo de uma célula do Excel em float.
    Devolve None para células que não são valores (fórmulas, textos sem número, vazias).
    """
    if valor is None:
        return None
    if isinstance(valor, bool):
        return None
    if isinstance(valor, (int, float)):
        return float(valor)
    if isinstance(valor, str):
        try:
            return parse_brl(valor)
        except ValueError:
            return None
    return None  # fórmulas (ArrayFormula, DataTableFormula) e outros tipos


def brl(v) -> str:
    """Formata número como moeda brasileira: R$ 1.234,56."""
    if v is None or pd.isna(v):
        return "-"
    s = f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"R$ {s}"


def pct(v) -> str:
    return f"{v:.1f}%".replace(".", ",")


def garantir_colunas(df: pd.DataFrame) -> pd.DataFrame:
    """Garante que todas as colunas esperadas existam, mesmo se a base estiver vazia."""
    df = df.copy()
    for col in COLUNAS_OBRIGATORIAS:
        if col not in df.columns:
            df[col] = pd.Series(dtype="object")
    return df


def categorizar_banco(descricao: str) -> str:
    d = normalizar(descricao)
    for categoria, padrao in REGRAS_CATEGORIA_BANCO:
        if re.search(padrao, d):
            return categoria
    return "Outros"


def categorizar_produto(nome: str) -> str:
    n = normalizar(nome)
    for categoria, padrao in REGRAS_PRODUTO:
        if re.search(padrao, n):
            return categoria
    return CATEGORIA_PRODUTO_PADRAO


# Parcela: "10 de 12", "3/10" ou "Parcela 3/10". Não pega datas como "02/10/2026".
RE_PARCELA = re.compile(r"(?<![\d/])(\d{1,2})\s*(?:de|/)\s*(\d{1,2})(?![\d/])", re.IGNORECASE)


def extrair_parcela(parcelamento, descricao):
    """Retorna (parcela_atual, parcela_total) se houver, senão (None, None)."""
    for texto in (str(parcelamento or ""), str(descricao or "")):
        for m in RE_PARCELA.finditer(texto):
            atual, total = int(m.group(1)), int(m.group(2))
            if 1 <= atual <= total and total >= 2:
                return atual, total
    return None, None


# ---------------------------------------------------------------
# Leitura das faturas e extratos
# ---------------------------------------------------------------
def ler_fatura_itau(caminho: Path) -> pd.DataFrame:
    """Lê uma fatura aberta do Itaú (.xlsx) e devolve os lançamentos."""
    wb = openpyxl.load_workbook(caminho, data_only=True)
    ws = wb.worksheets[0]
    linhas = list(ws.iter_rows(values_only=True))

    # Cabeçalho do cartão: "Cartão" na coluna B, valores na linha seguinte.
    nome_cartao, valor_parcial = None, None
    for i, row in enumerate(linhas):
        if row[1] == "Cartão":
            nome_cartao = str(linhas[i + 1][1])
            valor_parcial = valor_numerico(linhas[i + 1][6])
            break
    m = re.search(r"final\s+(\d{4})", nome_cartao or "")
    rotulo_cartao = f"Itaú final {m.group(1)}" if m else (nome_cartao or caminho.stem)

    # Cabeçalho da tabela de lançamentos.
    inicio = None
    for i, row in enumerate(linhas):
        if row[1] == "Data" and row[2] == "Lançamento":
            inicio = i + 1
            break
    if inicio is None:
        raise ValueError(f"Tabela de lançamentos não encontrada em {caminho.name}")

    registros = []
    ignorados = 0
    for row in linhas[inicio:]:
        if row[1] == "Importante saber":
            break
        if not isinstance(row[1], datetime):
            continue  # subtotal, linha vazia, etc.
        if row[4] is None:
            continue
        valor = valor_numerico(row[4])
        if valor is None:
            ignorados += 1  # célula com fórmula ou texto que não é número
            continue
        registros.append({
            "data": row[1],
            "descricao": str(row[2] or "").strip(),
            "parcelamento": str(row[3] or "").strip(),
            "valor": valor,
            "cartao": rotulo_cartao,
            "tipo_cartao": str(row[8] or "").strip(),
            "vencimento": None,
            "arquivo": caminho.name,
        })

    if ignorados:
        print(f"AVISO: {caminho.name}: {ignorados} linha(s) com valor que não é número foram ignoradas. "
              f"Confira essas linhas na planilha.")

    df = pd.DataFrame(registros)
    if valor_parcial is not None and not df.empty:
        soma = df.loc[df["valor"] > 0, "valor"].sum()
        if abs(soma - valor_parcial) > 0.05:
            print(f"AVISO: {caminho.name}: soma dos lançamentos ({brl(soma)}) "
                  f"difere do valor da fatura ({brl(valor_parcial)}).")
    return df


def ler_nubank(caminho: Path) -> pd.DataFrame:
    """Lê um extrato do Nubank (.csv) com colunas date, title, amount."""
    df_raw = pd.read_csv(caminho, dtype=str)
    df_raw.columns = [c.strip().lower() for c in df_raw.columns]
    return pd.DataFrame({
        "data": pd.to_datetime(df_raw["date"]),
        "descricao": df_raw["title"].str.strip(),
        "parcelamento": "",
        "valor": df_raw["amount"].map(parse_brl),
        "cartao": "Nubank",
        "tipo_cartao": "",
        "vencimento": None,
        "arquivo": caminho.name,
    })


def carregar_lancamentos(pasta: Path) -> pd.DataFrame:
    frames = []
    for f in sorted(pasta.glob("*.xlsx")):
        if f.name.startswith("~$"):
            continue
        if "fatura" in f.name.lower():
            frames.append(ler_fatura_itau(f))
    for f in sorted(pasta.glob("*.csv")):
        if "nubank" in f.name.lower():
            frames.append(ler_nubank(f))

    frames = [f for f in frames if not f.empty]
    if not frames:
        print("Nenhuma fatura (.xlsx) ou extrato Nubank (.csv) encontrado. Seguindo só com cupons.")
        return garantir_colunas(pd.DataFrame())
    df = pd.concat(frames, ignore_index=True)
    df["data"] = pd.to_datetime(df["data"])
    return df


def classificar_lancamentos(df: pd.DataFrame) -> pd.DataFrame:
    df = garantir_colunas(df)
    if df.empty:
        return df
    df["categoria"] = df["descricao"].map(categorizar_banco)
    df.loc[df["valor"] < 0, "categoria"] = "Pagamento/crédito"
    # Gasto = valor positivo que não seja pagamento de fatura (evita contar duas vezes).
    df["eh_gasto"] = (df["valor"] > 0) & (df["categoria"] != CATEGORIA_PAGAMENTO_FATURA)
    df["origem"] = "Fatura/extrato"
    df["loja"] = ""

    partes = [extrair_parcela(p, d) for p, d in zip(df["parcelamento"], df["descricao"])]
    df["parcela_atual"] = pd.to_numeric(pd.Series([t[0] for t in partes], index=df.index),
                                        errors="coerce")
    df["parcela_total"] = pd.to_numeric(pd.Series([t[1] for t in partes], index=df.index),
                                        errors="coerce")
    df["parcelas_restantes"] = (df["parcela_total"] - df["parcela_atual"]).fillna(0).clip(lower=0)
    df["saldo_restante"] = df["valor"] * df["parcelas_restantes"]
    return df


# ---------------------------------------------------------------
# Cupons fiscais: etapa 1 (extrair texto das fotos)
# ---------------------------------------------------------------
RE_PRECO = re.compile(r"(?<![\d,.])(-?\d{1,3}(?:\.\d{3})*,\d{2})(?!\d)")


def preparar_ocr():
    """
    Localiza o Tesseract, mostra versão e idiomas e escolhe o idioma a usar.
    Devolve o código do idioma (ex.: 'por') ou None se o OCR não puder rodar.
    """
    if pytesseract is None:
        print("AVISO: pytesseract não instalado. Os cupons serão ignorados.")
        print("   Instale com: pip install pytesseract")
        return None

    caminho = None
    if CAMINHO_TESSERACT.exists():
        caminho = str(CAMINHO_TESSERACT)
    else:
        achado = shutil.which("tesseract")
        if achado:
            caminho = achado
    if caminho is None:
        print(f"AVISO: Tesseract não encontrado em {CAMINHO_TESSERACT} nem no PATH.")
        print(f"   Instale: {LINK_TESSERACT}")
        print("   Se instalou em outro lugar, ajuste CAMINHO_TESSERACT no topo do script.")
        return None

    pytesseract.pytesseract.tesseract_cmd = caminho
    try:
        versao = pytesseract.get_tesseract_version()
        idiomas = pytesseract.get_languages(config="")
    except Exception as e:
        print(f"AVISO: o Tesseract não respondeu ({caminho}): {e}")
        return None

    print(f"Tesseract {versao} em: {caminho}")
    print(f"Idiomas instalados: {', '.join(idiomas)}")
    if IDIOMA_OCR in idiomas:
        return IDIOMA_OCR
    print(f"AVISO: idioma '{IDIOMA_OCR}' não está instalado no Tesseract. "
          f"Usando '{IDIOMA_FALLBACK}' (acentos e nomes de loja podem sair errados).")
    print(f"   Para corrigir: reinstale o Tesseract marcando 'Portuguese' em "
          f"Additional language data ({LINK_TESSERACT})")
    if IDIOMA_FALLBACK in idiomas:
        return IDIOMA_FALLBACK
    print("AVISO: nem o português nem o inglês estão instalados. Os cupons serão ignorados.")
    return None


def preparar_imagem(caminho: Path):
    """Abre a foto, corrige rotação, deixa em cinza, aumenta contraste e amplia se pequena."""
    img = Image.open(caminho)
    img = ImageOps.exif_transpose(img)          # corrige foto girada do celular
    img = ImageOps.grayscale(img)
    if img.width < LARGURA_MINIMA_OCR:          # fotos pequenas lêem melhor ampliadas
        escala = LARGURA_MINIMA_OCR / img.width
        img = img.resize((int(img.width * escala), int(img.height * escala)), FILTRO_LANCZOS)
    img = ImageOps.autocontrast(img, cutoff=1)
    img = img.filter(ImageFilter.MedianFilter(3))  # remove pontos de ruído da foto
    return img


def limiar_otsu(arr) -> int:
    """Limiar automático (Otsu) para separar tinta (escuro) do papel (claro)."""
    hist = np.bincount(arr.ravel(), minlength=256).astype(float)
    total = hist.sum()
    soma_tudo = float(np.dot(np.arange(256), hist))
    soma_b, peso_b = 0.0, 0.0
    melhor, limiar = -1.0, 127
    for t in range(256):
        peso_b += hist[t]
        if peso_b == 0:
            continue
        peso_f = total - peso_b
        if peso_f == 0:
            break
        soma_b += t * hist[t]
        media_b = soma_b / peso_b
        media_f = (soma_tudo - soma_b) / peso_f
        variancia = peso_b * peso_f * (media_b - media_f) ** 2
        if variancia > melhor:
            melhor, limiar = variancia, t
    return limiar


def corrigir_inclinacao(img):
    """Gira a foto para que as linhas de texto fiquem horizontais (testa de -6° a +6°)."""
    pequena = img.resize((max(1, img.width // 2), max(1, img.height // 2)), FILTRO_LANCZOS)
    melhor_ang, melhor_score = 0.0, -1.0
    for ang in np.arange(-6.0, 6.01, 0.5):
        girada = np.asarray(pequena.rotate(float(ang), resample=FILTRO_BICUBIC, fillcolor=255),
                            dtype=np.uint8)
        tinta = girada <= limiar_otsu(girada)
        # Linhas retas e horizontais dão um perfil de tinta com picos bem marcados (variância alta).
        score = float(np.var(tinta.sum(axis=1)))
        if score > melhor_score:
            melhor_ang, melhor_score = float(ang), score
    if abs(melhor_ang) < 0.25:
        return img
    print(f"   (foto inclinada: corrigida em {melhor_ang:.1f}°)")
    return img.rotate(melhor_ang, resample=FILTRO_BICUBIC, fillcolor=255)


def fatiar_linhas(img):
    """Corta a foto em faixas horizontais, uma por linha de texto."""
    arr = np.asarray(img, dtype=np.uint8)
    tinta = arr <= limiar_otsu(arr)
    perfil = tinta.sum(axis=1)
    minimo_tinta = max(3, int(arr.shape[1] * 0.004))  # ignora pontos isolados
    ativo = perfil > minimo_tinta
    faixas, ini = [], None
    for y, tem_tinta in enumerate(ativo):
        if tem_tinta and ini is None:
            ini = y
        elif not tem_tinta and ini is not None:
            faixas.append((ini, y))
            ini = None
    if ini is not None:
        faixas.append((ini, len(ativo)))
    altura_minima = max(12, int(arr.shape[0] * 0.006))
    margem = 6
    recortes = []
    for a, b in faixas:
        if b - a < altura_minima:
            continue
        recortes.append(img.crop((0, max(0, a - margem), img.width, min(img.height, b + margem))))
    return recortes


def ocr_por_linhas(img, idioma: str) -> str:
    """Lê cada faixa com psm 6 (bloco de texto), para não cortar as linhas dentro da faixa."""
    if pytesseract is None:  # confirmação explícita; preparar_ocr() já checou isso
        raise RuntimeError("pytesseract não está carregado")
    linhas = []
    for recorte in fatiar_linhas(img):
        texto = pytesseract.image_to_string(recorte, lang=idioma, config="--psm 6").strip()
        if texto:
            linhas.append(texto)
    return "\n".join(linhas)


def ocr_melhor(caminho: Path, idioma: str) -> str:
    """
    Versão 5.7: endireita a foto e lê linha por linha. Se a leitura por linhas
    achar poucos preços, usa a leitura do bloco inteiro e fica com a que achar mais preços.
    """
    if pytesseract is None:  # confirmação explícita; preparar_ocr() já checou isso
        raise RuntimeError("pytesseract não está carregado")
    img = corrigir_inclinacao(preparar_imagem(caminho))
    texto_linhas = ocr_por_linhas(img, idioma)
    pontos_linhas = len(RE_PRECO.findall(texto_linhas))
    if pontos_linhas >= 4:
        return texto_linhas
    texto_bloco = ocr_varias_configs(caminho, idioma)
    if pontos_linhas > len(RE_PRECO.findall(texto_bloco)):
        return texto_linhas
    return texto_bloco


def ocr_varias_configs(caminho: Path, idioma: str) -> str:
    """
    Roda o OCR em mais de um modo de leitura e devolve o texto com mais preços encontrados.
    Para cedo se já achar um bom número de preços.
    """
    if pytesseract is None:  # confirmação explícita; preparar_ocr() já checou isso
        raise RuntimeError("pytesseract não está carregado")
    img = preparar_imagem(caminho)
    melhor_texto, melhor_pontos = "", -1
    for cfg in CONFIGS_OCR:
        texto = pytesseract.image_to_string(img, lang=idioma, config=cfg)
        pontos = len(RE_PRECO.findall(texto))
        if pontos > melhor_pontos:
            melhor_texto, melhor_pontos = texto, pontos
        if pontos >= 4:
            break
    return melhor_texto


def listar_imagens(pasta: Path) -> list:
    """Fotos de cupom da pasta, em ordem alfabética."""
    return sorted(f for f in pasta.iterdir()
                  if f.is_file() and f.suffix.lower() in EXTENSOES_IMAGEM)


def caminho_texto_ocr(img: Path, saida: Path) -> Path:
    """Arquivo .txt onde fica o texto extraído de uma foto."""
    return saida / f"ocr_{img.stem}.txt"


def extrair_textos_cupons(pasta: Path, saida: Path, usar_salvos: bool = False) -> dict:
    """
    Etapa 1: extrai o texto de todas as fotos de cupom e salva em ocr_<foto>.txt.
    Se usar_salvos=True, não refaz o OCR de fotos que já têm o .txt salvo.
    Devolve {nome_da_foto: texto}.
    """
    textos = {}
    imagens = listar_imagens(pasta)
    if not imagens:
        print("Nenhuma foto de cupom (.jpg, .png...) encontrada na pasta.")
        return textos

    print(f"Fotos de cupom encontradas: {len(imagens)}")
    idioma, ocr_preparado = None, False
    for img in imagens:
        if img.name in CONFERIDOS:  # já conferida manualmente; não passa pelo OCR
            continue
        arq = caminho_texto_ocr(img, saida)
        if usar_salvos and arq.exists():
            textos[img.name] = arq.read_text(encoding="utf-8")
            print(f"{img.name}: usando texto já salvo ({arq.name})")
            continue

        if not ocr_preparado:
            idioma = preparar_ocr()
            ocr_preparado = True
        if idioma is None:
            print(f"AVISO: OCR indisponível; {img.name} não foi lido.")
            continue

        print(f"\nLendo {img.name} ...")
        try:
            texto = ocr_melhor(img, idioma)
        except Exception as e:  # imagem ilegível, Tesseract travou, etc.
            print(f"AVISO: não consegui ler {img.name}: {e}")
            continue

        arq.write_text(texto, encoding="utf-8")
        print(f"   Texto salvo em: {arq.name}")
        textos[img.name] = texto
    return textos


# ---------------------------------------------------------------
# Cupons fiscais: etapa 2 (interpretar o texto)
# ---------------------------------------------------------------

# Linha de quantidade e preços de um item de NFC-e, ex.: "1 Un 2,49 2,49" ou "0,160 Kg 29,90 4,78".
RE_QTD_PRECOS = re.compile(
    r"^\s*\d+(?:[.,]\d+)?\s*(?:[A-Za-z]{1,3})?\s+"
    r"(\d{1,3}(?:\.\d{3})*,\d{2})\s+"
    r"(\d{1,3}(?:\.\d{3})*,\d{2})\s*$")

# Data do cupom: 29-10-26, 29/10/26 ou 29/10/2026.
RE_DATA_CUPOM = re.compile(r"(?<!\d)(\d{2})[/-](\d{2})[/-](\d{4}|\d{2})(?!\d)")

# Linha que traz o total pago: "Valor Total R$", "Valor pago", "Total R$".
RE_TOTAL_LINHA = re.compile(r"va?lor\s+\w?otal|total\s*r\$|valor\s+pago", re.IGNORECASE)

# Desconto aplicado a um item: "PROD. COM DESCONTO ITEM 008 APLIC. SUBTOTAL -3,31".
RE_DESC_ITEM = re.compile(r"desconto\s+item\s+(\d{1,3})")

# Número do item no início da linha, com código de barras opcional.
RE_ID_ITEM = re.compile(r"^(\d{1,3})\s+(?:(\d{8,14})\s+)?(.*)$")

# Linhas que indicam o nome da loja e não são a loja em si.
RE_LOJA_EXCLUIR = (r"cnpj|cupom|fiscal|extrato|nfc|endereco|\brua\b|\bav\b|documento|"
                   r"auxiliar|fone|controle|eletronica|consumidor|imperatriz|leopoldo")


# ---------------------------------------------------------------
# Catálogo de produtos (nome correto por código)
# ---------------------------------------------------------------
NOME_CATALOGO = "produtos_catalogo.csv"
CATALOGO_INICIAL = [  # (código, nome correto, categoria) - conferidos na foto
    ("7898003541605", "Suvalan 200ml Uva", "Bebidas"),
    ("7898003541612", "Suvalan 200ml Laranja", "Bebidas"),
    ("0000000000017", "Pão Francês", "Padaria"),
    ("0000000003049", "Banana Caturra", "Hortifruti"),
    ("0000000003483", "Doguinho", "Carnes e aves"),
    ("0000000020039", "Donuts Chocolate", "Doces e snacks"),
]
CATALOGO = {}  # código normalizado -> (nome, categoria ou None)


def normalizar_codigo(codigo) -> str:
    """Só números e sem zeros à esquerda, para 0000000003483 e 3483 baterem."""
    digitos = re.sub(r"\D", "", str(codigo or ""))
    return str(int(digitos)) if digitos else ""


def carregar_catalogo(pasta: Path):
    """Lê produtos_catalogo.csv (cria com os produtos iniciais se não existir)."""
    arq = pasta / NOME_CATALOGO
    if not arq.exists():
        with arq.open("w", encoding="utf-8-sig", newline="") as f:
            w = csv.writer(f, delimiter=";")
            w.writerow(["codigo", "produto", "categoria"])
            for cod, nome, cat in CATALOGO_INICIAL:
                w.writerow([cod, nome, cat])
        print(f"Catálogo criado: {arq.name} (acrescente produtos novos nele)")
    CATALOGO.clear()
    with arq.open(encoding="utf-8-sig", newline="") as f:
        for row in csv.DictReader(f, delimiter=";"):
            cod = normalizar_codigo(row.get("codigo"))
            nome = (row.get("produto") or "").strip()
            cat = (row.get("categoria") or "").strip() or None
            if cod and nome:
                CATALOGO[cod] = (nome, cat)
    print(f"Produtos no catálogo: {len(CATALOGO)}")


def resolver_produto(codigo, nome_ocr: str):
    """
    Devolve (nome, categoria_ou_None, origem).
    origem: 'catálogo', 'aproximado' ou 'não reconhecido'.
    """
    cod = normalizar_codigo(codigo)
    if cod and cod in CATALOGO:
        nome, cat = CATALOGO[cod]
        return nome, cat, "catálogo"
    if CATALOGO and nome_ocr:
        por_chave = {normalizar(n): (n, c) for n, c in CATALOGO.values()}
        achados = difflib.get_close_matches(normalizar(nome_ocr), list(por_chave), n=1, cutoff=0.8)
        if achados:
            nome, cat = por_chave[achados[0]]
            return nome, cat, "aproximado"
    return f"{nome_ocr} (revisar)", None, "não reconhecido"


def limpar_linha_ocr(linha: str) -> str:
    # Erros comuns do OCR: letra O no lugar de zero dentro de números.
    linha = re.sub(r"(?<=[\d,])[Oo](?=[\d,])", "0", linha)
    # Preço com ponto decimal (14.99) vira 14,99. Não mexe em códigos longos.
    linha = re.sub(r"(?<![\d.])(\d{1,4})\.(\d{2})(?!\d)", r"\1,\2", linha)
    return linha.strip()


def separar_item(texto: str):
    """
    Separa 'NNN CODIGO DESCRICAO' em (numero_do_item, codigo_de_barras, descricao).
    Remove unidade (Kg, Un, k) e espaços sobrando.
    """
    m = RE_ID_ITEM.match(texto.strip())
    if m:
        numero, codigo, resto = int(m.group(1)), m.group(2) or "", m.group(3)
    else:
        numero, codigo, resto = None, "", texto.strip()
    resto = re.sub(r"\s*R\$\s*$", "", resto)
    resto = re.sub(r"\s+(kg|k|un|pc)\.?$", "", resto, flags=re.IGNORECASE)
    return numero, codigo, resto.strip(" -")


def parse_recibo(texto: str, nome_arquivo: str, data_padrao: datetime) -> dict:
    linhas = [limpar_linha_ocr(l) for l in texto.splitlines()]
    linhas = [l for l in linhas if l.strip()]
    avisos_leitura = []

    # Nome da loja: primeira linha com palavras e sem cabeçalhos fiscais.
    loja = Path(nome_arquivo).stem
    for l in linhas:
        if re.search(r"[A-Za-z]{4,}", l) and not re.search(RE_LOJA_EXCLUIR, l, re.IGNORECASE):
            loja = l.strip()
            break

    # Data da compra: aceita 29-10-26, 29/10/26 e 29/10/2026.
    data = data_padrao
    data_encontrada = False
    for m in RE_DATA_CUPOM.finditer(texto):
        dia, mes, ano = int(m.group(1)), int(m.group(2)), int(m.group(3))
        if ano < 100:
            ano += 2000
        try:
            data = datetime(ano, mes, dia)
            data_encontrada = True
            break
        except ValueError:
            continue  # OCR leu algo que não é data válida

    # Total pago: linha "Valor Total R$" (ou "Valor pago"), sem SUBTOTAL nem DESCONTO.
    total = None
    for l in linhas:
        if re.search(r"sub|desconto", l, re.IGNORECASE):
            continue
        if RE_TOTAL_LINHA.search(l):
            precos = RE_PRECO.findall(l)
            if precos:
                total = parse_brl(precos[-1])

    # Itens. Cada item pode vir em uma linha (nome + preço) ou em duas
    # (nome na primeira; "qtd unidade preço_unit total" na seguinte).
    itens = []
    descontos = {}           # numero do item -> valor do desconto (positivo)
    pendente = None          # (numero, codigo, descricao) aguardando a linha de preços
    for l in linhas:
        ln = normalizar(l)

        m_desc = RE_DESC_ITEM.search(ln)
        if m_desc:
            precos = RE_PRECO.findall(l)
            if precos:
                descontos[int(m_desc.group(1))] = abs(parse_brl(precos[-1]))
            continue

        if re.search(r"subtotal|desconto|total|dinheiro|troco|cartao|credito|debito|pix|"
                     r"valor pago|cnpj|cupom|fiscal|tributos|dados|pagamento|consumidor|"
                     r"chave|protocolo|nfc|serie|controle", ln):
            continue

        m_qp = RE_QTD_PRECOS.match(l)
        if m_qp:
            if pendente is not None:
                numero, codigo, desc = pendente
                itens.append({"numero": numero, "codigo": codigo, "produto": desc,
                              "valor": parse_brl(m_qp.group(2))})  # última coluna = total do item
                pendente = None
            continue

        precos = list(RE_PRECO.finditer(l))
        if precos:
            ultimo = precos[-1]
            numero, codigo, desc = separar_item(l[:ultimo.start()])
            if desc and re.search(r"[A-Za-z]{3,}", desc):
                itens.append({"numero": numero, "codigo": codigo, "produto": desc,
                              "valor": parse_brl(ultimo.group(1))})
                pendente = None
            continue

        numero, codigo, desc = separar_item(l)
        if desc and re.search(r"[A-Za-z]{3,}", desc):
            if pendente is not None:
                avisos_leitura.append(f"item sem preço: {pendente[2]}")
            pendente = (numero, codigo, desc)

    if pendente is not None:
        avisos_leitura.append(f"item sem preço: {pendente[2]}")

    # Aplica os descontos de item (valor líquido do item = total - desconto).
    for it in itens:
        if it["numero"] in descontos:
            it["valor"] = round(it["valor"] - descontos.pop(it["numero"]), 2)
    for numero in descontos:
        avisos_leitura.append(f"desconto do item {numero:03d} não associado a nenhum item")

    itens_limpos = []
    for i in itens:
        nome, cat, origem = resolver_produto(i["codigo"], i["produto"])
        if origem == "não reconhecido":
            avisos_leitura.append(f"produto não reconhecido (código {i['codigo'] or '-'}): {i['produto']}")
        itens_limpos.append({"produto": nome, "valor": i["valor"], "categoria": cat, "codigo": i["codigo"]})

    return {"loja": loja, "data": data, "data_encontrada": data_encontrada,
            "total": total, "itens": itens_limpos, "avisos_leitura": avisos_leitura}


def fechar_itens(itens: list, total):
    """
    Confere a soma dos itens com o total pago do cupom.
      - Se faltou ler algum item, cria a linha "Itens não identificados" com a diferença.
      - Se a soma passou do total, apenas avisa (não altera os valores lidos).
    Retorna (lista_de_itens, lista_de_avisos).
    """
    avisos = []
    pos = [dict(i) for i in itens]
    soma = round(sum(i["valor"] for i in pos), 2)

    if total is None:
        avisos.append(f"total não encontrado no cupom; itens somam {brl(soma)}")
    else:
        diff = round(total - soma, 2)
        if diff > 0.05:
            pos.append({"produto": "Itens não identificados (leitura)", "valor": diff})
            avisos.append(f"{brl(diff)} sem item identificado")
        elif diff < -0.05:
            avisos.append(f"itens somam {brl(soma)}, acima do total {brl(total)}; confira a foto")

    for i in pos:
        i["valor"] = round(i["valor"], 2)
        if i.get("categoria") is None:
            i["categoria"] = (CATEGORIA_PRODUTO_PADRAO
                              if i["produto"].startswith("Itens não identificados")
                              else categorizar_produto(i["produto"]))
    return pos, avisos


# ---------------------------------------------------------------
# Cupons conferidos manualmente (sem OCR)
# ---------------------------------------------------------------
NOME_CONFERIDOS = "cupons_conferidos.csv"
COLUNAS_CONFERIDOS = ["arquivo", "loja", "data", "total", "codigo", "produto", "qtd", "valor", "desconto", "emissao"]
EMISSAO_CONHECIDA = {  # cupons já conferidos (para preencher arquivos antigos)
    "andreazza_29-10-26.jpg": "29-10-26 17:14:37",
    "andreazza_28-10-26.jpg": "28-10-26 18:07:16",
}
CONFERIDOS_INICIAL = [  # itens exatamente como lidos nas fotos (valor = total da linha, sem desconto)
    # Andreazza 29-10-2026 17:14:37 - subtotal/total R$ 47,22
    ("andreazza_29-10-26.jpg", "Supermercados Andreazza", "29-10-2026", "47,22", "7898003541605", "Suvalan 200ml Uva", "1 Un", "2,49", "", "29-10-26 17:14:37"),
    ("andreazza_29-10-26.jpg", "Supermercados Andreazza", "29-10-2026", "47,22", "0000000000017", "Pao Frances", "0,340 Kg", "5,09", "", "29-10-26 17:14:37"),
    ("andreazza_29-10-26.jpg", "Supermercados Andreazza", "29-10-2026", "47,22", "0000000003049", "Banana Caturra", "1,175 Kg", "6,45", "", "29-10-26 17:14:37"),
    ("andreazza_29-10-26.jpg", "Supermercados Andreazza", "29-10-2026", "47,22", "0000000003483", "Doguinho", "0,160 Kg", "4,78", "", "29-10-26 17:14:37"),
    ("andreazza_29-10-26.jpg", "Supermercados Andreazza", "29-10-2026", "47,22", "7898003541605", "Suvalan 200ml Uva", "1 Un", "2,49", "", "29-10-26 17:14:37"),
    ("andreazza_29-10-26.jpg", "Supermercados Andreazza", "29-10-2026", "47,22", "0000000003483", "Doguinho", "0,144 Kg", "4,30", "", "29-10-26 17:14:37"),
    ("andreazza_29-10-26.jpg", "Supermercados Andreazza", "29-10-2026", "47,22", "0000000003483", "Doguinho", "0,172 Kg", "5,14", "", "29-10-26 17:14:37"),
    ("andreazza_29-10-26.jpg", "Supermercados Andreazza", "29-10-2026", "47,22", "0000000003483", "Doguinho", "0,144 Kg", "4,30", "", "29-10-26 17:14:37"),
    ("andreazza_29-10-26.jpg", "Supermercados Andreazza", "29-10-2026", "47,22", "7898003541612", "Suvalan 200ml Laranja", "1 Un", "2,49", "", "29-10-26 17:14:37"),
    ("andreazza_29-10-26.jpg", "Supermercados Andreazza", "29-10-2026", "47,22", "0000000020039", "Donuts Chocolate", "0,216 Kg", "9,69", "", "29-10-26 17:14:37"),
    # Andreazza 28-10-2026 18:07:16 - subtotal R$ 90,87, desconto R$ 6,62, total R$ 84,25
    ("andreazza_28-10-26.jpg", "Supermercados Andreazza", "28-10-2026", "84,25", "0000000003483", "Doguinho", "0,152 Kg", "4,54", "", "28-10-26 18:07:16"),
    ("andreazza_28-10-26.jpg", "Supermercados Andreazza", "28-10-2026", "84,25", "0000000003483", "Doguinho", "0,164 Kg", "4,90", "", "28-10-26 18:07:16"),
    ("andreazza_28-10-26.jpg", "Supermercados Andreazza", "28-10-2026", "84,25", "0000000003483", "Doguinho", "0,140 Kg", "4,18", "", "28-10-26 18:07:16"),
    ("andreazza_28-10-26.jpg", "Supermercados Andreazza", "28-10-2026", "84,25", "7896006779674", "Bisc Arroz Camil 150", "1 Un", "10,69", "", "28-10-26 18:07:16"),
    ("andreazza_28-10-26.jpg", "Supermercados Andreazza", "28-10-2026", "84,25", "7898681795826", "Cre Lte Dalja 300", "1 Un", "11,99", "", "28-10-26 18:07:16"),
    ("andreazza_28-10-26.jpg", "Supermercados Andreazza", "28-10-2026", "84,25", "0000000000017", "Pao Frances", "0,305 Kg", "4,56", "", "28-10-26 18:07:16"),
    ("andreazza_28-10-26.jpg", "Supermercados Andreazza", "28-10-2026", "84,25", "7891097000140", "Iog Batavo Grego 450", "1 Un", "14,99", "", "28-10-26 18:07:16"),
    ("andreazza_28-10-26.jpg", "Supermercados Andreazza", "28-10-2026", "84,25", "7891000248829", "Chocolate Kit Kat 41", "1 Un", "6,29", "3,31", "28-10-26 18:07:16"),
    ("andreazza_28-10-26.jpg", "Supermercados Andreazza", "28-10-2026", "84,25", "7891000248829", "Chocolate Kit Kat 41", "1 Un", "6,29", "3,31", "28-10-26 18:07:16"),
    ("andreazza_28-10-26.jpg", "Supermercados Andreazza", "28-10-2026", "84,25", "7896733400056", "Nata Holandes 220g", "1 Un", "9,99", "", "28-10-26 18:07:16"),
    ("andreazza_28-10-26.jpg", "Supermercados Andreazza", "28-10-2026", "84,25", "7898003541605", "Suvalan 200ml Uva", "2 Un", "4,98", "", "28-10-26 18:07:16"),
    ("andreazza_28-10-26.jpg", "Supermercados Andreazza", "28-10-2026", "84,25", "7898003541612", "Suvalan 200ml Laranja", "3 Un", "7,47", "", "28-10-26 18:07:16"),
]
CONFERIDOS = {}  # nome da foto -> lista de linhas do CSV


def carregar_conferidos(pasta: Path):
    """Lê cupons_conferidos.csv (cria com os cupons iniciais se não existir)."""
    arq = pasta / NOME_CONFERIDOS
    if not arq.exists():
        with arq.open("w", encoding="utf-8-sig", newline="") as f:
            w = csv.writer(f, delimiter=";")
            w.writerow(COLUNAS_CONFERIDOS)
            for linha in CONFERIDOS_INICIAL:
                w.writerow(linha)
        print(f"Arquivo criado: {arq.name} (itens conferidos; nomes das fotos: andreazza_29-10-26.jpg e andreazza_28-10-26.jpg)")
    CONFERIDOS.clear()
    linhas = ler_csv_dicts(arq)
    if linhas and "emissao" not in linhas[0]:
        for r in linhas:
            r["emissao"] = EMISSAO_CONHECIDA.get((r.get("arquivo") or "").strip(), "")
        gravar_csv(arq, COLUNAS_CONFERIDOS, linhas)
        print("Coluna 'emissao' adicionada ao cupons_conferidos.csv.")
    for row in linhas:
        nome = (row.get("arquivo") or "").strip()
        if nome and (row.get("produto") or "").strip():
            CONFERIDOS.setdefault(nome, []).append(row)
    if CONFERIDOS:
        print(f"Cupons conferidos (sem OCR): {', '.join(CONFERIDOS)}")


def montar_cupom_conferido(nome_arquivo: str) -> dict:
    """Monta o cupom a partir das linhas do CSV, no mesmo formato de parse_recibo()."""
    linhas = CONFERIDOS[nome_arquivo]
    cab = linhas[0]
    itens = []
    for r in linhas:
        bruto = parse_brl(r["valor"])
        desc = parse_brl(r["desconto"]) if (r.get("desconto") or "").strip() else 0.0
        cod = normalizar_codigo(r.get("codigo"))
        if cod in CATALOGO and CATALOGO[cod][1]:
            cat = CATALOGO[cod][1]
        else:
            cat = categorizar_produto(r["produto"])
        itens.append({"produto": r["produto"].strip(), "valor": round(bruto - desc, 2), "categoria": cat})
    return {"loja": cab["loja"].strip(),
            "data": datetime.strptime(cab["data"].strip(), "%d-%m-%Y"),
            "data_encontrada": True,
            "total": parse_brl(cab["total"]),
            "itens": itens,
            "avisos_leitura": []}


NOME_PENDENTES = "para_conferir.csv"
COLUNAS_PENDENTES = ["arquivo", "loja", "data", "total", "codigo", "produto", "valor", "status"]


def ler_csv_dicts(arq: Path) -> list:
    if not arq.exists():
        return []
    with arq.open(encoding="utf-8-sig", newline="") as f:
        return list(csv.DictReader(f, delimiter=";"))


RE_EMISSAO = re.compile(r"(?<!\d)(\d{2})[\s\-:./]{1,3}(\d{2})[\s\-:./]{1,3}(\d{2})\s+(\d{2})[\s\-:.]{1,2}(\d{2})[\s\-:.]{1,2}(\d{2})(?!\d)")


def emissao_do_texto(texto: str):
    """Data e hora de emissão lidas no texto (ex.: 29-10-26 17:14:37) como tupla de inteiros."""
    m = RE_EMISSAO.search(texto or "")
    return tuple(int(g) for g in m.groups()) if m else None


def emissao_texto(valor: str):
    """Converte '29-10-26 17:14:37' em tupla de inteiros para comparar."""
    m = RE_EMISSAO.search(valor or "")
    return tuple(int(g) for g in m.groups()) if m else None


def cupom_duplicado(emissao_ocr):
    """Nome do cupom conferido com a mesma emissão, ou None."""
    if emissao_ocr is None:
        return None
    for arq, linhas in CONFERIDOS.items():
        if emissao_texto(linhas[0].get("emissao", "")) == emissao_ocr:
            return arq
    return None


def gravar_csv(arq: Path, colunas: list, linhas: list):
    with arq.open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=colunas, delimiter=";", extrasaction="ignore")
        w.writeheader()
        for linha in linhas:
            w.writerow(linha)


NOME_COMPARACAO = "comparacao_ocr.csv"


def gerar_pendentes(pasta: Path, saida: Path):
    """
    Fotos sem conferência: lê o texto (OCR, ou o ocr_<foto>.txt já salvo) e grava as
    linhas em para_conferir.csv com status PENDENTE. Fotos já listadas não são relidas.
    """
    arq = pasta / NOME_PENDENTES
    existentes = ler_csv_dicts(arq)
    ja_listadas = {r["arquivo"] for r in existentes}
    pendentes_fotos = [i for i in listar_imagens(pasta)
                       if i.name not in CONFERIDOS and i.name not in ja_listadas]
    if not pendentes_fotos:
        return
    textos = extrair_textos_cupons(pasta, saida, usar_salvos=True)
    novas = []
    comparacoes = []
    for img in pendentes_fotos:
        texto = textos.get(img.name)
        if not texto or not texto.strip():
            print(f"AVISO: não consegui ler {img.name}. Veja o ocr_{img.stem}.txt, se existir.")
            continue
        rec = parse_recibo(texto, img.name, datetime.fromtimestamp(img.stat().st_mtime))
        emissao_ocr = emissao_do_texto(texto)
        dup = cupom_duplicado(emissao_ocr)
        if dup and emissao_ocr is not None:  # confirmação explícita; cupom_duplicado() só acha com emissão lida
            comparacoes.append({"foto": img.name, "emissao_ocr": f"{emissao_ocr[0]:02d}-{emissao_ocr[1]:02d}-{emissao_ocr[2]:02d} "
                f"{emissao_ocr[3]:02d}:{emissao_ocr[4]:02d}:{emissao_ocr[5]:02d}",
                "cupom_conferido": dup, "decisao": "DUPLICATA: já conferida, não entra de novo"})
            print(f"{img.name}: mesma emissão do cupom já conferido {dup}. Não será incluído.")
            continue
        comparacoes.append({"foto": img.name,
                            "emissao_ocr": "não lida" if emissao_ocr is None else
                            f"{emissao_ocr[0]:02d}-{emissao_ocr[1]:02d}-{emissao_ocr[2]:02d} "
                            f"{emissao_ocr[3]:02d}:{emissao_ocr[4]:02d}:{emissao_ocr[5]:02d}",
                            "cupom_conferido": "-",
                            "decisao": ("NOVA COMPRA: para conferir" if emissao_ocr else
                                        "NÃO VERIFICADO (sem data e hora no OCR): para conferir")})
        itens, avisos = fechar_itens(rec["itens"], rec["total"])
        total = rec["total"] if rec["total"] is not None else round(sum(i["valor"] for i in itens), 2)
        for i in itens:
            novas.append({
                "arquivo": img.name,
                "loja": rec["loja"],
                "data": rec["data"].strftime("%d-%m-%Y"),
                "total": f"{total:.2f}".replace(".", ","),
                "codigo": i.get("codigo", ""),
                "produto": i["produto"],
                "valor": f"{i['valor']:.2f}".replace(".", ","),
                "status": "PENDENTE",
            })
        print(f"Para conferir (não entra no dashboard ainda): {img.name} | {len(itens)} itens")
        for a in avisos:
            print(f"   AVISO ({img.name}): {a}")
    if comparacoes:
        gravar_csv(pasta / NOME_COMPARACAO, ["foto", "emissao_ocr", "cupom_conferido", "decisao"], comparacoes)
    if novas:
        gravar_csv(arq, COLUNAS_PENDENTES, existentes + novas)
        print(f"Abra {NOME_PENDENTES}, corrija o que estiver errado e troque o status para OK.")


def promover_aprovados(pasta: Path):
    """Fotos com todas as linhas em OK passam para cupons_conferidos.csv."""
    arq = pasta / NOME_PENDENTES
    pend = ler_csv_dicts(arq)
    if not pend:
        return
    por_foto = {}
    for r in pend:
        por_foto.setdefault(r["arquivo"], []).append(r)
    aprovadas = [f for f, rs in por_foto.items()
                 if all((r.get("status") or "").strip().upper() == "OK" for r in rs)]
    if not aprovadas:
        return
    conf = ler_csv_dicts(pasta / NOME_CONFERIDOS)
    for f in aprovadas:
        for r in por_foto[f]:
            conf.append({"arquivo": f, "loja": r["loja"], "data": r["data"], "total": r["total"],
                         "codigo": r.get("codigo", ""), "produto": r["produto"], "qtd": "",
                         "valor": r["valor"], "desconto": "", "emissao": EMISSAO_CONHECIDA.get(f, "")})
    gravar_csv(pasta / NOME_CONFERIDOS, COLUNAS_CONFERIDOS, conf)
    restantes = [r for r in pend if r["arquivo"] not in aprovadas]
    gravar_csv(arq, COLUNAS_PENDENTES, restantes)
    print(f"Aprovadas e incluídas no dashboard: {', '.join(aprovadas)}")


def ler_cupons(pasta: Path, saida: Path, usar_salvos: bool = False) -> list:
    """
    Só entram no dashboard os cupons conferidos (cupons_conferidos.csv).
    Não é preciso que a foto esteja na pasta: os dados conferidos já bastam.
    """
    cupons = []
    for nome, linhas in CONFERIDOS.items():
        rec = montar_cupom_conferido(nome)
        itens, avisos = fechar_itens(rec["itens"], rec["total"])
        rec["itens"], rec["avisos"], rec["arquivo"] = itens, avisos, nome
        cupons.append(rec)
        print(f"Cupom conferido: {nome} | {rec['loja']} | {rec['data']:%d/%m/%Y} | "
              f"total {brl(rec['total'])} | {len(itens)} itens")
        for a_ in avisos:
            print(f"   AVISO ({nome}): {a_}")
    return cupons


def chave_loja(loja: str) -> str:
    """Primeira palavra com 4+ letras do nome da loja, sem acentos (ex.: 'ANDREAZZA')."""
    palavras = [p for p in normalizar(loja).split() if len(p) >= 4 and p.isalpha()]
    return palavras[0] if palavras else ""


def integrar_cupons(df: pd.DataFrame, cupons: list):
    """
    Casa cada cupom com a linha da fatura (mesmo valor, data até 3 dias).
    Entre candidatos, prefere o que tem o nome da loja na descrição e, depois,
    a data mais próxima.
    Retorna (df_integrado, tabela_de_casamento).
    """
    df = garantir_colunas(df)
    casamentos = []
    if not cupons:
        return df, pd.DataFrame(casamentos)

    usados, remover, novos = set(), set(), []

    for c in cupons:
        status = "Sem fatura (outro pagamento)"
        cartao_item, data_item = "Cupom (sem fatura)", c["data"]

        if c["total"] is not None and not df.empty:
            cand = df[df["eh_gasto"].astype(bool)
                      & ~df.index.isin(usados)
                      & ((df["valor"].astype(float) - c["total"]).abs() <= 0.05)
                      & ((pd.to_datetime(df["data"]) - c["data"]).abs() <= pd.Timedelta(days=3))]
            if not cand.empty:
                chave = chave_loja(c["loja"])
                cand = cand.assign(
                    loja_bate=cand["descricao"].map(
                        lambda d: bool(chave) and chave in normalizar(d)),
                    distancia=(pd.to_datetime(cand["data"]) - c["data"]).abs(),
                )
                idx = cand.sort_values(by=["loja_bate", "distancia"],
                                       ascending=[False, True]).index[0]
                usados.add(idx)
                remover.add(idx)
                cartao_item = df.loc[idx, "cartao"]
                data_item = df.loc[idx, "data"]
                status = f"Casado com {cartao_item}"
                if not cand.loc[idx, "loja_bate"]:
                    status += " (loja não conferida)"

        for item in c["itens"]:
            novos.append({
                "data": data_item,
                "descricao": item["produto"],
                "parcelamento": "",
                "valor": item["valor"],
                "cartao": cartao_item,
                "tipo_cartao": "",
                "vencimento": None,
                "arquivo": c["arquivo"],
                "categoria": item["categoria"],
                "eh_gasto": True,
                "origem": f"Cupom: {c['loja']}",
                "loja": c["loja"],
            })

        casamentos.append({
            "Arquivo": c["arquivo"],
            "Loja": c["loja"],
            "Data": c["data"],
            "Total do cupom": c["total"] if c["total"] is not None else float("nan"),
            "Itens": len(c["itens"]),
            "Situação": status,
            "Avisos": "; ".join(c["avisos"]) if c["avisos"] else "-",
        })

    df = df.drop(index=list(remover))
    if novos:
        df = pd.concat([df, pd.DataFrame(novos)], ignore_index=True)
    df = garantir_colunas(df)
    df["valor"] = pd.to_numeric(df["valor"], errors="coerce")
    df["data"] = pd.to_datetime(df["data"])
    df["eh_gasto"] = df["eh_gasto"].astype(bool)
    df["parcelas_restantes"] = pd.to_numeric(df["parcelas_restantes"], errors="coerce").fillna(0)
    df["saldo_restante"] = pd.to_numeric(df["saldo_restante"], errors="coerce").fillna(0)
    df["loja"] = df["loja"].fillna("")
    return df, pd.DataFrame(casamentos)


# ---------------------------------------------------------------
# Gráficos
# ---------------------------------------------------------------
def fig_pizza(por_cat):
    f = px.pie(por_cat, names="categoria", values="total", hole=0.45,
               title="Distribuição dos gastos por categoria", template="plotly_white")
    # Percentual formatado por nós (vírgula decimal), igual às tabelas.
    f.update_traces(text=por_cat["percentual"].map(pct).tolist(), textinfo="label+text")
    return f


def fig_barras_categorias(por_cat):
    d = por_cat.sort_values(by="total")
    f = px.bar(d, x="total", y="categoria", orientation="h",
               title="Total gasto por categoria (R$)", template="plotly_white",
               labels={"total": "R$", "categoria": ""})
    f.update_traces(texttemplate="R$ %{x:,.2f}", textposition="outside")
    f.update_layout(showlegend=False, height=max(350, 40 * len(d) + 120))
    return f


def fig_cartoes(df_gastos):
    d = df_gastos.groupby(["cartao", "categoria"], as_index=False)["valor"].sum()
    f = px.bar(d, x="cartao", y="valor", color="categoria",
               title="Gastos por cartão/fonte, separados por categoria",
               template="plotly_white", labels={"valor": "R$", "cartao": ""})
    f.update_layout(barmode="stack")
    return f


def fig_sunburst(df_gastos):
    return px.sunburst(df_gastos, path=["cartao", "categoria"], values="valor",
                       title="Cartão → categoria", template="plotly_white")


def fig_diario(df_gastos):
    d = (df_gastos.groupby(df_gastos["data"].dt.normalize().rename("dia"))["valor"]
         .sum().reset_index())
    d["acumulado"] = d["valor"].cumsum()
    f = go.Figure()
    f.add_bar(x=d["dia"], y=d["valor"], name="Gasto no dia")
    f.add_scatter(x=d["dia"], y=d["acumulado"], name="Acumulado",
                  yaxis="y2", mode="lines+markers")
    f.update_layout(
        title="Gastos por dia e acumulado",
        template="plotly_white",
        yaxis=dict(title="R$"),
        yaxis2=dict(title="Acumulado (R$)", overlaying="y", side="right"),
        legend=dict(orientation="h"),
    )
    return f


def fig_potencial(por_cat):
    d = por_cat[por_cat["economizavel"]].sort_values(by="total")
    if d.empty:
        return None
    f = px.bar(d, x="total", y="categoria", orientation="h",
               title="Onde dá para cortar (categorias sugeridas para economia)",
               template="plotly_white", labels={"total": "R$", "categoria": ""})
    f.update_traces(texttemplate="R$ %{x:,.2f}", textposition="outside")
    f.update_layout(showlegend=False, height=max(300, 40 * len(d) + 120))
    return f


def fig_mercado_produtos(df_cupons):
    if df_cupons.empty:
        return None
    d = (df_cupons.groupby("categoria", as_index=False)["valor"].sum()
         .sort_values(by="valor"))
    f = px.bar(d, x="valor", y="categoria", orientation="h",
               title="Supermercado: gasto por tipo de produto (a partir dos cupons)",
               template="plotly_white", labels={"valor": "R$", "categoria": ""})
    f.update_traces(texttemplate="R$ %{x:,.2f}", textposition="outside")
    f.update_layout(showlegend=False, height=max(300, 40 * len(d) + 120))
    return f


# ---------------------------------------------------------------
# Montagem do HTML
# ---------------------------------------------------------------
CSS = """
body{font-family:Segoe UI,Arial,sans-serif;background:#f4f6fa;margin:0;color:#222}
header{background:#1f3a5f;color:#fff;padding:18px 28px}
header h1{margin:0;font-size:22px}
header p{margin:4px 0 0;opacity:.85;font-size:13px}
main{padding:20px 28px}
.kpis{display:flex;flex-wrap:wrap;gap:14px;margin-bottom:20px}
.kpi{background:#fff;border-radius:10px;padding:14px 18px;min-width:180px;box-shadow:0 1px 3px #0002}
.kpi span{font-size:12px;color:#666;text-transform:uppercase;letter-spacing:.4px}
.kpi b{display:block;font-size:22px;margin-top:4px}
.card{background:#fff;border-radius:10px;padding:16px;margin-bottom:20px;box-shadow:0 1px 3px #0002}
.card h2{font-size:17px;margin:0 0 10px;color:#1f3a5f}
.grid{display:grid;grid-template-columns:1fr 1fr;gap:20px}
@media(max-width:900px){.grid{grid-template-columns:1fr}}
table.tbl{border-collapse:collapse;width:100%;font-size:13px}
table.tbl th,table.tbl td{padding:6px 8px;border-bottom:1px solid #e5e8ef;text-align:left}
table.tbl th{background:#eef2f8}
table.tbl td{white-space:nowrap}
.resumo p{font-size:15px;line-height:1.5;margin:6px 0}
.alerta{color:#b03a2e;font-weight:600}
"""


def tabela(df_, formatadores=None):
    return df_.to_html(index=False, classes="tbl", border=0, na_rep="-",
                       formatters=formatadores or {}, escape=True)


def html_figura(fig, incluir_js):
    return fig.to_html(full_html=False, include_plotlyjs=incluir_js)


def montar_html(df, por_cat, por_fonte_total, parcelas, top, total, n, ticket,
                potencial, pix_total, assinaturas_total, saldo_restante_total,
                figs, pasta_origem, periodo, linha_mercado, tab_cupons,
                tab_produtos, tab_casamento_info):
    top3 = por_cat.head(3)
    frases_top = ", ".join(f"{r.categoria} ({brl(r.total)}, {pct(r.percentual)})"
                           for r in top3.itertuples())
    pot_pct = (potencial / total * 100) if total else 0

    resumo = f"""
    <p>Você gastou <b>{brl(total)}</b> em <b>{n}</b> lançamentos no período {periodo}
       (ticket médio de {brl(ticket)}).</p>
    <p>As categorias que mais pesam são: {frases_top}.</p>
    <p>Categorias sugeridas para corte somam <b>{brl(potencial)}</b>, ou {pct(pot_pct)} do total.</p>
    <p class="alerta">Pix no crédito: {brl(pix_total)}. Essa modalidade costuma ter custo
       financeiro alto; quitar à vista ou no débito costuma ser mais barato.</p>
    <p>Assinaturas e streaming: {brl(assinaturas_total)} (verifique se usa tudo).</p>
    <p>Parcelas ainda por pagar: <b>{brl(saldo_restante_total)}</b>.</p>
    <p>{linha_mercado}</p>
    """

    kpis = f"""
    <div class="kpis">
      <div class="kpi"><span>Total gasto</span><b>{brl(total)}</b></div>
      <div class="kpi"><span>Lançamentos</span><b>{n}</b></div>
      <div class="kpi"><span>Ticket médio</span><b>{brl(ticket)}</b></div>
      <div class="kpi"><span>Maior categoria</span><b>{por_cat.iloc[0]['categoria']}</b></div>
      <div class="kpi"><span>Potencial de corte</span><b>{brl(potencial)}</b></div>
      <div class="kpi"><span>Pix no crédito</span><b>{brl(pix_total)}</b></div>
      <div class="kpi"><span>Parcelas a pagar</span><b>{brl(saldo_restante_total)}</b></div>
    </div>
    """

    cat_show = por_cat.copy()
    cat_show["economizavel"] = cat_show["economizavel"].map({True: "Sim", False: "Não"})
    tab_cat = tabela(
        cat_show[["categoria", "lancamentos", "total", "percentual", "economizavel"]]
        .rename(columns={"categoria": "Categoria", "lancamentos": "Lançamentos",
                         "total": "Total", "percentual": "% do total",
                         "economizavel": "Potencial de corte"}),
        {"Total": brl, "% do total": pct},
    )

    tab_fonte = tabela(
        por_fonte_total.rename(columns={"cartao": "Cartão/fonte", "valor": "Total"}),
        {"Total": brl},
    )

    tab_parc = tabela(
        parcelas[["descricao", "cartao", "parcela_atual", "parcela_total",
                  "valor", "parcelas_restantes", "saldo_restante"]]
        .rename(columns={"descricao": "Estabelecimento", "cartao": "Cartão",
                         "parcela_atual": "Parcela atual", "parcela_total": "Total de parcelas",
                         "valor": "Valor da parcela", "parcelas_restantes": "Parcelas restantes",
                         "saldo_restante": "Saldo a pagar"}),
        {"Valor da parcela": brl, "Saldo a pagar": brl},
    ) if not parcelas.empty else "<p>Nenhuma parcela em andamento.</p>"

    tab_top = tabela(
        top[["data", "descricao", "cartao", "categoria", "valor"]]
        .assign(data=lambda d: d["data"].dt.strftime("%d/%m/%Y"))
        .rename(columns={"data": "Data", "descricao": "Descrição", "cartao": "Cartão",
                         "categoria": "Categoria", "valor": "Valor"}),
        {"Valor": brl},
    )

    tab_todos = tabela(
        df[["data", "descricao", "cartao", "categoria", "parcelamento", "origem", "valor"]]
        .assign(data=lambda d: d["data"].dt.strftime("%d/%m/%Y"))
        .rename(columns={"data": "Data", "descricao": "Descrição", "cartao": "Cartão",
                         "categoria": "Categoria", "parcelamento": "Parcelamento",
                         "origem": "Origem", "valor": "Valor"}),
        {"Valor": brl},
    )

    graficos = []
    for i, fig in enumerate(figs):
        if fig is None:
            graficos.append("")
            continue
        graficos.append(html_figura(fig, "cdn" if i == 0 else False))
    while len(graficos) < 7:
        graficos.append("")

    bloco_cupons = ""
    if tab_cupons is not None:
        bloco_cupons = f"""
  <div class="card"><h2>Cupons lidos e conferência com a fatura</h2>
    <p>{tab_casamento_info}</p>{tab_cupons}</div>
  <div class="grid">
    <div class="card">{graficos[5]}</div>
    <div class="card"><h2>Produtos mais comprados (cupons)</h2>{tab_produtos}</div>
  </div>
"""

    html = f"""<!doctype html>
<html lang="pt-BR">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Dashboard de finanças</title>
<style>{CSS}</style>
</head>
<body>
<header>
  <h1>Dashboard de finanças pessoais</h1>
  <p>Fonte: {pasta_origem} · Período: {periodo} · Gerado em {datetime.now():%d/%m/%Y %H:%M}</p>
</header>
<main>
  {kpis}

  <div class="card resumo">
    <h2>Para onde está indo o seu dinheiro</h2>
    {resumo}
  </div>

  <div class="grid">
    <div class="card">{graficos[0]}</div>
    <div class="card">{graficos[1]}</div>
  </div>

  <div class="card">{graficos[2]}</div>

  <div class="grid">
    <div class="card">{graficos[3]}</div>
    <div class="card">{graficos[4]}</div>
  </div>

  <div class="card"><h2>Resumo por categoria</h2>{tab_cat}</div>

  <div class="grid">
    <div class="card"><h2>Gastos por cartão</h2>{tab_fonte}</div>
    <div class="card"><h2>Parcelas em andamento</h2>{tab_parc}</div>
  </div>

  {bloco_cupons}

  <div class="card"><h2>15 maiores gastos</h2>{tab_top}</div>

  <div class="card"><h2>Todos os lançamentos</h2>{tab_todos}</div>
</main>
</body>
</html>
"""
    return html


# ---------------------------------------------------------------
# Dashboard em Excel (abas com tabelas e gráficos nativos)
# ---------------------------------------------------------------
COR_TITULO = "1F3A5F"
FMT_BRL = '"R$" #,##0.00'
FMT_PCT = '0.0"%"'
FMT_DATA = "dd/mm/yyyy"


def ordenar_desc(dfx, coluna):
    """Ordena do maior para o menor pela coluna indicada."""
    return dfx.iloc[np.argsort(-dfx[coluna].to_numpy())]


def escrever_df(ws, dfx, linha, formatos=None):
    """Escreve um DataFrame numa aba com cabeçalho estilizado. Devolve a última linha usada."""
    formatos = formatos or {}
    for j, col in enumerate(dfx.columns, start=1):
        c = ws.cell(linha, j, str(col))
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor=COR_TITULO)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for i, registro in enumerate(dfx.itertuples(index=False), start=linha + 1):
        for j, val in enumerate(registro, start=1):
            if val is None or (isinstance(val, float) and pd.isna(val)):
                val = None
            elif isinstance(val, pd.Timestamp):
                val = val.to_pydatetime()
            elif isinstance(val, (np.integer, np.floating)):  # tipos numpy -> Python
                val = val.item()
            c = ws.cell(i, j, val)
            fmt = formatos.get(dfx.columns[j - 1])
            if fmt:
                c.number_format = fmt
    return linha + len(dfx)


def ajustar_larguras(ws, maximo=45):
    for col_cells in ws.columns:
        tamanho = max((len(str(c.value)) for c in col_cells if c.value is not None), default=8)
        ws.column_dimensions[get_column_letter(col_cells[0].column)].width = min(tamanho + 2, maximo)


def salvar_excel(saida, df, por_cat, por_fonte_total, parcelas, top, tab_casamento,
                 cupons_rows, total, n, ticket, potencial, pix_total,
                 saldo_restante_total, periodo):
    wb = Workbook()

    # Aba Resumo: indicadores, tabela de categorias e gráficos
    ws = wb.worksheets[0]
    ws.title = "Resumo"
    ws["A1"] = "Dashboard de finanças pessoais"
    ws["A1"].font = Font(bold=True, size=16, color=COR_TITULO)
    ws["A2"] = f"Período: {periodo}"

    indicadores = [("Total gasto", total, FMT_BRL), ("Lançamentos", n, "0"),
                   ("Ticket médio", ticket, FMT_BRL), ("Potencial de corte", potencial, FMT_BRL),
                   ("Pix no crédito", pix_total, FMT_BRL),
                   ("Parcelas a pagar", saldo_restante_total, FMT_BRL)]
    for i, (rotulo, valor, fmt) in enumerate(indicadores):
        ws.cell(4 + i, 1, rotulo).font = Font(bold=True)
        c = ws.cell(4 + i, 2, round(float(valor), 2))
        c.number_format = fmt

    ws.cell(11, 1, "Gasto por categoria").font = Font(bold=True, size=12, color=COR_TITULO)
    cat = por_cat[["categoria", "lancamentos", "total", "percentual", "economizavel"]].copy()
    cat["economizavel"] = cat["economizavel"].map({True: "Sim", False: "Não"})
    cat.columns = ["Categoria", "Lançamentos", "Total", "% do total", "Potencial de corte"]
    ini = 12
    fim = escrever_df(ws, cat, ini, formatos={"Total": FMT_BRL, "% do total": FMT_PCT})

    rotulos = Reference(ws, min_col=1, min_row=ini + 1, max_row=fim)
    valores = Reference(ws, min_col=3, min_row=ini, max_row=fim)

    pizza = PieChart()
    pizza.title = "Distribuição por categoria"
    pizza.add_data(valores, titles_from_data=True)
    pizza.set_categories(rotulos)
    pizza.height, pizza.width = 9, 15
    ws.add_chart(pizza, "G3")

    barras = BarChart()
    barras.type = "bar"
    barras.title = "Total por categoria (R$)"
    barras.add_data(valores, titles_from_data=True)
    barras.set_categories(rotulos)
    barras.legend = None
    barras.height, barras.width = 9, 15
    ws.add_chart(barras, "G22")
    ajustar_larguras(ws)

    # Aba Cartões
    ws_c = wb.create_sheet("Cartões")
    fonte = por_fonte_total.rename(columns={"cartao": "Cartão/fonte", "valor": "Total"})
    fim_c = escrever_df(ws_c, fonte, 1, formatos={"Total": FMT_BRL})
    if fim_c > 1:
        grafico = BarChart()
        grafico.title = "Gastos por cartão (R$)"
        grafico.add_data(Reference(ws_c, min_col=2, min_row=1, max_row=fim_c), titles_from_data=True)
        grafico.set_categories(Reference(ws_c, min_col=1, min_row=2, max_row=fim_c))
        grafico.legend = None
        grafico.height, grafico.width = 9, 15
        ws_c.add_chart(grafico, "D2")
    ajustar_larguras(ws_c)

    # Aba Lançamentos (com filtro)
    ws_l = wb.create_sheet("Lançamentos")
    base = df[["data", "descricao", "cartao", "categoria", "parcelamento", "origem", "valor"]].copy()
    base["data"] = base["data"].dt.date
    base.columns = ["Data", "Descrição", "Cartão", "Categoria", "Parcelamento", "Origem", "Valor"]
    fim_l = escrever_df(ws_l, base, 1, formatos={"Data": FMT_DATA, "Valor": FMT_BRL})
    ws_l.auto_filter.ref = f"A1:G{fim_l}"
    ws_l.freeze_panes = "A2"
    ajustar_larguras(ws_l)

    # Aba Parcelas
    ws_p = wb.create_sheet("Parcelas")
    if parcelas.empty:
        ws_p["A1"] = "Nenhuma parcela em andamento."
    else:
        par = parcelas[["descricao", "cartao", "parcela_atual", "parcela_total",
                        "valor", "parcelas_restantes", "saldo_restante"]].copy()
        par.columns = ["Estabelecimento", "Cartão", "Parcela atual", "Total de parcelas",
                       "Valor da parcela", "Parcelas restantes", "Saldo a pagar"]
        escrever_df(ws_p, par, 1, formatos={"Valor da parcela": FMT_BRL, "Saldo a pagar": FMT_BRL})
    ajustar_larguras(ws_p)

    # Aba Cupons (conferência e itens)
    ws_k = wb.create_sheet("Cupons")
    if tab_casamento is not None and not tab_casamento.empty:
        fim_k = escrever_df(ws_k, tab_casamento, 1, formatos={"Total do cupom": FMT_BRL})
        if not cupons_rows.empty:
            itens = cupons_rows[["data", "loja", "descricao", "categoria", "valor"]].copy()
            itens["data"] = itens["data"].dt.date
            itens.columns = ["Data", "Loja", "Produto", "Categoria", "Valor"]
            escrever_df(ws_k, itens, fim_k + 3, formatos={"Data": FMT_DATA, "Valor": FMT_BRL})
    else:
        ws_k["A1"] = "Nenhum cupom de supermercado encontrado."
    ajustar_larguras(ws_k)

    # Aba Top 15
    ws_t = wb.create_sheet("Top 15 gastos")
    tp = top[["data", "descricao", "cartao", "categoria", "valor"]].copy()
    tp["data"] = tp["data"].dt.date
    tp.columns = ["Data", "Descrição", "Cartão", "Categoria", "Valor"]
    escrever_df(ws_t, tp, 1, formatos={"Data": FMT_DATA, "Valor": FMT_BRL})
    ajustar_larguras(ws_t)

    caminho = saida / "dashboard_financas.xlsx"
    try:
        wb.save(caminho)
    except PermissionError:
        print(f"AVISO: não consegui salvar {caminho.name}. Feche o arquivo no Excel e rode de novo.")
        return None
    return caminho


# ---------------------------------------------------------------
# Programa principal
# ---------------------------------------------------------------
def main():
    opcoes = {a for a in sys.argv[1:] if a.startswith("--")}
    argumentos = [a for a in sys.argv[1:] if not a.startswith("--")]
    so_texto = "--so-texto" in opcoes
    usar_textos = "--usar-textos" in opcoes

    entrada = Path(argumentos[0]) if len(argumentos) > 0 else PASTA_PADRAO
    saida = Path(argumentos[1]) if len(argumentos) > 1 else entrada

    if not entrada.exists():
        sys.exit(f"Pasta não encontrada: {entrada}")
    saida.mkdir(parents=True, exist_ok=True)
    carregar_conferidos(entrada)

    # Modo só texto: extrai o texto das fotos e para (para você revisar os .txt).
    if so_texto:
        extrair_textos_cupons(entrada, saida, usar_salvos=usar_textos)
        print("\nTextos extraídos. Abra os ocr_*.txt na pasta de saída para revisar.")
        print("Depois rode:  python dashboard_financas_v5_16.py --usar-textos")
        return

    # 1) Faturas e extratos
    df = classificar_lancamentos(carregar_lancamentos(entrada))

    # 2) Cupons: catálogo de produtos, texto das fotos (OCR ou textos salvos) e interpretação
    carregar_catalogo(entrada)
    promover_aprovados(entrada)
    carregar_conferidos(entrada)
    gerar_pendentes(entrada, saida)
    cupons = ler_cupons(entrada, saida)
    df, tab_casamento = integrar_cupons(df, cupons)

    if df.empty:
        print("\nNenhum lançamento confirmado ainda, então o dashboard não foi gerado.")
        if (entrada / NOME_PENDENTES).exists():
            print(f"Fotos aguardando conferência: abra {NOME_PENDENTES}, corrija os itens e troque o status para OK.")
        else:
            print("Coloque fotos de cupom na pasta ou adicione faturas/extratos.")
        return

    df = df.sort_values(by="data", ascending=False).reset_index(drop=True)
    gastos = df[df["eh_gasto"]].copy()
    if gastos.empty:
        print("\nNenhum gasto (valor positivo) encontrado; o dashboard não foi gerado.")
        return

    total = gastos["valor"].sum()
    n = len(gastos)
    ticket = total / n

    por_cat = (gastos.groupby("categoria")
               .agg(total=("valor", "sum"), lancamentos=("valor", "size"))
               .sort_values(by="total", ascending=False)
               .reset_index())
    por_cat["percentual"] = por_cat["total"] / total * 100
    por_cat["economizavel"] = por_cat["categoria"].isin(CATEGORIAS_ECONOMIZAVEIS)

    potencial = por_cat.loc[por_cat["economizavel"], "total"].sum()
    pix_total = por_cat.loc[por_cat["categoria"] == "Pix no crédito", "total"].sum()
    assinaturas_total = por_cat.loc[por_cat["categoria"] == "Assinaturas/streaming", "total"].sum()

    por_fonte_total = ordenar_desc(
        gastos.groupby("cartao", as_index=False)["valor"].sum(), "valor")

    parcelas = gastos[gastos["parcelas_restantes"] > 0].copy()
    saldo_restante_total = parcelas["saldo_restante"].sum()

    top = gastos.nlargest(15, "valor")
    periodo = f"{gastos['data'].min():%d/%m/%Y} a {gastos['data'].max():%d/%m/%Y}"

    # Blocos de supermercado (só se houver cupons)
    cupons_rows = gastos[gastos["origem"].astype(str).str.startswith("Cupom")]
    if cupons_rows.empty:
        linha_mercado = "Nenhum cupom de supermercado encontrado na pasta."
        tab_cupons = None
        tab_produtos = None
        tab_casamento_info = ""
        fig_prod = None
    else:
        val_cupons = cupons_rows["valor"].sum()
        linha_mercado = (f"Supermercado: {brl(val_cupons)} detalhados por produto a partir de "
                         f"{len(tab_casamento)} cupom(ns) lido(s), "
                         f"{len(cupons_rows)} itens no total.")
        tab_cupons = tabela(
            tab_casamento.assign(Data=lambda d: pd.to_datetime(d["Data"]).dt.strftime("%d/%m/%Y"))
            [["Arquivo", "Loja", "Data", "Total do cupom", "Itens", "Situação", "Avisos"]],
            {"Total do cupom": brl},
        )
        produtos = ordenar_desc(
            cupons_rows.groupby(["descricao", "categoria"], as_index=False)["valor"]
            .agg(total="sum", vezes="size"), "total").head(15)
        produtos = produtos.set_axis(["Produto", "Categoria", "Total", "Ocorrências"], axis=1)
        tab_produtos = tabela(produtos, {"Total": brl})
        tab_casamento_info = ("Cada cupom foi conferido com a fatura pelo valor e pela data, "
                              "com preferência para o nome da loja. Quando casa, a linha "
                              "genérica da fatura é substituída pelos itens do cupom, sem "
                              "duplicar o gasto.")
        fig_prod = fig_mercado_produtos(cupons_rows)

    figs = [
        fig_pizza(por_cat),
        fig_barras_categorias(por_cat),
        fig_cartoes(gastos),
        fig_diario(gastos),
        fig_potencial(por_cat),
        fig_prod,
        fig_sunburst(gastos),
    ]

    html = montar_html(
        df=df,
        por_cat=por_cat,
        por_fonte_total=por_fonte_total,
        parcelas=parcelas,
        top=top,
        total=total,
        n=n,
        ticket=ticket,
        potencial=potencial,
        pix_total=pix_total,
        assinaturas_total=assinaturas_total,
        saldo_restante_total=saldo_restante_total,
        figs=figs,
        pasta_origem=str(entrada),
        periodo=periodo,
        linha_mercado=linha_mercado,
        tab_cupons=tab_cupons,
        tab_produtos=tab_produtos,
        tab_casamento_info=tab_casamento_info,
    )

    caminho_html = saida / "dashboard_financas.html"
    caminho_html.write_text(html, encoding="utf-8")

    caminho_csv = saida / "transacoes_consolidadas.csv"
    export = df.copy()
    export["data"] = export["data"].dt.strftime("%d/%m/%Y")
    export.to_csv(caminho_csv, index=False, sep=";", decimal=",", encoding="utf-8-sig")

    caminho_xlsx = salvar_excel(
        saida, df, por_cat, por_fonte_total, parcelas, top, tab_casamento,
        cupons_rows, total, n, ticket, potencial, pix_total,
        saldo_restante_total, periodo)

    print("\nPronto!")
    print(f"  Total gasto       : {brl(total)} em {n} lançamentos")
    print(f"  Potencial de corte: {brl(potencial)}")
    print(f"  Cupons lidos      : {len(tab_casamento)}")
    print(f"  Dashboard         : {caminho_html}")
    print(f"  Planilha          : {caminho_csv}")
    if caminho_xlsx:
        print(f"  Dashboard Excel   : {caminho_xlsx}")


if __name__ == "__main__":
    main()
