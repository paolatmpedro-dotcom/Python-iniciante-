import os
import tkinter as tk
import xml.etree.ElementTree as ET
import zipfile
import tempfile
import shutil
import hashlib
import hmac
import secrets

from tkinter import messagebox, ttk


# ============================================================
# CONFIGURAÇÃO
# ============================================================

PASTA_DO_PROGRAMA = os.path.dirname(os.path.abspath(__file__))

ARQUIVO_EXCEL = os.path.join(
    PASTA_DO_PROGRAMA,
    "usuarios.xlsx"
)

CAMINHO_IMAGEM = os.path.join(
    PASTA_DO_PROGRAMA,
    "photo.jpg"
)


# ============================================================
# PILLOW
# ============================================================

try:
    from PIL import Image, ImageTk, ImageEnhance
    TEM_PILLOW = True
except ImportError:
    TEM_PILLOW = False


# ============================================================
# CONFIGURAÇÃO DO SISTEMA
# ============================================================

MAX_TENTATIVAS = 3

STATUS_ATIVO = "Ativo"
STATUS_INATIVO = "Inativo"
STATUS_BLOQUEADO = "Bloqueado"

NIVEL_USUARIO = "Usuário"
NIVEL_ADMINISTRADOR = "Administrador"

ITERACOES_HASH = 600_000


# ============================================================
# WANTED DESIGN SYSTEM
# ============================================================

COR_PRIMARY = "#E83E8C"
COR_PRIMARY_DARK = "#C92F72"
COR_PRIMARY_LIGHT = "#FFF0F7"

COR_SECONDARY = "#7656A6"
COR_SECONDARY_DARK = "#60448D"

COR_BACKGROUND = "#F8F5F7"
COR_SURFACE = "#FFFFFF"

COR_TEXT = "#2D2530"
COR_TEXT_SECONDARY = "#6F6470"
COR_TEXT_MUTED = "#958A95"

COR_BORDER = "#E7DDE4"

COR_SUCCESS = "#2E9B70"
COR_WARNING = "#D99128"
COR_ERROR = "#D9536F"

COR_DISABLED = "#C9C1C8"

COR_SHADOW = "#DCCED6"

COR_ROSA_1 = "#FBC2EB"
COR_ROSA_2 = "#EC6FAE"

# Compatibilidade
COR_ROSA = COR_PRIMARY
COR_ROSA_ESCURO = COR_PRIMARY_DARK
COR_ROSA_CLARO = COR_PRIMARY_LIGHT

COR_BRANCO = COR_SURFACE
COR_FUNDO = COR_BACKGROUND

COR_TEXTO = COR_TEXT
COR_MUTED = COR_TEXT_SECONDARY

COR_BORDA = COR_BORDER

COR_SUCESSO = COR_SUCCESS
COR_ERRO = COR_ERROR
COR_ALERTA = COR_WARNING
COR_EXCLUIR = "#C94F5D"

COR_SOMBRA = COR_SHADOW


# ============================================================
# HASH DE SENHA
# ============================================================

def gerar_hash_senha(senha):

    salt = secrets.token_bytes(16)

    hash_senha = hashlib.pbkdf2_hmac(
        "sha256",
        senha.encode("utf-8"),
        salt,
        ITERACOES_HASH
    )

    return (
        f"PBKDF2-SHA256$"
        f"{ITERACOES_HASH}$"
        f"{salt.hex()}$"
        f"{hash_senha.hex()}"
    )


def verificar_hash_senha(senha, senha_salva):

    try:

        partes = senha_salva.split("$")

        if len(partes) != 4:
            return False

        algoritmo = partes[0]
        iteracoes = int(partes[1])
        salt = bytes.fromhex(partes[2])
        hash_salvo = bytes.fromhex(partes[3])

        if algoritmo != "PBKDF2-SHA256":
            return False

        hash_calculado = hashlib.pbkdf2_hmac(
            "sha256",
            senha.encode("utf-8"),
            salt,
            iteracoes
        )

        return hmac.compare_digest(
            hash_calculado,
            hash_salvo
        )

    except Exception:
        return False


def senha_e_hash(senha):

    return senha.startswith("PBKDF2-SHA256$")


# ============================================================
# DEGRADÊ DINÂMICO
# ============================================================

def desenhar_degrade(
    canvas,
    cor1,
    cor2,
    largura,
    altura
):

    if largura <= 1 or altura <= 1:
        return

    try:

        r1, g1, b1 = canvas.winfo_rgb(cor1)
        r2, g2, b2 = canvas.winfo_rgb(cor2)

        canvas.delete("degrade")

        altura_real = max(altura - 1, 1)

        for i in range(altura):

            proporcao = i / altura_real

            r = int(
                r1 + (r2 - r1) * proporcao
            ) >> 8

            g = int(
                g1 + (g2 - g1) * proporcao
            ) >> 8

            b = int(
                b1 + (b2 - b1) * proporcao
            ) >> 8

            cor = f"#{r:02x}{g:02x}{b:02x}"

            canvas.create_line(
                0,
                i,
                largura,
                i,
                fill=cor,
                tags="degrade"
            )

        canvas.tag_lower("degrade")

    except tk.TclError:
        pass


# ============================================================
# RETÂNGULO ARREDONDADO
# ============================================================

def criar_retangulo_arredondado(
    canvas,
    x1,
    y1,
    x2,
    y2,
    raio=25,
    fill="white",
    outline=""
):

    raio = min(
        raio,
        abs(x2 - x1) / 2,
        abs(y2 - y1) / 2
    )

    pontos = [
        x1 + raio, y1,
        x2 - raio, y1,

        x2, y1,
        x2, y1 + raio,

        x2, y2 - raio,
        x2, y2,

        x2 - raio, y2,
        x1 + raio, y2,

        x1, y2,
        x1, y2 - raio,

        x1, y1 + raio,
        x1, y1
    ]

    return canvas.create_polygon(
        pontos,
        smooth=True,
        fill=fill,
        outline=outline
    )


# ============================================================
# LEITOR EXCEL
# ============================================================

class LeitorExcel:

    def __init__(self, arquivo):

        self.arquivo = arquivo

    def ler_planilha(self):

        if not os.path.exists(self.arquivo):

            raise FileNotFoundError(
                f"O arquivo:\n{self.arquivo}\n"
                "não foi encontrado."
            )

        try:

            with zipfile.ZipFile(
                self.arquivo,
                "r"
            ) as arquivo_zip:

                arquivos = arquivo_zip.namelist()

                caminho_planilha = None

                for nome in arquivos:

                    if (
                        nome.startswith(
                            "xl/worksheets/sheet"
                        )
                        and nome.endswith(".xml")
                    ):

                        caminho_planilha = nome
                        break

                if not caminho_planilha:

                    raise ValueError(
                        "O arquivo não possui "
                        "uma planilha válida."
                    )

                strings = []

                if "xl/sharedStrings.xml" in arquivos:

                    xml_strings = arquivo_zip.read(
                        "xl/sharedStrings.xml"
                    )

                    raiz_strings = ET.fromstring(
                        xml_strings
                    )

                    namespace = {
                        "main":
                        "http://schemas.openxmlformats.org/"
                        "spreadsheetml/2006/main"
                    }

                    for elemento in raiz_strings.findall(
                        ".//main:si",
                        namespace
                    ):

                        textos = [
                            texto.text or ""
                            for texto in elemento.findall(
                                ".//main:t",
                                namespace
                            )
                        ]

                        strings.append(
                            "".join(textos)
                        )

                xml_planilha = arquivo_zip.read(
                    caminho_planilha
                )

                raiz_planilha = ET.fromstring(
                    xml_planilha
                )

                namespace = {
                    "main":
                    "http://schemas.openxmlformats.org/"
                    "spreadsheetml/2006/main"
                }

                linhas = []

                for linha in raiz_planilha.findall(
                    ".//main:sheetData/main:row",
                    namespace
                ):

                    valores_linha = {}

                    for celula in linha.findall(
                        "main:c",
                        namespace
                    ):

                        referencia = celula.attrib.get(
                            "r",
                            ""
                        )

                        coluna = "".join(
                            c for c in referencia
                            if c.isalpha()
                        )

                        tipo = celula.attrib.get("t")

                        valor = celula.find(
                            "main:v",
                            namespace
                        )

                        valor_texto = ""

                        if tipo == "inlineStr":

                            textos = celula.findall(
                                ".//main:t",
                                namespace
                            )

                            valor_texto = "".join(
                                t.text or ""
                                for t in textos
                            )

                        elif (
                            tipo == "s"
                            and valor is not None
                        ):

                            indice = int(
                                valor.text
                            )

                            if indice < len(strings):

                                valor_texto = strings[
                                    indice
                                ]

                        elif valor is not None:

                            valor_texto = (
                                valor.text or ""
                            )

                        valores_linha[
                            coluna
                        ] = valor_texto

                    linhas.append(
                        valores_linha
                    )

                return linhas

        except Exception as erro:

            raise ValueError(
                "Não foi possível ler a planilha.\n"
                f"{erro}"
            )


# ============================================================
# ESCRITOR EXCEL
# ============================================================

class EscritorExcel:

    @staticmethod
    def salvar(
        arquivo,
        usuarios
    ):

        pasta = os.path.dirname(
            os.path.abspath(arquivo)
        )

        arquivo_temp = tempfile.NamedTemporaryFile(
            delete=False,
            suffix=".xlsx",
            dir=pasta
        )

        caminho_temp = arquivo_temp.name

        arquivo_temp.close()

        try:

            namespace = (
                "http://schemas.openxmlformats.org/"
                "spreadsheetml/2006/main"
            )

            rel_namespace = (
                "http://schemas.openxmlformats.org/"
                "officeDocument/2006/relationships"
            )

            ET.register_namespace(
                "",
                namespace
            )

            ET.register_namespace(
                "r",
                rel_namespace
            )

            # =================================================
            # CONTENT TYPES
            # =================================================

            content_types = ET.Element(
                "Types",
                xmlns=(
                    "http://schemas.openxmlformats.org/"
                    "package/2006/content-types"
                )
            )

            ET.SubElement(
                content_types,
                "Default",
                Extension="rels",
                ContentType=(
                    "application/"
                    "vnd.openxmlformats-package."
                    "relationships+xml"
                )
            )

            ET.SubElement(
                content_types,
                "Default",
                Extension="xml",
                ContentType="application/xml"
            )

            ET.SubElement(
                content_types,
                "Override",
                PartName="/xl/workbook.xml",
                ContentType=(
                    "application/"
                    "vnd.openxmlformats-officedocument."
                    "spreadsheetml.sheet.main+xml"
                )
            )

            ET.SubElement(
                content_types,
                "Override",
                PartName="/xl/worksheets/sheet1.xml",
                ContentType=(
                    "application/"
                    "vnd.openxmlformats-officedocument."
                    "spreadsheetml.worksheet+xml"
                )
            )

            ET.SubElement(
                content_types,
                "Override",
                PartName="/xl/styles.xml",
                ContentType=(
                    "application/"
                    "vnd.openxmlformats-officedocument."
                    "spreadsheetml.styles+xml"
                )
            )

            # =================================================
            # WORKBOOK
            # =================================================

            workbook = ET.Element(
                f"{{{namespace}}}workbook"
            )

            sheets = ET.SubElement(
                workbook,
                f"{{{namespace}}}sheets"
            )

            ET.SubElement(
                sheets,
                f"{{{namespace}}}sheet",
                {
                    "name": "Usuarios",
                    "sheetId": "1",
                    f"{{{rel_namespace}}}id":
                    "rId1"
                }
            )

            # =================================================
            # RELACIONAMENTOS
            # =================================================

            rel_namespace_package = (
                "http://schemas.openxmlformats.org/"
                "package/2006/relationships"
            )

            rels = ET.Element(
                "Relationships",
                xmlns=rel_namespace_package
            )

            ET.SubElement(
                rels,
                "Relationship",
                {
                    "Id": "rId1",
                    "Type": (
                        "http://schemas.openxmlformats.org/"
                        "officeDocument/2006/relationships/"
                        "officeDocument"
                    ),
                    "Target": "xl/workbook.xml"
                }
            )

            workbook_rels = ET.Element(
                "Relationships",
                xmlns=rel_namespace_package
            )

            ET.SubElement(
                workbook_rels,
                "Relationship",
                {
                    "Id": "rId1",
                    "Type": (
                        "http://schemas.openxmlformats.org/"
                        "officeDocument/2006/relationships/"
                        "worksheet"
                    ),
                    "Target": "worksheets/sheet1.xml"
                }
            )

            ET.SubElement(
                workbook_rels,
                "Relationship",
                {
                    "Id": "rId2",
                    "Type": (
                        "http://schemas.openxmlformats.org/"
                        "officeDocument/2006/relationships/"
                        "styles"
                    ),
                    "Target": "styles.xml"
                }
            )

            # =================================================
            # STYLES
            # =================================================

            styles = ET.Element(
                f"{{{namespace}}}styleSheet"
            )

            fonts = ET.SubElement(
                styles,
                f"{{{namespace}}}fonts",
                count="3"
            )

            fonte_padrao = ET.SubElement(
                fonts,
                f"{{{namespace}}}font"
            )

            ET.SubElement(
                fonte_padrao,
                f"{{{namespace}}}sz",
                val="11"
            )

            ET.SubElement(
                fonte_padrao,
                f"{{{namespace}}}name",
                val="Aptos"
            )

            fonte_cabecalho = ET.SubElement(
                fonts,
                f"{{{namespace}}}font"
            )

            ET.SubElement(
                fonte_cabecalho,
                f"{{{namespace}}}b"
            )

            ET.SubElement(
                fonte_cabecalho,
                f"{{{namespace}}}sz",
                val="11"
            )

            ET.SubElement(
                fonte_cabecalho,
                f"{{{namespace}}}color",
                rgb="FFFFFFFF"
            )

            ET.SubElement(
                fonte_cabecalho,
                f"{{{namespace}}}name",
                val="Aptos"
            )

            fonte_titulo = ET.SubElement(
                fonts,
                f"{{{namespace}}}font"
            )

            ET.SubElement(
                fonte_titulo,
                f"{{{namespace}}}b"
            )

            ET.SubElement(
                fonte_titulo,
                f"{{{namespace}}}sz",
                val="12"
            )

            ET.SubElement(
                fonte_titulo,
                f"{{{namespace}}}color",
                rgb="FFFFFFFF"
            )

            # =================================================
            # FILLS
            # =================================================

            fills = ET.SubElement(
                styles,
                f"{{{namespace}}}fills",
                count="4"
            )

            ET.SubElement(
                fills,
                f"{{{namespace}}}fill"
            )

            ET.SubElement(
                fills,
                f"{{{namespace}}}fill"
            )

            fill_rosa = ET.SubElement(
                fills,
                f"{{{namespace}}}fill"
            )

            pattern_rosa = ET.SubElement(
                fill_rosa,
                f"{{{namespace}}}patternFill",
                patternType="solid"
            )

            ET.SubElement(
                pattern_rosa,
                f"{{{namespace}}}fgColor",
                rgb="FFE83E8C"
            )

            ET.SubElement(
                pattern_rosa,
                f"{{{namespace}}}bgColor",
                indexed="64"
            )

            fill_claro = ET.SubElement(
                fills,
                f"{{{namespace}}}fill"
            )

            pattern_claro = ET.SubElement(
                fill_claro,
                f"{{{namespace}}}patternFill",
                patternType="solid"
            )

            ET.SubElement(
                pattern_claro,
                f"{{{namespace}}}fgColor",
                rgb="FFF8F5F7"
            )

            ET.SubElement(
                pattern_claro,
                f"{{{namespace}}}bgColor",
                indexed="64"
            )

            # =================================================
            # BORDERS
            # =================================================

            borders = ET.SubElement(
                styles,
                f"{{{namespace}}}borders",
                count="2"
            )

            ET.SubElement(
                borders,
                f"{{{namespace}}}border"
            )

            borda = ET.SubElement(
                borders,
                f"{{{namespace}}}border"
            )

            for lado in [
                "left",
                "right",
                "top",
                "bottom"
            ]:

                elemento = ET.SubElement(
                    borda,
                    f"{{{namespace}}}{lado}"
                )

                elemento.set(
                    "style",
                    "thin"
                )

                ET.SubElement(
                    elemento,
                    f"{{{namespace}}}color",
                    rgb="FFE7DDE4"
                )

            # =================================================
            # CELL STYLE XFS
            # =================================================

            cell_style_xfs = ET.SubElement(
                styles,
                f"{{{namespace}}}cellStyleXfs",
                count="1"
            )

            ET.SubElement(
                cell_style_xfs,
                f"{{{namespace}}}xf"
            )

            # =================================================
            # CELL XFS
            # =================================================

            cell_xfs = ET.SubElement(
                styles,
                f"{{{namespace}}}cellXfs",
                count="4"
            )

            ET.SubElement(
                cell_xfs,
                f"{{{namespace}}}xf",
                numFmtId="0",
                fontId="0",
                fillId="0",
                borderId="0",
                xfId="0"
            )

            ET.SubElement(
                cell_xfs,
                f"{{{namespace}}}xf",
                numFmtId="0",
                fontId="1",
                fillId="2",
                borderId="1",
                xfId="0",
                applyFont="1",
                applyFill="1",
                applyBorder="1"
            )

            ET.SubElement(
                cell_xfs,
                f"{{{namespace}}}xf",
                numFmtId="0",
                fontId="0",
                fillId="3",
                borderId="1",
                xfId="0",
                applyFill="1",
                applyBorder="1"
            )

            ET.SubElement(
                cell_xfs,
                f"{{{namespace}}}xf",
                numFmtId="0",
                fontId="2",
                fillId="2",
                borderId="1",
                xfId="0",
                applyFont="1",
                applyFill="1",
                applyBorder="1"
            )

            # =================================================
            # CELL STYLES
            # =================================================

            cell_styles = ET.SubElement(
                styles,
                f"{{{namespace}}}cellStyles",
                count="1"
            )

            ET.SubElement(
                cell_styles,
                f"{{{namespace}}}cellStyle",
                {
                    "name": "Normal",
                    "xfId": "0",
                    "builtinId": "0"
                }
            )

            # =================================================
            # PLANILHA
            # =================================================

            worksheet = ET.Element(
                f"{{{namespace}}}worksheet"
            )

            sheet_views = ET.SubElement(
                worksheet,
                f"{{{namespace}}}sheetViews"
            )

            ET.SubElement(
                sheet_views,
                f"{{{namespace}}}sheetView",
                {
                    "workbookViewId": "0",
                    "showGridLines": "0"
                }
            )

            ET.SubElement(
                worksheet,
                f"{{{namespace}}}sheetFormatPr",
                {
                    "defaultRowHeight": "22"
                }
            )

            # =================================================
            # COLUNAS
            # =================================================

            cols = ET.SubElement(
                worksheet,
                f"{{{namespace}}}cols"
            )

            larguras = {
                "A": "30",
                "B": "22",
                "C": "38",
                "D": "18",
                "E": "22",
                "F": "14"
            }

            for letra, largura in larguras.items():

                numero_coluna = (
                    ord(letra)
                    - ord("A")
                    + 1
                )

                ET.SubElement(
                    cols,
                    f"{{{namespace}}}col",
                    {
                        "min": str(numero_coluna),
                        "max": str(numero_coluna),
                        "width": largura,
                        "customWidth": "1"
                    }
                )

            # =================================================
            # DADOS
            # =================================================

            sheet_data = ET.SubElement(
                worksheet,
                f"{{{namespace}}}sheetData"
            )

            cabecalho = [
                "Nome",
                "Login",
                "Senha",
                "Status",
                "Nível",
                "Tentativas"
            ]

            linhas = [cabecalho]

            for usuario in usuarios:

                linhas.append([
                    usuario.get("Nome", ""),
                    usuario.get("Login", ""),
                    usuario.get("Senha", ""),
                    usuario.get("Status", ""),
                    usuario.get("Nível", ""),
                    str(
                        usuario.get(
                            "Tentativas",
                            0
                        )
                    )
                ])

            for numero_linha, valores in enumerate(
                linhas,
                start=1
            ):

                row = ET.SubElement(
                    sheet_data,
                    f"{{{namespace}}}row",
                    {
                        "r": str(numero_linha),
                        "ht": "24"
                    }
                )

                for indice, valor in enumerate(
                    valores
                ):

                    letra = chr(
                        ord("A") + indice
                    )

                    estilo = (
                        "1"
                        if numero_linha == 1
                        else "2"
                    )

                    celula = ET.SubElement(
                        row,
                        f"{{{namespace}}}c",
                        {
                            "r":
                            f"{letra}{numero_linha}",
                            "t": "inlineStr",
                            "s": estilo
                        }
                    )

                    is_element = ET.SubElement(
                        celula,
                        f"{{{namespace}}}is"
                    )

                    texto = ET.SubElement(
                        is_element,
                        f"{{{namespace}}}t"
                    )

                    texto.text = str(valor)

            # =================================================
            # FREEZE HEADER
            # =================================================

            sheet_views = worksheet.find(
                f"{{{namespace}}}sheetViews"
            )

            sheet_view = sheet_views.find(
                f"{{{namespace}}}sheetView"
            )

            ET.SubElement(
                sheet_view,
                f"{{{namespace}}}pane",
                {
                    "ySplit": "1",
                    "topLeftCell": "A2",
                    "activePane": "bottomLeft",
                    "state": "frozen"
                }
            )

            # =================================================
            # CRIAR XLSX
            # =================================================

            with zipfile.ZipFile(
                caminho_temp,
                "w",
                zipfile.ZIP_DEFLATED
            ) as arquivo_zip:

                arquivo_zip.writestr(
                    "[Content_Types].xml",
                    ET.tostring(
                        content_types,
                        encoding="utf-8",
                        xml_declaration=True
                    )
                )

                arquivo_zip.writestr(
                    "_rels/.rels",
                    ET.tostring(
                        rels,
                        encoding="utf-8",
                        xml_declaration=True
                    )
                )

                arquivo_zip.writestr(
                    "xl/workbook.xml",
                    ET.tostring(
                        workbook,
                        encoding="utf-8",
                        xml_declaration=True
                    )
                )

                arquivo_zip.writestr(
                    "xl/_rels/workbook.xml.rels",
                    ET.tostring(
                        workbook_rels,
                        encoding="utf-8",
                        xml_declaration=True
                    )
                )

                arquivo_zip.writestr(
                    "xl/styles.xml",
                    ET.tostring(
                        styles,
                        encoding="utf-8",
                        xml_declaration=True
                    )
                )

                arquivo_zip.writestr(
                    "xl/worksheets/sheet1.xml",
                    ET.tostring(
                        worksheet,
                        encoding="utf-8",
                        xml_declaration=True
                    )
                )

            shutil.move(
                caminho_temp,
                arquivo
            )

        except Exception:

            if os.path.exists(caminho_temp):
                os.remove(caminho_temp)

            raise


# ============================================================
# SISTEMA DE LOGIN
# ============================================================

class SistemaLogin:

    def __init__(
        self,
        arquivo
    ):

        self.arquivo = arquivo
        self.usuarios = []

    def carregar_usuarios(self):

        try:

            leitor = LeitorExcel(
                self.arquivo
            )

            linhas = leitor.ler_planilha()

            if not linhas:

                return False, "Planilha vazia."

            cabecalho = linhas[0]

            posicoes = {}

            for coluna in [
                "A",
                "B",
                "C",
                "D",
                "E",
                "F"
            ]:

                nome = cabecalho.get(
                    coluna,
                    ""
                ).strip()

                if nome:
                    posicoes[nome] = coluna

            self.usuarios = []

            for linha in linhas[1:]:

                nome = linha.get(
                    posicoes.get("Nome", "A"),
                    ""
                ).strip()

                login = linha.get(
                    posicoes.get("Login", "B"),
                    ""
                ).strip()

                senha = linha.get(
                    posicoes.get("Senha", "C"),
                    ""
                ).strip()

                status = linha.get(
                    posicoes.get("Status", "D"),
                    ""
                ).strip()

                nivel = linha.get(
                    posicoes.get("Nível", "E"),
                    ""
                ).strip()

                tentativas = linha.get(
                    posicoes.get("Tentativas", "F"),
                    "0"
                ).strip()

                if not nome and not login:
                    continue

                try:
                    tentativas = int(
                        tentativas
                    )
                except ValueError:
                    tentativas = 0

                self.usuarios.append({
                    "Nome": nome,
                    "Login": login,
                    "Senha": senha,
                    "Status": status,
                    "Nível": nivel,
                    "Tentativas": tentativas
                })

            return True, ""

        except Exception as erro:

            return False, str(erro)

    def salvar_usuarios(self):

        EscritorExcel.salvar(
            self.arquivo,
            self.usuarios
        )

    def buscar_por_login(
        self,
        login
    ):

        return next(
            (
                usuario
                for usuario in self.usuarios
                if usuario["Login"].lower()
                == login.strip().lower()
            ),
            None
        )

    def autenticar(
        self,
        login,
        senha
    ):

        usuario = self.buscar_por_login(
            login
        )

        if usuario is None:
            return "incorreto", None

        if (
            usuario["Status"].lower()
            == STATUS_BLOQUEADO.lower()
        ):

            return "bloqueado", usuario

        if (
            usuario["Status"].lower()
            != STATUS_ATIVO.lower()
        ):

            return "inativo", usuario

        senha_correta = False

        if senha_e_hash(
            usuario["Senha"]
        ):

            senha_correta = verificar_hash_senha(
                senha,
                usuario["Senha"]
            )

        else:

            senha_correta = (
                usuario["Senha"]
                == senha.strip()
            )

            if senha_correta:

                usuario["Senha"] = (
                    gerar_hash_senha(
                        senha
                    )
                )

        if not senha_correta:

            usuario["Tentativas"] += 1

            if (
                usuario["Tentativas"]
                >= MAX_TENTATIVAS
            ):

                usuario["Status"] = (
                    STATUS_BLOQUEADO
                )

                self.salvar_usuarios()

                return "bloqueado", usuario

            self.salvar_usuarios()

            return "incorreto", usuario

        usuario["Tentativas"] = 0

        self.salvar_usuarios()

        return "sucesso", usuario

    def cadastrar_usuario(
        self,
        nome,
        login,
        senha,
        status,
        nivel
    ):

        nome = nome.strip()
        login = login.strip()
        senha = senha.strip()

        if not nome:
            return False, "Informe o nome."

        if not login:
            return False, "Informe o login."

        if not senha:
            return False, "Informe a senha."

        if self.buscar_por_login(
            login
        ):

            return False, (
                "Já existe um usuário "
                "com esse login."
            )

        novo_usuario = {
            "Nome": nome,
            "Login": login,
            "Senha": gerar_hash_senha(
                senha
            ),
            "Status": status,
            "Nível": nivel,
            "Tentativas": 0
        }

        self.usuarios.append(
            novo_usuario
        )

        self.salvar_usuarios()

        return True, (
            "Usuário cadastrado com sucesso."
        )

    def cadastrar_autonomo(
        self,
        nome,
        login,
        senha
    ):

        return self.cadastrar_usuario(
            nome,
            login,
            senha,
            status=STATUS_ATIVO,
            nivel=NIVEL_USUARIO
        )

    def alterar_usuario(
        self,
        usuario_original,
        nome,
        login,
        senha,
        status,
        nivel
    ):

        nome = nome.strip()
        login = login.strip()
        senha = senha.strip()

        if not nome:
            return False, "Informe o nome."

        if not login:
            return False, "Informe o login."

        outro_usuario = next(
            (
                u
                for u in self.usuarios
                if u is not usuario_original
                and u["Login"].lower()
                == login.lower()
            ),
            None
        )

        if outro_usuario:

            return False, (
                "Esse login já está "
                "sendo utilizado."
            )

        usuario_original["Nome"] = nome
        usuario_original["Login"] = login
        usuario_original["Status"] = status
        usuario_original["Nível"] = nivel

        if senha:

            usuario_original["Senha"] = (
                gerar_hash_senha(
                    senha
                )
            )

        self.salvar_usuarios()

        return True, (
            "Usuário alterado com sucesso."
        )

    def excluir_usuario(
        self,
        usuario
    ):

        self.usuarios.remove(
            usuario
        )

        self.salvar_usuarios()

    def desbloquear_usuario(
        self,
        usuario
    ):

        usuario["Status"] = STATUS_ATIVO
        usuario["Tentativas"] = 0

        self.salvar_usuarios()


# ============================================================
# BOTÃO ARREDONDADO
# ============================================================

class BotaoArredondado(tk.Canvas):

    def __init__(
        self,
        parent,
        texto,
        comando,
        cor=COR_ROSA,
        cor_hover=COR_ROSA_ESCURO,
        largura=160,
        altura=44,
        raio=20,
        fonte=("Segoe UI", 10, "bold")
    ):

        super().__init__(
            parent,
            width=largura,
            height=altura,
            bg=parent.cget("bg"),
            highlightthickness=0,
            bd=0,
            cursor="hand2"
        )

        self.texto = texto
        self.comando = comando

        self.cor = cor
        self.cor_hover = cor_hover

        self.largura = largura
        self.altura = altura
        self.raio = raio
        self.fonte = fonte

        self.desenhar()

        self.bind(
            "<Enter>",
            self.entrar
        )

        self.bind(
            "<Leave>",
            self.sair
        )

        self.bind(
            "<Button-1>",
            self.clicar
        )

    def desenhar(
        self,
        cor=None
    ):

        if cor is None:
            cor = self.cor

        self.delete("all")

        criar_retangulo_arredondado(
            self,
            2,
            2,
            self.largura - 2,
            self.altura - 2,
            self.raio,
            fill=cor
        )

        self.create_text(
            self.largura // 2,
            self.altura // 2,
            text=self.texto,
            fill=COR_BRANCO,
            font=self.fonte
        )

    def entrar(
        self,
        event=None
    ):

        self.desenhar(
            self.cor_hover
        )

    def sair(
        self,
        event=None
    ):

        self.desenhar(
            self.cor
        )

    def clicar(
        self,
        event=None
    ):

        if self.comando:
            self.comando()

    def alterar_texto(
        self,
        texto
    ):

        self.texto = texto

        self.desenhar()


# ============================================================
# JANELA DE USUÁRIO
# ============================================================

class JanelaUsuario(tk.Toplevel):

    def __init__(
        self,
        master,
        sistema,
        usuario=None,
        callback=None,
        modo_autocadastro=False
    ):

        super().__init__(master)

        self.sistema = sistema
        self.usuario = usuario
        self.callback = callback
        self.modo_autocadastro = (
            modo_autocadastro
        )

        self.editando = (
            usuario is not None
        )

        titulo = "Novo usuário"

        if self.editando:
            titulo = "Alterar usuário"

        elif self.modo_autocadastro:
            titulo = "Criar conta"

        self.title(titulo)

        self.geometry(
            "480x680"
        )

        self.resizable(
            False,
            False
        )

        self.configure(
            bg=COR_ROSA_2
        )

        self.criar_interface()

        if self.editando:
            self.preencher_dados()

        self.grab_set()

    # ========================================================
    # INTERFACE
    # ========================================================

    def criar_interface(self):

        self.canvas = tk.Canvas(
            self,
            highlightthickness=0,
            bd=0
        )

        self.canvas.pack(
            fill="both",
            expand=True
        )

        self.bind(
            "<Configure>",
            self.atualizar_fundo
        )

        self.after(
            50,
            self.atualizar_fundo
        )

        # ====================================================
        # CARD
        # ====================================================

        self.card = tk.Frame(
            self,
            bg=COR_BRANCO
        )

        self.card.place(
            relx=0.5,
            rely=0.5,
            anchor="center",
            width=420,
            height=625
        )

        # ====================================================
        # TÍTULO
        # ====================================================

        titulo_texto = "Novo usuário"
        subtitulo_texto = (
            "Cadastre uma nova conta"
        )

        if self.editando:

            titulo_texto = "Alterar usuário"
            subtitulo_texto = (
                "Atualize os dados da conta"
            )

        elif self.modo_autocadastro:

            titulo_texto = "Cadastre-se"
            subtitulo_texto = (
                "Preencha os dados para criar sua conta"
            )

        tk.Label(
            self.card,
            text="✿",
            font=("Segoe UI Symbol", 28),
            bg=COR_BRANCO,
            fg=COR_ROSA
        ).pack(
            pady=(22, 0)
        )

        tk.Label(
            self.card,
            text=titulo_texto,
            font=("Segoe UI", 20, "bold"),
            bg=COR_BRANCO,
            fg=COR_TEXTO
        ).pack(
            pady=(2, 3)
        )

        tk.Label(
            self.card,
            text=subtitulo_texto,
            font=("Segoe UI", 9),
            bg=COR_BRANCO,
            fg=COR_MUTED
        ).pack(
            pady=(0, 22)
        )

        # ====================================================
        # FORMULÁRIO
        # ====================================================

        formulario = tk.Frame(
            self.card,
            bg=COR_BRANCO
        )

        formulario.pack(
            padx=42,
            fill="x"
        )

        # ====================================================
        # NOME
        # ====================================================

        self.criar_label(
            formulario,
            "Nome completo"
        )

        self.ent_nome = self.criar_entry(
            formulario
        )

        # ====================================================
        # LOGIN
        # ====================================================

        self.criar_label(
            formulario,
            "Login"
        )

        self.ent_login = self.criar_entry(
            formulario
        )

        # ====================================================
        # SENHA
        # ====================================================

        self.criar_label(
            formulario,
            "Senha"
        )

        frame_senha = tk.Frame(
            formulario,
            bg=COR_BRANCO
        )

        frame_senha.pack(
            fill="x",
            pady=(0, 17)
        )

        self.ent_senha = tk.Entry(
            frame_senha,
            font=("Segoe UI", 10),
            show="•",
            bg=COR_ROSA_CLARO,
            fg=COR_TEXTO,
            relief="flat",
            highlightthickness=1,
            highlightbackground=COR_BORDA,
            highlightcolor=COR_ROSA,
            insertbackground=COR_TEXTO
        )

        self.ent_senha.pack(
            side="left",
            fill="x",
            expand=True,
            ipady=9
        )

        self.mostrar_senha = tk.BooleanVar(
            value=False
        )

        tk.Checkbutton(
            frame_senha,
            text="Mostrar",
            variable=self.mostrar_senha,
            command=self.alternar_senha,
            bg=COR_BRANCO,
            fg=COR_MUTED,
            activebackground=COR_BRANCO,
            activeforeground=COR_TEXTO,
            selectcolor=COR_BRANCO,
            bd=0,
            font=("Segoe UI", 8)
        ).pack(
            side="right",
            padx=(8, 0)
        )

        # ====================================================
        # STATUS / NÍVEL
        # ====================================================

        if not self.modo_autocadastro:

            self.criar_label(
                formulario,
                "Status"
            )

            self.combo_status = ttk.Combobox(
                formulario,
                values=[
                    STATUS_ATIVO,
                    STATUS_INATIVO,
                    STATUS_BLOQUEADO
                ],
                state="readonly",
                font=("Segoe UI", 10)
            )

            self.combo_status.pack(
                fill="x",
                pady=(0, 17),
                ipady=6
            )

            self.combo_status.set(
                STATUS_ATIVO
            )

            self.criar_label(
                formulario,
                "Nível de acesso"
            )

            self.combo_nivel = ttk.Combobox(
                formulario,
                values=[
                    NIVEL_USUARIO,
                    NIVEL_ADMINISTRADOR
                ],
                state="readonly",
                font=("Segoe UI", 10)
            )

            self.combo_nivel.pack(
                fill="x",
                pady=(0, 20),
                ipady=6
            )

            self.combo_nivel.set(
                NIVEL_USUARIO
            )

        # ====================================================
        # BOTÃO
        # ====================================================

        btn_texto = "CADASTRAR USUÁRIO"

        if self.editando:
            btn_texto = "SALVAR ALTERAÇÕES"

        elif self.modo_autocadastro:
            btn_texto = "CRIAR MINHA CONTA"

        BotaoArredondado(
            self.card,
            btn_texto,
            self.salvar,
            largura=335,
            altura=46,
            raio=22
        ).pack(
            pady=5
        )

    # ========================================================
    # FUNDO
    # ========================================================

    def atualizar_fundo(
        self,
        event=None
    ):

        largura = self.canvas.winfo_width()
        altura = self.canvas.winfo_height()

        if largura <= 1 or altura <= 1:
            return

        desenhar_degrade(
            self.canvas,
            COR_ROSA_1,
            COR_ROSA_2,
            largura,
            altura
        )

        self.canvas.tag_lower(
            "degrade"
        )

    # ========================================================
    # LABEL
    # ========================================================

    def criar_label(
        self,
        parent,
        texto
    ):

        tk.Label(
            parent,
            text=texto,
            font=("Segoe UI", 9, "bold"),
            bg=COR_BRANCO,
            fg=COR_TEXTO,
            anchor="w"
        ).pack(
            fill="x",
            pady=(0, 6)
        )

    # ========================================================
    # ENTRY
    # ========================================================

    def criar_entry(
        self,
        parent
    ):

        entry = tk.Entry(
            parent,
            font=("Segoe UI", 10),
            bg=COR_ROSA_CLARO,
            fg=COR_TEXTO,
            relief="flat",
            highlightthickness=1,
            highlightbackground=COR_BORDA,
            highlightcolor=COR_ROSA,
            insertbackground=COR_TEXTO
        )

        entry.pack(
            fill="x",
            pady=(0, 17),
            ipady=9
        )

        return entry

    # ========================================================
    # PREENCHER
    # ========================================================

    def preencher_dados(self):

        self.ent_nome.insert(
            0,
            self.usuario["Nome"]
        )

        self.ent_login.insert(
            0,
            self.usuario["Login"]
        )

        if not self.modo_autocadastro:

            self.combo_status.set(
                self.usuario["Status"]
            )

            self.combo_nivel.set(
                self.usuario["Nível"]
            )

    # ========================================================
    # SENHA
    # ========================================================

    def alternar_senha(self):

        if self.mostrar_senha.get():

            self.ent_senha.config(
                show=""
            )

        else:

            self.ent_senha.config(
                show="•"
            )

    # ========================================================
    # SALVAR
    # ========================================================

    def salvar(self):

        nome = self.ent_nome.get()
        login = self.ent_login.get()
        senha = self.ent_senha.get()

        if self.modo_autocadastro:

            sucesso, mensagem = (
                self.sistema.cadastrar_autonomo(
                    nome,
                    login,
                    senha
                )
            )

        elif self.editando:

            status = self.combo_status.get()
            nivel = self.combo_nivel.get()

            sucesso, mensagem = (
                self.sistema.alterar_usuario(
                    self.usuario,
                    nome,
                    login,
                    senha,
                    status,
                    nivel
                )
            )

        else:

            status = self.combo_status.get()
            nivel = self.combo_nivel.get()

            sucesso, mensagem = (
                self.sistema.cadastrar_usuario(
                    nome,
                    login,
                    senha,
                    status,
                    nivel
                )
            )

        if sucesso:

            messagebox.showinfo(
                "Sucesso ✓",
                mensagem,
                parent=self
            )

            if self.callback:
                self.callback()

            self.destroy()

        else:

            messagebox.showerror(
                "Atenção",
                mensagem,
                parent=self
            )


# ============================================================
# PAINEL ADMINISTRATIVO
# ============================================================

class TelaAdministrador(tk.Toplevel):

    def __init__(
        self,
        master,
        sistema,
        administrador
    ):

        super().__init__(master)

        self.sistema = sistema
        self.administrador = administrador

        self.title(
            "Gerenciamento de Usuários"
        )

        self.geometry(
            "1100x700"
        )

        self.minsize(
            950,
            620
        )

        self.configure(
            bg=COR_BACKGROUND
        )

        self.criar_estilos()

        self.criar_interface()

        self.atualizar_tabela()

    # ========================================================
    # ESTILOS
    # ========================================================

    def criar_estilos(self):

        estilo = ttk.Style(self)

        try:
            estilo.theme_use("clam")
        except Exception:
            pass

        # ====================================================
        # COMBOBOX
        # ====================================================

        estilo.configure(
            "TCombobox",
            fieldbackground=COR_ROSA_CLARO,
            background=COR_SURFACE,
            foreground=COR_TEXTO,
            bordercolor=COR_BORDA,
            lightcolor=COR_BORDA,
            darkcolor=COR_BORDA,
            padding=6,
            arrowsize=14
        )

        # ====================================================
        # TREEVIEW
        # ====================================================

        estilo.configure(
            "Tabela.Treeview",
            background=COR_SURFACE,
            foreground=COR_TEXT,
            fieldbackground=COR_SURFACE,
            rowheight=42,
            borderwidth=0,
            relief="flat",
            font=("Segoe UI", 9)
        )

        # ====================================================
        # CABEÇALHO
        # ====================================================

        estilo.configure(
            "Tabela.Treeview.Heading",
            background=COR_PRIMARY,
            foreground=COR_SURFACE,
            relief="flat",
            borderwidth=0,
            padding=(12, 10),
            font=("Segoe UI", 9, "bold")
        )

        # ====================================================
        # SELEÇÃO
        # ====================================================

        estilo.map(
            "Tabela.Treeview",
            background=[
                ("selected", COR_PRIMARY)
            ],
            foreground=[
                ("selected", COR_SURFACE)
            ]
        )

        # ====================================================
        # SCROLLBAR
        # ====================================================

        estilo.configure(
            "Tabela.Vertical.TScrollbar",
            background=COR_PRIMARY,
            troughcolor=COR_BACKGROUND,
            bordercolor=COR_BORDER,
            arrowcolor=COR_SURFACE,
            relief="flat"
        )

        estilo.map(
            "Tabela.Vertical.TScrollbar",
            background=[
                ("active", COR_PRIMARY_DARK)
            ]
        )

    # ========================================================
    # INTERFACE
    # ========================================================

    def criar_interface(self):

        self.canvas = tk.Canvas(
            self,
            highlightthickness=0,
            bd=0
        )

        self.canvas.pack(
            fill="both",
            expand=True
        )

        self.bind(
            "<Configure>",
            self.atualizar_fundo
        )

        self.carregar_imagem()

        self.after(
            100,
            self.atualizar_fundo
        )

        # ====================================================
        # OVERLAY
        # ====================================================

        self.overlay = tk.Frame(
            self,
            bg=COR_SURFACE
        )

        self.overlay.place(
            relx=0.5,
            rely=0.5,
            anchor="center",
            relwidth=0.90,
            relheight=0.88
        )

        # ====================================================
        # CABEÇALHO
        # ====================================================

        cabecalho = tk.Frame(
            self.overlay,
            bg=COR_SURFACE
        )

        cabecalho.pack(
            fill="x",
            padx=38,
            pady=(28, 14)
        )

        tk.Label(
            cabecalho,
            text="Usuários",
            font=("Segoe UI", 25, "bold"),
            bg=COR_SURFACE,
            fg=COR_TEXTO
        ).pack(
            side="left"
        )

        tk.Label(
            cabecalho,
            text=(
                f"Olá, "
                f"{self.administrador['Nome']} 👋"
            ),
            font=("Segoe UI", 9),
            bg=COR_SURFACE,
            fg=COR_MUTED
        ).pack(
            side="right",
            pady=10
        )

        # ====================================================
        # LINHA
        # ====================================================

        tk.Frame(
            self.overlay,
            bg=COR_BORDER,
            height=1
        ).pack(
            fill="x",
            padx=38
        )

        # ====================================================
        # BOTÕES
        # ====================================================

        barra = tk.Frame(
            self.overlay,
            bg=COR_SURFACE
        )

        barra.pack(
            fill="x",
            padx=38,
            pady=18
        )

        BotaoArredondado(
            barra,
            "＋ Novo usuário",
            self.novo_usuario,
            largura=165,
            altura=42,
            raio=21
        ).pack(
            side="left",
            padx=(0, 8)
        )

        BotaoArredondado(
            barra,
            "✎ Alterar",
            self.alterar_usuario,
            cor=COR_SECONDARY,
            cor_hover=COR_SECONDARY_DARK,
            largura=140,
            altura=42,
            raio=21
        ).pack(
            side="left",
            padx=(0, 8)
        )

        BotaoArredondado(
            barra,
            "✕ Excluir",
            self.excluir_usuario,
            cor=COR_EXCLUIR,
            cor_hover="#A93F4D",
            largura=140,
            altura=42,
            raio=21
        ).pack(
            side="left",
            padx=(0, 8)
        )

        BotaoArredondado(
            barra,
            "🔓 Desbloquear",
            self.desbloquear_usuario,
            cor=COR_SUCCESS,
            cor_hover="#257D5A",
            largura=160,
            altura=42,
            raio=21
        ).pack(
            side="left"
        )

        # ====================================================
        # TABELA
        # ====================================================

        frame_tabela = tk.Frame(
            self.overlay,
            bg=COR_SURFACE,
            highlightbackground=COR_BORDER,
            highlightcolor=COR_BORDER,
            highlightthickness=1,
            bd=0
        )

        frame_tabela.pack(
            fill="both",
            expand=True,
            padx=38,
            pady=(4, 12)
        )

        colunas = (
            "Nome",
            "Login",
            "Status",
            "Nivel",
            "Tentativas"
        )

        self.tabela = ttk.Treeview(
            frame_tabela,
            columns=colunas,
            show="headings",
            selectmode="browse",
            style="Tabela.Treeview"
        )

        # ====================================================
        # CABEÇALHOS
        # ====================================================

        self.tabela.heading(
            "Nome",
            text="Nome"
        )

        self.tabela.heading(
            "Login",
            text="Login"
        )

        self.tabela.heading(
            "Status",
            text="Status"
        )

        self.tabela.heading(
            "Nivel",
            text="Nível"
        )

        self.tabela.heading(
            "Tentativas",
            text="Tentativas"
        )

        # ====================================================
        # COLUNAS
        # ====================================================

        self.tabela.column(
            "Nome",
            width=270,
            anchor="w"
        )

        self.tabela.column(
            "Login",
            width=200,
            anchor="w"
        )

        self.tabela.column(
            "Status",
            width=140,
            anchor="center"
        )

        self.tabela.column(
            "Nivel",
            width=180,
            anchor="center"
        )

        self.tabela.column(
            "Tentativas",
            width=110,
            anchor="center"
        )

        scrollbar = ttk.Scrollbar(
            frame_tabela,
            orient="vertical",
            command=self.tabela.yview,
            style="Tabela.Vertical.TScrollbar"
        )

        self.tabela.configure(
            yscrollcommand=scrollbar.set
        )

        self.tabela.pack(
            side="left",
            fill="both",
            expand=True
        )

        scrollbar.pack(
            side="right",
            fill="y"
        )

        self.tabela.bind(
            "<Double-1>",
            lambda event:
            self.alterar_usuario()
        )

        # ====================================================
        # RODAPÉ
        # ====================================================

        rodape = tk.Frame(
            self.overlay,
            bg=COR_SURFACE
        )

        rodape.pack(
            fill="x",
            padx=38,
            pady=(0, 18)
        )

        self.lbl_contador = tk.Label(
            rodape,
            text="Total de usuários: 0",
            font=("Segoe UI", 9, "bold"),
            bg=COR_SURFACE,
            fg=COR_MUTED
        )

        self.lbl_contador.pack(
            side="left"
        )

        tk.Button(
            rodape,
            text="Fechar",
            font=("Segoe UI", 9, "bold"),
            bg=COR_ROSA_CLARO,
            fg=COR_TEXTO,
            activebackground=COR_ROSA,
            activeforeground=COR_BRANCO,
            relief="flat",
            cursor="hand2",
            bd=0,
            command=self.destroy
        ).pack(
            side="right",
            ipadx=20,
            ipady=5
        )

    # ========================================================
    # CARREGAR IMAGEM
    # ========================================================

    def carregar_imagem(self):

        self.imagem_original = None
        self.imagem_fundo = None

        if not TEM_PILLOW:
            return

        if not os.path.exists(
            CAMINHO_IMAGEM
        ):
            return

        try:

            self.imagem_original = Image.open(
                CAMINHO_IMAGEM
            ).convert("RGB")

        except Exception:

            self.imagem_original = None

    # ========================================================
    # FUNDO
    # ========================================================

    def atualizar_fundo(
        self,
        event=None
    ):

        largura = self.canvas.winfo_width()
        altura = self.canvas.winfo_height()

        if largura <= 1 or altura <= 1:
            return

        # ====================================================
        # SEM IMAGEM
        # ====================================================

        if (
            not TEM_PILLOW
            or self.imagem_original is None
        ):

            desenhar_degrade(
                self.canvas,
                COR_ROSA_1,
                COR_ROSA_2,
                largura,
                altura
            )

            return

        try:

            imagem = self.imagem_original.copy()

            proporcao_largura = (
                largura /
                imagem.width
            )

            proporcao_altura = (
                altura /
                imagem.height
            )

            escala = max(
                proporcao_largura,
                proporcao_altura
            )

            nova_largura = max(
                1,
                int(
                    imagem.width *
                    escala
                )
            )

            nova_altura = max(
                1,
                int(
                    imagem.height *
                    escala
                )
            )

            imagem = imagem.resize(
                (
                    nova_largura,
                    nova_altura
                ),
                Image.LANCZOS
            )

            esquerda = max(
                0,
                (
                    nova_largura -
                    largura
                ) // 2
            )

            topo = max(
                0,
                (
                    nova_altura -
                    altura
                ) // 2
            )

            imagem = imagem.crop(
                (
                    esquerda,
                    topo,
                    esquerda + largura,
                    topo + altura
                )
            )

            imagem = ImageEnhance.Brightness(
                imagem
            ).enhance(0.65)

            self.imagem_fundo = ImageTk.PhotoImage(
                imagem
            )

            self.canvas.delete(
                "background"
            )

            self.canvas.create_image(
                0,
                0,
                image=self.imagem_fundo,
                anchor="nw",
                tags="background"
            )

            self.canvas.tag_lower(
                "background"
            )

        except Exception:

            desenhar_degrade(
                self.canvas,
                COR_ROSA_1,
                COR_ROSA_2,
                largura,
                altura
            )

    # ========================================================
    # ATUALIZAR TABELA
    # ========================================================

    def atualizar_tabela(self):

        if not hasattr(
            self,
            "tabela"
        ):
            return

        for item in self.tabela.get_children():

            self.tabela.delete(
                item
            )

        for indice, usuario in enumerate(
            self.sistema.usuarios
        ):

            item = self.tabela.insert(
                "",
                "end",
                iid=str(indice),
                values=(
                    usuario["Nome"],
                    usuario["Login"],
                    usuario["Status"],
                    usuario["Nível"],
                    usuario["Tentativas"]
                )
            )

            if (
                usuario["Status"].lower()
                == STATUS_BLOQUEADO.lower()
            ):

                self.tabela.item(
                    item,
                    tags=("bloqueado",)
                )

            elif (
                usuario["Status"].lower()
                == STATUS_INATIVO.lower()
            ):

                self.tabela.item(
                    item,
                    tags=("inativo",)
                )

            else:

                self.tabela.item(
                    item,
                    tags=("ativo",)
                )

        # ====================================================
        # CORES DAS LINHAS
        # ====================================================

        self.tabela.tag_configure(
            "bloqueado",
            background="#FCE8EE",
            foreground=COR_ERROR
        )

        self.tabela.tag_configure(
            "inativo",
            background="#FFF4E1",
            foreground=COR_WARNING
        )

        self.tabela.tag_configure(
            "ativo",
            background=COR_SURFACE,
            foreground=COR_TEXT
        )

        self.lbl_contador.config(
            text=(
                f"Total de usuários: "
                f"{len(self.sistema.usuarios)}"
            )
        )

    # ========================================================
    # USUÁRIO SELECIONADO
    # ========================================================

    def usuario_selecionado(self):

        selecao = self.tabela.selection()

        if not selecao:

            messagebox.showwarning(
                "Atenção",
                "Selecione um usuário na tabela.",
                parent=self
            )

            return None

        indice = int(
            selecao[0]
        )

        return self.sistema.usuarios[
            indice
        ]

    # ========================================================
    # NOVO
    # ========================================================

    def novo_usuario(self):

        JanelaUsuario(
            self,
            self.sistema,
            callback=self.atualizar_tabela
        )

    # ========================================================
    # ALTERAR
    # ========================================================

    def alterar_usuario(self):

        usuario = self.usuario_selecionado()

        if usuario is None:
            return

        JanelaUsuario(
            self,
            self.sistema,
            usuario=usuario,
            callback=self.atualizar_tabela
        )

    # ========================================================
    # EXCLUIR
    # ========================================================

    def excluir_usuario(self):

        usuario = self.usuario_selecionado()

        if usuario is None:
            return

        if usuario is self.administrador:

            messagebox.showwarning(
                "Operação não permitida",
                (
                    "Você não pode excluir "
                    "a própria conta."
                ),
                parent=self
            )

            return

        resposta = messagebox.askyesno(
            "Confirmar exclusão",
            (
                "Tem certeza que deseja "
                "excluir este usuário?\n\n"
                f"{usuario['Nome']}"
            ),
            parent=self
        )

        if not resposta:
            return

        try:

            self.sistema.excluir_usuario(
                usuario
            )

            messagebox.showinfo(
                "Sucesso ✓",
                "Usuário excluído com sucesso.",
                parent=self
            )

            self.atualizar_tabela()

        except Exception as erro:

            messagebox.showerror(
                "Erro",
                (
                    "Não foi possível excluir.\n"
                    f"{erro}"
                ),
                parent=self
            )

    # ========================================================
    # DESBLOQUEAR
    # ========================================================

    def desbloquear_usuario(self):

        usuario = self.usuario_selecionado()

        if usuario is None:
            return

        if (
            usuario["Status"].lower()
            != STATUS_BLOQUEADO.lower()
        ):

            messagebox.showinfo(
                "Desbloqueio",
                "Este usuário não está bloqueado.",
                parent=self
            )

            return

        resposta = messagebox.askyesno(
            "Desbloquear usuário",
            (
                f"Deseja desbloquear "
                f"{usuario['Nome']}?"
            ),
            parent=self
        )

        if not resposta:
            return

        try:

            self.sistema.desbloquear_usuario(
                usuario
            )

            messagebox.showinfo(
                "Sucesso ✓",
                (
                    "Usuário desbloqueado!\n\n"
                    "Ele já pode realizar login novamente."
                ),
                parent=self
            )

            self.atualizar_tabela()

        except Exception as erro:

            messagebox.showerror(
                "Erro",
                (
                    "Não foi possível desbloquear.\n"
                    f"{erro}"
                ),
                parent=self
            )


# ============================================================
# APLICAÇÃO PRINCIPAL
# ============================================================

class AppMobile(tk.Tk):

    def __init__(self):

        super().__init__()

        self.title(
            "Sistema de Login"
        )

        self.geometry(
            "430x720"
        )

        self.resizable(
            False,
            False
        )

        self.configure(
            bg=COR_ROSA_2
        )

        self.sistema = SistemaLogin(
            ARQUIVO_EXCEL
        )

        sucesso, erro = (
            self.sistema.carregar_usuarios()
        )

        if not sucesso:

            messagebox.showerror(
                "Erro",
                erro
            )

            self.destroy()

            return

        self.criar_layout()

    # ========================================================
    # LAYOUT
    # ========================================================

    def criar_layout(self):

        self.canvas = tk.Canvas(
            self,
            highlightthickness=0,
            bd=0
        )

        self.canvas.pack(
            fill="both",
            expand=True
        )

        self.bind(
            "<Configure>",
            self.atualizar_degrade
        )

        self.after(
            100,
            self.atualizar_degrade
        )

        # ====================================================
        # CARD
        # ====================================================

        self.card = tk.Frame(
            self.canvas,
            bg=COR_BRANCO
        )

        self.canvas.create_window(
            215,
            360,
            window=self.card,
            width=350,
            height=575
        )

        # ====================================================
        # ÍCONE
        # ====================================================

        tk.Label(
            self.card,
            text="🌸",
            font=("Segoe UI Emoji", 42),
            bg=COR_BRANCO,
            fg=COR_ROSA
        ).pack(
            pady=(18, 5)
        )

        # ====================================================
        # TÍTULO
        # ====================================================

        tk.Label(
            self.card,
            text="Bem-vindo",
            font=("Segoe UI", 22, "bold"),
            bg=COR_BRANCO,
            fg=COR_TEXTO
        ).pack()

        tk.Label(
            self.card,
            text="Acesse sua conta para continuar",
            font=("Segoe UI", 9),
            bg=COR_BRANCO,
            fg=COR_MUTED
        ).pack(
            pady=(5, 30)
        )

        # ====================================================
        # USUÁRIO
        # ====================================================

        tk.Label(
            self.card,
            text="Usuário",
            font=("Segoe UI", 9, "bold"),
            bg=COR_BRANCO,
            fg=COR_TEXTO,
            anchor="w"
        ).pack(
            fill="x",
            padx=20,
            pady=(0, 6)
        )

        self.ent_login = tk.Entry(
            self.card,
            font=("Segoe UI", 11),
            bg=COR_ROSA_CLARO,
            fg=COR_TEXTO,
            relief="flat",
            highlightthickness=1,
            highlightbackground=COR_BORDA,
            highlightcolor=COR_ROSA,
            insertbackground=COR_TEXTO
        )

        self.ent_login.pack(
            fill="x",
            padx=20,
            ipady=10,
            pady=(0, 20)
        )

        # ====================================================
        # SENHA
        # ====================================================

        tk.Label(
            self.card,
            text="Senha",
            font=("Segoe UI", 9, "bold"),
            bg=COR_BRANCO,
            fg=COR_TEXTO,
            anchor="w"
        ).pack(
            fill="x",
            padx=20,
            pady=(0, 6)
        )

        self.ent_senha = tk.Entry(
            self.card,
            font=("Segoe UI", 11),
            show="•",
            bg=COR_ROSA_CLARO,
            fg=COR_TEXTO,
            relief="flat",
            highlightthickness=1,
            highlightbackground=COR_BORDA,
            highlightcolor=COR_ROSA,
            insertbackground=COR_TEXTO
        )

        self.ent_senha.pack(
            fill="x",
            padx=20,
            ipady=10,
            pady=(0, 8)
        )

        # ====================================================
        # MOSTRAR SENHA
        # ====================================================

        self.mostrar_senha = tk.BooleanVar(
            value=False
        )

        tk.Checkbutton(
            self.card,
            text="Mostrar senha",
            variable=self.mostrar_senha,
            command=self.alternar_senha,
            font=("Segoe UI", 8),
            bg=COR_BRANCO,
            fg=COR_MUTED,
            activebackground=COR_BRANCO,
            activeforeground=COR_TEXTO,
            selectcolor=COR_BRANCO,
            bd=0
        ).pack(
            anchor="w",
            padx=20,
            pady=(0, 18)
        )

        # ====================================================
        # ENTRAR
        # ====================================================

        self.btn_entrar = BotaoArredondado(
            self.card,
            "ENTRAR",
            self.executar_login,
            largura=310,
            altura=44,
            raio=22,
            fonte=("Segoe UI", 10, "bold")
        )

        self.btn_entrar.pack(
            pady=(0, 8)
        )

        # ====================================================
        # CADASTRAR
        # ====================================================

        self.btn_cadastrar = BotaoArredondado(
            self.card,
            "CADASTRAR-SE",
            self.abrir_cadastro,
            cor=COR_SECONDARY,
            cor_hover=COR_SECONDARY_DARK,
            largura=310,
            altura=44,
            raio=22,
            fonte=("Segoe UI", 10, "bold")
        )

        self.btn_cadastrar.pack(
            pady=(0, 12)
        )

        # ====================================================
        # ESQUECI SENHA
        # ====================================================

        tk.Button(
            self.card,
            text="Esqueceu a senha?",
            font=("Segoe UI", 9),
            fg=COR_MUTED,
            bg=COR_BRANCO,
            activebackground=COR_BRANCO,
            activeforeground=COR_ROSA,
            bd=0,
            cursor="hand2",
            command=self.esqueci_senha
        ).pack()

        # ====================================================
        # ENTER
        # ====================================================

        self.bind(
            "<Return>",
            lambda event:
            self.executar_login()
        )

        self.ent_login.focus_set()

    # ========================================================
    # DEGRADÊ
    # ========================================================

    def atualizar_degrade(
        self,
        event=None
    ):

        largura = self.canvas.winfo_width()
        altura = self.canvas.winfo_height()

        if largura <= 1 or altura <= 1:
            return

        desenhar_degrade(
            self.canvas,
            COR_ROSA_1,
            COR_ROSA_2,
            largura,
            altura
        )

        self.canvas.tag_lower(
            "degrade"
        )

    # ========================================================
    # SENHA
    # ========================================================

    def alternar_senha(self):

        if self.mostrar_senha.get():

            self.ent_senha.config(
                show=""
            )

        else:

            self.ent_senha.config(
                show="•"
            )

    # ========================================================
    # ESQUECI SENHA
    # ========================================================

    def esqueci_senha(self):

        messagebox.showinfo(
            "Recuperação de senha",
            (
                "Entre em contato com o "
                "administrador do sistema "
                "para redefinir sua senha."
            ),
            parent=self
        )

    # ========================================================
    # CADASTRO
    # ========================================================

    def abrir_cadastro(self):

        JanelaUsuario(
            self,
            self.sistema,
            modo_autocadastro=True
        )

    # ========================================================
    # LOGIN
    # ========================================================

    def executar_login(self):

        login = self.ent_login.get().strip()
        senha = self.ent_senha.get()

        if not login:

            messagebox.showwarning(
                "Atenção",
                "Informe o usuário.",
                parent=self
            )

            self.ent_login.focus_set()

            return

        if not senha:

            messagebox.showwarning(
                "Atenção",
                "Informe a senha.",
                parent=self
            )

            self.ent_senha.focus_set()

            return

        self.btn_entrar.alterar_texto(
            "Entrando..."
        )

        self.update_idletasks()

        self.after(
            120,
            lambda:
            self.finalizar_login(
                login,
                senha
            )
        )

    # ========================================================
    # FINALIZAR LOGIN
    # ========================================================

    def finalizar_login(
        self,
        login,
        senha
    ):

        status, usuario = (
            self.sistema.autenticar(
                login,
                senha
            )
        )

        # ====================================================
        # SUCESSO
        # ====================================================

        if status == "sucesso":

            self.btn_entrar.cor = COR_SUCESSO
            self.btn_entrar.cor_hover = COR_SUCESSO

            self.btn_entrar.alterar_texto(
                "✓ Acesso realizado!"
            )

            self.update_idletasks()

            self.after(
                600,
                lambda:
                self.abrir_sistema(
                    usuario
                )
            )

            return

        # ====================================================
        # RESTAURAR
        # ====================================================

        self.btn_entrar.cor = COR_ROSA
        self.btn_entrar.cor_hover = (
            COR_ROSA_ESCURO
        )

        self.btn_entrar.alterar_texto(
            "ENTRAR"
        )

        # ====================================================
        # BLOQUEADO
        # ====================================================

        if status == "bloqueado":

            messagebox.showerror(
                "Conta bloqueada",
                (
                    "Sua conta foi bloqueada "
                    f"após {MAX_TENTATIVAS} tentativas.\n\n"
                    "Procure o administrador "
                    "para realizar o desbloqueio."
                ),
                parent=self
            )

            self.ent_senha.delete(
                0,
                tk.END
            )

            return

        # ====================================================
        # INATIVO
        # ====================================================

        if status == "inativo":

            messagebox.showwarning(
                "Usuário inativo",
                (
                    "Usuário desativado. "
                    "Procure seu líder.\n\n"
                    f"Nome: {usuario['Nome']}"
                ),
                parent=self
            )

            return

        # ====================================================
        # INCORRETO
        # ====================================================

        tentativas = 0

        if usuario:

            tentativas = usuario[
                "Tentativas"
            ]

        restantes = (
            MAX_TENTATIVAS -
            tentativas
        )

        messagebox.showerror(
            "Login incorreto",
            (
                "Login ou senha incorretos.\n\n"
                f"Tentativas restantes: {restantes}"
            ),
            parent=self
        )

        self.ent_senha.delete(
            0,
            tk.END
        )

        self.ent_senha.focus_set()

    # ========================================================
    # ABRIR SISTEMA
    # ========================================================

    def abrir_sistema(
        self,
        usuario
    ):

        self.ent_senha.delete(
            0,
            tk.END
        )

        messagebox.showinfo(
            "Login bem-sucedido! ✓",
            (
                f"Seja bem-vindo(a), "
                f"{usuario['Nome']}!"
            ),
            parent=self
        )

        if (
            usuario["Nível"].lower()
            == NIVEL_ADMINISTRADOR.lower()
        ):

            TelaAdministrador(
                self,
                self.sistema,
                usuario
            )


# ============================================================
# INICIAR SISTEMA
# ============================================================

if __name__ == "__main__":

    app = AppMobile()

    app.mainloop()