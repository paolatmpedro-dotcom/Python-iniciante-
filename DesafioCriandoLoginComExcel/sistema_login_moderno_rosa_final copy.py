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
# TENTA IMPORTAR PILLOW
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
# CORES
# ============================================================

COR_ROSA_1 = "#fbc2eb"
COR_ROSA_2 = "#ec6fae"

COR_ROSA = "#ec6fae"
COR_ROSA_ESCURO = "#c94f88"
COR_ROSA_CLARO = "#fff1f8"

COR_BRANCO = "#ffffff"
COR_FUNDO = "#fff8fc"

COR_TEXTO = "#3d2935"
COR_MUTED = "#927582"

COR_BORDA = "#f1c9dc"

COR_SUCESSO = "#42a878"
COR_ERRO = "#d95d75"
COR_ALERTA = "#d99a35"
COR_EXCLUIR = "#c94f5d"

COR_SOMBRA = "#dba8c1"


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
# DEGRADÊ
# ============================================================

def desenhar_degrade(canvas, cor1, cor2, largura, altura):

    if largura <= 0 or altura <= 0:
        return

    r1, g1, b1 = canvas.winfo_rgb(cor1)
    r2, g2, b2 = canvas.winfo_rgb(cor2)

    r_ratio = (r2 - r1) / max(altura, 1)
    g_ratio = (g2 - g1) / max(altura, 1)
    b_ratio = (b2 - b1) / max(altura, 1)

    canvas.delete("degrade")

    for i in range(altura):

        nr = int(r1 + r_ratio * i) >> 8
        ng = int(g1 + g_ratio * i) >> 8
        nb = int(b1 + b_ratio * i) >> 8

        cor = f"#{nr:02x}{ng:02x}{nb:02x}"

        canvas.create_line(
            0,
            i,
            largura,
            i,
            fill=cor,
            tags="degrade"
        )

    canvas.tag_lower("degrade")


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
    def salvar(arquivo, usuarios):

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

            rels = ET.Element(
                "Relationships",
                xmlns=(
                    "http://schemas.openxmlformats.org/"
                    "package/2006/relationships"
                )
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
                xmlns=(
                    "http://schemas.openxmlformats.org/"
                    "package/2006/relationships"
                )
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
                count="1"
            )

            ET.SubElement(
                fonts,
                f"{{{namespace}}}font"
            )

            fills = ET.SubElement(
                styles,
                f"{{{namespace}}}fills",
                count="2"
            )

            ET.SubElement(
                fills,
                f"{{{namespace}}}fill"
            )

            ET.SubElement(
                fills,
                f"{{{namespace}}}fill"
            )

            borders = ET.SubElement(
                styles,
                f"{{{namespace}}}borders",
                count="1"
            )

            ET.SubElement(
                borders,
                f"{{{namespace}}}border"
            )

            cell_style_xfs = ET.SubElement(
                styles,
                f"{{{namespace}}}cellStyleXfs",
                count="1"
            )

            ET.SubElement(
                cell_style_xfs,
                f"{{{namespace}}}xf"
            )

            cell_xfs = ET.SubElement(
                styles,
                f"{{{namespace}}}cellXfs",
                count="1"
            )

            ET.SubElement(
                cell_xfs,
                f"{{{namespace}}}xf",
                numFmtId="0",
                fontId="0",
                fillId="0",
                borderId="0"
            )

            # =================================================
            # PLANILHA
            # =================================================

            worksheet = ET.Element(
                f"{{{namespace}}}worksheet"
            )

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
                    r=str(numero_linha)
                )

                for indice, valor in enumerate(
                    valores
                ):

                    letra = chr(
                        ord("A") + indice
                    )

                    celula = ET.SubElement(
                        row,
                        f"{{{namespace}}}c",
                        {
                            "r":
                            f"{letra}{numero_linha}",
                            "t": "inlineStr"
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

    def __init__(self, arquivo):

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
                    tentativas = int(tentativas)
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

    def buscar_por_login(self, login):

        return next(
            (
                usuario
                for usuario in self.usuarios
                if usuario["Login"].lower()
                == login.strip().lower()
            ),
            None
        )

    def autenticar(self, login, senha):

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

        if senha_e_hash(usuario["Senha"]):

            senha_correta = verificar_hash_senha(
                senha,
                usuario["Senha"]
            )

        else:

            # Compatibilidade com senha antiga
            senha_correta = (
                usuario["Senha"]
                == senha.strip()
            )

            if senha_correta:

                usuario["Senha"] = gerar_hash_senha(
                    senha
                )

        if not senha_correta:

            usuario["Tentativas"] += 1

            if (
                usuario["Tentativas"]
                >= MAX_TENTATIVAS
            ):

                usuario["Status"] = STATUS_BLOQUEADO

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

        if self.buscar_por_login(login):

            return False, (
                "Já existe um usuário "
                "com esse login."
            )

        novo_usuario = {
            "Nome": nome,
            "Login": login,
            "Senha": gerar_hash_senha(senha),
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

    def cadastrar_autonomo(self, nome, login, senha):
        """Cadastra novos usuários via tela de login com perfil padrão"""
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
                gerar_hash_senha(senha)
            )

        self.salvar_usuarios()

        return True, (
            "Usuário alterado com sucesso."
        )

    def excluir_usuario(self, usuario):

        self.usuarios.remove(
            usuario
        )

        self.salvar_usuarios()

    def desbloquear_usuario(self, usuario):

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

    def desenhar(self, cor=None):

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
            fill="white",
            font=self.fonte
        )

    def entrar(self, event):

        self.desenhar(
            self.cor_hover
        )

    def sair(self, event):

        self.desenhar(
            self.cor
        )

    def clicar(self, event):

        if self.comando:
            self.comando()

    def alterar_texto(self, texto):

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
        self.modo_autocadastro = modo_autocadastro

        self.editando = usuario is not None

        titulo = "Novo usuário"
        if self.editando:
            titulo = "Alterar usuário"
        elif self.modo_autocadastro:
            titulo = "Criar Conta"

        self.title(titulo)

        self.geometry(
            "460x650"
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

    def criar_interface(self):

        canvas = tk.Canvas(
            self,
            highlightthickness=0
        )

        canvas.pack(
            fill="both",
            expand=True
        )

        desenhar_degrade(
            canvas,
            COR_ROSA_1,
            COR_ROSA_2,
            460,
            650
        )

        card = tk.Frame(
            canvas,
            bg=COR_BRANCO
        )

        canvas.create_window(
            230,
            325,
            window=card,
            width=390,
            height=590
        )

        titulo_texto = "Novo usuário"
        subtitulo_texto = "Cadastre uma nova conta"

        if self.editando:
            titulo_texto = "Alterar usuário"
            subtitulo_texto = "Atualize os dados da conta"
        elif self.modo_autocadastro:
            titulo_texto = "Cadastre-se"
            subtitulo_texto = "Preencha os dados para criar sua conta"

        tk.Label(
            card,
            text=titulo_texto,
            font=("Segoe UI", 20, "bold"),
            bg=COR_BRANCO,
            fg=COR_TEXTO
        ).pack(
            pady=(25, 5)
        )

        tk.Label(
            card,
            text=subtitulo_texto,
            font=("Segoe UI", 9),
            bg=COR_BRANCO,
            fg=COR_MUTED
        ).pack(
            pady=(0, 20)
        )

        formulario = tk.Frame(
            card,
            bg=COR_BRANCO
        )

        formulario.pack(
            padx=30,
            fill="x"
        )

        self.criar_label(
            formulario,
            "Nome"
        )

        self.ent_nome = self.criar_entry(
            formulario
        )

        self.criar_label(
            formulario,
            "Login"
        )

        self.ent_login = self.criar_entry(
            formulario
        )

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
            pady=(0, 15)
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
            highlightcolor=COR_ROSA
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
            selectcolor=COR_BRANCO,
            bd=0
        ).pack(
            side="right",
            padx=(8, 0)
        )

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
                pady=(0, 15),
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

        btn_texto = "CADASTRAR USUÁRIO"
        if self.editando:
            btn_texto = "SALVAR ALTERAÇÕES"
        elif self.modo_autocadastro:
            btn_texto = "CRIAR MINHA CONTA"

        BotaoArredondado(
            card,
            btn_texto,
            self.salvar,
            largura=320,
            altura=46
        ).pack(
            pady=5
        )

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
            fill="x"
        )

    def criar_entry(self, parent):

        entry = tk.Entry(
            parent,
            font=("Segoe UI", 10),
            bg=COR_ROSA_CLARO,
            fg=COR_TEXTO,
            relief="flat",
            highlightthickness=1,
            highlightbackground=COR_BORDA,
            highlightcolor=COR_ROSA
        )

        entry.pack(
            fill="x",
            pady=(2, 15),
            ipady=9
        )

        return entry

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

    def alternar_senha(self):

        if self.mostrar_senha.get():

            self.ent_senha.config(
                show=""
            )

        else:

            self.ent_senha.config(
                show="•"
            )

    def salvar(self):

        nome = self.ent_nome.get()
        login = self.ent_login.get()
        senha = self.ent_senha.get()

        if self.modo_autocadastro:
            sucesso, mensagem = self.sistema.cadastrar_autonomo(
                nome,
                login,
                senha
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
            "1050x680"
        )

        self.minsize(
            950,
            620
        )

        self.criar_interface()

        self.atualizar_tabela()

    def criar_interface(self):

        self.canvas = tk.Canvas(
            self,
            highlightthickness=0
        )

        self.canvas.pack(
            fill="both",
            expand=True
        )

        self.bind(
            "<Configure>",
            self.atualizar_fundo
        )

        # =====================================================
        # IMAGEM DE FUNDO
        # =====================================================

        self.carregar_imagem()

        # =====================================================
        # CAMADA
        # =====================================================

        self.overlay = tk.Canvas(
            self.canvas,
            highlightthickness=0
        )

        self.overlay.place(
            relx=0,
            rely=0,
            relwidth=1,
            relheight=1
        )

        # =====================================================
        # CARD PRINCIPAL
        # =====================================================

        self.card = tk.Frame(
            self.overlay,
            bg=COR_BRANCO
        )

        self.overlay.create_window(
            525,
            340,
            window=self.card,
            width=970,
            height=610
        )

        # =====================================================
        # CABEÇALHO
        # =====================================================

        cabecalho = tk.Frame(
            self.card,
            bg=COR_BRANCO
        )

        cabecalho.pack(
            fill="x",
            padx=35,
            pady=(25, 10)
        )

        tk.Label(
            cabecalho,
            text="Usuários",
            font=("Segoe UI", 24, "bold"),
            bg=COR_BRANCO,
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
            bg=COR_BRANCO,
            fg=COR_MUTED
        ).pack(
            side="right",
            pady=8
        )

        # =====================================================
        # BOTÕES
        # =====================================================

        barra = tk.Frame(
            self.card,
            bg=COR_BRANCO
        )

        barra.pack(
            fill="x",
            padx=35,
            pady=10
        )

        BotaoArredondado(
            barra,
            "＋ Novo usuário",
            self.novo_usuario,
            largura=160,
            altura=42
        ).pack(
            side="left",
            padx=(0, 8)
        )

        BotaoArredondado(
            barra,
            "✎ Alterar",
            self.alterar_usuario,
            cor="#8e70a8",
            cor_hover="#72578b",
            largura=140,
            altura=42
        ).pack(
            side="left",
            padx=(0, 8)
        )

        BotaoArredondado(
            barra,
            "✕ Excluir",
            self.excluir_usuario,
            cor=COR_EXCLUIR,
            cor_hover="#a93f4d",
            largura=140,
            altura=42
        ).pack(
            side="left",
            padx=(0, 8)
        )

        BotaoArredondado(
            barra,
            "🔓 Desbloquear",
            self.desbloquear_usuario,
            cor=COR_SUCESSO,
            cor_hover="#32865f",
            largura=160,
            altura=42
        ).pack(
            side="left"
        )

        # =====================================================
        # TABELA
        # =====================================================

        frame_tabela = tk.Frame(
            self.card,
            bg=COR_BRANCO
        )

        frame_tabela.pack(
            fill="both",
            expand=True,
            padx=35,
            pady=10
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
            selectmode="browse"
        )

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

        self.tabela.column(
            "Nome",
            width=250
        )

        self.tabela.column(
            "Login",
            width=180
        )

        self.tabela.column(
            "Status",
            width=120,
            anchor="center"
        )

        self.tabela.column(
            "Nivel",
            width=170,
            anchor="center"
        )

        self.tabela.column(
            "Tentativas",
            width=100,
            anchor="center"
        )

        scrollbar = ttk.Scrollbar(
            frame_tabela,
            orient="vertical",
            command=self.tabela.yview
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
            lambda event: self.alterar_usuario()
        )

        # =====================================================
        # RODAPÉ
        # =====================================================

        rodape = tk.Frame(
            self.card,
            bg=COR_BRANCO
        )

        rodape.pack(
            fill="x",
            padx=35,
            pady=(5, 20)
        )

        self.lbl_contador = tk.Label(
            rodape,
            text="Total de usuários: 0",
            font=("Segoe UI", 9, "bold"),
            bg=COR_BRANCO,
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
            activeforeground="white",
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
    # IMAGEM
    # ========================================================

    def carregar_imagem(self):

        self.imagem_original = None
        self.imagem_fundo = None

        if TEM_PILLOW and os.path.exists(
            CAMINHO_IMAGEM
        ):

            try:

                self.imagem_original = Image.open(
                    CAMINHO_IMAGEM
                )

            except Exception:

                self.imagem_original = None

    def atualizar_fundo(self, event=None):

        largura = self.winfo_width()
        altura = self.winfo_height()

        if (
            largura <= 1
            or altura <= 1
        ):
            return

        self.canvas.delete("fundo")

        if (
            TEM_PILLOW
            and self.imagem_original is not None
        ):

            try:

                imagem = self.imagem_original.copy()

                imagem.thumbnail(
                    (largura, altura),
                    Image.LANCZOS
                )

                fundo = Image.new(
                    "RGB",
                    (largura, altura),
                    "#f7dce9"
                )

                x = (
                    largura
                    - imagem.width
                ) // 2

                y = (
                    altura
                    - imagem.height
                ) // 2

                fundo.paste(
                    imagem,
                    (x, y)
                )

                fundo = ImageEnhance.Brightness(
                    fundo
                ).enhance(0.75)

                self.imagem_fundo = ImageTk.PhotoImage(
                    fundo
                )

                self.canvas.create_image(
                    0,
                    0,
                    image=self.imagem_fundo,
                    anchor="nw",
                    tags="fundo"
                )

            except Exception:

                self.canvas.create_rectangle(
                    0,
                    0,
                    largura,
                    altura,
                    fill=COR_ROSA_CLARO,
                    outline="",
                    tags="fundo"
                )

        else:

            self.canvas.create_rectangle(
                0,
                0,
                largura,
                altura,
                fill=COR_ROSA_CLARO,
                outline="",
                tags="fundo"
            )

        self.canvas.tag_lower(
            "fundo"
        )

    # ========================================================
    # TABELA
    # ========================================================

    def atualizar_tabela(self):

        if not hasattr(
            self,
            "tabela"
        ):
            return

        for item in self.tabela.get_children():

            self.tabela.delete(item)

        for indice, usuario in enumerate(
            self.sistema.usuarios
        ):

            self.tabela.insert(
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

        self.lbl_contador.config(
            text=(
                f"Total de usuários: "
                f"{len(self.sistema.usuarios)}"
            )
        )

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
                "Você não pode excluir "
                "a própria conta.",
                parent=self
            )

            return

        resposta = messagebox.askyesno(
            "Confirmar exclusão",
            (
                "Tem certeza que deseja excluir "
                "este usuário?\n\n"
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
                f"Não foi possível excluir.\n{erro}",
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
                f"Não foi possível desbloquear.\n{erro}",
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
            highlightthickness=0
        )

        self.canvas.pack(
            fill="both",
            expand=True
        )

        self.bind(
            "<Configure>",
            self.atualizar_degrade
        )

        # =====================================================
        # CARD PRINCIPAL
        # =====================================================

        criar_retangulo_arredondado(
            self.canvas,
            35,
            65,
            395,
            655,
            35,
            fill=COR_BRANCO
        )

        self.card = tk.Frame(
            self.canvas,
            bg=COR_BRANCO
        )

        self.canvas.create_window(
            215,
            360,
            window=self.card,
            width=310,
            height=530
        )

        # =====================================================
        # ÍCONE
        # =====================================================

        tk.Label(
            self.card,
            text="🌸",
            font=("Segoe UI Emoji", 42),
            bg=COR_BRANCO,
            fg=COR_ROSA
        ).pack(
            pady=(5, 5)
        )

        # =====================================================
        # TÍTULO
        # =====================================================

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
            pady=(3, 25)
        )

        # =====================================================
        # USUÁRIO
        # =====================================================

        tk.Label(
            self.card,
            text="Usuário",
            font=("Segoe UI", 9, "bold"),
            bg=COR_BRANCO,
            fg=COR_TEXTO,
            anchor="w"
        ).pack(
            fill="x"
        )

        self.ent_login = tk.Entry(
            self.card,
            font=("Segoe UI", 11),
            bg=COR_ROSA_CLARO,
            fg=COR_TEXTO,
            relief="flat",
            highlightthickness=1,
            highlightbackground=COR_BORDA,
            highlightcolor=COR_ROSA
        )

        self.ent_login.pack(
            fill="x",
            ipady=10,
            pady=(5, 18)
        )

        # =====================================================
        # SENHA
        # =====================================================

        tk.Label(
            self.card,
            text="Senha",
            font=("Segoe UI", 9, "bold"),
            bg=COR_BRANCO,
            fg=COR_TEXTO,
            anchor="w"
        ).pack(
            fill="x"
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
            highlightcolor=COR_ROSA
        )

        self.ent_senha.pack(
            fill="x",
            ipady=10,
            pady=(5, 8)
        )

        # =====================================================
        # MOSTRAR SENHA
        # =====================================================

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
            selectcolor=COR_BRANCO,
            bd=0
        ).pack(
            anchor="w",
            pady=(0, 18)
        )

        # =====================================================
        # BOTÃO ENTRAR
        # =====================================================

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

        # =====================================================
        # BOTÃO CADASTRAR-SE (NOVO)
        # =====================================================

        self.btn_cadastrar = BotaoArredondado(
            self.card,
            "CADASTRAR-SE",
            self.abrir_cadastro,
            cor="#8e70a8",
            cor_hover="#72578b",
            largura=310,
            altura=44,
            raio=22,
            fonte=("Segoe UI", 10, "bold")
        )

        self.btn_cadastrar.pack(
            pady=(0, 12)
        )

        # =====================================================
        # ESQUECI
        # =====================================================

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

        # =====================================================
        # ENTER
        # =====================================================

        self.bind(
            "<Return>",
            lambda event: self.executar_login()
        )

        self.ent_login.focus_set()

    # ========================================================
    # DEGRADÊ
    # ========================================================

    def atualizar_degrade(self, event=None):

        desenhar_degrade(
            self.canvas,
            COR_ROSA_1,
            COR_ROSA_2,
            self.winfo_width(),
            self.winfo_height()
        )

        self.canvas.tag_lower(
            "degrade"
        )

    # ========================================================
    # MOSTRAR SENHA
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
    # CADASTRO DE USUÁRIO
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

        # ====================================================
        # EFEITO ENTRANDO...
        # ====================================================

        self.btn_entrar.alterar_texto(
            "Entrando..."
        )

        self.update_idletasks()

        self.after(
            120,
            lambda: self.finalizar_login(
                login,
                senha
            )
        )

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
                lambda: self.abrir_sistema(
                    usuario
                )
            )

            return

        # ====================================================
        # RESTAURA BOTÃO
        # ====================================================

        self.btn_entrar.cor = COR_ROSA
        self.btn_entrar.cor_hover = COR_ROSA_ESCURO

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
            MAX_TENTATIVAS
            - tentativas
        )

        messagebox.showerror(
            "Erro",
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

    def abrir_sistema(self, usuario):

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
# INICIAR
# ============================================================

if __name__ == "__main__":

    app = AppMobile()

    app.mainloop()