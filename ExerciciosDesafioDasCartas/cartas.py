import tkinter as tk
from tkinter import messagebox
import os
import json
import random
import zipfile
import xml.etree.ElementTree as ET

# ============================================================
# CONFIGURAÇÕES E DESIGN SYSTEM
# ============================================================

PASTA_PROGRAMA = os.path.dirname(os.path.abspath(__file__))
ARQUIVO_USUARIOS = os.path.join(PASTA_PROGRAMA, "usuarios.xlsx")
ARQUIVO_ESTATISTICAS = os.path.join(PASTA_PROGRAMA, "estatisticas.json")
PASTA_CARTAS = os.path.join(PASTA_PROGRAMA, "cartas")

COLOR_BG = "#0D1117"
COLOR_CANVAS = "#161B22"
COLOR_FELT = "#0F382C"
COLOR_FELT_LIGHT = "#154D3D"
COLOR_ACCENT = "#238636"
COLOR_GOLD = "#F1E05A"
COLOR_TEXT = "#E6EDE3"
COLOR_TEXT_MUTED = "#8B949E"
COLOR_DANGER = "#DA3633"

FONT_PRIMARY = ("Segoe UI", 10)
FONT_BOLD = ("Segoe UI", 10, "bold")
FONT_TITLE = ("Segoe UI", 18, "bold")
FONT_CARDS = ("Segoe UI", 14, "bold")


# ============================================================
# LEITOR DE EXCEL SEM OPENPYXL/PANDAS
# ============================================================

class LeitorExcel:
    NS = {
        "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
        "rel": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    }

    def __init__(self, arquivo):
        self.arquivo = arquivo

    def coluna_para_numero(self, referencia):
        letras = ""
        for caractere in referencia:
            if caractere.isalpha():
                letras += caractere
            else:
                break
        numero = 0
        for letra in letras:
            numero = numero * 26 + ord(letra.upper()) - ord("A") + 1
        return numero - 1

    def ler_planilha(self):
        if not os.path.exists(self.arquivo):
            raise FileNotFoundError(f"Arquivo não encontrado:\n{self.arquivo}")

        with zipfile.ZipFile(self.arquivo, "r") as arquivo_zip:
            nomes = arquivo_zip.namelist()
            strings = []

            if "xl/sharedStrings.xml" in nomes:
                raiz = ET.fromstring(arquivo_zip.read("xl/sharedStrings.xml"))
                for si in raiz:
                    texto = ""
                    for item in si.iter("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t"):
                        texto += item.text or ""
                    strings.append(texto)

            caminho_planilha = "xl/worksheets/sheet1.xml"
            if caminho_planilha not in nomes:
                planilhas = [n for n in nomes if n.startswith("xl/worksheets/") and n.endswith(".xml")]
                if not planilhas:
                    raise ValueError("Não foi possível encontrar uma planilha no Excel.")
                caminho_planilha = planilhas[0]

            raiz = ET.fromstring(arquivo_zip.read(caminho_planilha))
            linhas = []
            namespace = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
            sheet_data = raiz.find(f"{namespace}sheetData")

            if sheet_data is None:
                return []

            for linha_xml in sheet_data:
                valores = {}
                for celula in linha_xml:
                    referencia = celula.attrib.get("r", "")
                    indice = self.coluna_para_numero(referencia)
                    tipo = celula.attrib.get("t")
                    valor_xml = celula.find(f"{namespace}v")
                    valor = valor_xml.text if valor_xml is not None else ""

                    if tipo == "s":
                        try:
                            valor = strings[int(valor)]
                        except IndexError:
                            valor = ""
                    elif tipo == "inlineStr":
                        elemento = celula.find(f"{namespace}is/{namespace}t")
                        if elemento is not None:
                            valor = elemento.text or ""

                    valores[indice] = valor

                if valores:
                    maior_coluna = max(valores.keys())
                    linha = [valores.get(i, "") for i in range(maior_coluna + 1)]
                    linhas.append(linha)

        return linhas


# ============================================================
# GERENCIADOR DE USUÁRIOS E ESTATÍSTICAS
# ============================================================

class UsuarioManager:
    def __init__(self, arquivo):
        self.arquivo = arquivo

    def carregar_usuarios(self):
        leitor = LeitorExcel(self.arquivo)
        linhas = leitor.ler_planilha()
        if not linhas:
            return []

        cabecalho = [str(valor).strip().lower() for valor in linhas[0]]
        usuarios = []

        for linha in linhas[1:]:
            dados = {}
            for indice, nome_coluna in enumerate(cabecalho):
                if indice < len(linha):
                    dados[nome_coluna] = str(linha[indice]).strip()
            if dados:
                usuarios.append(dados)

        return usuarios

    def autenticar(self, login, senha):
        usuarios = self.carregar_usuarios()
        for usuario in usuarios:
            login_excel = usuario.get("login") or usuario.get("usuario") or ""
            senha_excel = usuario.get("senha") or ""
            status = (usuario.get("status") or "ativo").lower()

            if login_excel.lower() == login.lower() and senha_excel == senha:
                if status not in ("ativo", "active", "sim", "1"):
                    return {"sucesso": False, "mensagem": "Usuário desativado."}
                return {"sucesso": True, "usuario": usuario}

        return {"sucesso": False, "mensagem": "Login ou senha incorretos."}


class Estatisticas:
    def __init__(self, arquivo):
        self.arquivo = arquivo
        self.dados = {}
        self.carregar()

    def carregar(self):
        if not os.path.exists(self.arquivo):
            self.dados = {}
            return
        try:
            with open(self.arquivo, "r", encoding="utf-8") as file:
                self.dados = json.load(file)
        except Exception:
            self.dados = {}

    def salvar(self):
        with open(self.arquivo, "w", encoding="utf-8") as file:
            json.dump(self.dados, file, indent=4, ensure_ascii=False)

    def garantir_usuario(self, login):
        if login not in self.dados:
            self.dados[login] = {"vitorias": 0, "derrotas": 0, "partidas": 0}
            self.salvar()

    def registrar_vitoria(self, login):
        self.garantir_usuario(login)
        self.dados[login]["vitorias"] += 1
        self.dados[login]["partidas"] += 1
        self.salvar()

    def registrar_derrota(self, login):
        self.garantir_usuario(login)
        self.dados[login]["derrotas"] += 1
        self.dados[login]["partidas"] += 1
        self.salvar()

    def pegar(self, login):
        self.garantir_usuario(login)
        return self.dados[login]


# ============================================================
# ENGINE DE REGRAS DE TRUCO
# ============================================================

class Carta:
    def __init__(self, valor, naipe):
        self.valor = valor
        self.naipe = naipe

    def nome(self):
        return f"{self.valor}_{self.naipe}"


class Baralho:
    VALORES = ["4", "5", "6", "7", "q", "j", "k", "a", "2", "3"]
    NAIPES = ["ouros", "copas", "espadas", "paus"]

    def __init__(self):
        self.cartas = [Carta(v, n) for v in self.VALORES for n in self.NAIPES]

    def embaralhar(self):
        random.shuffle(self.cartas)

    def distribuir(self):
        self.embaralhar()
        return self.cartas[:3], self.cartas[3:6], self.cartas[6]


VALOR_BASE = {"4": 1, "5": 2, "6": 3, "7": 4, "q": 5, "j": 6, "k": 7, "a": 8, "2": 9, "3": 10}
ORDEM_NAIPE = {"paus": 4, "copas": 3, "espadas": 2, "ouros": 1}


def valor_manilha(carta, vira):
    valores = list(VALOR_BASE.keys())
    indice = valores.index(vira.valor)
    proximo = valores[(indice + 1) % len(valores)]

    if carta.valor == proximo:
        return 100 + ORDEM_NAIPE[carta.naipe]
    return VALOR_BASE[carta.valor]


def comparar_cartas(carta1, carta2, vira):
    v1, v2 = valor_manilha(carta1, vira), valor_manilha(carta2, vira)
    return 1 if v1 > v2 else (-1 if v1 < v2 else 0)


class JogoTruco:
    def __init__(self, jogador):
        self.jogador = jogador
        self.placar_jogador = 0
        self.placar_bot = 0
        self.valor_mao = 1
        self.jogador_mao = []
        self.bot_mao = []
        self.vira = None
        self.cartas_mesa = []
        self.quedas_jogador = 0
        self.quedas_bot = 0
        self.vez = "jogador"
        self.jogo_encerrado = False

    def nova_mao(self):
        baralho = Baralho()
        self.jogador_mao, self.bot_mao, self.vira = baralho.distribuir()
        self.cartas_mesa = []
        self.quedas_jogador = 0
        self.quedas_bot = 0
        self.valor_mao = 1
        self.vez = random.choice(["jogador", "bot"])

    def jogar_carta_jogador(self, indice):
        if self.vez != "jogador" or indice >= len(self.jogador_mao):
            return None
        carta = self.jogador_mao.pop(indice)
        self.cartas_mesa.append(("jogador", carta))
        self.vez = "bot"
        return carta

    def jogar_carta_bot(self):
        if not self.bot_mao:
            return None
        cartas_ordenadas = sorted(self.bot_mao, key=lambda c: valor_manilha(c, self.vira))
        carta = cartas_ordenadas[-1] if random.random() < 0.65 else random.choice(self.bot_mao)
        self.bot_mao.remove(carta)
        self.cartas_mesa.append(("bot", carta))
        self.vez = "jogador"
        return carta

    def resolver_rodada(self):
        if len(self.cartas_mesa) < 2:
            return None

        j_carta = next((c for d, c in self.cartas_mesa if d == "jogador"), None)
        b_carta = next((c for d, c in self.cartas_mesa if d == "bot"), None)

        res = comparar_cartas(j_carta, b_carta, self.vira)
        if res > 0:
            self.quedas_jogador += 1
            vencedor = "jogador"
        elif res < 0:
            self.quedas_bot += 1
            vencedor = "bot"
        else:
            vencedor = "empate"

        self.cartas_mesa = []

        if self.quedas_jogador >= 2:
            self.placar_jogador += self.valor_mao
            return "jogador_mao"
        if self.quedas_bot >= 2:
            self.placar_bot += self.valor_mao
            return "bot_mao"

        return vencedor

    def evoluir_truco(self):
        proximos = {1: 3, 3: 6, 6: 9, 9: 12}
        if self.valor_mao in proximos:
            self.valor_mao = proximos[self.valor_mao]
            return True
        return False

    def terminou(self):
        return self.placar_jogador >= 12 or self.placar_bot >= 12


# ============================================================
# COMPONENTES VISUAIS REUTILIZÁVEIS
# ============================================================

class ModernButton(tk.Button):
    def __init__(self, parent, text, command, bg=COLOR_ACCENT, fg=COLOR_TEXT, state="normal", **kwargs):
        super().__init__(
            parent,
            text=text,
            command=command,
            bg=bg,
            fg=fg if state == "normal" else COLOR_TEXT_MUTED,
            activebackground=COLOR_CANVAS,
            activeforeground=COLOR_GOLD,
            font=FONT_BOLD,
            relief="flat",
            bd=0,
            padx=12,
            pady=8,
            state=state,
            cursor="hand2" if state == "normal" else "arrow",
            **kwargs
        )


# ============================================================
# INTERFACE GRÁFICA PRINCIPAL
# ============================================================

class InterfaceGrafica:
    def __init__(self, root):
        self.root = root
        self.root.title("Truco Paulista - Edição Moderna")
        self.root.geometry("1100x750")
        self.root.minsize(950, 650)
        self.root.configure(bg=COLOR_BG)

        self.usuario_manager = UsuarioManager(ARQUIVO_USUARIOS)
        self.estatisticas = Estatisticas(ARQUIVO_ESTATISTICAS)
        self.usuario_atual = None
        self.jogo = None
        self.imagens = {}

        self.mostrar_login()

    def limpar_tela(self):
        for widget in self.root.winfo_children():
            widget.destroy()

    # --- TELA DE LOGIN ---
    def mostrar_login(self):
        self.limpar_tela()

        card = tk.Frame(self.root, bg=COLOR_CANVAS, padx=40, pady=40)
        card.place(relx=0.5, rely=0.5, anchor="center")

        tk.Label(card, text="♠ ♥ ♣ ♦", bg=COLOR_CANVAS, fg=COLOR_GOLD, font=("Segoe UI", 28)).pack()
        tk.Label(card, text="TRUCO", bg=COLOR_CANVAS, fg=COLOR_TEXT, font=FONT_TITLE).pack(pady=(0, 20))

        tk.Label(card, text="USUÁRIO", bg=COLOR_CANVAS, fg=COLOR_TEXT_MUTED, font=FONT_BOLD).pack(anchor="w")
        self.entry_login = tk.Entry(card, width=28, bg=COLOR_BG, fg=COLOR_TEXT, insertbackground=COLOR_TEXT, relief="flat", font=FONT_PRIMARY)
        self.entry_login.pack(ipady=8, pady=(4, 15))

        tk.Label(card, text="SENHA", bg=COLOR_CANVAS, fg=COLOR_TEXT_MUTED, font=FONT_BOLD).pack(anchor="w")
        self.entry_senha = tk.Entry(card, width=28, show="•", bg=COLOR_BG, fg=COLOR_TEXT, insertbackground=COLOR_TEXT, relief="flat", font=FONT_PRIMARY)
        self.entry_senha.pack(ipady=8, pady=(4, 25))

        btn = ModernButton(card, text="INICIAR PARTIDA", command=self.realizar_login, width=24)
        btn.pack()

        self.entry_login.bind("<Return>", lambda e: self.realizar_login())
        self.entry_senha.bind("<Return>", lambda e: self.realizar_login())
        self.entry_login.focus()

    def realizar_login(self):
        login = self.entry_login.get().strip()
        senha = self.entry_senha.get()

        if not login or not senha:
            messagebox.showwarning("Atenção", "Preencha usuário e senha.")
            return

        try:
            resultado = self.usuario_manager.autenticar(login, senha)
        except Exception as erro:
            messagebox.showerror("Erro", f"Erro na leitura dos dados.\n{erro}")
            return

        if not resultado["sucesso"]:
            messagebox.showerror("Login", resultado["mensagem"])
            return

        self.usuario_atual = resultado["usuario"]
        login_real = self.usuario_atual.get("login") or login
        self.estatisticas.garantir_usuario(login_real)
        self.iniciar_jogo()

    # --- INICIALIZAÇÃO DA PARTIDA ---
    def iniciar_jogo(self):
        self.jogo = JogoTruco(self.usuario_atual)
        self.jogo.nova_mao()
        self.mostrar_jogo()

    # --- TELA DE JOGO PRINCIPAL ---
    def mostrar_jogo(self):
        self.limpar_tela()

        # BARRA DE STATUS / CABEÇALHO
        top_bar = tk.Frame(self.root, bg=COLOR_CANVAS, height=60, padx=20)
        top_bar.pack(fill="x")

        nome = self.usuario_atual.get("nome") or self.usuario_atual.get("login") or "Jogador"
        dados = self.estatisticas.pegar(self.usuario_atual.get("login", ""))

        tk.Label(top_bar, text=f"👤 {nome}", bg=COLOR_CANVAS, fg=COLOR_TEXT, font=FONT_BOLD).pack(side="left")
        tk.Label(top_bar, text=f"W: {dados['vitorias']}  L: {dados['derrotas']}", bg=COLOR_CANVAS, fg=COLOR_TEXT_MUTED, font=FONT_PRIMARY).pack(side="left", padx=15)

        self.lbl_placar = tk.Label(top_bar, text="", bg=COLOR_CANVAS, fg=COLOR_GOLD, font=("Segoe UI", 16, "bold"))
        self.lbl_placar.pack(side="right")

        # ÁREA DE JOGO
        self.mesa = tk.Frame(self.root, bg=COLOR_FELT)
        self.mesa.pack(fill="both", expand=True, padx=20, pady=15)

        # MÃO DO BOT
        self.frame_bot = tk.Frame(self.mesa, bg=COLOR_FELT)
        self.frame_bot.pack(pady=10)

        # MESA CENTRAL
        self.center_area = tk.Frame(self.mesa, bg=COLOR_FELT)
        self.center_area.pack(expand=True)

        self.frame_vira = tk.Frame(self.center_area, bg=COLOR_FELT)
        self.frame_vira.pack(side="left", padx=30)

        self.frame_mesa = tk.Frame(self.center_area, bg=COLOR_FELT_LIGHT, padx=20, pady=15)
        self.frame_mesa.pack(side="right", padx=30)

        # STATUS DO JOGO
        self.lbl_status = tk.Label(self.mesa, text="Sua vez de jogar!", bg=COLOR_FELT, fg=COLOR_GOLD, font=FONT_BOLD)
        self.lbl_status.pack(pady=5)

        # MÃO DO JOGADOR
        self.frame_jogador = tk.Frame(self.mesa, bg=COLOR_FELT)
        self.frame_jogador.pack(pady=10)

        # CONTROLES & AÇÕES
        self.frame_acoes = tk.Frame(self.root, bg=COLOR_CANVAS, padx=15, pady=12)
        self.frame_acoes.pack(fill="x", side="bottom")

        self.atualizar_tela()

        if self.jogo.vez == "bot":
            self.lbl_status.config(text="O computador está pensando...")
            self.root.after(1000, self.bot_jogar)

    # --- RENDERIZAÇÃO E AÇÕES ---
    def atualizar_tela(self):
        self.desenhar_cartas_bot()
        self.desenhar_vira()
        self.desenhar_mesa()
        self.desenhar_cartas_jogador()
        self.criar_acoes()
        self.lbl_placar.config(text=f"VOCÊ {self.jogo.placar_jogador}  x  {self.jogo.placar_bot} BOT")

    def desenhar_cartas_bot(self):
        for w in self.frame_bot.winfo_children():
            w.destroy()
        for _ in self.jogo.bot_mao:
            lbl = tk.Label(self.frame_bot, text="🂠", bg=COLOR_CANVAS, fg=COLOR_TEXT_MUTED, font=("Segoe UI", 24), width=4, height=2, relief="groove")
            lbl.pack(side="left", padx=4)

    def desenhar_vira(self):
        for w in self.frame_vira.winfo_children():
            w.destroy()
        tk.Label(self.frame_vira, text="VIRA", bg=COLOR_FELT, fg=COLOR_GOLD, font=FONT_BOLD).pack()
        visual = self.criar_visual_carta(self.frame_vira, self.jogo.vira)
        visual.pack()

    def desenhar_mesa(self):
        for w in self.frame_mesa.winfo_children():
            w.destroy()

        if not self.jogo.cartas_mesa:
            tk.Label(self.frame_mesa, text="Aguardando jogada...", bg=COLOR_FELT_LIGHT, fg=COLOR_TEXT_MUTED, font=FONT_PRIMARY).pack(padx=20, pady=10)
            return

        for dono, carta in self.jogo.cartas_mesa:
            box = tk.Frame(self.frame_mesa, bg=COLOR_FELT_LIGHT)
            box.pack(side="left", padx=10)
            tk.Label(box, text="Você" if dono == "jogador" else "Bot", bg=COLOR_FELT_LIGHT, fg=COLOR_TEXT_MUTED, font=FONT_PRIMARY).pack()
            visual = self.criar_visual_carta(box, carta)
            visual.pack()

    def desenhar_cartas_jogador(self):
        for w in self.frame_jogador.winfo_children():
            w.destroy()

        for idx, carta in enumerate(self.jogo.jogador_mao):
            btn = self.criar_visual_carta(self.frame_jogador, carta, comando=lambda i=idx: self.jogar_carta(i))
            btn.pack(side="left", padx=8)

    def criar_visual_carta(self, parent, carta, comando=None):
        caminho_png = os.path.join(PASTA_CARTAS, f"{carta.nome()}.png")
        caminho_gif = os.path.join(PASTA_CARTAS, f"{carta.nome()}.gif")

        imagem = None
        try:
            if os.path.exists(caminho_png):
                imagem = tk.PhotoImage(file=caminho_png)
            elif os.path.exists(caminho_gif):
                imagem = tk.PhotoImage(file=caminho_gif)
        except Exception:
            imagem = None

        if imagem:
            self.imagens[carta.nome()] = imagem
            if comando:
                return tk.Button(parent, image=imagem, command=comando, relief="flat", bg=COLOR_FELT, activebackground=COLOR_FELT, cursor="hand2")
            return tk.Label(parent, image=imagem, bg=COLOR_FELT)

        simbolos = {"ouros": "♦", "copas": "♥", "espadas": "♠", "paus": "♣"}
        cor = COLOR_DANGER if carta.naipe in ("ouros", "copas") else "#111111"
        texto_carta = f"{carta.valor.upper()}\n{simbolos[carta.naipe]}"

        if comando:
            return tk.Button(parent, text=texto_carta, command=comando, bg="#FFFFFF", fg=cor, font=FONT_CARDS, relief="flat", width=5, height=3, cursor="hand2")
        return tk.Label(parent, text=texto_carta, bg="#FFFFFF", fg=cor, font=FONT_CARDS, relief="flat", width=5, height=3)

    def criar_acoes(self):
        for w in self.frame_acoes.winfo_children():
            w.destroy()

        tk.Label(
            self.frame_acoes, 
            text=f"VALOR DA MÃO: {self.jogo.valor_mao} PTS", 
            bg=COLOR_CANVAS, 
            fg=COLOR_GOLD, 
            font=FONT_BOLD
        ).pack(side="left", padx=10)

        proximos_nomes = {1: "TRUCO! (3)", 3: "PEDIR 6!", 6: "PEDIR 9!", 9: "PEDIR 12!"}
        
        if self.jogo.valor_mao in proximos_nomes and not self.jogo.jogo_encerrado:
            pode_trucar = (self.jogo.vez == "jogador")
            
            btn_truco = ModernButton(
                self.frame_acoes, 
                text=f"🔥 {proximos_nomes[self.jogo.valor_mao]}", 
                command=self.trucar, 
                bg=COLOR_ACCENT if pode_trucar else COLOR_CANVAS,
                state="normal" if pode_trucar else "disabled"
            )
            btn_truco.pack(side="left", padx=5)

        btn_sair = ModernButton(
            self.frame_acoes, 
            text="DESISTIR DA MÃO", 
            command=self.confirmar_saida, 
            bg=COLOR_DANGER
        )
        btn_sair.pack(side="right", padx=5)

    # --- FLUXO DA RODADA ---
    def jogar_carta(self, indice):
        if self.jogo.vez != "jogador" or self.jogo.jogo_encerrado:
            return

        carta = self.jogo.jogar_carta_jogador(indice)
        if not carta:
            return

        self.atualizar_tela()

        if len(self.jogo.cartas_mesa) == 2:
            self.lbl_status.config(text="Avaliando rodada...")
            self.root.after(500, self.processar_fim_rodada)
        else:
            self.lbl_status.config(text="Computador está jogando...")
            self.root.after(600, self.bot_jogar)

    def bot_jogar(self):
        if self.jogo.jogo_encerrado or self.jogo.vez != "bot":
            return

        if not self.jogo.bot_mao and len(self.jogo.cartas_mesa) < 2:
            self.processar_fim_rodada()
            return

        self.jogo.jogar_carta_bot()
        self.atualizar_tela()

        if len(self.jogo.cartas_mesa) == 2:
            self.lbl_status.config(text="Avaliando rodada...")
            self.root.after(500, self.processar_fim_rodada)
        else:
            self.lbl_status.config(text="Sua vez!")

    def processar_fim_rodada(self):
        if self.jogo.jogo_encerrado:
            return

        resultado = self.jogo.resolver_rodada()
        self.atualizar_tela()

        if resultado == "jogador":
            self.lbl_status.config(text="Você levou a queda! Sua vez.")
            self.jogo.vez = "jogador"
            self.criar_acoes()
        elif resultado == "bot":
            self.lbl_status.config(text="O Computador levou a queda!")
            self.jogo.vez = "bot"
            self.criar_acoes()
            self.root.after(800, self.bot_jogar)
        elif resultado == "empate":
            self.lbl_status.config(text="Queda empatada!")
            self.criar_acoes()
        elif resultado in ("jogador_mao", "bot_mao"):
            vencedor_str = "Você" if resultado == "jogador_mao" else "Computador"
            messagebox.showinfo("Fim da Mão", f"{vencedor_str} venceu a mão e levou {self.jogo.valor_mao} ponto(s)!")

            if self.jogo.terminou():
                self.finalizar_partida("jogador" if self.jogo.placar_jogador >= 12 else "bot")
            else:
                self.jogo.nova_mao()
                self.mostrar_jogo()

    def trucar(self):
        if self.jogo.vez != "jogador":
            return

        val_antigo = self.jogo.valor_mao
        if not self.jogo.evoluir_truco():
            return

        if random.random() < 0.25:
            messagebox.showinfo("Truco!", "O Computador CORREU! Você ganha os pontos.")
            self.jogo.placar_jogador += val_antigo
            if self.jogo.terminou():
                self.finalizar_partida("jogador")
            else:
                self.jogo.nova_mao()
                self.mostrar_jogo()
        else:
            messagebox.showinfo("Truco!", f"O Computador ACEITOU!\nA mão agora vale {self.jogo.valor_mao} pontos.")
            self.atualizar_tela()

    def finalizar_partida(self, vencedor):
        self.jogo.jogo_encerrado = True
        login = self.usuario_atual.get("login", "")

        if vencedor == "jogador":
            self.estatisticas.registrar_vitoria(login)
            msg = "🏆 PARABÉNS! Você venceu a partida!"
        else:
            self.estatisticas.registrar_derrota(login)
            msg = "🤖 FIM DE JOGO! O computador venceu."

        if messagebox.askyesno("Fim de Jogo", f"{msg}\n\nDeseja jogar novamente?"):
            self.iniciar_jogo()
        else:
            self.mostrar_login()

    def confirmar_saida(self):
        if messagebox.askyesno("Desistir", "Deseja abandonar esta mão? O ponto irá para o adversário."):
            self.jogo.placar_bot += self.jogo.valor_mao
            if self.jogo.terminou():
                self.finalizar_partida("bot")
            else:
                self.jogo.nova_mao()
                self.mostrar_jogo()


# ============================================================
# INICIALIZAÇÃO DA APLICAÇÃO
# ============================================================

if __name__ == "__main__":
    if not os.path.exists(PASTA_CARTAS):
        os.makedirs(PASTA_CARTAS)

    root = tk.Tk()
    app = InterfaceGrafica(root)
    root.mainloop()