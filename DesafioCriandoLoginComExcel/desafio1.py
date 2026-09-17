usuario = "aaaaa"
senha = "0000"

admin = "admin"
senha_admin = "1234"

tentativas = 0
ativo = True

while ativo:

    while tentativas < 3:

        print("\n--- LOGIN ---")
        nome = input("\nUsuário: ")
        while len(usuario) == 0:
            print("Usuario não pode ficar em branco")

        senha_digitada = input("Senha: ")    
        while len(senha) == 0:    
             print("Senha não pode ficar em branco")

        if nome == usuario and senha_digitada == senha:
            print("\nAcesso autorizado!")
            ativo = False
            break

        elif nome != usuario:
            print("Usuário incorreto!")

        else:
            print("Senha incorreta!")

        tentativas += 1
        print(f"Tentativa {tentativas}/3")

    if tentativas == 3:

        print("\n--- CONTA BLOQUEADA ---")
        admin_tentativas = 0

        while admin_tentativas < 2:

            nome_admin = input("\nAdmin: ")
            senha_digitada = input("Senha: ")

            if nome_admin == admin and senha_digitada == senha_admin:
                print("\nConta desbloqueada!")
                tentativas = 0
                break

            elif nome_admin != admin:
                print("Administrador incorreto!")

            else:
                print("Senha do administrador incorreta!")

            admin_tentativas += 1
            print(f"Tentativa {admin_tentativas}/2")
            print("\n------")

        if admin_tentativas == 2:
            print("\nConta continua bloqueada!")
            ativo = False