nomeUsuario = input("Digite o nome de usuário: ")
senha = input("Digite a senha: ") 
situacaoLogin = input("Você está ativo (sim/não): ")

if (nomeUsuario == "admin" and senha == "1234") and situacaoLogin == "sim":
    print("Acesso autorizado")
else:
    print("Acesso negado") 

