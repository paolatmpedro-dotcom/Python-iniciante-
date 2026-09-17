try:
    ano_nascimento = int(input("Digite o ano em que você nasceu: "))
except ValueError:
    print("Por favor, digite apenas números para o ano.")
else:
    print(f"O ano do seu nascimento é: {ano_nascimento}")