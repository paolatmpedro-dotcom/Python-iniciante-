idade = float(input("Digite sua idade: "))

rendaMensal = float(input("Digite sua renda mensal: "))

if idade >= 18 and rendaMensal <= 2500.00:
    print("Pode solicitar o benefício.")
else:
    print("Não atende aos critérios")

