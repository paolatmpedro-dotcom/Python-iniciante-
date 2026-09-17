nota1 = float(input("Digite a nota 1: "))
nota2 = float(input("Digite a nota 2: "))

media = (nota1 + nota2) / 2

if media > 6:
    print("Aprovado, a nota foi", media)
else:
    print("Reprovado, a nota foi", media)    