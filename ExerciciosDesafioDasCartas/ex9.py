def maior_de_tres(numero1, numero2, numero3):
    if numero1 > numero2 and numero1 > numero3:
        return numero1
    elif numero2 > numero1 and numero2 > numero3:
        return numero2
    else:
        return numero3


numero1 = int(input("Digite o primeiro número: "))
numero2 = int(input("Digite o segundo número: "))
numero3 = int(input("Digite o terceiro número: "))

vencedor = maior_de_tres(numero1, numero2, numero3)

print("O maior número é:", vencedor)