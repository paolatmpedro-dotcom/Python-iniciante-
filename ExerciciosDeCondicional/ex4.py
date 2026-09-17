numero1 = float(input("Digite um numero1: "))

numero2 = float(input("\nDigite um numero2: "))

print("\nOpções")

print("1 - soma")
print("2 - subtracao")
print("3 - multiplicação")
print("4 - divisao")

opcao = float(input("\nEscolha uma opção: "))

if opcao == 1 :
    resultado = numero1 + numero2
elif opcao == 2 :
    resultado = numero1 - numero2
elif opcao == 3 :
    resultado = numero1 * numero2
elif opcao == 4 :
    resultado = numero1 / numero2
else:
    resultado = "opção inválida"

print(f"\nResultado: {resultado}\n")
