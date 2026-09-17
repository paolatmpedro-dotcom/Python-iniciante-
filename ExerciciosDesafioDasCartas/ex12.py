def celsius_para_fahrenheit(celsius):
    return (celsius * 9 / 5) + 32


def fahrenheit_para_celsius(fahrenheit):
    return (fahrenheit - 32) * 5 / 9


print("===== CONVERSOR DE TEMPERATURA =====")
print("1 - Celsius para Fahrenheit")
print("2 - Fahrenheit para Celsius")

opcao = input("Escolha uma opção: ")

try:
    temperatura = float(input("Digite a temperatura: "))

    if opcao == "1":
        resultado = celsius_para_fahrenheit(temperatura)
        print("Temperatura em Fahrenheit:", resultado)

    elif opcao == "2":
        resultado = fahrenheit_para_celsius(temperatura)
        print("Temperatura em Celsius:", resultado)

    else:
        print("Opção inválida.")

except ValueError:
    print("Digite uma temperatura válida.")