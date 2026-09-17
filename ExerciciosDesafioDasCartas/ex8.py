def calcular_area_retangulo(base, altura):
    area = base * altura
    return area


base = int(input("Digite a base do retângulo: "))
altura = int(input("Digite a altura do retângulo: "))

resultado = calcular_area_retangulo(base, altura)

print("A área total do retângulo é:", resultado)