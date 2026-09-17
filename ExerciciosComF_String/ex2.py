idade = int(input("Digite sua idade: "))
deficiencia = input("Tem alguma deficiência (sim/não): ").lower()

if idade >= 60 or deficiencia == "sim":
    print("Atendimento prioritário")
else:
    print("Atendimento normal")
    