nome = input("Digite seu nome: ")
idade = float(input("Digite sua idade: "))

nomeResponsavel = ""

if idade >= 18:
    acesso = "Acesso liberado"

elif idade < 18:
    nomeResponsavel = input("Digite o nome do responsável: ").strip()

    if nomeResponsavel:
        acesso = "Acesso liberado com responsável"
    else:
        acesso = "Acesso negado: responsável não informado"


print("\n--- Resultado ---")
print(f"Nome: {nome}")
print(f"Idade: {idade} anos")

if nomeResponsavel:
    print("Possui responsável: Sim")
    print(f"Responsável: {nomeResponsavel}")
else:
    print("Possui responsável: Não")

print(f"Situação: {acesso}")
