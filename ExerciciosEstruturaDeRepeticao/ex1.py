senha_correta = "4321"
senha_digitada = ""
contador = 0


while senha_digitada != senha_correta and contador < 3:
        senha_digitada = input("Digite a senha:  ")
        contador +=1


if senha_digitada == senha_correta:
    print("Acesso permitido, você usou", contador, "tentativas")   
else:
    print("Senha Bloqueada, você usou todas as tentativas")

