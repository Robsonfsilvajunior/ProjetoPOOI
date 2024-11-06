import pandas as pd
import os

output_dir = "C:/Users/Ryan/Desktop/ProjetoPOOI"

# Função para validar entrada de dados
def matricula_valid():
    while True:
        matricula = input("Digite a matrícula do aluno (máximo 6 dígitos): ")
        if matricula.isdigit() and len(matricula) == 6:
            return matricula
        print("Matrícula inválida. Tente novamente.")

def nome_valid():
    while True:
        nome = input("Digite o nome do aluno: ")
        if nome.isalpha():
            return nome
        print("Nome inválido. Use apenas letras.")
def nota_valid(disciplina):
    while True:
        try:
            nota = float(input(f"Digite a nota de {disciplina} (0 a 10): "))
            if nota < 0 or nota > 10:
                print("Nota fora do intervalo permitido. Tente novamente.")
                nota = float(input(f"Digite a nota de {disciplina} (0 a 10): "))
            else:
                return nota
        except ValueError:
            print("Entrada inválida. Digite um número.")


# pega os dados e valida e adiciona ao dicionário
alunos = []
while True:
    matricula = matricula_valid()
    nome = nome_valid()
    nota_pvb = nota_valid("PVB")
    nota_paw = nota_valid("PAW")
    nota_bd = nota_valid("BD")
    nota_pooi = nota_valid("POOI")
    
    #média 
    media = ((nota_pvb + nota_paw + nota_bd + nota_pooi) / 4)
    
    #situação do aluno
    if media >= 6:
        situacao = "APROVADO"
    elif media < 6 and media >= 3.75:
        situacao = "EXAME FINAL"
    else:
        situacao = "RETIDO"
    
    # Adiciona os dados ao dicionário
    alunos.append({
        "Matricula": matricula,
        "Aluno": nome,
        "PVB": nota_pvb,
        "PAW": nota_paw,
        "BD": nota_bd,
        "POOI": nota_pooi,
        "Média": media,
        "Situação": situacao
    })
    
    
    if input("Deseja adicionar outro aluno? (s/n): ").lower() != 's':
        break


df = pd.DataFrame(alunos)
df.to_excel("notasfinaisalunos.xlsx", index=False)
print(f"Arquivo Excel gerado em: {"notasfinaisalunos.xlsx"}")

def gerar_html(aluno):
    matricula = aluno["Matricula"]
    nome = aluno["Aluno"]
    pvb = aluno["PVB"]
    paw = aluno["PAW"]
    bd = aluno["BD"]
    pooi = aluno["POOI"]
    media = aluno["Média"]
    situacao = aluno["Situação"]
    
    #cor de fundo
    if situacao == "APROVADO":
        cor_fundo = "rgb(0, 104, 139)"
    elif situacao == "EXAME FINAL":
        cor_fundo = "rgb(255, 215, 0)"
    else:
        cor_fundo = "red"
    
    #html
    html_content = f"""
        <html>
        <head><title>Notas do Aluno {nome}</title></head>
        <body style="text-align: center;">
            <tr><p>MATRICULA: {matricula}</p></tr>
            
            <!-- Nome do aluno com fundo colorido -->
            <tr><p>ALUNO: <b style="background-color: red; color: black; padding: 5px; border-radius: 5px;">{nome}</b></p></tr>
            
            <table border="0.2" style="width: 50%; margin: 0 auto; background-color: {cor_fundo}; color: white;">
                <tr><th>PVB</th><th>PAW</th><th>BD</th><th>POOI</th><th>MEDIA</th></tr>
                <tr><th>{pvb}</th><th>{paw}</th><th>{bd}</th><th>{pooi}</th><th>{media}</th></tr>
            </table>

            <!-- Situacao com fundo colorido -->
            <tr><p>Situacao Final: <b style="background-color: {cor_fundo}; color: black; padding: 5px; border-radius: 5px;">{situacao}</b></p></tr>
        </body>
        </html>
        """

    
    #salva html
    html_path = os.path.join(output_dir, f"{matricula}.html")
    with open(html_path, "w") as file:
        file.write(html_content)
    print(f"Arquivo HTML gerado do aluno {nome} em: {html_path}")

if input("Deseja gerar arquivos HTML para cada aluno? (s/n): ").lower() == 's':
    for aluno in alunos:
            gerar_html(aluno)
