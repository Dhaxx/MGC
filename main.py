import openpyxl as pyxl
import datetime
from random import randint
import subprocess

'''
1. Criar função de contagem regressiva de dias;
2. Criar laço para inputar os dados de controle dos dias;
3. Criar função para cálculo da qtd de insulina;
4. Inserir as linhas na planilha.
'''

def main():
    i = 6
    planilha = pyxl.load_workbook('Glicemia.xlsx')
    primeira_tabela = planilha['Planilha1']
    segunda_tabela = planilha['Planilha2']

    for x, dia in enumerate(calcula_dias()):
        if x <= 14:
            print(f"==={dia}===")
            preenche_valor(primeira_tabela, dia, i)
            i += 2
            continue

        i = 6 if x == 15 else i
        if x >= 15:
            print(f"==={dia}===")
            preenche_valor(segunda_tabela, dia, i)
            i += 2

    planilha.save('PLANILHA_PREENCHIDA.xlsx')
    subprocess.Popen(['start', 'PLANILHA_PREENCHIDA.xlsx'], shell=True)

def calcula_dias():
    data_atual = datetime.datetime.now()
    data_formatada = data_atual.strftime('%d/%m/%Y')
    data_final = data_atual - datetime.timedelta(days=29)
    dias = []

    while data_atual >= data_final:
        data_formatada = data_atual.strftime('%d/%m/%Y')
        data_atual -= datetime.timedelta(days=1)
        dias.append(data_formatada)
    return dias

def calcula_glicemia(x):
    medicao = x
    
    if medicao <= 100:
        return "0 UI"
    if medicao <= 140:
        return "2 UI"
    if medicao <= 160:
        return "4 UI"
    if medicao <= 180:
        return "6 UI"
    if medicao <= 200:
        return "8 UI"
    if medicao <= 260:
        return "10 UI"
    if medicao <= 280:
        return "12 UI"
    if medicao > 200:
        return "14 UI"

def preenche_valor(planilha, dia, i):
    glicemia_jejum = ''
    glicemia_apos_almoco = ''
    glicemia_antes_de_dormir = ''

    while type(glicemia_jejum) != int:
        glicemia_jejum = input("Glicemia em jejum: ")
        if glicemia_jejum.isdigit():
            glicemia_jejum = int(glicemia_jejum)
        elif glicemia_jejum == '':
            glicemia_jejum = randint(60, 450)
        else: 
            print("VALOR INVÁLIDO")
            continue    
    while type(glicemia_apos_almoco) != int:
        glicemia_apos_almoco = input("Glicemia em 2H após o Almoço: ")
        if glicemia_apos_almoco.isdigit():
            glicemia_apos_almoco = int(glicemia_apos_almoco)
        elif glicemia_apos_almoco == '':
            glicemia_apos_almoco = randint(60, 450)
        else: 
            print("VALOR INVÁLIDO")
            continue    
    while type(glicemia_antes_de_dormir) != int:
        glicemia_antes_de_dormir = input("Glicemia antes de Dormir: ")
        if glicemia_antes_de_dormir.isdigit():
            glicemia_antes_de_dormir = int(glicemia_antes_de_dormir)
        elif glicemia_antes_de_dormir == '':
            glicemia_antes_de_dormir = randint(60, 450)
        else: 
            print("VALOR INVÁLIDO")
            continue   

    planilha[f'A{i}'] = dia
    planilha[f'C{i}'] = glicemia_jejum
    planilha[f'E{i}'] = calcula_glicemia(glicemia_jejum)
    planilha[f'G{i}'] = glicemia_apos_almoco
    planilha[f'H{i}'] = calcula_glicemia(glicemia_apos_almoco)
    planilha[f'M{i}'] = glicemia_antes_de_dormir
    planilha[f'N{i}'] = calcula_glicemia(glicemia_antes_de_dormir)

if __name__ == "__main__":
    main()