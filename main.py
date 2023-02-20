import openpyxl as pyxl
import datetime

'''
1. Criar função de contagem regressiva de dias;
2. Criar laço para inputar os dados de controle dos dias;
3. Criar função para cálculo da qtd de insulina;
4. Inserir as linhas na planilha.
'''

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

def preeche_planilha():
    i = 6
    planilha = pyxl.load_workbook('Glicemia.xlsx')['Planilha1']
    for x, dia in enumerate(calcula_dias()):
        if x <= 14:
            print(f"==={dia}===")
            glicemia_jejum = int(input("Glicemia em jejum: "))
            glicemia_apos_almoco = int(input("Glicemia em 2H após o Almoço: "))
            glicemia_antes_de_dormir = int(input("Glicemia antes de Dormir: "))
            planilha[f'A{i}'] = dia
            planilha[f'C{i}'] = glicemia_jejum 
            planilha[f'E{i}'] = calcula_glicemia(glicemia_jejum)
            planilha[f'G{i}'] = glicemia_apos_almoco 
            planilha[f'H{i}'] = calcula_glicemia(glicemia_apos_almoco)
            planilha[f'M{i}'] = glicemia_antes_de_dormir 
            planilha[f'N{i}'] = calcula_glicemia(glicemia_antes_de_dormir)
            i += 2
            planilha.parent.save(f'Primeiros_15dias.xlsx')
            continue
        i = 6 if x == 15 else i
        if x >= 15:
            print(f"==={dia}===")
            glicemia_jejum = int(input("Glicemia em jejum: "))
            glicemia_apos_almoco = int(input("Glicemia em 2H após o Almoço: "))
            glicemia_antes_de_dormir = int(input("Glicemia antes de Dormir: "))
            planilha[f'A{i}'] = dia
            planilha[f'C{i}'] = glicemia_jejum 
            planilha[f'E{i}'] = calcula_glicemia(glicemia_jejum)
            planilha[f'G{i}'] = glicemia_apos_almoco 
            planilha[f'H{i}'] = calcula_glicemia(glicemia_apos_almoco)
            planilha[f'M{i}'] = glicemia_antes_de_dormir 
            planilha[f'N{i}'] = calcula_glicemia(glicemia_antes_de_dormir)
            i += 2
            planilha.parent.save(f'Últimos_15dias.xlsx')

if __name__ == "__main__":
    preeche_planilha()