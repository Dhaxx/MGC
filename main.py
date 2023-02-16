import openpyxl as pyxl
import datetime

planilha = pyxl.load_workbook('Glicemia.xlsx')['Planilha1']

def preenche_planilha():
    data_atual = datetime.datetime.now()
    data_formatada = data_atual.strftime('%d/%m/%Y')
    dias_preenchimento = int(input(f"==={data_formatada}===. Qual o período que deseja informar: "))
    data_final = data_atual - datetime.timedelta(days=dias_preenchimento)

    while data_atual >= data_final:
        data_formatada = data_atual.strftime('%d/%m/%Y')
        print(f"Data atual: {data_formatada}")
        data_atual -= datetime.timedelta(days=1)


if __name__ == "__main__":
    preenche_planilha()