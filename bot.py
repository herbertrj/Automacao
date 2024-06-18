import datetime

import openpyxl

from Mensagem.mensageria import MSG

Tipo_mensagem = MSG()

def carregar_planilha(nome_arquivo, nome_planilha):
    """Carrega a planilha especificada do arquivo Excel."""
    try:
        workbook = openpyxl.load_workbook(nome_arquivo)
        return workbook[nome_planilha]
    except Exception as e:
        print(f"Erro ao carregar a planilha: {e}")
        return None

def formatar_data(data_vencimento_excel):
    """Formata a data de vencimento para o formato desejado e ajusta a hora atual."""
    agora = datetime.datetime.now()
    data_vencimento_excel = data_vencimento_excel.replace(hour=agora.hour, minute=agora.minute, second=agora.second)
    data_vencimento_formatada = data_vencimento_excel.strftime('%d/%m/%Y')
    return data_vencimento_excel, data_vencimento_formatada

def calcular_dias_para_vencimento(data_vencimento_excel):
    """Calcula o número de dias entre hoje e a data de vencimento."""
    agora = datetime.datetime.now()
    return (data_vencimento_excel - agora).days

def processar_clientes(pagina_clientes):
    """Processa cada cliente na planilha e envia a mensagem apropriada."""
    num_linhas = pagina_clientes.max_row
    for num_linha in range(2, num_linhas + 1):
        linha = pagina_clientes[num_linha]
        nome = linha[0].value
        valor = linha[1].value
        data_vencimento_excel = linha[2].value
        telefone = linha[3].value

        if data_vencimento_excel is not None:
            data_vencimento_excel, data_vencimento_formatada = formatar_data(data_vencimento_excel)
            calc_dias = calcular_dias_para_vencimento(data_vencimento_excel)

            if -1 <= calc_dias < 0:
                Tipo_mensagem.vencimento_atual(nome, telefone, data_vencimento_formatada, valor)
            elif calc_dias < -1:
                Tipo_mensagem.vencido(nome, abs(calc_dias), data_vencimento_formatada, valor, telefone)
            elif -1 < calc_dias < 3:
                Tipo_mensagem.vencimento_futuro(nome, data_vencimento_formatada, valor, telefone)
        else:
            break

def main():
    NOME_ARQUIVO = 'clientesNovo.xlsx'
    NOME_PLANILHA = 'Planilha1'
    
    pagina_clientes = carregar_planilha(NOME_ARQUIVO, NOME_PLANILHA)
    if pagina_clientes:
        processar_clientes(pagina_clientes)

if __name__ == "__main__":
    main()
