import datetime
import os
import webbrowser
from time import sleep
from urllib.parse import quote

import openpyxl
import pyautogui

agora = datetime.datetime.now()
hora = agora.hour
periodo_dia = "manhã" if 6 <= hora < 12 else ("tarde" if 12 <= hora < 18 else "noite")

class MSG:
    nome = None
    vencimento = None
    valor = None
    telefone = None

    # webbrowser.open('https://web.whatsapp.com/')
    # sleep(10)
    
    def vencimento_atual(self,pNome, pTelefone, pVencimento, pValor):
        self.nome = pNome
        self.telefone = pTelefone
        self.vencimento = pVencimento
        self.valor = pValor

        print('******** Hoje *********')
        msg = (f'Olá {pNome}, você tem um boleto com vencimento para hoje dia: {pVencimento}, no valor de R$ {pValor}.\n'
        f'Pague seu boleto ainda hoje para não ter o seu acesso bloqueado.\n'
        f'Se já pagou, desconsidere essa mensagem. Tenha uma Boa {periodo_dia}.\n')
        # print(msg)
        self.enviar(pNome, pTelefone, msg)

    def vencido(self, pNome, diferenca_dia,  vencimento, valor, pTelefone):
        # print('******** vencido *********')
        self.nome = pNome
        self.vencimento = vencimento
        self.valor = valor
        
        msg = (f'Olá {pNome}, você tem boleto(s) vencidos a {diferenca_dia} dias. vencido dia {vencimento}, no valor de R$ {valor}.\n'
        f'Pague seu boleto o quanto antes e evite eventuais problemas.\n'
        f'Se já pagou, desconsidere essa mensagem. Tenha uma Boa {periodo_dia}.\n')
        print(msg)
        self.enviar(pNome, pTelefone, msg)

    def vencimento_futuro(self, pNome, vencimento_format, valor, pTelefone):
        self.nome = pNome
        self.vencimento = vencimento_format
        self.valor = valor

        # print('******** Vencerá *********')
        msg = (f'Olá {pNome}, estamos enviando esta mensagem para lembrá-lo de que você tem boleto(s) que vencerá(ão) no dia {vencimento_format}, no valor de R$ {valor}.\n'
        f'Se já pagou, desconsidere esta mensagem. Tenha uma Boa {periodo_dia}.\n')
        self.enviar(pNome, pTelefone, msg)

    def enviar(self, pNome, telefone, tipo_mensagem):
        try:
            link_mensagem_whatsapp = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(tipo_mensagem)}'
            webbrowser.open(link_mensagem_whatsapp)
            sleep(12)
            # print("Tentando localizar a imagem do botão minimizar...")
            # Captura de tela para verificar o que está visível
            # pyautogui.screenshot('screenshot.png')
            seta = pyautogui.locateCenterOnScreen('seta.png')
            sleep(7)
            pyautogui.click(seta[0],seta[1])
            sleep(3)
            pyautogui.hotkey('ctrl','w')
            sleep(3)
        except:
            print(f'Não foi possível enviar mensagem para {pNome}')
            with open('erros.csv','a',newline='',encoding='utf-8') as arquivo:
                arquivo.write(f'{pNome},{telefone}{os.linesep}')
            