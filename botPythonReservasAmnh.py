# _*_ coding: utf-8 _*_

import pandas as pd
import datetime
import pyautogui
import time
import os
import re
import unidecode

hoje = datetime.date.today()
amanha = hoje + datetime.timedelta(+1)
seteDias = hoje + datetime.timedelta(+7)
oitoDias = hoje + datetime.timedelta(+8)


#Entrar no Stays e baixar as 2 planilhas ->
pyautogui.PAUSE = 3
time.sleep(4)
pyautogui.rightClick(x=1313, y=1054)
pyautogui.click(x=1302, y=934)
pyautogui.hotkey('alt', 'space')
pyautogui.press('x')
pyautogui.click(x=1114, y=52)
pyautogui.click(x=1114, y=52)
pyautogui.press("backspace")
pyautogui.write("https://xxxxxxxxxxxxxx.com.br")
pyautogui.press("enter")
time.sleep(10)
pyautogui.click(x=230, y=500)
pyautogui.write("xxx")
pyautogui.click(x=225, y=567)
pyautogui.write("xxxxx")
pyautogui.click(x=220, y=635)
time.sleep(18)
pyautogui.click(x=1118, y=336)
pyautogui.click(x=10, y=353)
pyautogui.click(x=97, y=459)
time.sleep(2)
pyautogui.click(x=1745, y=146)
pyautogui.click(x=1121, y=648)
time.sleep(36)
pyautogui.click(x=215, y=289)
pyautogui.click(x=101, y=361)
pyautogui.click(x=306, y=624)
time.sleep(2)
pyautogui.click(x=1745, y=146)
pyautogui.click(x=1121, y=648)
time.sleep(36)
pyautogui.click(x=1896, y=10)


#Lembretes Mãe ->
lembretes = [[], ["- 403/510 Energia"], [], [], [], ["- 403/614 Aluguel", "- Vivi e-social"], [], [], [], [], ["- 510 Internet"], [], [], [], ["- 807 Energia"], ["- 510/807 Alugel", "- Quinzena: Vivi e matheus"], [], ["- Pagar Booking"], [], [], ["- 807 Internet"], [], [], [], [], ["- 215/403/614 Internet"], [], [], ["- 215/614 Energia"], [], ["- 215 Aluguel", "- Quinzena: Vivi e matheus"], []]

lembrete_amanha = ''
for i in lembretes:
    lembretes_index = lembretes.index(i)
    if (amanha.day == lembretes_index):
        lembrete_amanha = lembretes[lembretes_index]


#ABRIR EXCEL E PEGAR OS DADOS DE ENTRADA E DE SAÍDA DE AMANHA ->

#Transformando o mês em string por causa do excel
meses = ["jan", "fev", "mar", "abr", "jun", "jul", "ago", "set", "out", "nov", "dez"]
mes = ''
for i in meses:
    if (amanha.month == meses.index(i)):
        mes = meses[(meses.index(i)) - 1]

#Formatando o dia para ficar mais preciso, pois o excel é: dia '02' e não dia 2
dia_amanha = 0
if (amanha.day < 10):
        dia_amanha = '0{}'.format(amanha.day)
else:
    dia_amanha = amanha.day

#Colocando os dados no formato do excel para realizar a busca (ex.: 03 abr 2023)->
chegadaEsaida = '{} {} {}'.format(dia_amanha, mes, amanha.year)

#Leitura das tabelas e separando os dados de chegadas e saidas ->
tabelachegadas = pd.read_excel("fromX_toX_propX_ownerX.xlsx")
tabelasaidas = pd.read_excel(f'{hoje}_{seteDias}_propX_ownerX.xlsx')

dados_de_chegadas = tabelachegadas.loc[tabelachegadas["Chegada"] == chegadaEsaida, "Mês":"Total da Reserva"]
dados_de_saidas = tabelasaidas.loc[tabelasaidas["Data de Saída"] == chegadaEsaida, "Mês":"Total da Reserva"]



#Dados das chegadas por apt ->
def localizar_dados_chegadas(apt):
    return dados_de_chegadas.loc[dados_de_chegadas["Nome Interno do Anúncio"] == apt, "Mês":"Total da Reserva"]

dados_de_chegadas215 = localizar_dados_chegadas("Neo2-215")
dados_de_chegadas403 = localizar_dados_chegadas("Portville-403")
dados_de_chegadas510 = localizar_dados_chegadas("Ametista6-510")
dados_de_chegadas614 = localizar_dados_chegadas("Neo1-614")
dados_de_chegadas807 = localizar_dados_chegadas("Ametista6 - 807")


#Dados das saídas por apt ->
def localizar_dados_saidas(apt):
    return dados_de_saidas.loc[dados_de_saidas["Nome Interno do Anúncio"] == apt, "Mês":"Total da Reserva"]

dados_de_saidas215 = localizar_dados_saidas("Neo2-215")
dados_de_saidas403 = localizar_dados_saidas("Portville-403")
dados_de_saidas510 = localizar_dados_saidas("Ametista6-510")
dados_de_saidas614 = localizar_dados_saidas("Neo1-614")
dados_de_saidas807 = localizar_dados_saidas("Ametista6 - 807")


#ORGAZINHANDO MENSAGENS DE CHEGADAS E DE SAÍDAS POR APT. E FORMATANDO ELAS ->
cgdmsg215 = []  #chegadas do 215... (['Anuncio: 215', 'Hospede: X', 'R$: 111', 'Canal: Airbnb'])
cgdmsg403 = []
cgdmsg510 = []
cgdmsg614 = []
cgdmsg807 = []

sdsmsg215 = [] #saidas do 215...
sdsmsg403 = []
sdsmsg510 = []
sdsmsg614 = []
sdsmsg807 = []

def gerar_array_com_infos_de_chegadas_e_saidas_formatados(relacao, apt):
    nome_hosp = unidecode.unidecode(row["Nome do Hóspede"])     #Tirando a acentuanção
    arr_palavras = nome_hosp.split()    #Separando as palavras do nome em um array para em nomes grandes ficar só 3
    nome_hosp_formatado = ''
    if len(arr_palavras) > 3:
        nome_hosp_formatado = '{} {} {}'.format(arr_palavras[0], arr_palavras[1], arr_palavras[2])
    else:
        nome_hosp_formatado = nome_hosp

    qtd_hospedes = 'Qtd. Hospd: {}'.format(row["Total de Hóspedes"])

    ta = re.compile('airbnb')   #Regexp, biblioteca re, salvando a palavra 'airbnb' de uma forma q a biblio entende
    tb = re.compile('booking.com')
    checkair = ta.findall(str(row["Canal"]))    #Chechando se o canal é Airbnb ou Booking
    checkboo = tb.findall(str(row["Canal"]))
    canal = ''
    nao_e_airbnb = False
    if (checkair):
        canal = 'Canal: Airbnb'
    elif (checkboo):
        canal = 'Canal: Booking'
    else:
        canal = 'Canal: WhatsApp'
    if not (checkair):
        nao_e_airbnb = True

    #Apartir daqui é apenas adicionando as infos geradas nos arrays de acordo com entrada e saida ->
    if (relacao == 'entrada'):  #A função é extend é apenas para adicionar mais de um valor simultaniamente no array
        if (apt == '215'):
            cgdmsg215.extend(['Anuncio: {}'.format(apt), 'Hospede: {}'.format(nome_hosp_formatado), canal, qtd_hospedes])
            if (nao_e_airbnb):
                cgdmsg215.extend(['Telefone: {}'.format(row["Número de telefone"]), 'R$: {}'.format(row["Total da Reserva"])])

        elif (apt == '403'):
            cgdmsg403.extend(['Anuncio: {}'.format(apt), 'Hospede: {}'.format(nome_hosp_formatado), canal, qtd_hospedes])
            if (nao_e_airbnb):
                cgdmsg403.extend(['Telefone: {}'.format(row["Número de telefone"]), 'R$: {}'.format(row["Total da Reserva"])])

        elif (apt == '510'):
            cgdmsg510.extend(['Anuncio: {}'.format(apt), 'Hospede: {}'.format(nome_hosp_formatado), canal, qtd_hospedes])
            if (nao_e_airbnb):
                cgdmsg510.extend(['Telefone: {}'.format(row["Número de telefone"]), 'R$: {}'.format(row["Total da Reserva"])])

        elif (apt == '614'):
            cgdmsg614.extend(['Anuncio: {}'.format(apt), 'Hospede: {}'.format(nome_hosp_formatado), canal, qtd_hospedes])
            if (nao_e_airbnb):
                cgdmsg614.extend(['Telefone: {}'.format(row["Número de telefone"]), 'R$: {}'.format(row["Total da Reserva"])])

        elif (apt == '807'):
            cgdmsg807.extend(['Anuncio: {}'.format(apt), 'Hospede: {}'.format(nome_hosp_formatado), canal, qtd_hospedes])
            if (nao_e_airbnb):
                cgdmsg807.extend(['Telefone: {}'.format(row["Número de telefone"]), 'R$: {}'.format(row["Total da Reserva"])])

    elif (relacao == 'saida'):  #Nas Saídas Vou deixar apenas o Apt/nome do hospede mesmo
        if (apt == '215'):
            sdsmsg215.extend(['Anuncio: {}'.format(apt), 'Hospede: {}'.format(nome_hosp_formatado)])
            #sdsmsg215.append(canal)
            # if (nao_e_airbnb):
            #     sdsmsg215.extend(['Telefone: {}'.format(row["Número de telefone"]), 'R$: {}'.format(row["Total da Reserva"])])

        elif (apt == '403'):
            sdsmsg403.extend(['Anuncio: {}'.format(apt), 'Hospede: {}'.format(nome_hosp_formatado)])
            
        elif (apt == '510'):
            sdsmsg510.extend(['Anuncio: {}'.format(apt), 'Hospede: {}'.format(nome_hosp_formatado)])
            
        elif (apt == '614'):
            sdsmsg614.extend(['Anuncio: {}'.format(apt), 'Hospede: {}'.format(nome_hosp_formatado)])
            
        elif (apt == '807'):
            sdsmsg807.extend(['Anuncio: {}'.format(apt), 'Hospede: {}'.format(nome_hosp_formatado)])
            


#Gerando mensagens de entrada (utilizando a função acima) ->
for i, row in dados_de_chegadas215.iterrows():
    gerar_array_com_infos_de_chegadas_e_saidas_formatados('entrada', '215')

for i, row in dados_de_chegadas403.iterrows():
    gerar_array_com_infos_de_chegadas_e_saidas_formatados('entrada', '403')

for i, row in dados_de_chegadas510.iterrows():
    gerar_array_com_infos_de_chegadas_e_saidas_formatados('entrada', '510')

for i, row in dados_de_chegadas614.iterrows():
    gerar_array_com_infos_de_chegadas_e_saidas_formatados('entrada', '614')

for i, row in dados_de_chegadas807.iterrows():
    gerar_array_com_infos_de_chegadas_e_saidas_formatados('entrada', '807')


#Gerando mensagens de saida ->
for i, row in dados_de_saidas215.iterrows():
    gerar_array_com_infos_de_chegadas_e_saidas_formatados('saida', '215')

for i, row in dados_de_saidas403.iterrows():
    gerar_array_com_infos_de_chegadas_e_saidas_formatados('saida', '403')

for i, row in dados_de_saidas510.iterrows():
    gerar_array_com_infos_de_chegadas_e_saidas_formatados('saida', '510')

for i, row in dados_de_saidas614.iterrows():
    gerar_array_com_infos_de_chegadas_e_saidas_formatados('saida', '614')

for i, row in dados_de_saidas807.iterrows():
    gerar_array_com_infos_de_chegadas_e_saidas_formatados('saida', '807')

msgs_chegadas = [cgdmsg215, cgdmsg403, cgdmsg510, cgdmsg614, cgdmsg807]
msgs_saidas = [sdsmsg215, sdsmsg403, sdsmsg510, sdsmsg614, sdsmsg807]

# for i in msgs_chegadas:
#     print(i)

# for i in msgs_saidas:
#     print(i)


# FUNÇÃO PARA GERAR RELATORIO QUE SERÁ ENVIADO NO WPP ->
def pular_linha():
    pyautogui.hotkey('shift', 'enter')

def gerar_relatorio_pai_e_mae():
    resumo_entrada = []
    resumo_saida = []
    pyautogui.write('------RELATORIO DIARIO------')
    pyautogui.PAUSE = 0.2
    pular_linha()
    pyautogui.write('---------------------------')
    pular_linha()
    pular_linha()

    #Escrever Entradas ->
    pyautogui.write('Chegadas amanha ->')
    pular_linha()
    pular_linha()
    total_chegadas = 0
    for mensagems in msgs_chegadas:  # -> [[cgdmsg215], [cgdmsg403], [cgdmsg510], [cgdmsg614], [cgdmsg807]]
        if mensagems != []:
            total_chegadas += 1
            arr_msg = mensagems[0].split()      #Mensagems[0] = 'Anuncio: 215' então o split[1] é para pegar só '215'
            resumo_entrada.append(arr_msg[1])   #Para aqui colocar nos array de resumos, tipo: ['215', '614']
            pyautogui.write('---------------------------')
            pular_linha()
            for msg in mensagems:   # -> [cgdmsg215] -> ['Anuncio: xxx', 'Hosp: xxxx xx xxxx', 'Canal: Airbnb']
                pyautogui.write(msg)
                pular_linha()
    pyautogui.write('---------------------------')
    pular_linha()
    pyautogui.write(f'TOTAL DE CHEGADAS: {total_chegadas}')
    pular_linha()
    pyautogui.write('---------------------------')
    pular_linha()
    pyautogui.write('---------------------------')
    pular_linha()
    pular_linha()

    #Escrever Saidas ->
    pyautogui.write('Saidas Amanha ->')
    pular_linha()
    pular_linha()
    total_saidas = 0
    for mensagems in msgs_saidas:
        if mensagems != []:
            total_saidas += 1
            arr_msg = mensagems[0].split()
            resumo_saida.append(arr_msg[1])
            pyautogui.write('---------------------------')
            pular_linha()
            pyautogui.write(mensagems[0])
            pular_linha()
            pyautogui.write(mensagems[1])
            pular_linha()
    pyautogui.write('---------------------------')
    pular_linha()
    pyautogui.write(f'TOTAL DE SAIDAS: {total_saidas}')
    pular_linha()
    pyautogui.write('---------------------------')
    pular_linha()
    pyautogui.write('---------------------------')
    pular_linha()
    pular_linha()

    #Escrever Resumo
    pyautogui.write('Resumo Geral ->')
    pular_linha()
    pular_linha()

    #Entrada ->
    str_resumo_entr = ''
    tam = len(resumo_entrada)
    num = 0

    for i in resumo_entrada:
        num += 1
        if num == tam:
            str_resumo_entr += i
        else:
            str_resumo_entr += '{}/'.format(i)
    pyautogui.write('Entradas: ({})'.format(str_resumo_entr))
    pular_linha()

    #Saida ->
    str_resumo_sd = ''
    tam = len(resumo_saida)
    num = 0

    for i in resumo_saida:
        num += 1
        if num == tam:
            str_resumo_sd += i
        else:
            str_resumo_sd += '{}/'.format(i)
    pyautogui.write('Saidas: ({})'.format(str_resumo_sd))
    pular_linha()
    pular_linha()
    pyautogui.write('---------------------------')
    pular_linha()
    pyautogui.write('---------------X---------------')
    pyautogui.press('enter')
    pyautogui.PAUSE = 1.5
    time.sleep(1)


#ENVIAR NO WHATSAPP ->
pyautogui.PAUSE = 1.5
pyautogui.press("win")
pyautogui.write("Whatsap")
pyautogui.press("enter")

#Mãe ->
pyautogui.click(x=184, y=119)
pyautogui.write("Mae")
pyautogui.click(x=168, y=185)
gerar_relatorio_pai_e_mae()
if (lembrete_amanha):   # +- linha 53 do código
        pyautogui.write(f'Lembretes Amanha {amanha.day}/{amanha.month} ->')
        pular_linha()
        pular_linha()
        for lmbrt in lembrete_amanha:
            pyautogui.write(lmbrt)
            pular_linha()
        pyautogui.press('enter')
else:
    pyautogui.write('Amanha nao tem Lembretes.\n')

#Pai ->
pyautogui.click(x=184, y=119)
pyautogui.write("Pai")
pyautogui.click(x=168, y=185)
gerar_relatorio_pai_e_mae()

#Vivi ->
pyautogui.click(x=184, y=119)
pyautogui.write("Vivi")
pyautogui.click(x=168, y=185)
pyautogui.PAUSE = 0.2
pyautogui.write('Saidas Amanha ->')
pular_linha()
pular_linha()
total_saidas = 0
for mensagems in msgs_saidas:
    if mensagems != []:
        total_saidas += 1
        pyautogui.write('---------------------------')
        pular_linha()
        pyautogui.write(mensagems[0])
        pular_linha()
        pyautogui.write(mensagems[1])
        pular_linha()
pyautogui.write('---------------------------')
pular_linha()
pyautogui.write(f'Total de saidas: {total_saidas}\n')
time.sleep(1)
pyautogui.click(x=1896, y=10)


#Deletar Arquivos ->
os.remove("fromX_toX_propX_ownerX.xlsx")
os.remove(f'{hoje}_{seteDias}_propX_ownerX.xlsx')