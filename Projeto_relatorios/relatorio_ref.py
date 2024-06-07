import pandas as pd
import numpy as np
from datetime import datetime
import emoji

relatorio = 'relatorio_ref.xlsx'
df = pd.read_excel(relatorio,sheet_name='Planilha1')

with open('relatorio_ref.txt', 'w') as file:
    file.write(relatorio)

previstogeral = df.loc[0, 'previsto']
previstogeralf = int(previstogeral * 100)
realgeral = df.loc[0, 'real']
realgeralf = int(realgeral * 100)
terminoprevisto = df.loc[0, 'termino previsto']
terminoreal = df.loc[0, 'termino real']
desviog = df.loc[0, 'desviog']

print('*Segue report referente redução de fluxo 04/06/2024* \n\U0001F4C5Data: 04/06/2024')
print('\U0001F642Previsto: {}%'.format(previstogeralf))
#Código para avaliar qual emoticon será utilizado com base nas % de previsto e real
if realgeralf >= previstogeralf:
    print('\U0001F600Real: {}%'.format(realgeralf))
else:
    print('\U0001FAE4Real: {}%'.format(realgeralf))
print('\U0001F550Término previsto: {}'.format(terminoprevisto))
print('\U0001F550Término real: {}'.format(terminoreal))
print('Desvio: {}h'.format(desviog))

print('\U0001FE0F Highlights \U0001FE0F')

area1 = df.loc[0, 'area']
previsto1 = df.loc[0, 'preevisto']
previsto1f = int(previsto1 * 100)
real1 = df.loc[0, 'reeal']
real1f = int(real1 * 100)
iniciolb1 = df.loc[0, 'inicio linha de base']
inicioreal1 = df.loc[0, 'inicio real']
terminolb1 = df.loc[0, 'termino linha de base']
terminoreal1 = df.loc[0, 'term real']
desvio1 = df.loc[0, 'desvio']

print(area1)
print('\U0001F642Previsto: {}%'.format(previsto1f))
if real1 > previsto1f:
    print('\U0001F642Real: {}%'.format(real1f))
else:
    print('\U0001FAE4Real: {}%'.format(real1f))
print('Início LB: {}'.format(iniciolb1))
print('Início real: {}'.format(inicioreal1))
print('Término LB: {}'.format(terminolb1))
print('Término real: {}'.format(terminoreal1))
print('Desvio: {}h'.format(desvio1))

area2 = df.loc[1, 'area']
previsto2 = df.loc[1, 'preevisto']
previsto2f = int(previsto2 * 100)
real2 = df.loc[1, 'reeal']
real2f = int(real2 * 100)
iniciolb2 = df.loc[1, 'inicio linha de base']
inicioreal2 = df.loc[1, 'inicio real']
terminolb2 = df.loc[1, 'termino linha de base']
terminoreal2 = df.loc[1, 'term real']
desvio2 = df.loc[1, 'desvio']

print(area2)
print('\U0001F642Previsto: {}%'.format(previsto2f))
if real2f >= previsto2f:
    print('\U0001F642Real: {}%'.format(real2f))
else:
    print('\U0001FAE4Real: {}%'.format(real2f))
print('Início LB: {}'.format(iniciolb2))
print('Início real: {}'.format(inicioreal2))
print('Término LB: {}'.format(terminolb2))
print('Término real: {}'.format(terminoreal2))
print('Desvio: {}h'.format(desvio2))
print('Impacto: Atraso no início da atividade de desmontagem de parafusos do 030-RE-035')

area3 = df.loc[2, 'area']
previsto3 = df.loc[2, 'preevisto']
previsto3f = int(previsto3 * 100)
real3 = df.loc[2, 'reeal']
real3f = int(real3 * 100)
iniciolb3 = df.loc[2, 'inicio linha de base']
inicioreal3 = df.loc[2, 'inicio real']
terminolb3 = df.loc[2, 'termino linha de base']
terminoreal3 = df.loc[2, 'term real']
desvio3 = df.loc[2, 'desvio']

print(area3)
print('\U0001F642Previsto: {}%'.format(previsto3f))
if real3f >= previsto3f:
    print('\U0001F642Real: {}%'.format(real3f))
else:
    print('\U0001FAE4Real: {}%'.format(real3f))
print('Início LB: {}'.format(iniciolb3))
print('Início real: {}'.format(inicioreal3))
print('Término LB: {}'.format(terminolb3))
print('Término real: {}'.format(terminoreal3))
print('Desvio: {}h'.format(desvio3))
print('Impacto na duração de execução da atividade XXXXXX, gerando impacto de 40 minutos no término das ativades. Não impacta no tempo final da obra.')

area4 = df.loc[3, 'area']
previsto4 = df.loc[3, 'preevisto']
previsto4f = int(previsto4 * 100)
real4 = df.loc[3, 'reeal']
real4f = int(real4 * 100)
iniciolb4 = df.loc[3, 'inicio linha de base']
inicioreal4 = df.loc[3, 'inicio real']
terminolb4 = df.loc[3, 'termino linha de base']
terminoreal4 = df.loc[3, 'term real']
desvio4 = df.loc[3, 'desvio']

print(area4)
print('\U0001F642Previsto: {}%'.format(previsto4f))
if real4f >= previsto4f:
    print('\U0001F642Real: {}%'.format(real4f))
else:
    print('\U0001FAE4Real: {}%'.format(real4f))
print('Início LB: {}'.format(iniciolb4))
print('Início real: {}'.format(inicioreal4))
print('Término LB: {}'.format(terminolb4))
print('Término real: {}'.format(terminoreal4))
print('Desvio: {}h'.format(desvio4))
print('Atividade de Limpeza no bocal da válvula de alimentação do FL 91 ainda não iniciada, devido a não realização do bloqueio. Previsão para o bloqueio ser realizado as 12:20. Não impacta no término das ativiades pois não é caminho crítico')

