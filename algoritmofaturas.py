import pandas as pd

tabela = pd.read_excel("03.Fatura.xlsx")

#display(tabela)
#item = tabela.loc=[linha,coluna]
#Coletando UC, Referência, taxa de ICMS e valor total da fatura
UC = tabela.loc[0,'UC']
REF = tabela.loc[0,'REF']
ICMS = tabela.loc[5,'VALOR']
totalFatura = tabela.loc[0,'VALOR']

#Coletando Consumo Ponta, Fora Ponta e Reservado
if 'Consumo em kWh - Ponta' in tabela.columns:
    CP = tabela.loc[0,'Consumo em kWh - Ponta']
    rates_CP = tabela.loc[2,'Consumo em kWh - Ponta']
    values_CP = tabela.loc[3,'Consumo em kWh - Ponta']
else:
    CP = 0
    rates_CP = 0
    values_CP = 0
if 'Consumo em kWh - Fora Ponta' in tabela.columns:
    CFP = tabela.loc[0,'Consumo em kWh - Fora Ponta']
    rates_CFP = tabela.loc[2,'Consumo em kWh - Fora Ponta']
    values_CFP = tabela.loc[3,'Consumo em kWh - Fora Ponta']
else:
    CFP = 0
    rates_CFP = 0
    values_CFP = 0
if 'Consumo em kWh Reservado' in tabela.columns:
    CR = tabela.loc[0,'Consumo em kWh Reservado']
    rates_CR = tabela.loc[2,'Consumo em kWh Reservado']
    values_CR = tabela.loc[3,'Consumo em kWh Reservado']
else:
    CR = 0
    rates_CR = 0
    values_CR = 0
CT = CP + CFP + CR #Cálculo do Consumo Total
values_CT = values_CP + values_CFP + values_CR #Cálculo do Consumo Total em Reais

#Coletando Energia Reativa Ponta, Fora Ponta e Reservado
if 'Energia Reativa Exced em KWh - Ponta' in tabela.columns:
    ERP = tabela.loc[0,'Energia Reativa Exced em KWh - Ponta']
    rates_ERP = tabela.loc[2,'Energia Reativa Exced em KWh - Ponta']
    values_ERP = tabela.loc[3,'Energia Reativa Exced em KWh - Ponta']
else:
    ERP = 0
    rates_ERP = 0
    values_ERP = 0
if 'Energia Reativa Exced em KWh - Fponta' in tabela.columns:
    ERFP = tabela.loc[0,'Energia Reativa Exced em KWh - Fponta']
    rates_ERFP = tabela.loc[2,'Energia Reativa Exced em KWh - Fponta']
    values_ERFP = tabela.loc[3,'Energia Reativa Exced em KWh - Fponta']
else:
    ERFP = 0
    rates_ERFP = 0
    values_ERFP = 0
if 'Energia Reativa Exced Reservado - F Ponta' in tabela.columns:
    ERR = tabela.loc[0,'Energia Reativa Exced Reservado - F Ponta']
    rates_ERR = tabela.loc[2,'Energia Reativa Exced Reservado - F Ponta']
    values_ERR = tabela.loc[3,'Energia Reativa Exced Reservado - F Ponta']
else:
    ERR = 0
    rates_ERR = 0
    values_ERR = 0
ERT = ERFP + ERP + ERR #Cálculo da Energia Reativa Total
values_ERT = values_ERFP + values_ERP + values_ERR #Cálculo da Energia Reativa Total em Reais
   
#Coletando a Demanda Contratada, Faturada, Medida, Não Consumida e Ultrapassagem Fora Ponta
if 'Demanda Contratada FP' in tabela.columns:
    demandaContratFP = tabela.loc[0,'Demanda Contratada FP']
else:
    demandaContratFP = 0
if 'Demanda Fat FP' in tabela.columns:
    demandaFatFP = tabela.loc[0,'Demanda Fat FP']
else:
    demandaFatFP = 0
if 'Demanda de Potência Medida - Fora Ponta' in tabela.columns:
    demandaMedidaFP = tabela.loc[0,'Demanda de Potência Medida - Fora Ponta']
    rates_demandaFP = tabela.loc[2,'Demanda de Potência Medida - Fora Ponta']
    values_demandaFP = tabela.loc[3,'Demanda de Potência Medida - Fora Ponta']
else:
    demandaMedidaFP = 0
    rates_demandaFP = 0
    values_demandaFP = 0
if 'Demanda Potência Não Consumida - F Ponta' in tabela.columns:
    demandaNaoConsFP = tabela.loc[0,'Demanda Potência Não Consumida - F Ponta']
    rates_demandaNaoConsFP = tabela.loc[2,'Demanda Potência Não Consumida - F Ponta']
    values_demandaNaoConsFP = tabela.loc[3,'Demanda Potência Não Consumida - F Ponta']
else:
    demandaNaoConsFP = 0
    rates_demandaNaoConsFP = 0
    values_demandaNaoConsFP = 0
if 'Demanda Potência Ativa - Ultrap - F Ponta' in tabela.columns:
    demandaUltrapFP = tabela.loc[0,'Demanda Potência Ativa - Ultrap - F Ponta']
    rates_demandaUltrapFP = tabela.loc[2,'Demanda Potência Ativa - Ultrap - F Ponta']
    values_demandaUltrapFP = tabela.loc[3,'Demanda Potência Ativa - Ultrap - F Ponta']
else:
    demandaUltrapFP = 0
    rates_demandaUltrapFP = 0
    values_demandaUltrapFP = 0
    
#Coletando Demanda Reativa Ponta e Fora Ponta
if 'Demanda Potência Reativa Exced - Ponta' in tabela.columns:
    demandaReatP = tabela.loc[0,'Demanda Potência Reativa Exced - Ponta']
    rates_demandaReatP = tabela.loc[2,'Demanda Potência Reativa Exced - Ponta']
    values_demandaReatP = tabela.loc[3,'Demanda Potência Reativa Exced - Ponta']
else:
    demandaReatP = 0
    rates_demandaReatP = 0
    values_demandaReatP = 0
if 'Demanda Potência Reativa Exced - F Ponta' in tabela.columns:
    demandaReatFP = tabela.loc[0,'Demanda Potência Reativa Exced - F Ponta']
    rates_demandaReatFP = tabela.loc[2,'Demanda Potência Reativa Exced - F Ponta']
    values_demandaReatFP = tabela.loc[3,'Demanda Potência Reativa Exced - F Ponta']
else:
    demandaReatFP = 0
    rates_demandaReatFP = 0
    values_demandaReatFP = 0
demandaReatTotal = demandaReatP + demandaReatFP #Demanda Reativa Total
values_demandaReatTotal = values_demandaReatP + values_demandaReatFP #Demanda Reativa Total em Reais

#Verificando se houve cobrança de Demanda Complementar
if 'Demanda Complementar - F. Ponta' in tabela.columns:
    DC = tabela.loc[0,'Demanda Complementar - F. Ponta']
    rates_DC = tabela.loc[2,'Demanda Complementar - F. Ponta']
    values_DC = tabela.loc[3,'Demanda Complementar - F. Ponta']
else:
    DC = 0
    rates_DC = 0
    values_DC = 0
    
#Bandeira Vermelha e Amarela
if 'Adic. B. Vermelha' in tabela.columns:
    values_BV = tabela.loc[3,'Adic. B. Vermelha']
else:
    values_BV = 0
if 'Adic. B. Amarela' in tabela.columns:
    values_BA = tabela.loc[3,'Adic. B. Amarela']
else:
    values_BA = 0

#Outros valores
if 'Contrib de Ilum Pub' in tabela.columns:
    values_diversos1 = tabela.loc[3,'Contrib de Ilum Pub']
else:
    values_diversos1 = 0
if 'Devolução Subsídio' in tabela.columns:
    values_diversos2 = tabelab.loc[3,'Devolução Subsídio']
else:
    values_diversos2 = 0
values_diversos = values_diversos1 + values_diversos2

#Valores de créditos
if 'xxx' in tabela.columns:
    values_creditos1 = tabela.loc[3,'xxx']
else:
    values_creditos1 = 0
if 'xxxx' in tabela.columns:
    values_creditos2 = tabelab.loc[3,'xxxx']
else:
    values_creditos2 = 0
values_creditos = values_creditos1 + values_creditos2   


#Imprimindo as informações
print(' UC: ',UC,
      '\n Subgrupo: ',0,
      '\n basic_clientName: ',0,
      '\n custom_dateReference: ',REF,
      '\n basic_modality: ',0,
      '\n rates_icms: ',ICMS/100,
      '\n rates_pisCofins: ',0,
      '\n measures_energyOffPeak: ',CFP,
      '\n measures_energyInjectedConsumedTotal: ',0,
      '\n measures_energyIntermediate: ',0,
      '\n measures_energyPeak: ',CP,
      '\n measures_energyReserved: ',CR,
      '\n measures_energyTotal: ',CT,
      '\n measures_energyReactiveOffPeak: ',ERFP,
      '\n measures_energyReactivePeak: ',ERP,
      '\n measures_energyReactiveReserved: ',ERR,
      '\n measures_energyInjectedTotal: ',0,
      '\n measures_demandComplementaryTotal: ',DC,
      '\n measures_demandContractOffPeak: ',demandaContratFP,
      '\n measures_demandContractPeak: ',0,
      '\n measures_demandBilledPeak: ',0,
      '\n measures_demandBilledOffPeak: ',demandaFatFP,
      '\n measures_demandUnusedOffPeak: ',demandaNaoConsFP,
      '\n measures_demandUnusedPeak: ',0,
      '\n measures_demandMeasuredOffPeak: ',demandaMedidaFP,
      '\n measures_demandMeasuredPeak: ',0,
      '\n measures_demandReactivePeak: ',demandaReatP,
      '\n measures_demandReactiveOffPeak: ',demandaReatFP,
      '\n measures_demandExcessOffPeak: ',demandaUltrapFP,
      '\n measures_demandExcessPeak: ',0,
      '\n values_balanceOffPeak: ',0,
      '\n values_balancePeak: ',0,
      '\n rates_aclTotal: ',0,
      '\n rates_aclOffPeak: ',0,
      '\n rates_aclPeak: ',0,
      '\n rates_energyOffPeakTe: ',0,
      '\n rates_energyOffPeakTotal: ',rates_CFP,
      '\n rates_energyOffPeakTusd: ',0,
      '\n rates_energyInjectedConsumedTotal: ',0,
      '\n rates_energyIntermediateTotal: ',0,
      '\n rates_energyPeakTe: ',0,
      '\n rates_energyPeakTotal: ',rates_CP,
      '\n rates_energyPeakTusd: ',0,
      '\n rates_energyReservedTotal: ',rates_ERR,
      '\n rates_energyReactiveOffPeak: ',rates_ERFP,
      '\n rates_energyReactivePeak: ',rates_ERP,
      '\n rates_energyReactiveReserved: ',rates_ERR,
      '\n rates_energyInjectedTotal: ',0,
      '\n rates_energyInjectedOffPeakTotal: ',0,  
      '\n rates_energyInjectedPeakTotal: ',0,
      '\n rates_energyInjectedTotalTusd: ',0,
      '\n rates_energyInjectedTotalTe: ',0,
      '\n rates_availability: ',0,
      '\n rates_demandBilledOffPeak: ',rates_demandaFP,
      '\n rates_demandBilledPeak: ',0,
      '\n rates_demandComplementaryTotal: ',rates_DC,
      '\n rates_demandExcessOffPeak: ',rates_demandaUltrapFP,
      '\n rates_demandExcessPeak: ',0,
      '\n rates_demandReactiveOffPeak: ',rates_demandaReatFP,
      '\n rates_demandReactivePeak: ',rates_demandaReatP,
      '\n rates_demandUnusedOffPeak: ',rates_demandaNaoConsFP,
      '\n rates_demandUnusedPeak: ',0,
      '\n values_autoProduction: ',0,
      '\n values_tusdDiscountCorrectionOffPeak: ',0,
      '\n values_tusdDiscountCorrectionPeak: ',0,
      '\n values_yellowFlag: ',values_BA,
      '\n values_redFlagP1: ',values_BV,
      '\n values_redFlagP2: ',0,
      '\n values_shortageFlag: ',0,
      '\n values_otherCharges: ',values_diversos,
      '\n values_energyOffPeakTe: ',0,
      '\n values_energyOffPeakTotal: ',values_CFP,
      '\n values_energyOffPeakTusd: ',0,
      '\n values_energyInjectedConsumedTotal: ',0,
      '\n values_energyIntermediateTotal: ',0,
      '\n values_energyPeakTe: ',0,
      '\n values_energyPeakTusd: ',0,
      '\n values_energyPeakTotal: ',values_CP,
      '\n values_energyReservedTotal: ',values_CR,
      '\n values_qualityCredit: ',values_creditos,
      '\n values_demandComplementaryTotal: ',0,
      '\n values_demandBilledOffPeak: ',values_demandaFP,
      '\n values_demandBilledPeak: ',0,
      '\n values_demandBilledTotal: ',0,
      '\n values_demandUnusedOffPeak: ',values_demandaNaoConsFP,
      '\n values_demandUnusedPeak: ',0,
      '\n values_demandUnusedTotal: ',values_demandaNaoConsFP,
      '\n values_demandReactiveOffPeak: ',values_demandaReatFP,
      '\n basic_billedDaysCount: ',0,
      '\n values_demandReactivePeak: ',values_demandaReatP,
      '\n values_demandReactiveTotal: ',values_demandaReatTotal,
      '\n values_demandExcessOffPeak: ',values_demandaUltrapFP,
      '\n values_demandExcessPeak: ',0,
      '\n values_demandExcessTotal: ',values_demandaUltrapFP,
      '\n values_autoProductionDiscount: ',0,
      '\n values_tusdDiscountPeak: ',0,
      '\n values_tusdDiscountOffPeak: ',0,
      '\n values_availability: ',0,
      '\n values_energyReactiveOffPeak: ',values_ERFP,
      '\n values_energyReactivePeak: ',values_ERP,
      '\n values_energyReactiveReserved: ',values_ERR,
      '\n values_energyReactiveTotal: ',values_ERT,
      '\n values_energyInjectedTotalTe: ',0,
      '\n values_energyInjectedTotalTusd: ',0,
      '\n values_icms: ',0,
      '\n values_icmsTaxable: ',0,
      '\n values_icmsDiscount: ',0,
      '\n values_otherCorrections: ',0,
      '\n values_pisCofins: ',0,
      '\n values_services: ',0,
      '\n values_subsidy: ',0,
      '\n values_subsidyCharges: ',0,
      '\n values_subsidyDiscount: ',0,
      '\n values_subvention: ',0,
      '\n values_totalCharges: ',totalFatura)
