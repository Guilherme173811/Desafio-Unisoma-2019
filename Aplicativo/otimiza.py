import numpy as np
import datetime
from datetime import time
from ortools.linear_solver import pywraplp
import pandas as pd
import openpyxl
from reportlab.pdfgen import canvas
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, PageBreak, Spacer
from reportlab.lib.styles import getSampleStyleSheet 
from reportlab.lib.units import mm, inch
import os
import zipfile

def main(planilhas,path):
    INCONSISTENCIAS1=[]
    INCONSISTENCIAS2=[]
    INCONSISTENCIAS3=[]
    INCONSISTENCIAS4=[]
    
    
    hora1=[time(i,00) for i in range(8,17) ]
    hora2=[time(i,30) for i in range(7,17) ]
    horarios=sorted(hora1+hora2)
    dias=['Segunda','Terça','Quarta','Quinta','Sexta']
    discretiza=pd.DataFrame(columns=dias,index=horarios)
    discretiza=discretiza.fillna(0)
    i=1
    for coluna in list(discretiza):    
        for indx in discretiza.index:
            discretiza.loc[indx,coluna]+=i
            i+=1
    discretiza.loc[time(17,00)]=0
    
   

    #Checa tratamentos disponíveis, agenda da semana passada e disponibilidade dos funcionários
    
    #nome das planilhas dos funcionarios
    nomes_planilhas =[planilha for planilha in planilhas.sheet_names if planilha != 'Auxiliar' and  planilha !='Cadastro da Criança' and planilha !='Disponibilidade da Criança' and  planilha !='Atendimento Regular' and planilha !='Atendimento Esporádico'] 
    #planilhas de horário de cada funcionário
    funcionarios=[planilhas.parse(sheet_name=nome) for nome in nomes_planilhas ] 
    # nome de cada funcionário
    nome_func = [nome for nome in nomes_planilhas]
    # todos os atendimentos disponíveis
    atendimento =[funcionarios[i].columns[0] for i in range(len(funcionarios))]
    
    dicio_funcionarios = {list(zip(nome_func,atendimento))[i]:funcionarios[i] for i in range(len(nomes_planilhas)) }
    
    semana = ['SEG','TER', 'QUA', 'QUI', 'SEX']
    d_semana ={'SEG':'Segunda','TER':'Terça', 'QUA':'Quarta', 'QUI':'Quinta', 'SEX':'Sexta'}
    
    #Quantidade de horários livres por funcionário
    Quantidade_hora_livre_funcionario=[(f[semana].isnull().sum(),f.columns[0]) for f in funcionarios]
    #DataFrame com todos os horários ocupados
    Disponibilidades1 = pd.DataFrame(columns=['Criança Atendida','Dia','Horário','Tratamento','Funcionário','Horário Discretizado'])
        
    semana_nutri = ['SEG.1','TER.1', 'QUA.1', 'QUI.1', 'SEX.1']
    d_semana_nutri ={'SEG.1':'Segunda','TER.1':'Terça', 'QUA.1':'Quarta', 'QUI.1':'Quinta', 'SEX.1':'Sexta'}
    Semana_anterior_nutri = pd.DataFrame(columns=['Criança Atendida','Dia','Horário','Tratamento','Funcionário','Horário Discretizado'])
    
    
    
    for tup_func,df in dicio_funcionarios.items():
        df=df.fillna(0)
        for dia in semana:
            for indx in df.index:
                if df.loc[indx,dia] != 0:
                    Disponibilidades1.loc[int(np.random.randint(0,1000000000))]=[df.loc[indx,dia],d_semana[dia],df.loc[indx,tup_func[1]],tup_func[1],tup_func[0],0]
        
        if tup_func[1]=='Nutrição':
            for dia in semana_nutri:
                for indx in df.index:
                    if df.loc[indx,dia] != 0:
                        Semana_anterior_nutri.loc[int(np.random.randint(0,1000000000))]=[df.loc[indx,dia],d_semana_nutri[dia],df.loc[indx,tup_func[1]],tup_func[1],tup_func[0],0]
    
    Semana_anterior_nutri=Semana_anterior_nutri.set_index('Criança Atendida')  
    for j in range(len(Semana_anterior_nutri.index)):
        Semana_anterior_nutri.iloc[j]['Horário Discretizado']=discretiza.loc[Semana_anterior_nutri.iloc[j]['Horário']][Semana_anterior_nutri.iloc[j]['Dia']] 
    
    Disponibilidades1=Disponibilidades1.set_index('Criança Atendida')  
    for j in range(len(Disponibilidades1.index)):
        Disponibilidades1.iloc[j]['Horário Discretizado']=discretiza.loc[Disponibilidades1.iloc[j]['Horário']][Disponibilidades1.iloc[j]['Dia']] 
    

    #"Divide cada aba do arquivo carregado em diferentes DataFrames do pandas" 
    dispo=planilhas.parse(sheet_name='Disponibilidade da Criança',index_col='IDENTIFICAÇÃO')
    atendR=planilhas.parse(sheet_name='Atendimento Regular',index_col='IDENTIFICAÇÃO')
    atendE=planilhas.parse(sheet_name='Atendimento Esporádico',index_col='IDENTIFICAÇÃO')
    cadastro=planilhas.parse(sheet_name='Cadastro da Criança',index_col='IDENTIFICAÇÃO')
    #"Cria um set com todas as crianças. Retiradas da aba de Disponibilidade" 
    crianças=[i for i in set(dispo.index)]

    #"Cria um DataFrame onde iremos compilar todos os dados 
    #referentes a cada criança. Cada linha é uma criança    
    #e cada coluna é um grupo de informações"
    compilado = pd.DataFrame(columns=['Disponibilidade','Tratamentos','Pode reagendar','Frequencia'],index=crianças)
    
    
    #"Completa o DataFrame "compilado", adicionando um dicionário na coluna "Tratamentos" 
    #para cada criança. Neste dicionário, as keys são os tratamentos 
    #que a criança precisa receber e os values são a quantidade de vezes 
    #por semana de cada tratamento. Caso a criança não esteja em nenhuma 
    #das planilhas, printa mensagem avisando"
    for criança in crianças:
        if criança in atendR.index:
            if type(atendR.loc[criança,'TIPO DE ATENDIMENTO']) == str:
                dicio_atendsR={atendR.loc[criança,'TIPO DE ATENDIMENTO']:atendR.loc[criança,'QTD. DE ATENDIMENTO SEMANAL']}
                dicio_atendsR_agend={atendR.loc[criança,'TIPO DE ATENDIMENTO']:atendR.loc[criança,'PODE SER REAGENDADO?']}
                dicio_atendsR_freq={atendR.loc[criança,'TIPO DE ATENDIMENTO']:atendR.loc[criança,'FREQUENCIA DE ATENDIMENTO']}            
            else:
                dicio_atendsR=dict(zip(atendR.loc[criança,'TIPO DE ATENDIMENTO'],atendR.loc[criança,'QTD. DE ATENDIMENTO SEMANAL']))
                dicio_atendsR_agend=dict(zip(atendR.loc[criança,'TIPO DE ATENDIMENTO'],atendR.loc[criança,'PODE SER REAGENDADO?']))
                dicio_atendsR_freq=dict(zip(atendR.loc[criança,'TIPO DE ATENDIMENTO'],atendR.loc[criança,'FREQUENCIA DE ATENDIMENTO']))
    
            compilado.loc[criança]['Tratamentos']=dicio_atendsR
            compilado.loc[criança]['Pode reagendar']=dicio_atendsR_agend
            compilado.loc[criança]['Frequencia']=dicio_atendsR_freq
            
        elif criança in atendE.index:
            dicio_atendsE={atendimento:1 for atendimento in pd.Series(atendE.loc[criança,'TIPO DE ATENDIMENTO'])}
            compilado.loc[criança]['Tratamentos']=dicio_atendsE
            compilado.loc[criança]['Pode reagendar']=dict(zip(compilado.loc[criança]['Tratamentos'].keys(),['Sim' for i in range(len(compilado.loc[criança]['Tratamentos'].keys()))]))
            compilado.loc[criança]['Frequencia']=dict(zip(compilado.loc[criança]['Tratamentos'].keys(),['Semanal' for i in range(len(compilado.loc[criança]['Tratamentos'].keys()))]))
        else:
            INCONSISTENCIAS1.append('ERRO 0 - A criança '+str(criança)+' não consta nas planilhas "Atendimento Regular" e "Atendimento Esporádico"')
    
    
 
    #"Completa o DataFrame "compilado", adicionando uma lista na coluna "Disponibilidade" 
    #para cada criança. Esta lista é o set de disponibilidade da criança"
                
    def ConstroiSet(day,periodo,h_inicial=0,h_final=0):
         if day == 'Todos':
                if periodo == 'Ambos':
                    return [i for i in range(1,96)]
                elif periodo == 'Tarde':
                    return [i for i in range(10,20)]+[i for i in range(29,39)]+[i for i in range(48,58)]+[i for i in range(67,77)]+[i for i in range(86,96)]
                elif periodo == 'Manhã':
                    return [i for i in range(1,10)]+[i for i in range(20,29)]+[i for i in range(39,48)]+[i for i in range(58,67)]+[i for i in range(77,86)]
                else:
                    lista_dispo=discretiza.loc[h_inicial:h_final].values.tolist()
                    lista_dispo.remove(lista_dispo[-1])
                    lista_dispo=[numero for listinha in lista_dispo for numero in listinha]
                    return lista_dispo
         else:
                if periodo == 'Ambos':           
                    return list(discretiza.loc[time(7,30):time(16,30)][day])
                elif periodo == 'Manhã':
                    return list(discretiza.loc[time(7,30):time(11,30)][day])
                elif periodo == 'Tarde': 
                    return list(discretiza.loc[time(12,00):time(16,30)][day])
                else:
                    lista_dispo=list(discretiza.loc[h_inicial:h_final][day])
                    lista_dispo.remove(lista_dispo[-1])
                    return lista_dispo
    
    for criança in crianças:
        if type(dispo.loc[criança]['DIA DA SEMANA']) == str:  
            compilado.loc[criança]['Disponibilidade']=ConstroiSet(dispo.loc[criança]['DIA DA SEMANA'],dispo.loc[criança]['PERÍODO'],dispo.loc[criança]['HORA INICIAL'],dispo.loc[criança]['HORA FINAL'])
        else:
            disponibilidade_final=[]
            for j in range(len(dispo.loc[criança].index)):
                serie=dispo.loc[criança].iloc[j,:]
                disponibilidade_final=disponibilidade_final + ConstroiSet(serie['DIA DA SEMANA'],serie['PERÍODO'],serie['HORA INICIAL'],serie['HORA FINAL'])
            compilado.loc[criança]['Disponibilidade']=disponibilidade_final
    
    for row in compilado.index:
        compilado.loc[row]['Disponibilidade']=list(set(compilado.loc[row]['Disponibilidade']))
    
    compilado=compilado.fillna(0)       

    #"Checa se criança tem mais demanda de tratamentos do que horário disponível"
    for row in compilado.index:
        if compilado.loc[row]['Tratamentos'] != 0:
            tamanho_dispo=len(compilado.loc[row]['Disponibilidade'])
            quant_atend=sum(compilado.loc[row]['Tratamentos'].values())
            if quant_atend >= tamanho_dispo:
                INCONSISTENCIAS2.append('ERRO 1 - A criança '+str(row)+' apresenta mais demanda de tratamentos do que horários disponíveis')
 
    #Remove crianças cadastradas porém sem demanda de atendimento
    compilado1=compilado.query('Tratamentos!=0')
    
    #Para crianças com 4 ou 5 atendimentos semanais, este valor se torna 3 
    #Para crinças com valores maiores do que 1 na demanda de atendimento de nutrição
    for kid in compilado1.index:
        for trat,num in compilado1.loc[kid]['Tratamentos'].items():
            if num > 3:
                compilado1.loc[kid]['Tratamentos'][trat]=3
            if trat == 'Nutrição ':
                compilado1.loc[kid]['Tratamentos'][trat]=1
                compilado1.loc[kid]['Tratamentos']['Nutrição']=compilado1.loc[kid]['Tratamentos'].pop('Nutrição ')
        for trat,num in compilado1.loc[kid]['Pode reagendar'].items():
            if trat == 'Nutrição ':
                compilado1.loc[kid]['Pode reagendar']['Nutrição']=compilado1.loc[kid]['Pode reagendar'].pop('Nutrição ')
        for trat,num in compilado1.loc[kid]['Frequencia'].items():
            if trat == 'Nutrição ':
                compilado1.loc[kid]['Frequencia']['Nutrição']=compilado1.loc[kid]['Frequencia'].pop('Nutrição ')
                 

    #Cria dataframe contendo info geral dos atendimentos por especialidade
    Set_Tratamentos=dict(dicio_funcionarios.keys())
    Tratamentos=list(set(Set_Tratamentos.values()))
    
    especialidades_df=pd.DataFrame(index=Tratamentos,columns=['Demanda Total','Total de Funcionários','Funcionários'])
    especialidades_df=especialidades_df.fillna(0)
    for dicio in compilado1['Tratamentos']:
        for trat,qnt in dicio.items():
            if trat in Tratamentos:
                especialidades_df.loc[trat]['Demanda Total']+=qnt
    
    for func,trat in Set_Tratamentos.items():
        especialidades_df.loc[trat]['Total de Funcionários']+=1
    
    funcios=[]
    for trats in especialidades_df.index:
        funcios.append([k for k,v in Set_Tratamentos.items() if v == trats])
    
    especialidades_df['Funcionários']=funcios
        
        
    print(especialidades_df)
    

    #Set_Tratamentos=['Neurologia',  'Fisioterapia',  'Terapia Ocupacional',  'Psicologia', 'Fonoaudiologia', 'Pedagogia',  'Nutrição ']
    Set_Tratamentos=Set_Tratamentos=dict(dicio_funcionarios.keys())
    Tratamentos=list(set(Set_Tratamentos.values()))
    M=200
    # Create the mip solver with the CBC backend.
    solver = pywraplp.Solver('simple_mip_program',
                                 pywraplp.Solver.CBC_MIXED_INTEGER_PROGRAMMING)
    
    variaveis=[(criança,horari,trat,func) for criança in compilado1.index for horari in compilado1.loc[criança]['Disponibilidade'] for func,trat in Set_Tratamentos.items() if trat in compilado1.loc[criança]['Tratamentos'].keys()]
    
    X= {vari:solver.IntVar(0,1,str(vari)) for vari in variaveis}
    
    
    S={c:solver.IntVar(0,1,'S ' + str(c)) for c in compilado1.index}
    T={c:solver.IntVar(0,1,'T ' + str(c)) for c in compilado1.index}
    QA={c:solver.IntVar(0,1,'QA ' + str(c)) for c in compilado1.index}
    QI={c:solver.IntVar(0,1,'QI ' + str(c)) for c in compilado1.index}
    SE={c:solver.IntVar(0,1,'SE ' + str(c)) for c in compilado1.index}
    
    DST={(c,t):solver.IntVar(0,1,'ST ' + str(c) +' ' +str(t)) for c in compilado1.index for t in list(set(Set_Tratamentos.values())) }
    DTQA={(c,t):solver.IntVar(0,1,'TQA ' + str(c) +' ' +str(t)) for c in compilado1.index for t in list(set(Set_Tratamentos.values()))}
    DQAI={(c,t):solver.IntVar(0,1,'QAI ' + str(c) +' ' +str(t)) for c in compilado1.index for t in list(set(Set_Tratamentos.values()))}
    DQIS={(c,t):solver.IntVar(0,1,'QIS ' + str(c) +' ' +str(t)) for c in compilado1.index for t in list(set(Set_Tratamentos.values()))}
   
    CT={f:solver.IntVar(0,100,str(f)) for f in Set_Tratamentos.keys() }
    CTmax={t:solver.IntVar(0,100,'max'+str(t)) for t in set(Set_Tratamentos.values()) }
    CTmin={t:solver.IntVar(0,100,'min'+str(t)) for t in set(Set_Tratamentos.values()) }
    
    
    for f,t in Set_Tratamentos.items():
        carga_total=[]
        for c in compilado1.index:
            if t in compilado1.loc[c]['Tratamentos'].keys():
                for h in compilado1.loc[c]['Disponibilidade']:
                    carga_total.append(X[(c,h,t,f)])         
        solver.Add(solver.Sum(carga_total)==CT[f])
        
    for f,t in Set_Tratamentos.items():
        solver.Add(CTmax[t]>=CT[f])
        solver.Add(CTmin[t]<=CT[f])
        
        
    print('Number of variables =', solver.NumVariables())
    
    for c in compilado1.index:
        for h in compilado1.loc[c]['Disponibilidade']:
            solver.Add(solver.Sum(X[(c,h,t,f)] for t in Tratamentos if t in compilado1.loc[c]['Tratamentos'].keys() for f in [k for k,v in Set_Tratamentos.items() if v == t] )<= 1)
    
    for c in compilado1.index:
        for t in Tratamentos:
            if t in compilado1.loc[c]['Tratamentos'].keys():
                solver.Add(solver.Sum(X[(c,h,t,f)] for h in compilado1.loc[c]['Disponibilidade'] for f in [k for k,v in Set_Tratamentos.items() if v == t] )<= compilado1.loc[c]['Tratamentos'][t])
    
    for t in Tratamentos:
       for h in range(1,96):
           for f in [k for k,v in Set_Tratamentos.items() if v == t]:
               listona=[]       
               for key,value in X.items():
                   if key[1]==h and key[2]==t and key[3]==f :
                       listona.append(value)
               solver.Add(solver.Sum(listona)<= 1)
      
    #Restrições dos Períodos     
    for c in compilado1.index:
        #Segunda
        solver.Add(solver.Sum(X[(c,h,t,f)] for h in range (1,10) if h in compilado1.loc[c]['Disponibilidade'] for t in Tratamentos if t in compilado1.loc[c]['Tratamentos'].keys() for f in [k for k,v in Set_Tratamentos.items() if v == t] )<= M*(1-S[c]))
        solver.Add(solver.Sum(X[(c,h,t,f)] for h in range (10,20) if h in compilado1.loc[c]['Disponibilidade'] for t in Tratamentos if t in compilado1.loc[c]['Tratamentos'].keys() for f in [k for k,v in Set_Tratamentos.items() if v == t] )<= M*S[c])
        #Terça
        solver.Add(solver.Sum(X[(c,h,t,f)] for h in range (20,29) if h in compilado1.loc[c]['Disponibilidade'] for t in Tratamentos if t in compilado1.loc[c]['Tratamentos'].keys() for f in [k for k,v in Set_Tratamentos.items() if v == t] )<= M*(1-T[c]))
        solver.Add(solver.Sum(X[(c,h,t,f)] for h in range (29,39) if h in compilado1.loc[c]['Disponibilidade'] for t in Tratamentos if t in compilado1.loc[c]['Tratamentos'].keys() for f in [k for k,v in Set_Tratamentos.items() if v == t] )<= M*T[c])
        #Quarta
        solver.Add(solver.Sum(X[(c,h,t,f)] for h in range (39,48) if h in compilado1.loc[c]['Disponibilidade']for t in Tratamentos if t in compilado1.loc[c]['Tratamentos'].keys() for f in [k for k,v in Set_Tratamentos.items() if v == t] )<= M*(1-QA[c]))
        solver.Add(solver.Sum(X[(c,h,t,f)] for h in range (48,58) if h in compilado1.loc[c]['Disponibilidade'] for t in Tratamentos if t in compilado1.loc[c]['Tratamentos'].keys() for f in [k for k,v in Set_Tratamentos.items() if v == t] )<= M*QA[c])
        #Quinta
        solver.Add(solver.Sum(X[(c,h,t,f)] for h in range (58,67) if h in compilado1.loc[c]['Disponibilidade'] for t in Tratamentos if t in compilado1.loc[c]['Tratamentos'].keys() for f in [k for k,v in Set_Tratamentos.items() if v == t] )<= M*(1-QI[c]))
        solver.Add(solver.Sum(X[(c,h,t,f)] for h in range (67,77) if h in compilado1.loc[c]['Disponibilidade'] for t in Tratamentos if t in compilado1.loc[c]['Tratamentos'].keys() for f in [k for k,v in Set_Tratamentos.items() if v == t] )<= M*QI[c])
        #Sexta
        solver.Add(solver.Sum(X[(c,h,t,f)] for h in range (77,86) if h in compilado1.loc[c]['Disponibilidade'] for t in Tratamentos if t in compilado1.loc[c]['Tratamentos'].keys() for f in [k for k,v in Set_Tratamentos.items() if v == t] )<= M*(1-SE[c]))
        solver.Add(solver.Sum(X[(c,h,t,f)] for h in range (86,96) if h in compilado1.loc[c]['Disponibilidade'] for t in Tratamentos if t in compilado1.loc[c]['Tratamentos'].keys() for f in [k for k,v in Set_Tratamentos.items() if v == t] )<= M*SE[c])
    
    #Restrições intervalo de 1 dia
    for c in compilado1.index:
        for t in Tratamentos:
            if t in compilado1.loc[c]['Tratamentos'].keys():
                #Segunda-Terça
                solver.Add(solver.Sum(X[(c,h,t,f)] for h in range (1,20) if h in compilado1.loc[c]['Disponibilidade'] for f in [k for k,v in Set_Tratamentos.items() if v == t])<= (1-DST[(c,t)]))
                solver.Add(solver.Sum(X[(c,h,t,f)] for h in range (20,39) if h in compilado1.loc[c]['Disponibilidade']for f in [k for k,v in Set_Tratamentos.items() if v == t])<= DST[(c,t)])
                #Terça-Quarta
                solver.Add(solver.Sum(X[(c,h,t,f)] for h in range (20,39) if h in compilado1.loc[c]['Disponibilidade']for f in [k for k,v in Set_Tratamentos.items() if v == t] )<= (1-DTQA[(c,t)]))
                solver.Add(solver.Sum(X[(c,h,t,f)] for h in range (39,58) if h in compilado1.loc[c]['Disponibilidade']for f in [k for k,v in Set_Tratamentos.items() if v == t])<= DTQA[(c,t)])
                #Quarta-Quinta
                solver.Add(solver.Sum(X[(c,h,t,f)] for h in range (39,58) if h in compilado1.loc[c]['Disponibilidade']for f in [k for k,v in Set_Tratamentos.items() if v == t] )<= (1-DQAI[(c,t)]))
                solver.Add(solver.Sum(X[(c,h,t,f)] for h in range (58,77) if h in compilado1.loc[c]['Disponibilidade']for f in [k for k,v in Set_Tratamentos.items() if v == t])<= DQAI[(c,t)])
                #Quinta-Sexta
                solver.Add(solver.Sum(X[(c,h,t,f)] for h in range (58,77) if h in compilado1.loc[c]['Disponibilidade']for f in [k for k,v in Set_Tratamentos.items() if v == t] )<= (1-DQIS[(c,t)]))
                solver.Add(solver.Sum(X[(c,h,t,f)] for h in range (77,96) if h in compilado1.loc[c]['Disponibilidade']for f in [k for k,v in Set_Tratamentos.items() if v == t])<= DQIS[(c,t)])
             
           
    print('Number of constraints =', solver.NumConstraints())
    
    #solver.Maximize(solver.Sum(variavel for variavel in X.values()))
    
    #FO
    objective = solver.Objective()
    for variavel in X.values():
        objective.SetCoefficient(variavel, 1)
    #Função de balanceamento de carga de trabalho entre funcionários
    for variavel in CTmax.values():
        objective.SetCoefficient(variavel, -1)
    for variavel in CTmin.values():
        objective.SetCoefficient(variavel, 1)
      
    objective.SetMaximization()
    
    def ChecaRestr(criança,hora,tratamento,funcionario):
        if criança == 'Indisponível':
            indisp=[]
            for indices,vari in X.items():
                if indices[3]==funcionario and indices[1]==hora:
                    indisp.append(vari)
            solver.Add(solver.Sum(indisp)<= 0)            
        elif criança not in set(cadastro.index):
            INCONSISTENCIAS1.append('Erro 1 - A criança '+str(criança)+' não consta na planilha de atendimentos')
        else:
            #Lógica da Nutrição
            if (criança,hora,tratamento,funcionario) in X.keys() and tratamento == 'Nutrição' and compilado1.loc[criança,'Frequencia']['Nutrição']=='Quinzenal':
                for nome,vari in X.items():
                    if nome[0]==criança and nome[2]=='Nutrição':
                        vari.SetUb(0)
            if (criança,hora,tratamento,funcionario) in X.keys() and tratamento in compilado1.loc[criança]['Pode reagendar'].keys() and compilado1.loc[criança]['Pode reagendar'][tratamento]=='Não':
                objective.SetCoefficient(X[(criança,hora,tratamento,funcionario)], 1000)
    
    for l in range(len(Disponibilidades1.index)):
        ChecaRestr(Disponibilidades1.iloc[l].name,Disponibilidades1.iloc[l]['Horário Discretizado'],Disponibilidades1.iloc[l]['Tratamento'],Disponibilidades1.iloc[l]['Funcionário'])
    
    
    
    def ChecaRestr_n(criança,hora,tratamento,funcionario):
        if criança not in set(cadastro.index) and criança !='Indisponível':
            INCONSISTENCIAS1.append('Erro 3 - A criança '+str(criança)+' não consta na planilha de atendimentos')
        else:
            #Lógica da Nutrição
            if (criança,hora,tratamento,funcionario) in X.keys() and tratamento == 'Nutrição':
                objective.SetCoefficient(X[(criança,hora,tratamento,funcionario)], 1000)
    
    for l in range(len(Semana_anterior_nutri.index)):
        ChecaRestr_n(Semana_anterior_nutri.iloc[l].name,Semana_anterior_nutri.iloc[l]['Horário Discretizado'],Semana_anterior_nutri.iloc[l]['Tratamento'],Semana_anterior_nutri.iloc[l]['Funcionário'])
    

    
    result_status = solver.Solve()
    # The problem has an optimal solution.
    assert result_status == pywraplp.Solver.OPTIMAL
    
    assert solver.VerifySolution(1e-7, True)
    
    print('Objective value =', solver.Objective().Value())
    
    
    var_basicas = [nome for nome,vari in X.items() if vari.solution_value() ==1]
    df_c_atend_OR=pd.DataFrame(var_basicas,columns=['Crianças Atendidas','Horário','Tratamento','Funcionário'])
    df_c_atend_OR=df_c_atend_OR.set_index('Crianças Atendidas')
    
    Cria_n_atend=[]
    for kid in compilado1.index: 
        if kid not in df_c_atend_OR.index:
            INCONSISTENCIAS3.append('AVISO - A '+str(kid) + ' não foi atendida')
            Cria_n_atend.append(kid)
        elif type(df_c_atend_OR.loc[kid]['Horário'])== np.int64:
            if sum(compilado1.loc[kid]['Tratamentos'].values()) != 1:
                if 'Nutrição' in compilado1.loc[kid]['Tratamentos'].keys() and compilado1.loc[kid]['Frequencia']['Nutrição'] == 'Quinzenal' and 'Nutrição' in list(Disponibilidades1.loc[kid]['Tratamento']):
                    asd=10
                else:
                    INCONSISTENCIAS4.append('AVISO -A criança ' + str(kid) + ' não teve todos seus atendimentos alocados')
                    Cria_n_atend.append(kid)
        else:
            if sum(compilado1.loc[kid]['Tratamentos'].values()) != len(df_c_atend_OR.loc[kid]['Tratamento']):
                if 'Nutrição' in compilado1.loc[kid]['Tratamentos'].keys() and compilado1.loc[kid]['Frequencia']['Nutrição'] == 'Quinzenal' and 'Nutrição' in list(Disponibilidades1.loc[kid]['Tratamento']):
                    asd=10
                else:
                    INCONSISTENCIAS4.append('AVISO - A criança ' + str(kid) + ' não teve todos seus atendimentos alocados')
                    Cria_n_atend.append(kid)
                
   
    df_c_atend_OR['Criança']=df_c_atend_OR.index

    dias=[]
    horas=[]
    for linha in range(len(df_c_atend_OR.index)):
        hor=df_c_atend_OR.iloc[linha]['Horário']
        coord=np.where(discretiza.values==hor)
        horas.append(str(discretiza.index[coord[0]][0]))
        dias.append(discretiza.columns[coord[1]][0])
    
    df_c_atend_OR['Dia']=dias
    df_c_atend_OR['Hora']=horas
    
    plenamente=0
    parcial=0
    nao_atend=0
    for kid in compilado1.index:
        if 'Nutrição' in compilado1.loc[kid]['Tratamentos'].keys() and compilado1.loc[kid]['Frequencia']['Nutrição'] == 'Quinzenal' and kid in Disponibilidades1.index and  'Nutrição' in list(Disponibilidades1.loc[kid]['Tratamento']):
            demanda=sum(compilado1.loc[kid]['Tratamentos'].values())-1
        else:
            demanda=sum(compilado1.loc[kid]['Tratamentos'].values())
        if kid in df_c_atend_OR.index:
            atendimentos=len(df_c_atend_OR.loc[kid]['Tratamento'])
            if demanda > atendimentos:
                parcial+=1
            elif atendimentos==0:
                nao_atend+=1
            else:
                plenamente+=1
        else:
            nao_atend+=1
            
                
    #GERA ARQUIVO EXCEL COM RESULTADOS
    wb=openpyxl.load_workbook(path,keep_vba=True)
    abas_fixas=['Auxiliar','Cadastro da Criança','Atendimento Regular','Atendimento Esporádico','Disponibilidade da Criança']
    hora1=[time(i,00) for i in range(8,17) ]
    hora2=[time(i,30) for i in range(7,17) ]    
    horarios=sorted(hora1+hora2)
    coordenadas={'Segunda':'B','Terça':'C','Quarta':'D','Quinta':'E','Sexta':'F'}
    coordenadas2={str(key):ind+2 for ind,key in enumerate(horarios)}
    coordenadas_nutri={'Segunda':'H','Terça':'I','Quarta':'J','Quinta':'K','Sexta':'L'}
    desloca={'H':'B','I':'C','J':'D','K':'E','L':'F'}
    
    for inx,nome in enumerate(wb.sheetnames):
        if nome not in abas_fixas:
            df=df_c_atend_OR[df_c_atend_OR['Funcionário']==nome]
            if not df.empty and df['Tratamento'][0]=='Nutrição':
                for i in coordenadas_nutri.values():
                    for j in coordenadas2.values():
                        wb.worksheets[inx][str(i)+str(j)]=''
                        wb.worksheets[inx][str(i)+str(j)]=wb.worksheets[inx][str(desloca[i])+str(j)].value
    
                for i in coordenadas.values():
                    for j in coordenadas2.values():
                        if wb.worksheets[inx][str(i)+str(j)].value != 'Indisponível':
                            wb.worksheets[inx][str(i)+str(j)]=''
                for ind,criança in enumerate(df.index):
                    hora=df.iloc[ind]['Hora']
                    dia=df.iloc[ind]['Dia']
                    wb.worksheets[inx][str(coordenadas[dia])+str(coordenadas2[hora])]=criança  
    
               
            
            else:
                for i in coordenadas.values():
                    for j in coordenadas2.values():
                        if wb.worksheets[inx][str(i)+str(j)].value != 'Indisponível':
                            wb.worksheets[inx][str(i)+str(j)]=''
                for ind,criança in enumerate(df.index):
                    hora=df.iloc[ind]['Hora']
                    dia=df.iloc[ind]['Dia']
                    wb.worksheets[inx][str(coordenadas[dia])+str(coordenadas2[hora])]=criança  
    
    
    wb.save('modelo-de-dados_ccp_Result.xlsm')
    
    
    #GERA PASTA COM TODOS OS PDFs
    
    def separa_crianca(data2,i):
        crianca = data2.index
        cod_criaça = list(set(crianca))    
        selecao = data2.index == i          
        return data2[selecao]
    
    def formata_df(dados):
        dados["Crianças"] = dados.index
        dados = dados[["Crianças","Tratamento","Funcionário","Dia","Hora","Horário"]]
        dados.columns = ["CRIANÇA","TRATAMENTOS","FUNCIONÁRIO","DIA","HORA","HORÁRIO"]
        
        lendados = dados.columns
        novas_colunas = range(len(lendados))
        df = pd.Series(novas_colunas)
        df = pd.DataFrame(df)
        df = df.T
        df.columns = dados.columns
        df.loc[1] = dados.columns
        df = df.drop(0)
        dados = pd.concat([df,dados]) 
        return dados


    def constroi_pdf(dados,cont,i):
        nome = str(i)
        # pagesize = (140 * mm, 216 *mm ) # largura, altura 
        pdf = SimpleDocTemplate (nome+'.pdf') #Criação de um PDF
    
        flowables = [] #Lista de elementos que irão compor o PDF
        dados = np.array(dados).tolist()
    
        sample_style_sheet = getSampleStyleSheet() #Estilos de escrita
    
        paragraph_1 = Paragraph("Relatório de atendimentos da "+ nome, sample_style_sheet['Heading2'])
        paragraph_2 = Paragraph("Data:"+datetime.date.today().strftime("%d/%m/%Y"),sample_style_sheet['BodyText'])
        tabela= Table(dados, spaceBefore = 30)
    
        flowables.append(paragraph_1)
        flowables.append(paragraph_2)
        # flowables.append(Spacer(0*mm,5*mm))
        flowables.append(tabela)
        
        pdf.build (flowables)

    dire=os.getcwd()
    os.chdir(str(dire)+'\\PDF_agenda')
    dire_PDF=os.getcwd()
    for cont,i in enumerate(set(df_c_atend_OR.index)):
        x = separa_crianca(df_c_atend_OR,i)
        y = formata_df(x)
        constroi_pdf(y,cont,i)
    
    os.chdir(dire)
    with zipfile.ZipFile('PDFs.zip', 'w',zipfile.ZIP_DEFLATED) as zipObj:
        for folderName, subfolders, filenames in os.walk(dire_PDF):
            for filename in filenames:
                filePath = os.path.join(folderName, filename)
                zipObj.write( os.path.relpath(filePath, os.path.join(dire_PDF, '..')) )
            
    for f in os.listdir(dire_PDF):
        os.remove(os.path.join(dire_PDF,f))
    
    Disponibilidades1=  Disponibilidades1.drop("Horário Discretizado",axis=1)
    Disponibilidades1['Criança Atendida'] = Disponibilidades1.index  
    Disponibilidades1=Disponibilidades1.drop('Indisponível')
    df_c_atend_OR=df_c_atend_OR.drop('Horário',axis=1)
    return df_c_atend_OR, {'Plenamente': plenamente, 'Parcial': parcial, 'N Atendidas':nao_atend}, Disponibilidades1, INCONSISTENCIAS1,INCONSISTENCIAS2,INCONSISTENCIAS3,INCONSISTENCIAS4

