library(DT)
library(shiny)
library(graphics)
library(shinydashboard)
library(reticulate)
library(ggplot2)
library(readxl)



pandas <- import("pandas")
source_python("otimiza.py")




paleta=function(df_c_atend_OR){
  trats<-unique(df_c_atend_OR['Tratamento'])
  cbp2 <- c( "#E69F00", "#56B4E9", "#009E73","#F0E442", "#0072B2", "#D55E00", "#CC79A7","#000000")
  dicio_cores<-c(cbp2[1:length(unlist(trats))])
  names(dicio_cores)<-trats[[1]]
  return(dicio_cores)}


paleta_func=function(df_c_atend_OR,dicio_cores){
  funcs<-c(unique(df_c_atend_OR[c('Tratamento','Funcionário')]))
  dicio_func<-funcs[[1]]
  names(dicio_func)<-funcs[[2]]
  
  cores_func<-dicio_cores[dicio_func]
  names(cores_func)<-names(dicio_func)
  
  return(cores_func)
}




#NÃºmero de atendimentos
Num_atend = function(df_c_atend_OR){ return(nrow(df_c_atend_OR))}

#NÃºmero de crianÃ§as atendidas
Num_kids_atend = function(df_c_atend_OR){return(length(unique(df_c_atend_OR[['Criança']])))}


#quantidade de funcionarios
Num_func = function(df_c_atend_OR){return(length(unique(df_c_atend_OR[['Tratamento']])))}

#GrÃ¡fico FuncionÃ¡rios
Plot_funcio = function(df_c_atend_OR){
  df_func <-data.frame(table(unlist(df_c_atend_OR["Funcionário"])))
  colnames(df_func)<-c("Funcionário",'Atendimentos')
  cores<-paleta(df_c_atend_OR)
  cores_funci<-sort(paleta_func(df_c_atend_OR,cores))
  
  funcs<-c(unique(df_c_atend_OR[c('Tratamento','Funcionário')]))
  dicio_func<-funcs[[2]]
  names(dicio_func)<-funcs[[1]]
  
  print(dicio_func[unique(names(dicio_func))])
  
  plot_funcio<-ggplot(data=df_func, aes(x=reorder(Funcionário,Atendimentos), y=Atendimentos,fill=Funcionário))+scale_fill_manual( name='Especialidades', breaks=dicio_func[unique(names(dicio_func))], labels=unique(names(dicio_func)), values = cores_funci)+
    geom_bar(stat="identity",width=0.6) +
    geom_col(aes(y=95),fill='grey80',colour='black')+
    geom_col(colour='black')+
    geom_hline(yintercept = mean(df_func[['Atendimentos']]), color="red")+
    theme(panel.background = element_blank(),
          axis.text=element_text(size=rel(1.9)),
          axis.title = element_text(size=30),
          legend.title = element_text(size=25),
          legend.text = element_text(size=19),
          axis.line.x = element_line()
    )+
    labs(x = "Funcionário")+
    labs(y = "Quantidade de Atendimentos")+
    geom_text(aes(label=Atendimentos, y=Atendimentos-2),size=5)
  
  
  
  return(plot_funcio + coord_flip())
}


#GrÃ¡fico Tratamentos
Plot_trat=function(df_c_atend_OR){
  df_trat <-data.frame(table(unlist(df_c_atend_OR["Tratamento"])))
  colnames(df_trat)<-c("Especialidade",'Atendimentos')
  cores<-paleta(df_c_atend_OR)
  print(cores)
  plot_trat<-ggplot(data=df_trat, aes(x=reorder(Especialidade,Atendimentos), y=Atendimentos,fill=Especialidade))+
    geom_bar(stat="identity",width=0.6,colour='black')+
    geom_hline(yintercept = mean(df_trat[['Atendimentos']]), color="red")+
    scale_fill_manual(values = cores)+
    theme(panel.background = element_blank(),
          axis.text=element_text(size=rel(1.9)),
          axis.title = element_text(size=30),
          legend.position = 'none',
          axis.line.x = element_line())+
    labs(x = "Especialidade")+
    labs(y = "Quantidade de Atendimentos")+
    geom_text(aes(label=Atendimentos, y=Atendimentos-2),size=5)
  return(plot_trat +coord_flip())
}

#GrÃ¡fico Dias
Plot_dias=function(df_c_atend_OR){
  df_dias <-data.frame(table(unlist(df_c_atend_OR["Dia"])))
  colnames(df_dias)<-c("Dia",'Atendimentos')
  plot_dias<-ggplot(data=df_dias, aes(x=Dia, y=Atendimentos))+
    scale_fill_manual(breaks=c('Segunda','Terça','Quarta','Quinta','Sexta'), labels=df_dias['Dia'])+
    geom_bar(stat="identity", fill="#009E73", width=0.5,colour='black')+ 
    geom_hline(yintercept = mean(df_dias[['Atendimentos']]), color="red")+
    geom_text(aes(label=Atendimentos, y=Atendimentos+4),size=5)+
    labs(x = "Dia da semana")+
    labs(y = "Total de atendimentos")+
    scale_x_discrete(limits=c('Segunda','Terça','Quarta','Quinta','Sexta'))+
    theme(panel.background = element_blank(),
          axis.text=element_text(size=rel(1.7)),
          axis.title = element_text(size=30),
          legend.position = 'none',
          axis.line.y = element_line()
    )
  
  return(plot_dias)}



#GrÃ¡fico Categoria

Plot_cat=function(dados_atendimentos){
  df_kids<-data.frame(unlist(dados_atendimentos))
  colnames(df_kids)<-c('Crianças')
  df_kids['Categorias']<-c('Plenamente Atendidas','Parcialmente Atendidas','Não Atendidas')
  plot_kids<-ggplot(data=df_kids, aes(x=Categorias, y=Crianças,fill=Categorias))+  
    scale_fill_manual(labels=c('Não Atendidas','Parcialmente Atendidas','Plenamente Atendidas'),values = c("red1", "gold1", "green1"))+
    geom_bar(stat="identity", width=0.5,colour='black')+
    geom_text(aes(label=Crianças, y=Crianças+4),size=5)+
    theme(panel.background = element_blank(),
          axis.text=element_text(size=rel(1.6)),
          axis.title = element_text(size=30),
          legend.title = element_text(size=25),
          legend.text = element_text(size=19),          
          axis.line.y = element_line())+
    labs(x = "  ")+
    labs(y = "Quantidade de Crianças")
  
  return(plot_kids)}

options(shiny.maxRequestSize = 200*1024^2)

dh<- dashboardHeader(title = "Menu", dropdownMenu(type = "notifications",
                                                 messageItem(
                                                   from = "Erros",
                                                   message = "Ler a aba das Inconsistencias após o programa ser otimizado ."
                                                 )))
ds<-  dashboardSidebar(
  sidebarMenu(
    #menuItem("Inserir dados do excel", tabName = "dados", icon = icon("file-excel")),
    menuItem("Cadastro das Crianças",tabName = "cadastro",icon = icon("list-ol")),
    #menuItem("Revisão do Excel (Profissionais)", tabName = "pes_prof",icon = icon ("address-card")),
    #menuItem("Revisão do Excel (Crianças)", tabName = "criancas", icon= icon ('calendar-alt')),
    menuItem("Resultados Gerais", tabName = "informacoes", icon = icon("star")),
    menuItem("Pesquisa", tabName = "pesquisa", icon = icon("search")),
    menuItem("Inconsistências", tabName = "erro", icon = icon("exclamation-triangle"))
    
  ))


arquivo<- fluidRow(
  # fluidRow(column(width=7,offset=2,box(title = "ROTINA INTELIGENTE", width = NULL, solidHeader = TRUE, status = "primary",
  #                                      fluidRow(column(2,img(src='Logo.jpg',height='145',width='145')),
  #                                               column(8,offset=1,wellPanel(width=10, fileInput(inputId = "arq", label="Insira o arquivo Excel da Semana Anterior",multiple = FALSE,accept = c('.xlsx','.xls','.csv','.xlsm'),width ='80%'))
  #                                                      
  #                                               ))),
  #                 fluidRow(column(8,offset=5,uiOutput('ui.action1'),
  #                                 textOutput('text'),
  #                                 textOutput('text3'))),
  #                 fluidRow(column(8,offset=5,uiOutput('ui.action2'),
  #                                 textOutput('text2'))))
  #          
  #          
  # )
  )

# 
# cad<- fluidRow(fluidRow(valueBoxOutput("kid_atend1"),valueBoxOutput("Tot_atendimentos1"),valueBoxOutput('Tot_especialidades')),
#                fluidRow (column(12,box(title = "Atendimentos",width = 12, status = "warning", solidHeader = TRUE,
#                              "",selectInput("tipo_atend", "Tipos de Atendimentos:", choices=c("Atendimento Regular","Atendimento Esporádico"))))),
#                fluidRow(valueBoxOutput("kid_atend4"), valueBoxOutput('Tot_atendimentos4')),
#                fluidRow(box(title='Atendimentos Demandados',status = "primary", solidHeader = TRUE,collapsible = TRUE, dataTableOutput("tabela_atends2"))))
# 
# 


cad<- fluidRow(fluidRow(valueBoxOutput('Tot_especialidades'),valueBoxOutput("kid_atend1"),valueBoxOutput("Tot_atendimentos1")),
               fluidRow (box(title = "Atendimentos",width = 4, status = "primary", solidHeader = TRUE,
                             "",selectInput("tipo_atend", "Tipos de Atendimentos:", choices=c("Atendimento Regular","Atendimento Esporádico")))
                         ,valueBoxOutput("kid_atend4"), valueBoxOutput('Tot_atendimentos4')),
               fluidRow(column(12,box(title='Tabela de Atendimentos',width = 12,status = "primary", solidHeader = TRUE,collapsible = TRUE, dataTableOutput("tabela_atends2"))))
)

infos_gerais <- fluidRow( 
 fluidRow(valueBoxOutput("kid_atend"),valueBoxOutput("Tot_atendimentos")),
  box(width=11,height=12,title = "RESULTADOS", status = "success", solidHeader = TRUE,
      tabBox( id = "tabset1",width = 12,
              tabPanel("Alocação de crianças", h2("Número de atendimentos por categoria ") ,h4("Crianças PARCIALMENTE atendidas são aquelas que tiveram PARTE das consultas que demandavam atendidas."),br(), h4("Crianças PLENAMENTE atendidas são aquelas que tiveram TODAS das consultas que demandavam atendidas."),br(), h4("Crianças NÃO atendidas são aquelas que não foram atendidas em nenhuma de suas demandas."),br(),
                       plotOutput("plot1")),
              tabPanel("Atendimentos por Especialidade/Funcionários", h2("Número de atendimentos semanais por Funcionário") , h4("A linha vermelha representa a média de atendimentos em ambos os gráficos.")
                       , fluidRow(plotOutput("plot4",height = 300)),br(),br(),br(),h2("Número de atendimentos semanais por Especialidade"),h4("A linha vermelha representa a média de atendimentos em ambos os gráficos."),fluidRow(plotOutput("plot2",height = 300))),
              tabPanel("Atendimentos por dia da semana", h2("Número total de atendimentos por dia da semana") , h4("A linha vermelha representa a média de atendimentos")
                       , plotOutput("plot3")) 
  )))


pesquisa<-fluidRow(fluidRow(
  column(3,offset= 3,box(width=NULL, title = "Faça o download do Excel", status = "warning",downloadButton("downloadexcel", "Download Excel"))),
  column(3,offset= 0,box(width=NULL,title = "Faça o download dos PDF's ", status = "warning",downloadButton('pdf','Gerar PDFs')))),
  fluidRow(column(12,box(title = "PESQUISA","Pesquise pelo nome da criança e consulte seus horários",
                         width=12, status = "success", solidHeader = TRUE,dataTableOutput('tabela')))))

inconsistencias<- fluidRow (fluidRow(
  box(title = " Crianças que não constam nas planilhas de Atendimento ",icon = icon("file-excel"), status = "success", solidHeader = TRUE,collapsible = TRUE,
      icon = icon("file-excel"),fill = TRUE,verbatimTextOutput('incon1')),
  box(title = "Crianças que apresentam mais demanda de tratamentos do que horários disponíveis", status = "success", solidHeader = TRUE,collapsible = TRUE,
      verbatimTextOutput('incon2')),
  box(title = "Crianças que foram não atendidas", status = "danger", solidHeader = TRUE,collapsible = TRUE,
      verbatimTextOutput('incon3')),
  box(title = "Crianças que não tiveram todos seus atendimentos alocados", status = "danger", solidHeader = TRUE,collapsible = TRUE,
      verbatimTextOutput('incon4')))) 

db<- dashboardBody(
  tabItems(#O conteÃºdo de cada aba Ã© criado dentro da funÃ§Ã£o tabItem()
    #tabItem(tabName = "dados",arquivo),
    tabItem(tabName = "cadastro", cad ),
    # tabItem(tabName = "pes_prof",rev_func),
    # tabItem(tabName = "criancas", rev_criança),
    tabItem(tabName = "informacoes", infos_gerais),
    tabItem(tabName = "pesquisa",pesquisa),
    tabItem(tabName = "erro",inconsistencias)))  


ui <- fluidPage(fluidRow(class='myrow',tags$head(tags$style(".myrow{background-color:white;}")),
  column(11,offset=0,titlePanel(img(src='logo.jpg',height='60',width='60','Casa da Criança Paralítica',windowtitle='Casa da Criança Paralítica'))),HTML('<br/>'),HTML('<br/>'),textOutput('bem'),tags$head(tags$style("#bem{color: black;font-size: 23px;text-align: left;}"))),
  conditionalPanel(condition ='typeof input.gatilho == "undefined"',
   #                fluidRow(column(1,offset=0,img(src='CasaCCPHome.jpg',height='810',width='1670')),
                   fluidRow(column(1,offset=0,align="left",div(class='fundo' ,img(src='casa2.jpg',height='740',width='1709')),tags$style(".fundo{margin-left:-30px;} ")),
                            HTML('<br/>'),HTML('<br/>'),HTML('<br/>'),HTML('<br/>'),HTML('<br/>'),HTML('<br/>'),HTML('<br/>'),HTML('<br/>'),HTML('<br/>'),HTML('<br/>'),
  
                            column(8,offset=1,align="center",
                                   fluidRow( textOutput('nome7'), tags$head(tags$style("#nome7{color: white;font-size: 60px;font-weight: bold;}"))),
                                   
                                   fluidRow(textOutput('descri7'),tags$head(tags$style("#descri7{color: white;font-size: 30px;}"))),
                                   HTML('<br/>'),HTML('<br/>'),
                                   wellPanel(width=10, fileInput(inputId = "arq", label="Insira o arquivo Excel da Semana Anterior",multiple = FALSE,accept = c('.xlsx','.xls','.csv','.xlsm'),width ='80%')),
                                   fluidRow(column(8,offset=2,align="center",uiOutput('ui.action1'),textOutput('text'),textOutput('text3'),
                                                   tags$head(tags$style("#text3{color: white;font-size: 25px;}")),
                                                   tags$head(tags$style("#text{color: white;font-size: 25px;}"))
                                                   )),
                                   fluidRow(column(8,offset=2,align="center",uiOutput('ui.action2'),textOutput('text2'),uiOutput('ui.action3'),tags$head(tags$style("#text2{color: white;font-size: 25px;}"))))               
                            ))
                   )

  ,
  conditionalPanel(condition='typeof input.gatilho != "undefined"',
  div(class='dash7',dashboardPage(title = "Menu", dh, ds, db, skin='red')),tags$style(".dash7{margin-left:-26px;margin-right:-14px;} "))
  
  )

server = function(input,output){
  # output$front<-renderUI(column(12,img(src='casa.jpg',height='850',width='1700')),
  #                        column(12,wellPanel(width=10, fileInput(inputId = "arq", label="Insira o arquivo Excel da Semana Anterior",multiple = FALSE,accept = c('.xlsx','.xls','.csv','.xlsm'),width ='80%'))
  #                        ))
  output$descri7<-renderText('Planeje a semana da CCP')
  output$nome7<-renderText('EASY WEEK')
  output$bem<-renderText('Bem-Vindo!')
  output$ui.action1 <- renderUI({
    if (is.null(input$arq))
      return()
    actionButton('valida','Confirmar Arquivo',class = "btn-primary")
  })
  
  observeEvent(input$valida,{
    print('validando arquivo')
    
    planilhas <<- pandas$ExcelFile(input$arq$datapath)
    
    # nome_func_dt<-c()
    # for (nome in excel_sheets(input$arq$datapath)){
    #   if (nome !="Auxiliar")
    #     if (nome !="Cadastro da Criança" )
    #       if (nome !="Atendimento Regular")
    #         if (nome !="Atendimento Esporádico")
    #           if (nome !="Disponibilidade da Criança")
    #             nome_func_dt[[nome]] <- read_excel(input$arq$datapath,sheet = nome)
    # }
    # 
    # func_espec<-c()
    # for (nome in names(nome_func_dt)){
    #   func_espec<-c(func_espec,nome)
    # }
    
    
    print('Arquivo pronto')
    
    output$text<-renderText({'Arquivo aprovado, prossiga com a otimização.'})
    output$text3<-renderText({'Clique em otimizar'})
    output$ui.action2 <- renderUI(actionButton('SOLVE','Otimizar',class = "btn-primary"))
    #output$text2<-renderText({'A otimização estará completa quando os gráficos da aba "informações gerais" estiverem disponíveis.'})
    output$text2<-renderText({'Isto pode levar alguns minutos...'})
  })



  # gatilin<-reactive(0)
  # observeEvent(input$SOLVE,{
  #   gatilin()<-1
  #     })
  
  # if(gatilin()==1){
  #   output$text2<-renderText({'Isto pode levar alguns minutos...'})
  # }

  
  observeEvent(input$SOLVE, {
    
    print('Rodando solver')
    Retorno<-main(planilhas,input$arq$datapath)
    df_c_atend_OR<-data.frame(Retorno[1])
    dados_atendimentos<-Retorno[2]
    Disponibilidade<-data.frame(Retorno[3])
    inconsistencias1<-Retorno[4]
    inconsistencias2<-Retorno[5]
    inconsistencias3<-Retorno[6]
    inconsistencias4<-Retorno[7]
    
    output$incon1<-renderPrint(inconsistencias1,width = 12)
    output$incon2<-renderPrint(inconsistencias2,width = 12)
    output$incon3<-renderPrint(inconsistencias3,width = 12)
    output$incon4<-renderPrint(inconsistencias4,width = 12)
    
    
    # output$ui.colunas_funcionarios<- renderUI(selectInput("sele_func", "Nome:",choices = func_espec))
    # 
    # output$tab_profi <- renderDataTable(nome_func_dt[[input$sele_func]])
    
    atendReg <- read_excel(input$arq$datapath,sheet = "Atendimento Regular")
    atendEsp<- read_excel(input$arq$datapath,sheet = "Atendimento Esporádico")
    #quantidades de crianÃ§as atendidas (atendR) 
    quant_cri_atendR = length(unique(atendReg[['IDENTIFICAÇÃO']]))
    #quantidades de crianÃ§as atendidas (atendE)
    quant_cri_atendE = length(unique(atendEsp[['IDENTIFICAÇÃO']]))
    
    #quantidades de atendimentos total (atendR)
    quant_atendR=0
    for(linha in atendReg[["QTD. DE ATENDIMENTO SEMANAL"]]){
      quant_atendR=quant_atendR+linha
    }
    #quantidades de atendimentos total (atendE) 
    quant_atendE = length(atendEsp[['IDENTIFICAÇÃO']])
    
    #tabela_cad<-c("Atendimento Regular"=atendReg,"Atendimento Esporádico"=atendEsp)
    vb1<-c("Atendimento Regular"=quant_cri_atendR,"Atendimento Esporádico"=quant_cri_atendE)
    vb2<-c("Atendimento Regular"=quant_atendR,"Atendimento Esporádico"=quant_atendE)
    vb3<-list("Atendimento Regular"=atendReg,"Atendimento Esporádico"=atendEsp)
    
    
    #Número de atendimentos (pré - resolução)
    quant_atendimentos_cri_cadastradas = quant_atendR+quant_atendE
    
    #Número de crianças atendimentas (pré - resolução)
    quant_cri_cadastradas = quant_cri_atendR+quant_cri_atendE
    
    
    output$kid_atend <- renderValueBox(valueBox(Num_kids_atend(df_c_atend_OR),"Crianças Atendidas",width=20,icon = icon("child"),color = "green"))
    output$Tot_atendimentos <- renderValueBox(valueBox(Num_atend(df_c_atend_OR),"Total de Atendimentos",width=20,icon = icon("user-md"),color = "green"))
    
    output$kid_atend1 <- renderValueBox(valueBox(quant_cri_cadastradas,"Total de Crianças cadastradas",width=20,icon = icon("address-card"),color = "yellow"))
    output$Tot_atendimentos1 <- renderValueBox(valueBox(quant_atendimentos_cri_cadastradas,"Total de Atendimentos reservados",width=20,icon = icon("calendar-alt"),color = "yellow"))
    
    output$kid_atend4 <- renderValueBox(valueBox(vb1[[input$tipo_atend]],paste("Crianças que Demandam",toString(input$tipo_atend)),icon = icon("child"),color = "blue"))
    output$Tot_atendimentos4 <- renderValueBox(valueBox(vb2[[input$tipo_atend]],paste("Total de Atendimentos",toString(input$tipo_atend)),icon = icon("diagnoses"),color = "blue"))
    output$Tot_especialidades <- renderValueBox(valueBox(Num_func (df_c_atend_OR),"Tipos de Atendimentos",width=20,icon = icon("user-md"),color = "yellow"))    
    
    output$tabela_atends2 <-renderDataTable(vb3[[input$tipo_atend]])
    
    
    #output$disponibilidades<- renderDataTable(Disponibilidady)
    
    
    # output$kid_atend2 <- renderValueBox(valueBox(Num_kids_atend(df_c_atend_OR),"Crianças Atendidas",width=20))
    # output$Tot_atendimentos2 <- renderValueBox(valueBox(Num_atend(df_c_atend_OR),"Total de Atendimentos",width=20))
    
    output$plot4 <- renderPlot(Plot_funcio(df_c_atend_OR))
    output$plot2 <- renderPlot(Plot_trat(df_c_atend_OR))
    output$plot3 <- renderPlot(Plot_dias(df_c_atend_OR))
    output$plot1 <- renderPlot(Plot_cat(dados_atendimentos))
    output$tabela <- renderDataTable(df_c_atend_OR)
    print('Otimização Completa')
    
    output$ui.action3 <- renderUI(actionButton('gatilho','gal'))
  })
  
  output$downloadexcel <- downloadHandler(
    filename<-function(){paste('modelo-de-dados_ccp_Result','xlsm',sep ='.')},
    content<-function(file){file.copy('modelo-de-dados_ccp_Result.xlsm', file)} 
  )
  output$pdf <- downloadHandler(
    filename<-function(){paste('PDFs','zip',sep ='.')},
    content<-function(file){file.copy('PDFS.zip', file)}
  )
  
}


shinyApp(ui = ui, server = server)
