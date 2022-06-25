# Leitura de DataSet e Tratamento dos dados
Projeto realizado para obtênção de nota para conclusão da disciplina de Linguagem de Programação Aplicada R. <br>
Este programa lê um arquivo Excel denominado escolas_media_alunos_turma_2010.xls 
e gera um resumo estatístico para cada aba da referida planilha com seus respectivos gráficos.<br>
O resumo estatístico supracitado é gerado em Excel e os gráficoem PDF.<br>
Foram usadas as funções barplot e ggplot para geração dos gráficos<br>

### Remove os warnings do Console do R Studio
```
options(warn=-1)
```


### Inclusão das bibliotecas e pacotes

```
if(!require(tidyverse))
{
  install.packages("tidyverse");
  require(tidyverse)
}
library(readxl)    #install.packages("readxl")
library(ggplot2)   #install.packages("ggplot2")
library(dplyr)
library(tibble)
library(writexl)   #install.packages("writexl")
library(outliers)  #install.packages("outliers")
```


### Variáveis Globais e Constantes
```
colunaInicial <- 16   # Coluna P -1º Ano
colunaFinal   <- 24     # Coluna X - 8ª série/ 9° ano
linhaHeader  <- 8   #Linha 09 da Planilha
linhaInicial <- 9   #Linha 10 da Planilha  
diretorio <- "C:\\Trabalho\\"
nome_arquivo <- "escolas_media_alunos_turma_2010.xls"
escolas_media_alunos_turma_2010 <- paste(diretorio,nome_arquivo, sep="")
extensao_arq_excel <- ".xlsx"
extensao_arq_pdf_tipo1  <- "_v1.pdf"
extensao_arq_pdf_tipo2  <- "_v2.pdf"
nomeRegioes <- c("Norte", "Nordeste01", "Nordeste02", "Sudeste", "Sul", "CentroOeste")
qtd_regioes <- 6
regioes <- c("Norte" =  paste(diretorio, nomeRegioes[1], extensao_arq_excel, sep=""), 
             "Nordeste - EXT MA E BA" =  paste(diretorio, nomeRegioes[2], extensao_arq_excel, sep=""),
             "Nordeste - SOMENTE MA E BA" = paste(diretorio, nomeRegioes[3], extensao_arq_excel, sep=""),
             "Sudeste" =  paste(diretorio, nomeRegioes[4], extensao_arq_excel, sep=""),
             "Sul" =  paste(diretorio, nomeRegioes[5], extensao_arq_excel, sep=""),
             "Centro-Oeste" =  paste(diretorio, nomeRegioes[6], extensao_arq_excel, sep="") )
Ctrs <- regioes
series=c("1ºAno","2ºAno","3ºAno","4ºAno","5ºAno","6ºAno","7ºAno","8ºAno","9ºAno")
cores = c("#0000FF", "#F4A460", "#00FF00", "#8A2BE2", 
          "#FF00FF", "#DC143C", "#FFFF00", "#A9A9A9",
          "#66CDAA")
```


### Define a parte Client da exibição da aplicação Web de exibição dos dados
```
ui <- fluidPage(
  
  # Application title
  titlePanel("Média dos Alunos da turma de 2020"),
  
  # Sidebar with dropdown
  sidebarLayout(
    sidebarPanel(
      selectInput(inputId = "selects", 
                  choices = Ctrs,
                  label = "Selecione a Região", multiple = FALSE),
      selectInput(inputId = "tipo_grafico", 
                  choices = c("Tipo 1", "Tipo 2"),
                  label = "Selecione o Tipo Gráfico", 
                  selected = "Tipo 1",
                  multiple = FALSE)
    ),
    
    # Show a plot of the generated distribution
    mainPanel(
      plotOutput("Plot")
    )
  )
)
```


### Define a parte Server da exibição da aplicação Web de exibição dos dados

```
server <- function(input, output) {
  
  grafico = reactive({
    read_excel(input$selects, sheet=1, .name_repair = "minimal", range = cell_cols("A:B"), skip=1)
  })
  
  plot1 <- reactive({
    # this should be a complete plot image
    mydata = grafico()
    ggplot(mydata,
           aes(x = series, y = `Média das média dos alunos`)) +
      geom_col() +
      scale_y_continuous(limits = c(0, 30))+
      geom_text(aes(label = `Média das média dos alunos`), 
                vjust = -1) +
      labs(title = "Media de Alunos da Região", x="Série", y="Média") +
      theme_grey(base_size = 10) 
  })
  
  plot2 <- reactive({
    # this should be a complete plot image
    mydata = grafico()
    grf_barra = barplot(mydata[[2]],
                        main = "Media de Alunos da Região",
                        xlab = "Série",
                        ylab = "Média",
                        names.arg = series,
                        col = cores,
                        horiz = FALSE,
                        space = 0.2)
    text(x = grf_barra, y = mydata[[2]]-2, labels=mydata[[2]])
  })
  
  # Return the requested graph
  graphInput <- reactive({
    if ( input$tipo_grafico == "Tipo 1") {
      plot1() }
    else  {  
      plot2() }
  })
  
  output$Plot <- renderPlot({ 
    graphInput()
  })
}
```


### Função f_obtem_valores
```
f_obtem_valores <- function(df_regiao, coluna)
{
  # Variáveis 
  x <- 0
  j <- 1
  contador <- linhaInicial
  soma_valor <- 0
  qtd <- 0
  ponto_medio <- 0
  media <- 0
  minimo <- 9999
  maximo <- 0
  mediana <- 0
  desvio_padrao <- 0
  outlier_inferior <- 0
  outlier_superior <- 0
  linha <- coluna - 15
  vetor_auxiliar <- rep(0,100000)
  vetor_dados <- c(media, minimo, maximo, desvio_padrao, mediana,
                   0, 0, 0, 0, outlier_inferior, outlier_superior)   #quarties e outliers
  
  while (x == 0)
  {
    valor <- df_regiao[contador, coluna]
    if (is.na(valor) == TRUE) 
    {
      x <- 1
    }
    
    valor_numerico = as.numeric(valor)
    
    if (is.na(valor_numerico) == FALSE)
    {
      soma_valor <- soma_valor + valor_numerico
      qtd <- qtd + 1
      vetor_auxiliar[j] <- valor_numerico  
      j <- j + 1
      
      # identifica o valor mínimo
      if (valor_numerico < minimo)
      {
        minimo <- valor_numerico  
      }
      
      # identifica o valor maximo
      if (valor_numerico > maximo)
      {
        maximo <- valor_numerico  
      }
      
    }  # fim do if (is.na(valor_numerico) == FALSE)
    
    contador <- contador + 1
  } # fim do while
  
  media <-  soma_valor / qtd
  vetor_dados[1] <- round(media,2)
  vetor_dados[2] <- minimo
  vetor_dados[3] <- maximo
  
  # Preenche o vetor_sd de tamanho definido pela quantidade de
  # registros numericos das colunas da planilha (variável qtd)
  vetor_sd <- rep(0, qtd)
  for (y in 1:qtd)
  {  
    vetor_sd[y] <- vetor_auxiliar[y]
  }  
  
  # Cálculo do desvio padrão
  vetor_dados[4] <- sd(vetor_sd)
  
  # Ordena o vetor para cálculo da Mediana
  vetor_ordenado <- sort(vetor_sd, decreasing = FALSE)
  # Mediana
  vetor_dados[5] <- median(vetor_ordenado)
  
  # Quartis
  quarties = quantile(vetor_ordenado)
  vetor_dados[6] <- quarties[[2]]  # Primeiro quartil - 25%
  vetor_dados[7] <- quarties[[3]]  # Segundo  quartil - 50%
  vetor_dados[8] <- quarties[[4]]  # Terceiro quartil - 75%
  vetor_dados[9] <- quarties[[5]]  # Primeiro quartil - 100%
  
  # outlier - inferior
  outlier_inferior <- outlier(vetor_ordenado, opposite = TRUE)
  vetor_dados[10] <- outlier_inferior
  
  # outlier - Superior
  outlier_superior <- outlier(vetor_ordenado, opposite = FALSE)
  vetor_dados[11] <- outlier_superior
  
  return (vetor_dados)
} # fim da função f_obtem_valores
```



## PROGRAMA PRINCIPAL

```
# Limpa o console do R
cat("\014") 

for (id_regiao in 1:qtd_regioes)
{
  
  Msglog  = paste("Processando dados da região ", nomeRegioes[id_regiao], sep="")
  print(Msglog)   
  nome_arquivo_excel <- paste(diretorio,nomeRegioes[id_regiao],extensao_arq_excel, sep="")
  grafico_tipo1 <- paste(diretorio,nomeRegioes[id_regiao],extensao_arq_pdf_tipo1, sep="")
  grafico_tipo2 <- paste(diretorio,nomeRegioes[id_regiao],extensao_arq_pdf_tipo2, sep="")
  
  ##----------------------------------------------------
  ## Passo 1 - Importar a Planilha em Data Frames
  ## Leitura do Arquivo de Dados
  ##----------------------------------------------------
  
  df_excel <- read_excel(escolas_media_alunos_turma_2010, sheet=id_regiao, .name_repair = "minimal")
  
  ##----------------------------------------------------
  ## Passo 2 - Calcule as médias por ano (série) para 
  ##           cada uma das regiões e apresente um 
  ##           gráfico de barras utilizando as colunas
  ##           referentes as séries
  ## Passo 3 - Análise exploratória dos dados das colunas de cada região 
  ## Dados: média, mínimo, máximo, desvio-padrão, mediana, quartis e Outliers.
  ##----------------------------------------------------
  
  
  # Inicialização do data frame da análise exploratória
  vetor_serie=c("1º Ano", 
                "1ª série/ 2° ano", 
                "2ª série/ 3° ano", 
                "3ª série/ 4° ano",
                "4ª série/ 5° ano",
                "5ª série/ 6° ano",
                "6ª série/ 7° ano",
                "7ª série/ 8° ano",
                "8ª série/ 9° ano")
  vetor_media=c(0,0,0,0,0,0,0,0,0)
  vetor_minimo=c(0,0,0,0,0,0,0,0,0)
  vetor_maximo=c(0,0,0,0,0,0,0,0,0)
  vetor_desvio_padrao=c(0,0,0,0,0,0,0,0,0)
  vetor_mediana=c(0,0,0,0,0,0,0,0,0)
  vetor_quartil1=c(0,0,0,0,0,0,0,0,0)
  vetor_quartil2=c(0,0,0,0,0,0,0,0,0)
  vetor_quartil3=c(0,0,0,0,0,0,0,0,0)
  vetor_quartil4=c(0,0,0,0,0,0,0,0,0)
  vetor_outlier_inferior=c(0,0,0,0,0,0,0,0,0)
  vetor_outlier_superior=c(0,0,0,0,0,0,0,0,0)
  df_analise<-data.frame(vetor_serie,
                         vetor_media, 
                         vetor_minimo,
                         vetor_maximo,
                         vetor_desvio_padrao,
                         vetor_mediana,
                         vetor_quartil1,
                         vetor_quartil2,
                         vetor_quartil3,
                         vetor_quartil4,
                         vetor_outlier_inferior,
                         vetor_outlier_superior)
  
  linha <- 1
  for (coluna in colunaInicial:colunaFinal) 
  {
    vetor_retorno <- f_obtem_valores(df_excel, coluna)  
    df_analise[linha,2]  = round(vetor_retorno[1],2)        # Média
    df_analise[linha,3]  = round(vetor_retorno[2],2)        # Mínimo
    df_analise[linha,4]  = round(vetor_retorno[3],2)        # Máximo
    df_analise[linha,5]  = round(vetor_retorno[4],2)        # Desvio Padrão
    df_analise[linha,6]  = vetor_retorno[5]                 # Mediana
    df_analise[linha,7]  = vetor_retorno[6]                 # Primeiro Quartil
    df_analise[linha,8]  = vetor_retorno[7]                 # Segundo Quartil
    df_analise[linha,9]  = vetor_retorno[8]                 # Terceiro Quartil
    df_analise[linha,10] = vetor_retorno[9]                 # Quarto Quartil
    df_analise[linha,11] = vetor_retorno[10]                # Outlier inferior
    df_analise[linha,12] = vetor_retorno[11]                # Outlier superior
    linha <- linha + 1 
  }
  
  series=c("1ºAno","2ºAno","3ºAno","4ºAno","5ºAno","6ºAno","7ºAno","8ºAno","9ºAno")
  medias= df_analise[,2]
  
  # Altera o Header das colunas
  names(df_analise)[names(df_analise) == "vetor_serie"] <- "Série / Ano"
  names(df_analise)[names(df_analise) == "vetor_media"] <- "Média das média dos alunos"
  names(df_analise)[names(df_analise) == "vetor_minimo"] <- "Menor média"
  names(df_analise)[names(df_analise) == "vetor_maximo"] <- "Maior Média"
  names(df_analise)[names(df_analise) == "vetor_desvio_padrao"] <- "Desvio Padrão"
  names(df_analise)[names(df_analise) == "vetor_mediana"] <- "Mediana"
  names(df_analise)[names(df_analise) == "vetor_quartil1"] <- "Primeiro Quartil"
  names(df_analise)[names(df_analise) == "vetor_quartil2"] <- "Segundo Quartil"
  names(df_analise)[names(df_analise) == "vetor_quartil3"] <- "Terceiro Quartil"
  names(df_analise)[names(df_analise) == "vetor_quartil4"] <- "Quarto Quartil"
  names(df_analise)[names(df_analise) == "vetor_outlier_inferior"] <- "Outlier Inferior"
  names(df_analise)[names(df_analise) == "vetor_outlier_superior"] <- "Outlier Superior"
  
  arquivo_saida = nome_arquivo_excel
  write_xlsx(df_analise, arquivo_saida)
  
  # Data Frame do Gráfico
  df_grafico01 = data_frame(series, medias)
  names(df_grafico01)[names(df_grafico01) == "series"] <- "Série"
  names(df_grafico01)[names(df_grafico01) == "medias"] <- "Média"
  
  # Salva o gráfico em PDF
  pdf(grafico_tipo1, width = 20, height = 14)
  
  strMsg = paste("Média de Alunos da Região", nomeRegioes[id_regiao])
  
  
  grf_barra = barplot(medias,
          main = strMsg,
          xlab = "Série",
          ylab = "Média",
          names.arg = series,
          col = cores,
          horiz = FALSE,
          space = 0.2)
  text(x = grf_barra, y = medias-2, labels=medias)
  dev.off()
  
  ## Gráfico de Barras - Tipo 2
  df_grafico02 = data_frame(series, medias)
  
  ## Salva o gráfico em PDF
  pdf(grafico_tipo2, width = 20, height = 14)
  
  print(
    ggplot(data = df_grafico02,
           aes(x = series, y = round(medias,2))) +
      geom_col() +
      scale_y_continuous(limits = c(0, 30))+
      geom_text(aes(label = round(medias,2)), 
                vjust = -1) +
      labs(title = "Media de Alunos da Região", x="Série", y="Média") +
      theme_grey(base_size = 10) 
  )
  dev.off()
  
  
} # fim da leitura da regioes - for (id_regiao in 1:qtd_regioes)
```

### Executa a aplicação Web
```
shinyApp(ui = ui, server = server)
```
