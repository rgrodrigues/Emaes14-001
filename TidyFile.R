##########Organiza��o do ficheiro de dados para an�lise##########

##instala��o de pacotes necess�rios
if("dplyr" %in% rownames(installed.packages()) == FALSE) {install.packages("dplyr")}
if("DataCombine" %in% rownames(installed.packages()) == FALSE) {install.packages("DataCombine")}
if("rJava" %in% rownames(installed.packages()) == FALSE) {install.packages("rJava")}
if("xlsxjars" %in% rownames(installed.packages()) == FALSE) {install.packages("xlsxjars")}
if("xlsx" %in% rownames(installed.packages()) == FALSE) {install.packages("xlsx")}
if("XLConnect" %in% rownames(installed.packages()) == FALSE) {install.packages("XLConnect")}
if("foreign" %in% rownames(installed.packages()) == FALSE) {install.packages("foreign")}
if("openxlsx" %in% rownames(installed.packages()) == FALSE) {install.packages("openxlsx")}
if("Hmisc" %in% rownames(installed.packages()) == FALSE) {install.packages("Hmisc")}
if("devtools" %in% rownames(installed.packages()) == FALSE) {install.packages("devtools")}

#carregar os pacotes necess�rios

Sys.setenv(JAVA_HOME='C:\\Program Files\\Java\\jre1.8.0_25') #set environment variable
        library(rJava)
        library(xlsxjars)
        library(xlsx)
        library(XLConnect)
        library(foreign)
        library(openxlsx)
        library(Hmisc)
        library(devtools)

##ler o ficheiro Excel
        setwd("D:/Dropbox/_Transfer�ncia/2014-2016/Fresenius/An�liseDados")
        file <- ("DADOS REINFUSAO r02.xlsx")
        data <- read.xlsx(file, 1) #ler a primeira folha do livro excel

#vari�veis e vetores operacionais
        vars <- names(data) #nomes das vari�veis
        meses <- c(2:12) #folhas restantes do livro
        outputfile <- c("DadosFresenius.xlsx")

for(i in meses) {
        data2 <- read.xlsx(file, i)
        data <- merge(data,data2, all.x=TRUE, all.y=TRUE)
}
        rm(data2)

#Mudar todos os valores NULL para NA

data <- as.data.frame(lapply(data, function(x){replace(x, x =="NULL",NA)}))

#gravar o ficheiro em excel para ultrapassar o separador decimal errado. n�o fazendo isto h� confus�o com as casa decimais.
write.xlsx2(data,outputfile) 

#ler de novo
data <- read.xlsx(outputfile)

###manipula��o de vari�veis

        #criar labels com base nos nomes presentes no excel original
        labels <- names(data)
        
        #mudar os nomes das vari�veis para nomes operacionais
        names <- c("id","ano","mes","clinica","doente","hmg","ferrit","reinfusao",
           "epo","Fe","protocolo")

        #mudar o nome das vari�veis
        names(data) <- names

        #mudar de factor para numeric, vari�veis 6 e 9 (7 e 10, com nova vari�vel ID)
        for(i in c(7,10)) {data[,i] <- as.numeric(data[,i])}

        #vari�vel protocolo como factor
        protocoloLevels <- c(0,1)
        protocoloLabels <- c("Com protocolo","Sem protocolo")
        data$protocolo <- factor(data$protocolo,
                                    levels=protocoloLevels, 
                                    labels=protocoloLabels)
        rm(protocoloLevels); rm(protocoloLabels)

        #vari�vel m�s como factor
        meses <- format(ISOdate(2004,1:12,1),"%B")
        data$mes <- factor(data$mes,
                           levels=1:12,
                           labels=meses)

        #vari�vel cl�nica com factor
        data$clinica <- factor(data$clinica)

        #vari�vel doente como factor
        data$doente <- factor(data$doente)


#gravar ficheiro final

write.xlsx2(data,outputfile,row.names=FALSE) #gravar ficheiro xlsx completo

##gravar em formato spss (de http://r4stats.com/examples/data-export/)

        #nome e caminho dos ficheiros

        wd <- paste(getwd(),"/", sep="") #criar o path
        wd <- gsub("/","\\",wd, fixed=TRUE) #mudar / para \ no caminho
        datafile <- paste(wd,"dadosFresenius.txt", sep="") #ficheiro com os dados
        codefile <- paste(wd,"sintaxeFresenius.sps", sep="") #ficheiro de sintaxe

        #labels das vari�veis
        for(i in ncol(data)) {
                label(data[i]) <- labels[i]
        }


write.foreign(data,
              datafile,
              codefile,
              package  = "SPSS",
              variable.labels=labels)

