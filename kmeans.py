import tkinter as tk 
from tkinter import filedialog, ttk
import numpy as np
import pandas as pd
from pandas import DataFrame
import matplotlib.pyplot as plt
from sklearn.cluster import KMeans
from sklearn.preprocessing import MinMaxScaler
import xlsxwriter


root= tk.Tk() #atribuindo uma nova instância de Tk à uma variável chamada root
#Tk é a classe principal que retorna a janela principal. Portanto a variável root fará referência à janela principal

root.title("K-means") #titulo da janela principal

canvas1 = tk.Canvas(root, width = 400, height = 330,  relief = 'raised')
canvas1.pack() #pack é um gerenciador de layout. Ele alinha os compontes verticalmente ou horizontalmente

label1 = tk.Label(root, text='Clusterização pelo k-Means')
label1.config(font=('helvetica', 14))
canvas1.create_window(200, 25, window=label1)

label2 = tk.Label(root, text='')
label2.config(font=('helvetica', 8),fg='red')
canvas1.create_window(200, 100, window=label2)

label3 = tk.Label(root, text='Digite a quantidade de Clusters:')
label3.config(font=('helvetica', 10))
canvas1.create_window(200, 185, window=label3) 

entry1 = tk.Entry (root,width=6)
canvas1.create_window(180, 210, window=entry1) 

label4 = tk.Label(root, text='Escolha o eixo x:')
label4.config(font=('helvetica', 10))
canvas1.create_window(120, 130, window=label4) 

entry2 = ttk.Combobox(root, width=15, textvariable = tk.StringVar())
canvas1.create_window(120, 150, window=entry2) 

label5 = tk.Label(root, text='Escolha o eixo y:')
label5.config(font=('helvetica', 10))
canvas1.create_window(280, 130, window=label5) 

entry3 = ttk.Combobox(root, width=15, textvariable = tk.StringVar())  
canvas1.create_window(280, 150, window=entry3) 

label6 = tk.Label(root, text='')
label6.config(font=('helvetica', 8),fg='red')
canvas1.create_window(200, 300, window=label6)



#função obterDadosExcel ----------------------------------------------------------------------------------------------
def obterDadosExcel ():
    
    global df
    global numFig
    global tabela
    global arrayColunas
    numFig= 0
    caminho_do_arquivo = filedialog.askopenfilename()
    ler_arquivo = pd.read_excel (caminho_do_arquivo)
    df = DataFrame(ler_arquivo) 
    colunas = ler_arquivo.columns.values # nome de todas as colunas que tem no arquivo
    quant = colunas.size
    tabela=ler_arquivo .iloc[:,0:quant].values
    listaColunas = ""

    for coluna in colunas:
        listaColunas = listaColunas + "," + coluna # insere o nome de cada coluna numa mesma string, listaColunas, inserindo virgula entre cada elemento
    
    #valores para a combobox de x
    entry2['values'] = (listaColunas.split(","))# listaColunas é cortada a cada virgula deixando essas partes independentes
    #valores para a combobox de y
    entry3['values'] = (listaColunas.split(",")) 
    #valor do texto após importar um arquivo excel
    label2["text"] = "Arquivo importado com sucesso!" #altera o valor de text do label2 após a execução das linhas acima
    arrayColunas = (listaColunas.split(","))

# a função obterDadosExcel é invocada para o clique do botão a baixo 
# command vincula uma função a um botão quando clicado    
botaoExcel = tk.Button(text=" Importar arquivo Excel ", command=obterDadosExcel, bg='green', fg='white', font=('helvetica', 10, 'bold'))
canvas1.create_window(200, 70, window=botaoExcel)




#função gerarElbow -----------------------------------------------------------------------------------------------
def gerarElbow ():

    global df
    global numFig

    x = entry2.get() #retorna o que foi escolhido dentro de entry2, que é o valor do eixo x
    y = entry3.get() #retorna o que foi escolhido dentro de entry3, que é o valor do eixo y

    # Normalizar os valores - colocar todos na mesma escala
    scaler = MinMaxScaler()
    scaler.fit(df[[y]])
    df[y] = scaler.transform(df[[y]])

    scaler.fit(df[[x]])
    df[x] = scaler.transform(df[[x]])

    inertia = [] #Criada uma lista vazia chamada inertia
    for i in range(1,11): # para i de 1 a 10
        kmeans = KMeans(n_clusters=i, init ='k-means++', random_state=1234) 
        kmeans.fit(df[[x,y]])
        inertia.append(kmeans.inertia_) #adicionamos na lista  os valores das inércias 
        #kmeans.interta_  conferirmos o momento de inércia total que foi calculado.

    #gráfico da lista com os valores das inércias
    numFig += 1 #incrementa +1 para o número da figura toda vez que entra nesta função
    plt.figure(numFig) # cria nova figura toda vez que entra nesta função
    plt.plot(range(1,11),inertia) # plota o gráfico quantidade de clusters x  inércia
    plt.suptitle('Gáfico Elbow', fontsize=16)
    plt.title(x+' x '+y)
    plt.xlabel('Número de Clusters')
    plt.ylabel('Inércia total')
    plt.show()

# a função gerarElbow é invocada para o clique do botão a baixo 
# command vincula uma função a um botão quando clicado    
botaoElbow = tk.Button(text="Elbow", command=gerarElbow, bg='gray', fg='white', font=('helvetica', 8, 'bold'))
canvas1.create_window(230, 210, window=botaoElbow)




#função gerarKMeans ----------------------------------------------------------------------------------------------
def gerarKMeans ():

    global df
    global numFig
    global labels
    global numeroDeClusters

    numeroDeClusters = int(entry1.get()) #entry1.get() retorna o que foi digitado dentro de entry1, que é a quantidade de clusters

    
    x = entry2.get() #retorna o que foi escolhido dentro de entry2, que é o valor do eixo x
    y = entry3.get() #retorna o que foi escolhido dentro de entry3, que é o valor do eixo y

    # Normalizar os valores - colocar todos na mesma escala
    scaler = MinMaxScaler()
    scaler.fit(df[[y]])
    df[y] = scaler.transform(df[[y]])

    scaler.fit(df[[x]])
    df[x] = scaler.transform(df[[x]])
    
    #iniciando o k-means do sklearn e aplicando ele aos dados ".fit(df)". Feito isso ele realiza todas as contas.
    kmeans = KMeans(n_clusters=numeroDeClusters, init='k-means++',random_state=0)
    #kmeans.fit(df) 
    kmeans.fit_predict(df[[x,y]])
    centroides = kmeans.cluster_centers_
    labels = kmeans.labels_

    # Define vetor de cores para os clusters do gráfico
    LABEL_COLOR_MAP = {0 : 'Red',1 : 'Lime', 2 : 'Blue',3 : 'Yellow',4: 'Cyan',5: 'Pink',6: 'Gray',7: 'Green',8: 'Purple',9: 'Orange'}
    label_color = [LABEL_COLOR_MAP[l] for l in labels] # atribui as cores para cada cluster seguindo a ordem no vetor
 
    numFig += 1 #incrementa +1 para o número da figura toda vez que entra nesta função
    plt.figure(numFig) # cria nova figura toda vez que entra nesta função
    plt.suptitle('Gráfico Kmeans', fontsize=16)
    plt.title(x+' x '+y)
    plt.Figure(figsize=(4,3), dpi=100) 
    plt.subplot(111) 
    #plt.scatter(df[x], df[y], c= kmeans.labels_.astype(float), s=50, alpha=0.5)
    plt.scatter(df[x], df[y], c=label_color, s=50, alpha=0.5)
    plt.scatter(centroides[:, 0], centroides[:, 1], c='black', s=50, marker='x')
    plt.xlabel(x)
    plt.ylabel(y)
    plt.show()

# a função gerarKMeans é invocada para o clique do botão a baixo    
processButton = tk.Button(text=' Gerar Gráfico ', command=gerarKMeans, bg='brown', fg='white', font=('helvetica', 10, 'bold'))
canvas1.create_window(120, 270, window=processButton)




#função gerarPlanilha ----------------------------------------------------------------------------------------
def gerarPlanilha ():

    global arrayColunas
    global tabela
    global labels
    global numeroDeClusters

    # Create a workbook and add a worksheet.
    workbook = xlsxwriter.Workbook('PlanilhaKmeans.xlsx')
    worksheet = workbook.add_worksheet()

    # Definir formatações
    bold = workbook.add_format({'bold': True})
    center = workbook.add_format({'align': 'center', 'border': 1, 'bg_color': '#ADD8E6'})
    boldCenter = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': '#6495ED', 'text_wrap':True})

    TotalLinhas= len(tabela[:,1])
    TotalColunas = len(arrayColunas)-1

    # insere filtros nas colunas: cordenada da primeira celula e cordenada da ultima celula
    worksheet.autofilter(0, 0, TotalLinhas, TotalColunas)

    # insere os nomes das colunas
    i=0
    for coluna in arrayColunas:
        if coluna !="" :
            worksheet.write(0,i, coluna, boldCenter)
            i=i+1
    worksheet.write(0,i,'Cluster', boldCenter)  
    
    # insere os dados 
    total = np.zeros(numeroDeClusters)
    i=0
    while(i<TotalLinhas):   # enquanto i for menor que o total de linhas
        j=0
        while(j<TotalColunas): # enquanto j for menor que o total de colunas
            dado = tabela[i,j]
            worksheet.write(i+1, j, dado, center)
            j=j+1
        worksheet.write(i+1, j, labels[i], center)
        total[labels[i]] =  total[labels[i]] +1 #guarda o total de pessoas de cada cluster
        i=i+1    

   # Montagem do gráfico
    colInicio=  j+3 # coluna inicial - terceira coluna após a ultima coluna da tabela
    col= colInicio # coluna que será incrementada a partir da coluna inicial
    linCluster = 1 # linha onde ficarão os nomes dos clusters
    linValores = 2  # linha onde ficarão os valores dos clusters
    linGraf = 4 # linha onde iniciará o gráfico

    worksheet.write(linCluster, colInicio-1, 'Cluster',boldCenter) #insere  titulo da linha dos clusters
    worksheet.write(linValores, colInicio-1, 'Pessoas',boldCenter) ##insere  titulo da linha dos valores
    for k in range(numeroDeClusters):
        worksheet.write(linCluster, col, k,boldCenter) #insere a linha com os nomes dos clusters 
        worksheet.write(linValores, col, total[k],center) #insere  a linha com os valores 
        col=col+1 # vai para a próxima coluna
    
    col=col-1
    
    # Cria o gráfico do tipo colunas
    chart1 = workbook.add_chart({'type': 'column'})
    
    # Define vetor com as mesmas cores dos clusters do gráfico
    LABEL_COLOR_MAP = {0 : 'red',1 : 'lime', 2 : 'blue',3 : 'yellow',4: 'cyan',5: 'pink',6: 'gray',7: 'green',8: 'purple',9: 'orange'}
    
    # Lê os dados para o gráfico
    chart1.add_series({
        'categories': ['Sheet1', linCluster, colInicio, linCluster, col], # lê os valores para x
        'values':     ['Sheet1', linValores, colInicio, linValores, col], # lê os valores para y
        'data_labels': {'value': True},
        'points':   [{'fill':{'color':LABEL_COLOR_MAP[l]}}for l in range(numeroDeClusters)], # atribui as cores para cada coluna seguindo a ordem no vetor
    })

    # Adiciona o título do gráfico.
    chart1.set_title({'name': 'Resultado do Agrupamento'})

    #Adiciona nome para o eixo x
    chart1.set_x_axis({'name': 'Clusters'})

    #Adiciona nome para o eixo y
    chart1.set_y_axis({'name': 'Quantidade de Pessoas'})

    # Desativa a legenda do gráfico
    chart1.set_legend({'none': True})

    # Insere o gráfico na planilha.
    worksheet.insert_chart(linGraf,colInicio-1, chart1) 
    
    label6["text"] = "Arquivo PlanilhaKmeans.xlsx salvo com sucesso!" 
    workbook.close()

# a função gerarPlanilha é invocada para o clique do botão a baixo 
# command vincula uma função a um botão quando clicado    
botaoGerarPlanilha = tk.Button(text=" Gerar Planilha ", command=gerarPlanilha, bg='brown', fg='white', font=('helvetica', 10, 'bold'))
canvas1.create_window(280, 270, window=botaoGerarPlanilha)



root.mainloop() # root, que é a referência para janela principal,  invoca a função mainloop 
# mainloop é um laço de repetição que será executado enquanto a janela principal, root, estiver sendo exibida.
# mainloop interrompe a execução do código, que vier após, enquanto a janela estiver sendo exibida