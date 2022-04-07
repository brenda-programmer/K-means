# Apresentação do Projeto

O presente projeto foi criado utilizando  a linguagem Python. Trata-se de uma aplicação desktop com a finalidade de gerar um gráfico K-Means baseado em uma determinada base de dados escolhida pelo usuário.

# Como rodar o projeto

- Para rodar o projeto você deverá ter o Python instalado na sua máquina;
- Após descompactar o projeto, baixado daqui do Git, clicar no arquivo kmeans.py, e o projeto será aberto;



# O Projeto K-Means

- A interface do projeto exibirá as opções de importar um arquivo excel, escolher o eixo x e y, digitar a quantidade de clusters, Elbow, gerar gráfico e gerar planilha:

![image](https://user-images.githubusercontent.com/54628539/162088467-f11bc52f-3d73-4371-b277-32791e2fe1fe.png)

- O botão **Importar um arquivo Excel**, possibilitará que o usuário escolha a base de dados em que o K-Means irá se basear. Essa base deverá estar contida numa planilha excel. A planilha de exemplo, *Mall_Customers.xlsx*, está contida na pasta junto à aplicação e ela contém dados de compradores de um shopping.

- Os campos **Escolha o eixo x** e **Escolha o eixo y**, apresentarão uma listagem com as colunas da planilha que foi carregada. Essas são as opções de eixo para o gráfico K-Means que será gerado:

![image](https://user-images.githubusercontent.com/54628539/162090940-792dfa6c-ddfe-4d3e-8f8a-31df265e967a.png)

- O campo **Digite a quantidade de Clusters**, é destinado para que o usuário digite a quantidade de clusters desejada para o gráfico K-Means.

- O botão **Elbow**, pode ser acionado pelo usuário caso ele não saiba que quantidade de clusters colocar. Este botão apresentará um gráfico Elbow, um gráfico de linha, que indicará uma melhor quantidade de clusters para o usuário bem no ponto em que se forma um "cotovelo". A partir desse valor indicado o usuário poderá colocá-lo no campo em que se pede a quantidade de clusters:

![image](https://user-images.githubusercontent.com/54628539/162093002-e9b7a72e-267c-4fbe-aa5c-97b4058cb6de.png)

- O botão **Gerar Gráfico**, irá gerar o gráfico K-Means de acordo com os dados fornecidos na base de dados carregada, tomando como parâmetro os eixos x e y escolhidos e agrupando conforme a quantidade de clusters que foi digitada:

![image](https://user-images.githubusercontent.com/54628539/162093776-a1bb563d-4201-43c1-b79b-59f539d418d5.png)

- O botão **Gerar Planilha**, irá gerar uma palnilha excel, que será salva automáticamente na pasta do projeto com o nome *PlanilhaKmeans.xlsx*, que exibirá os detalhes do gráfico K-Means - mostrará em que cluster está cada um dos dados da da base de dados e também um gráfico de barras que exibe a quantidade de dados em cada cluster:

![image](https://user-images.githubusercontent.com/54628539/162094420-064de9bf-b087-4bc9-81e3-b394b52e1b16.png)

