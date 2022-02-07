# rest-ans
Serviço REST a ser disponibilizado em ambiente Windows para consumir os recursos do pacote pywin32, com a finalidade de converter a fonte de dados em XLS para XLSX.
Isso é necessário para que a partir da descompactação do XLSX seja possível acessar os dados em cache das tabelas dinâmicas do arquivo original e oferecê-los para download no format .csv.

Os endpoints expostos por esse serviço têm a finalidade de serem consumidos pelo pipeline construído para [Apache Airflow](https://github.com/marquesini/etl-ans/tree/master) consumí-los e armarzenar os dados trarados em uma base PostgreSQL.

**Endpoints**
* http://192.168.1.6:5000/oil
* http://192.168.1.6:5000/diesel

**Comandos para ativá-lo**
Python <= 3.8
* pip install -r requirements.txt
* python .\convert.py

A extração da base em cache foi realizada com base no código sugerido em uma thread do StackOverflow:
[Extracting data from excel pivot table spreadsheet in linux](https://stackoverflow.com/questions/4433952/extracting-data-from-excel-pivot-table-spreadsheet-in-linux)

Apesar desse script conseguir extrair os dados em cache, a base resultante não representa a realidade dos totais exibidos na tabela dinâmica do arquivo original.

Não consegui via scripts encontrar uma maneira de acessar os dados de maneira confiável. Também tentei converter o XLS para o XLSX diretamente pelo Python em container Linux utilizando o LibreOffice, mas os dados ficaram ainda mais divergentes.

Uma solução confiável que eu adotaria para realizar o Extract das bases em cache, seria utilizar uma ferramenta RPA como o UIPath. O robô faria a seguinte jornada:
* Abrir o XLS original;
* Clicar com o botão direito na tabela de **Vendas, pelas distribuidoras¹, dos derivados combustíveis de petróleo por Grande Região e produto - 2000-2020 (m3)**;
* Clicar na opção **Mostrar lista de campos**;
* Clicar na caixa colunas na coluna **ANO**;
* Clicar na opção **Remover Campo**;
* Dar um duplo clique no total da tabela dinâmica.

Depois eu repetiria os passos acima para a tabela **Vendas, pelas distribuidoras¹, de óleo diesel por tipo e Unidade da Federação - 2013-2020 (m3)**.

Essas etapas resultariam na exibição da fonte de dados original da tabela dinâmica e essas informações poderiam ser exportadas no formato .csv para o ambiente [ETL Apache Airflow](https://github.com/marquesini/etl-ans/tree/master) concluir o Transform e o Load.
