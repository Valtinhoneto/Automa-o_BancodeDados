## Relatório de Vendas por Loja

Este script em Python utiliza a biblioteca pandas para processar e analisar dados de vendas de uma tabela em formato Excel. O objetivo é gerar um relatório de vendas por loja, contendo informações como faturamento, quantidade de produtos vendidos e ticket médio por produto em cada loja. Além disso, o script utiliza a biblioteca win32com para enviar o relatório por e-mail utilizando o Microsoft Outlook.

# Funcionalidades
Leitura do arquivo "Tabela_Dados.xlsx" contendo os dados de vendas por loja.
Cálculo do faturamento total por loja.
Cálculo da quantidade total de produtos vendidos por loja.
Cálculo do ticket médio por produto em cada loja.
Envio do relatório por e-mail com os resultados obtidos.

# Requisitos
Python 3.x
Bibliotecas pandas e win32com instaladas (pode ser instaladas via pip)

# Utilização
Certifique-se de ter o arquivo "Tabela_Dados.xlsx" no mesmo diretório do script.
Execute o script em um ambiente Python compatível.
Verifique a saída no console, que mostrará o faturamento, quantidade de produtos vendidos e ticket médio por produto em cada loja.
O relatório será enviado por e-mail para o endereço especificado no código.

# Personalização
Caso queira utilizar um arquivo Excel diferente, certifique-se de que o nome e formato do arquivo estejam corretos, e substitua o nome do arquivo na linha tabela_vendas = pd.read_excel('Tabela_Dados.xlsx').
Para alterar o endereço de e-mail de destino do relatório, modifique a linha mail.To = 'valterneto123456789@outlook.com'.
O corpo do e-mail pode ser personalizado conforme necessário, utilizando a variável mail.HTMLBody.

# Observações
Certifique-se de que o Microsoft Outlook esteja instalado e configurado corretamente em seu sistema.
Este script foi desenvolvido para Windows, utilizando a biblioteca win32com para integração com o Outlook. Caso esteja utilizando um sistema operacional diferente, talvez seja necessário adaptar o código para funcionar com a biblioteca de e-mail correspondente.
