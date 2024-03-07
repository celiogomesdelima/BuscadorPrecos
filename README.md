Claro, aqui está a descrição do programa no formato Markdown:

---

### Descrição do Programa

Este programa em Python utiliza a biblioteca Selenium para automatizar a busca de produtos em sites de compras online, como Google Shopping e Buscapé, com base em uma lista de produtos fornecida em um arquivo Excel. Abaixo está uma breve descrição das principais etapas:

1. **Importações de Bibliotecas**: O programa importa as bibliotecas necessárias, incluindo Selenium para automação web, Pandas para manipulação de dados e Time para pausas na execução.

2. **Leitura da Base de Dados**: Ele lê os dados da planilha Excel que contém os produtos a serem pesquisados, bem como informações sobre termos banidos, preço mínimo e máximo.

3. **Funções de Verificação**: São definidas duas funções para verificar se um produto contém termos banidos ou se possui todos os termos requeridos.

4. **Funções de Busca**: Duas funções são definidas para buscar produtos nos sites Google Shopping e Buscapé, filtrando-os com base nos critérios especificados.

5. **Iteração sobre os Produtos**: O programa itera sobre os produtos da base de dados, realizando buscas e salvando os resultados filtrados em um novo DataFrame.

6. **Exportação de Resultados**: Os resultados filtrados são exportados para um novo arquivo Excel chamado "Ofertas.xlsx".

7. **Envio por E-mail**: Se houver produtos encontrados, o programa utiliza a biblioteca win32com para enviar um e-mail com os resultados para um endereço especificado.

8. **Encerramento do Navegador**: Após a conclusão das operações, o navegador é fechado.

Em resumo, este programa automatiza o processo de pesquisa e filtragem de produtos em sites de compras online, permitindo economizar tempo e esforço ao usuário.

---
