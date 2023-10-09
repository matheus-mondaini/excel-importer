# Excel Importer - Conversor de dados de arquivos Xlsx para SQL Server

O Excel Importer é uma versátil ferramenta em C# com interface gráfica em Windows Forms que simplifica a importação de dados de planilhas Excel para bancos de dados SQL Server. Este projeto foi criado como um desafio para lidar com a importação de dados de várias planilhas em arquivos Excel e inseri-los de forma transparente no banco de dados SQL Server do usuário. Com o Excel Importer, os usuários podem selecionar planilhas específicas, garantindo flexibilidade e precisão no manuseio dos dados. Seja lidando com dados financeiros complexos ou listas simples, o Excel Importer simplifica o processo, tornando a migração de dados simples e eficaz. Experimente agora para otimizar o fluxo de trabalho de importação de dados do Excel para o SQL Server!

## Requisitos

Antes de começar, certifique-se de que atende aos seguintes requisitos:

- Estar conectado a um servidor SQL Server com o nome `Gemini\SQL2019`, que contenha a tabela `Desafio_Planilha`.
- Ser um usuário permitido neste domínio.

## Como Usar

Siga estas etapas para importar dados da planilha para o banco de dados:

1. Clone este repositório para o seu ambiente local:

   ```bash
   git clone https://github.com/seu-usuario/seu-projeto.git

2. Execute o atalho do `DesafioImportaExcel.exe`, que se encontra logo na pasta principal do projeto, e siga as instruções para selecionar a planilha e o tipo de dados a serem importados (Cliente ou Debitos).

3. O aplicativo importará os dados da planilha para o banco de dados no servidor Gemini\SQL2019, caso assim desejar.

4. Os dados ficarão disponíveis no banco de dados para consultas e análises.

- Observação: Este projeto é atualmente para uso interno e requer uma configuração específica do servidor SQL Server. Certifique-se de que seu ambiente atenda a esses requisitos antes de prosseguir.

## Referências:

### Referências e discussões que auxiliaram na escolha da biblioteca para leitura de arquivos .xlsx:

- [pt.stackoverflow.com - Qual a melhor forma de fazer a leitura de um arquivo .xls?](https://pt.stackoverflow.com/questions/15590/qual-a-melhor-forma-de-fazer-a-leitura-de-um-arquivo-xls)

- [pt.stackoverflow.com - Importar Excel para SQL Server (C#)](https://pt.stackoverflow.com/questions/121767/importar-excel-para-sql-server-c)

- [stackoverflow.com - Quais são as diferenças entre as bibliotecas EPPlus e ClosedXML para trabalhar com arquivos Excel?](https://stackoverflow.com/questions/10501528/what-are-the-differences-between-the-epplus-and-closedxml-libraries-for-working)

- Também acessei os sites e informações de outras bibliotecas mencionadas nas fontes, mas como não as utilizei no código, não há necessidade de fornecer links para todas as bibliotecas comparadas.

### Para escolher a biblioteca, utilizei os seguintes links para entender melhor seu funcionamento:

- [Recursos do EPPlus](https://epplussoftware.com/pt/Developers/Features)

- [Documentação da API EPPlus](https://epplussoftware.com/docs/5.8/api/OfficeOpenXml.html)

- [Repositório EPPlus no GitHub](https://github.com/JanKallman/EPPlus)

- [EPPlus no GitHub](https://github.com/EPPlusSoftware/EPPlus)

### Outros:

- [ConnectionStrings.com](https://www.connectionstrings.com/)

- [Exceção de Licença do EPPlus](https://www.epplussoftware.com/en/Developers/LicenseException)