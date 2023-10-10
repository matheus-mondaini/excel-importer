# Excel Importer - Xlsx to SQL Server data converter

Excel Importer is a versatile C# Windows Forms tool that simplifies data import from Excel spreadsheets into SQL Server databases. This project was created as a challenge to handle data import from multiple worksheets within Excel files and seamlessly insert it into the user's SQL Server database. With Excel Importer, users can select specific worksheets, ensuring flexibility and precision in data handling. Whether you're dealing with complex financial data or simple lists, Excel Importer streamlines the process, making data migration a breeze. Try it now to optimize your Excel-to-SQL data import workflow!
## Requirements

Before you begin, make sure you meet the following requirements:

- Be connected to a SQL Server, which contains a database called `Desafio_Planilha` for importing.
- Be an authorized user in this domain.

## How to Use

Follow these steps to import data from the spreadsheet into the database:

1. Clone this repository to your local environment:

   ```bash
   git clone https://github.com/your-username/your-project.git
   ```

2. Set the SQL Connection informations in the connectionString.txt file, replacing the respective fields with your available SQL connection, following the [Requirements].

3. Run the `DesafioImportaExcel.exe` shortcut, located in the main project folder, and follow the instructions to select the spreadsheet and the type of data to import (Cliente or Debitos).

4. The application will import the data from the spreadsheet into a database on the server, if desired.

5. The data will be available in the database for queries and analysis.

- Note: This project is currently for internal use and requires specific SQL Server configuration. Ensure that your environment meets these requirements before proceeding.

## References:

### References and discussions that helped in choosing the library for .xlsx file reading:

- [pt.stackoverflow.com - Qual a melhor forma de fazer a leitura de um arquivo .xls?](https://pt.stackoverflow.com/questions/15590/qual-a-melhor-forma-de-fazer-a-leitura-de-um-arquivo-xls)

- [pt.stackoverflow.com - Importar Excel para SQL Server (C#)](https://pt.stackoverflow.com/questions/121767/importar-excel-para-sql-server-c)

- [stackoverflow.com - What are the differences between the EPPlus and ClosedXML libraries for working with Excel files?](https://stackoverflow.com/questions/10501528/what-are-the-differences-between-the-epplus-and-closedxml-libraries-for-working)

- I also accessed the websites and information of other libraries mentioned in the sources, but since I didn't use them in the code, there is no need to provide links to all compared libraries.

### To choose the library, I used the following links to better understand its operation:

- [EPPlus Features](https://epplussoftware.com/pt/Developers/Features)

- [EPPlus API Documentation](https://epplussoftware.com/docs/5.8/api/OfficeOpenXml.html)

- [EPPlus GitHub Repository](https://github.com/JanKallman/EPPlus)

- [EPPlus on GitHub](https://github.com/EPPlusSoftware/EPPlus)

### Other:

- [ConnectionStrings.com](https://www.connectionstrings.com/)

- [EPPlus License Exception](https://www.epplussoftware.com/en/Developers/LicenseException)