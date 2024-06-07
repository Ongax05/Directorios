using System;
using System.IO;
using OfficeOpenXml;

class Program
{
    static void Main()
    {
        string rootDirectory = @"C:\Users\sopor\OneDrive\Escritorio\REPORTES"; // Especifica la ruta del directorio raíz

        try
        {
            string filePath = @"C:\Users\sopor\OneDrive\Escritorio\PORTAFOLIOS DE REPORTES.xlsx";

            // Asegúrate de que EPPlus tenga licencia
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using var package = new ExcelPackage(new FileInfo(filePath));
            var worksheet = package.Workbook.Worksheets[0];
            int rowCount = worksheet.Dimension?.Rows ?? 0; // Utilizando operador de nulabilidad

            string[] files = Directory.GetFiles(rootDirectory, "*.*", SearchOption.AllDirectories);

            foreach (var file in files)
            {
                if (worksheet.Dimension != null && rowCount >= worksheet.Dimension.End.Row)
                {
                    int newRow = rowCount + 1;
                    worksheet.InsertRow(newRow, 1);

                    var fileExtension = Path.GetExtension(file);
                    if (!fileExtension.Equals(".txt"))
                    {
                        var fileRelativePath = Path.GetRelativePath(rootDirectory, file);
                        var fileName = Path.GetFileName(file);

                        worksheet.Cells[newRow, 4].Value = fileName; // Asigna nombre de archivo a la columna 4
                        worksheet.Cells[newRow, 5].Value = fileRelativePath; // Asigna ruta relativa a la columna 5
                        rowCount++; // Incrementa el contador de filas
                    }
                }
            }
            package.Save();
            Console.WriteLine("Filas añadidas correctamente.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Se produjo un error: {ex.Message}");
        }
    }
}
