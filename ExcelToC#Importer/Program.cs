using ExcelImporterMethod;
using Persons.Models;


namespace ExcelToC_Importer
{
    internal class Program
    {
        static void Main(string[] args)
        {

            string filePath = @"C:\Study\print.xlsx"; 

            List<Person> dataList = ExcelImporter.ReadExcelDataUsingPropertyNames<Person>(filePath);

            Console.WriteLine("Data from Excel file:");
            Console.WriteLine("------------------------------------------------------------------------------------------------");
            Console.WriteLine("| Name                 | Age   | Email               | Title          |  Experience  | Idea    |");
            Console.WriteLine("------------------------------------------------------------------------------------------------");

            foreach (var person in dataList)
            {
                Console.WriteLine($"| {person.Name,-20} | {person.Age,5} | {person.Email,-19} | {person.Title,-14} | {person.Experience,12} | {person.idea, 7} |");
            }

            Console.WriteLine("-----------------------------------------------------------------------------------------------");
        }
    }
}
