using BroggiSoft.OfficeExport; // Import the BroggiSoft.OfficeExport library

namespace OfficeExportTestProject
{
    class Program
    {
        static void Main(string[] args)
        {
            // Define a model with properties that correspond to the placeholders in your Word document
            var model = new
            {
                name = "John Doe",
                address = "123 Main Street",
                city = "Anytown",
                publishingCity = "PublishingCity",
                publishingDate = DateTime.Now.ToString("d"),
                subject = "Test Subject",
                lines = new List<dynamic>
                {
                    new { year = "2023", sales = "10000", netIncome = "5000" },
                    new { year = "2024", sales = "12000", netIncome = "6000" }
                },
                logo = Convert.ToBase64String(File.ReadAllBytes(@".\Resources\logo.png")) // replace with the actual path to your logo file in the Resources folder
            };

            // Use the ExportFromObject method of the OfficeExport class to replace the placeholders in the template with the values from the model
            string sourcePathFile = @".\Resources\WordSource.docx";
            string destinationPathFile = @".\WordResult.docx";
            OfficeExport.ExportFromObject(sourcePathFile, destinationPathFile, model);

            Console.WriteLine("Export successful!");
        }
    }
}
