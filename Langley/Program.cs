using Langley.Services;
using System.CommandLine;

namespace Langley
{
    internal class Program
    {
        static async Task Main(string[] args)
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            var rootCommand = new RootCommand("langley");

            var creepCommand = new Command("creep",
                """
                Define a set of included directories to search for words coming from a specific column on an Excel file. 
                This command will create a new Excel file that will list the words in one column and to what file path the
                word can be found.
                """);

            var excelInColumnOption = new Option<string>("--excel-in-column", "Column Letter the word list appear in the Excel.");
            excelInColumnOption.Arity = ArgumentArity.ExactlyOne;
            excelInColumnOption.IsRequired = true;

            var excelInColumnRowOption = new Option<int>("--excel-in-row", "Row Number word list starts to appear in the Excel.");
            excelInColumnRowOption.Arity = ArgumentArity.ExactlyOne;
            excelInColumnRowOption.IsRequired = true;

            var excelInOption = new Option<string>("--excel-in", "The complete file path to Excel Input.");
            excelInOption.Arity = ArgumentArity.ExactlyOne;
            excelInOption.IsRequired = true;

            var excelOutOption = new Option<string>("--excel-out", "The complete file path to Excel Output.");
            excelOutOption.Arity = ArgumentArity.ExactlyOne;
            excelOutOption.IsRequired = true;

            var creepCommandIncludeDirOption = new Option<List<string>>("--include-dir", "Includes a directory path to the search path.");
            creepCommandIncludeDirOption.Arity = ArgumentArity.OneOrMore;
            creepCommandIncludeDirOption.IsRequired = true;

            creepCommand.AddOption(excelInColumnOption);
            creepCommand.AddOption(excelInColumnRowOption);
            creepCommand.AddOption(excelInOption);
            creepCommand.AddOption(excelOutOption);
            creepCommand.AddOption(creepCommandIncludeDirOption);
            
            creepCommand.SetHandler(async (
                excelInColumnValue,
                excelInRowValue,
                excelInValue,
                excelOutValue,
                creepCommandIncludeValue) =>
            {
                Console.WriteLine($"excelInColumnValue {excelInColumnValue}");
                Console.WriteLine($"excelInColumnValue {excelInRowValue}");
                Console.WriteLine($"excelInColumnValue {excelInValue}");
                Console.WriteLine($"excelInColumnValue {excelOutValue}");
                foreach (var value in creepCommandIncludeValue)
                {
                    Console.WriteLine($"creepCommandIncludeValue {value}");
                }

                var creepService = new CreepService()
                {
                    IncludedDirectories = creepCommandIncludeValue,
                    ExcelInputRow = excelInRowValue,
                    ExcelInputColumn = excelInColumnValue,
                    Output = excelOutValue,
                    ExcelInput = excelInValue,
                };

                creepService.Run();
            },
            excelInColumnOption,
            excelInColumnRowOption,
            excelInOption,
            excelOutOption,
            creepCommandIncludeDirOption
            );

            rootCommand.Add(creepCommand);
            rootCommand.SetHandler(() =>
            {
                Console.WriteLine("Hello world!");
            });

            await rootCommand.InvokeAsync(args);
        }
    }
}
