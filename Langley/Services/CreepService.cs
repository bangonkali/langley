using ExcelDataReader;
using System.Text;

namespace Langley.Services
{
    public class CreepService
    {
        public List<string> ExcludeDirs { get; set; } = [
            ".git",
            ".vs",
            "obj",
            "bin",
            "wwwroot",
            "release",
            "debug",
        ];

        public List<string> ExcludeFileExtensions { get; set; } = [
            ".docx",
            ".gitignore",
            ".ico",
            ".resx",
            ".xlsx",
            ".mp3",
            ".mp4",
            ".mov",
            ".gif",
            ".ttf",
            ".woff",
            ".woff2",
            ".xsd",
            ".png",
            ".bmp",
            ".jpeg",
            ".jpg",
            ".zip",
            ".rar",
            ".7z",
            ".db",
            ".dll",
            ".obj",
            ".exe",
            ".svg",
            ".so",
            ".pdf"];

        /// <summary>
        /// The output findings are stored here.
        /// </summary>
        public CreepFindings Findings { get; set; } = new CreepFindings();

        /// <summary>
        /// The source of words to search recursively from directory files.
        /// </summary>
        public List<string> WordList { get; set; } = new List<string>();

        /// <summary>
        /// The directory paths to search words from.
        /// </summary>
        public List<string> IncludedDirectories { get; set; } = new List<string>();

        public int ExcelInputRow { get; set; } = 1;

        public string ExcelInputColumn { get; set; } = "A";

        /// <summary>
        /// The path to the csv to write the output to.
        /// </summary>
        public string Output { get; set; } = "";

        public string ExcelInput { get; set; } = "";

        public void Run()
        {

            LoadWords();

            Recurse();

            WriteFindingsToFile();
        }

        /// <summary>
        /// Writes the findings to the output path defined by Output property of this class. The format
        /// of the output is as follows:
        /// 
        /// ```
        /// Word, Line Number, Line Row, Path
        /// AAA, 34, 2, C:\some\file.a
        /// AAA, 34, 9, C:\some\file.a
        /// AAB, 39, 1, C:\some\file-2.a
        /// ```
        /// </summary>
        private void WriteFindingsToFile()
        {
            // Ensure the directory exists before writing the file.
            var directory = Path.GetDirectoryName(Output);
            if (directory != null && !Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            using var writer = new StreamWriter(Output, false, Encoding.UTF8);
            writer.WriteLine("Word, Line Number, Line Row, File Name, File Extension, File Full Path");

            foreach (var wordFinding in Findings.WordFindings)
            {
                foreach (var finding in wordFinding.Findings)
                {
                    var line = new StringBuilder();
                    line.Append(wordFinding.Word); // Word
                    line.Append(", ");
                    line.Append(finding.Line); // Line Number
                    line.Append(", ");
                    line.Append(finding.LineColumn); // Line Row
                    line.Append(", ");
                    line.Append($"{finding.File.Name}"); // File Name
                    line.Append(", ");
                    line.Append($"{finding.File.Extension}"); // Extension Name
                    line.Append(", ");
                    line.Append($"\"{finding.File.FullName}\""); // Path, enclosed in quotes to handle commas in paths
                    writer.WriteLine(line.ToString());
                }
            }
        }

        /// <summary>
        /// Recursively find all files in all directories from IncludedDirectories. For each file, find 
        /// all occurences of any word in memory from WordList. If word is found, log finding in to memory
        /// on Findings. Each word should only appear on on Findings but multiple instances of a word can
        /// be found on a file or on multiple files based on CreepFindingEntry within Findings.
        /// </summary>
        private void Recurse()
        {
            foreach (var directory in IncludedDirectories)
            {
                var dir = new DirectoryInfo(directory);
                RecurseDirectory(dir);
            }
        }

        private void RecurseDirectory(DirectoryInfo directory)
        {
            try
            {
                // List all files in the current directory
                foreach (FileInfo file in directory.GetFiles())
                {
                    if (!ExcludeFileExtensions.Any(x => x.Equals(file.Extension, StringComparison.CurrentCultureIgnoreCase)))
                    {
                        FindWordsInFile(file);
                    }
                }

                // Recursively call this method for each subdirectory
                foreach (DirectoryInfo subDirectory in directory.GetDirectories())
                {

                    if (!ExcludeDirs.Any(x => x.Equals(subDirectory.Name, StringComparison.CurrentCultureIgnoreCase)))
                    {
                        RecurseDirectory(subDirectory);
                    }
                }
            }
            catch (UnauthorizedAccessException e)
            {
                Console.WriteLine("Access denied to directory: " + directory.FullName + " - " + e.Message);
            }
            catch (System.IO.DirectoryNotFoundException e)
            {
                Console.WriteLine("Directory not found: " + directory.FullName + " - " + e.Message);
            }
        }

        private void FindWordsInFile(FileInfo file)
        {
            Console.WriteLine($"Findings words from file {file.FullName}");

            int lineNumber = 0;
            foreach (var line in File.ReadLines(file.FullName, Encoding.UTF8))
            {
                lineNumber++; // Increment line number as we go through each line

                foreach (var word in WordList)
                {
                    int columnIndex = line.IndexOf(word, StringComparison.InvariantCulture); // case sensitive search
                    while (columnIndex != -1) // Check for multiple occurrences of the word in the same line
                    {
                        LogWordFinding(word, new CreepFindingEntry()
                        {
                            File = file,
                            Line = lineNumber,
                            LineColumn = columnIndex + 1,
                        });

                        //LogWordFinding(word, file.FullName, lineNumber, columnIndex + 1); // Adjusting columnIndex to be 1-based

                        // Find next occurrence, if any, by starting search right after the current found word
                        columnIndex = line.IndexOf(word, columnIndex + word.Length, StringComparison.OrdinalIgnoreCase);
                    }
                }
            }
        }

        private void LogWordFinding(string word, CreepFindingEntry findingEntry)
        {
            Console.WriteLine($"Findings words from file {findingEntry.File.FullName} - Found {word} on {findingEntry.Line}:{findingEntry.LineColumn}");
            var wordFinding = Findings.WordFindings.FirstOrDefault(wf => wf.Word.Equals(word));
            if (wordFinding == null)
            {
                wordFinding = new CreepWordFindings
                {
                    Word = word,
                    Findings = new List<CreepFindingEntry>()
                };
                Findings.WordFindings.Add(wordFinding);
            }

            wordFinding.Findings.Add(findingEntry);
        }


        /// <summary>
        /// Load all words from Excel sheet in to Memory
        /// </summary>
        private void LoadWords()
        {
            var stringType = typeof(string);
            var count = 0;

            using var stream = File.Open(ExcelInput, FileMode.Open, FileAccess.Read);
            using var reader = ExcelReaderFactory.CreateReader(stream);
            do
            {
                while (reader.Read())
                {
                    count++;

                    // Offset rows with ExcelInputRow
                    if (count <= ExcelInputRow)
                    {
                        continue;
                    }

                    var valueType = reader.GetFieldType(0);
                    if (valueType != null && valueType == stringType)
                    {
                        var value = reader.GetString(0);
                        Console.WriteLine($"{count} Found {value}");
                        WordList.Add(value);
                    }
                }
            } while (reader.NextResult());
        }
    }

    public class CreepFindings
    {
        /// <summary>
        /// The list of Unique Words and Findings for each word.
        /// </summary>
        public List<CreepWordFindings> WordFindings { get; set; } = [];
    }

    public class CreepWordFindings
    {
        /// <summary>
        /// The word.
        /// </summary>
        public string Word { get; set; }

        /// <summary>
        /// The instances of findings for a specific word.
        /// </summary>
        public List<CreepFindingEntry> Findings { get; set; } = [];
    }

    public class CreepFindingEntry
    {
        /// <summary>
        /// The path of the file in which the word is found at.
        /// </summary>
        public FileInfo File { get; set; }

        /// <summary>
        /// The line or file row on which the word is found at.
        /// </summary>
        public int Line { get; set; }

        /// <summary>
        /// The offset from start of file which the word is found at.
        /// </summary>
        public int LineColumn { get; set; }
    }
}
