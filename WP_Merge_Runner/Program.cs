using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WordPerfect;
using System.Runtime.InteropServices;


string datTemplatePath = "wp_dat_template.dat";
if (! File.Exists(datTemplatePath))
{
    return ErrorOut("wp_dat_template.dat was not found in the program's root directory");
}

string tempDatPath = $@"{Path.GetFullPath(".")}\temp\dat_{DateTime.Now.ToString("MMddyyhhmmss")}.dat";
if (! Directory.Exists("temp"))
{
    Directory.CreateDirectory("temp");
}

if (args.Length != 3)
{
    Console.WriteLine("Invalid command line arguments.");
    Console.WriteLine($"Expected number of arguments: 3.  Actual number of arguments: {args.Length}");
    Console.WriteLine("Expected arguments:");
    Console.WriteLine("1. Path to WordPerfect template");
    Console.WriteLine("2. Path to new-line-delimited text file");
    Console.WriteLine("3. Path where the output file should be placed");
    return ErrorOut("");
}

string templatePath = args[0];
if (!File.Exists(templatePath))
{
    return ErrorOut($"No word perfect template exists at '{templatePath}'");
}

string dataPath = args[1];
if (!File.Exists(dataPath))
{
    return ErrorOut($"No data file exists at '{templatePath}'");
}

string outputPath = args[2];
if (File.Exists(outputPath))
{
    return ErrorOut($"A file already exists at '{outputPath}'");
}

string directoryPortionOfOutputPath = Path.GetDirectoryName(outputPath);
if (! Directory.Exists(directoryPortionOfOutputPath))
{
    return ErrorOut($"The output directory does not exist '{outputPath}'");
}

StreamReader reader = new StreamReader(dataPath);
string dataText = reader.ReadToEnd();
string[] inputDataLines = dataText.Split(Environment.NewLine);

string[] outputDataLines = new string[1000];

for (int i = 0; i < inputDataLines.Length; i++)
{
    outputDataLines[i] = inputDataLines[i];
}

for (int i = (inputDataLines.Length - 1); i < 1000; i++)
{
    outputDataLines[i] = "";
}

File.Copy(datTemplatePath, tempDatPath);

PerfectScript perfectScript = new PerfectScript();

perfectScript.FileOpen(tempDatPath);
perfectScript.PosLineEnd();
perfectScript.MergeEndRecord();

foreach (string mergeField in outputDataLines)
{
    perfectScript.KeyType(mergeField);
    perfectScript.MergeEndField();
}

perfectScript.SaveAll();
perfectScript.Close();

perfectScript.MergeRun(
    _MergeRun_FormFileType_enum.FormFile_MergeRun_FormFileType,
    templatePath,
    _MergeRun_DataFileType_enum.DataFile_MergeRun_DataFileType,
    tempDatPath,
    _MergeRun_OutputFileType_enum.ToFile_MergeRun_OutputFileType,
    outputPath
);

Marshal.ReleaseComObject(perfectScript);

return 1;

int ErrorOut(string message)
{
    return -1;
}