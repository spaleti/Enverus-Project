using System;
using System.Configuration;
using System.Net;
using System.IO;
using System.Globalization;
using System.Linq;

using System.Net.Http;

using GroupDocs.Conversion;
using GroupDocs.Conversion.FileTypes;
using GroupDocs.Conversion.Options.Convert;

//downloads file but it wont open it - 0bytes
/*try
{
    using (var client = new WebClient())
    {
        //.DownloadFile vrti u nedogled
        client.DownloadFileAsync(new Uri("https://bakerhughesrigcount.gcs-web.com/intl-rig-count?c=79687&p=irol-rigcountsintl"), "Worldwide Rig Count Jan 2023.xlsx");
    }
    Console.WriteLine("Download complete.");
}
catch (WebException ex)
{
    Console.WriteLine("An error occurred while downloading the file: " + ex.Message);
}
catch (TaskCanceledException ex)
{
    Console.WriteLine("The download was cancelled: " + ex.Message);
}*/


//the rest works if there is xlsx file

string currentDirectory = Directory.GetCurrentDirectory();
string excelFilePath = Path.Combine(currentDirectory, "Worldwide Rig Count Jan 2023.xlsx");
string csvFilePath = Path.Combine(currentDirectory, "Worldwide Rig Count Jan 2023.csv");

// stop program if file does not exist
if (!File.Exists(excelFilePath))
{
    Console.WriteLine("Need to put file there manually");
    return;
}

using (Converter converter = new Converter(excelFilePath))
{
    SpreadsheetConvertOptions options = new SpreadsheetConvertOptions
    {
        PageNumber = 2,
        PagesCount = 1,
        Format = SpreadsheetFileType.Csv // Specify the conversion format
    };
    converter.Convert(csvFilePath, options);
}

List<string> linesToWrite = new List<string>(); // list for 2022 i 2021 year

using (var streamReader = new StreamReader(csvFilePath))
{
    //skip header + new lines + 2023 year
    for (int i = 0; i < 21; i++)
    {
        streamReader.ReadLine();
    }
    //write in list 2022 and 2021 years
    for (int i = 0; i < 30; i++)
    {
        linesToWrite.Add(streamReader.ReadLine());
    }
}
//delete all content of csv file
System.IO.File.WriteAllText(csvFilePath, string.Empty);

//write in content in csv for 2022 and 2021 year
using (StreamWriter writer = new StreamWriter(csvFilePath))
{
    foreach (string line in linesToWrite)
    {
        writer.WriteLine(line);
    }
}