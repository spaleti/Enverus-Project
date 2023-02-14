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

//OVO JE NAJBLIZE DO SADA - skine fajl ali nece da ga otvori kaze 0bytes
/*try
{
    using (var client = new WebClient())
    {
        //DownloadFile vrti u nedogled
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


//OSTATAK RADI AKO POSTOJI XLSX dokument

string currentDirectory = Directory.GetCurrentDirectory();
string excelFilePath = Path.Combine(currentDirectory, "Worldwide Rig Count Jan 2023.xlsx");
string csvFilePath = Path.Combine(currentDirectory, "Worldwide Rig Count Jan 2023.csv");

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

List<string> linesToWrite = new List<string>(); // lista za 2022 i 2021

using (var streamReader = new StreamReader(csvFilePath))
{
    //preskociti pocetak i 2023 god
    for (int i = 0; i < 21; i++)
    {
        streamReader.ReadLine();
    }
    //upisati u listu 2022 i 2021
    for (int i = 0; i < 30; i++)
    {
        linesToWrite.Add(streamReader.ReadLine());
    }
}
//brisem sadrzaj celog csv fajla
System.IO.File.WriteAllText(csvFilePath, string.Empty);

//upisujem 2022 2021 
using (StreamWriter writer = new StreamWriter(csvFilePath))
{
    foreach (string line in linesToWrite)
    {
        writer.WriteLine(line);
    }
}