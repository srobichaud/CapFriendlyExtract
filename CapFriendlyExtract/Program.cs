using ClosedXML.Excel;
using HtmlAgilityPack;

string param = Environment.GetCommandLineArgs()[1];

string url = param;

decimal resultsPerPage = 50;

HtmlWeb web = new HtmlWeb();
var htmlDoc = web.Load(url);

var node = htmlDoc.DocumentNode.SelectSingleNode("//div[@id='browse_a']/h5");
string results = node.InnerText.Replace("RESULTS (", "");
results = results.Replace(")", "");
results = results.Replace(",", "");

decimal totalResults = Convert.ToDecimal(results);

decimal t = totalResults / resultsPerPage;
decimal amountOfPages = Math.Ceiling(t);

using (var workbook = new XLWorkbook())
{
    var worksheet = workbook.Worksheets.Add("data");
    string symbol = "?";
    if (url.Contains("?"))
    {
        symbol = "&";
    }
    for (int i = 0; i < amountOfPages; i++)
    {
        int page = i + 1;
        string urlPage = url + symbol + "pg=" + page;
        var htmlDoc2 = web.Load(urlPage);
        Console.WriteLine(urlPage);

        var rows = htmlDoc2.DocumentNode.SelectNodes("//table[@id='brwt']/tbody/tr");
        int k = (int)(i * resultsPerPage);
        foreach (var tr in rows)
        {
            int j = 0;
            int row = (int)(k + 1);

            foreach (var td in tr.ChildNodes)
            {
                int column = j + 1;
                worksheet.Cell(row, column).Value = td.InnerText;
                j++;
            }
            k++;
        }
    }
    workbook.SaveAs("cap_friendly_" + DateTime.Now.Ticks + ".xlsx");
}

Console.WriteLine("Complété...");
Console.ReadLine();