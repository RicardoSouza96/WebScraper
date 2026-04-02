using Microsoft.Playwright;
using ClosedXML.Excel;
using System.Globalization;

public class DadosClimaticos
{
    public string Data {get; set;}
    public string Hora { get; set; }
    public string Temperatura { get; set; }
    public string Umidade { get; set; }
}

class Program
{
    public static async Task Main()
    {
        Console.Write("Código da estação: ");
        string stationId = Console.ReadLine();
        Console.Write("Data de início (ano-mês-dia): ");
        DateTime.TryParse(Console.ReadLine(), out DateTime dataInicio);
        Console.Write("Data de fim (ano-mês-dia): ");
        DateTime.TryParse(Console.ReadLine(), out DateTime dataFim);
        Console.Clear();
        var relatorio = new List<DadosClimaticos>();

        using var playwright = await Playwright.CreateAsync();
        await using var browser = await playwright.Chromium.LaunchAsync(new BrowserTypeLaunchOptions {Headless = true});
        var page = await browser.NewPageAsync();

        while(dataInicio <= dataFim)
        {
            string dataFormatada = dataInicio.ToString("yyyy-MM-dd");
            string url = $"https://www.wunderground.com/dashboard/pws/{stationId}/table/{dataFormatada}/{dataFormatada}/daily";
            Console.WriteLine($"\nProcessando dia: {dataFormatada}...");

            try
            {
                await page.GotoAsync(url, new PageGotoOptions { Timeout = 60000 });

                var seletorTabela = "table.history-table.desktop-table";
                await page.WaitForSelectorAsync(seletorTabela, new PageWaitForSelectorOptions {Timeout = 10000});

                var linhas = await page.Locator($"{seletorTabela} tbody tr").AllAsync();
                foreach (var linha in linhas)
                {
                    if (await linha.Locator("td").CountAsync() > 0)
                    {
                        var hora = await linha.Locator("td:nth-child(1)").InnerTextAsync();
                        
                        var tempLocator = linha.Locator("td:nth-child(2) .wu-value");
                        var temperatura = await tempLocator.CountAsync() > 0 ? await tempLocator.First.InnerTextAsync() : "N/A";

                        var umidLocator = linha.Locator("td:nth-child(4) .wu-value");
                        var umidade = await umidLocator.CountAsync() > 0 ? await umidLocator.First.InnerTextAsync() : "N/A";

                        relatorio.Add(new DadosClimaticos
                        {
                            Data = dataFormatada,
                            Hora = hora,
                            Temperatura = temperatura,
                            Umidade = umidade
                        });
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Erro no dia {dataFormatada}: {ex.Message}");
            }

            dataInicio = dataInicio.AddDays(1);
            await Task.Delay(1000);

            GerarExcel(relatorio);
        }
        Console.WriteLine("Exportação finalizada. Pressione qualquer tecla para sair...");
        Console.ReadKey();
    }

    static void GerarExcel(List<DadosClimaticos> dados)
    {
        string nomePasta = "Relatorio";
        string nomeArquivo = "DadosClimaticos.xlsx";

        if(!Directory.Exists(nomePasta))
            Directory.CreateDirectory(nomePasta);
        
        string caminhoCompleto = Path.Combine(nomePasta, nomeArquivo);

        using (var workbook = new XLWorkbook())
        {
            var worksheet = workbook.Worksheets.Add("Historico");

            // Cria o Cabeçalho
            worksheet.Cell(1, 1).Value = "Data";
            worksheet.Cell(1, 2).Value = "Hora";
            worksheet.Cell(1, 3).Value = "Temperatura (°C)";
            worksheet.Cell(1, 4).Value = "Umidade (%)";

            // Deixa o cabeçalho em Negrito
            var rangeCabecalho = worksheet.Range("A1:D1");
            rangeCabecalho.Style.Font.Bold = true;
            rangeCabecalho.Style.Fill.BackgroundColor = XLColor.LightGray;

            // Preenche os dados
            int linhaAtual = 2;
            foreach (var item in dados)
            {
                worksheet.Cell(linhaAtual, 1).Value = item.Data;
                worksheet.Cell(linhaAtual, 2).Value = item.Hora;
                
                if (double.TryParse(item.Temperatura, NumberStyles.Any, CultureInfo.InvariantCulture, out double tempFahrenheit))
                {
                    double tempCelsiusExata = (tempFahrenheit - 32) / 1.8;
                    double tempCelsiusFinal = Math.Round(tempCelsiusExata, 1, MidpointRounding.AwayFromZero);
                    worksheet.Cell(linhaAtual, 3).Value = tempCelsiusFinal;
                }
                else
                    worksheet.Cell(linhaAtual, 3).Value = item.Temperatura;

                if (double.TryParse(item.Umidade, NumberStyles.Any, CultureInfo.InvariantCulture, out double umidNum))
                    worksheet.Cell(linhaAtual, 4).Value = umidNum;
                else
                    worksheet.Cell(linhaAtual, 4).Value = item.Umidade;

                linhaAtual++;
            }

            worksheet.Columns().AdjustToContents();

            workbook.SaveAs(caminhoCompleto);
        }
    }
}
