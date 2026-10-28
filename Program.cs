using Microsoft.Playwright;
using ClosedXML.Excel;
using System.Globalization;
using System.Text.RegularExpressions;

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
        string stationId = Console.ReadLine()?.Trim() ?? "";
        Console.Write("Data de início (ano-mês-dia): ");
        bool inicioValido = DateTime.TryParse(Console.ReadLine(), out DateTime dataInicio);
        Console.Write("Data de fim (ano-mês-dia): ");
        bool fimValido = DateTime.TryParse(Console.ReadLine(), out DateTime dataFim);
        if (stationId == "" || !inicioValido || !fimValido)
        {
            Console.WriteLine("Código da estação ou data inválida.");
            return;
        }
        Console.Clear();
        var relatorio = new List<DadosClimaticos>();

        using var playwright = await Playwright.CreateAsync();
        await using var browser = await playwright.Chromium.LaunchAsync(new BrowserTypeLaunchOptions {Headless = true});
        var context = await browser.NewContextAsync();

        while(dataInicio <= dataFim)
        {
            string dataFormatada = dataInicio.ToString("yyyy-MM-dd");
            Console.WriteLine($"\nProcessando dia: {dataFormatada}...");

            try
            {
                var linhas = await ColetarDia(context, stationId, dataFormatada);
                if (linhas.Length == 0)
                    Console.WriteLine($"Nenhum registro no dia {dataFormatada} (a estação pode ter ficado sem enviar dados).");

                foreach (var celulas in linhas)
                {
                    // Colunas da tabela: 1 Hora, 2 Temperatura, 3 Ponto de orvalho, 4 Umidade
                    relatorio.Add(new DadosClimaticos
                    {
                        Data = dataFormatada,
                        Hora = celulas[0],
                        Temperatura = celulas[1],
                        Umidade = celulas[3]
                    });
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Erro no dia {dataFormatada}: {ex.Message}");
            }

            dataInicio = dataInicio.AddDays(1);
            await Task.Delay(1000);

            GerarExcel(relatorio, stationId);
        }
        Console.WriteLine("Exportação finalizada. Pressione qualquer tecla para sair...");
        Console.ReadKey();
    }

    static async Task<string[][]> ColetarDia(IBrowserContext context, string stationId, string data)
    {
        var page = await context.NewPageAsync();
        try
        {
            // O site ignora a data da URL: o período exibido vem do localStorage, o mesmo que o seletor "Custom" da página grava
            await page.AddInitScriptAsync($$"""
                try {
                    localStorage.setItem('wu-date-range', 'custom');
                    localStorage.setItem('wu-date-custom-start', '{{data}}');
                    localStorage.setItem('wu-date-custom-end', '{{data}}');
                } catch {}
                """);

            await page.GotoAsync($"https://www.wunderground.com/dashboard/pws/{stationId}", new PageGotoOptions { Timeout = 60000 });

            // Espera os componentes carregarem e o seletor de datas aplicar o período (ele preenche o rótulo da data ao terminar).
            // Antes disso o site passa por uma verificação anti-bot que recarrega a página, por isso o timeout longo
            var periodo = await page.WaitForFunctionAsync(@"() => {
                const datas = document.querySelector('pws-history-dates');
                const pronto = customElements.get('dashboard-body-view') && customElements.get('pws-history-table') && datas?.querySelector('.date-label')?.textContent;
                return pronto ? datas.startDate : false;
            }", null, new PageWaitForFunctionOptions { Timeout = 60000 });
            string dataExibida = await periodo.JsonValueAsync<string>();
            if (dataExibida != data)
                throw new Exception($"o site exibiu o dia {dataExibida} em vez de {data}");

            // A tabela só é preenchida depois de trocar a visão "Graph" por "Table"
            await page.Locator("button.toggle-btn[data-view='table']").DispatchEventAsync("click");

            var seletorLinhas = "pws-history-table table tbody tr";
            try
            {
                await page.Locator(seletorLinhas).First.WaitForAsync(new LocatorWaitForOptions { Timeout = 20000 });
            }
            catch (TimeoutException)
            {
                // A tabela fica vazia quando a estação não enviou dados no dia
                return Array.Empty<string[]>();
            }

            return await page.Locator(seletorLinhas).EvaluateAllAsync<string[][]>(
                "linhas => linhas.map(linha => Array.from(linha.cells, celula => celula.innerText.trim()))");
        }
        finally
        {
            await page.CloseAsync();
        }
    }

    static void GerarExcel(List<DadosClimaticos> dados, string stationId)
    {
        string nomePasta = "Relatorio";
        string nomeArquivo = $"DadosClimaticos_{stationId}.xlsx";

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
                
                var temperatura = ExtrairNumero(item.Temperatura);
                if (temperatura.HasValue)
                {
                    // O site exibe °C ou °F conforme a preferência de unidade; só converte quando vier em °F
                    if (item.Temperatura.Contains("°F"))
                        temperatura = Math.Round((temperatura.Value - 32) / 1.8, 1, MidpointRounding.AwayFromZero);
                    worksheet.Cell(linhaAtual, 3).Value = temperatura.Value;
                }
                else
                    worksheet.Cell(linhaAtual, 3).Value = "N/A";

                var umidade = ExtrairNumero(item.Umidade);
                if (umidade.HasValue)
                    worksheet.Cell(linhaAtual, 4).Value = umidade.Value;
                else
                    worksheet.Cell(linhaAtual, 4).Value = "N/A";

                linhaAtual++;
            }

            worksheet.Columns().AdjustToContents();

            workbook.SaveAs(caminhoCompleto);
        }
    }

    // Os valores vêm com a unidade junto e ponto decimal (ex.: "20.4°C", "57%"); célula sem dado vem como "—"
    static double? ExtrairNumero(string texto)
    {
        var numero = Regex.Match(texto, @"-?\d+(\.\d+)?");
        return numero.Success ? double.Parse(numero.Value, CultureInfo.InvariantCulture) : null;
    }
}
