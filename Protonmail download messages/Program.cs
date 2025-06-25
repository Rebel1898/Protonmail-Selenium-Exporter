using Microsoft.Win32;
using OpenQA.Selenium;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.Support.UI;
using Protonmail_Selenium_Exporter;
using SeleniumExtras.WaitHelpers;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using WebDriverManager;
using WebDriverManager.DriverConfigs.Impl;
using WindowsInput;
using System.Linq;
using System.Text.RegularExpressions;
using System.Globalization;
using System.Diagnostics;

namespace Protonmail_download_messages
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                InputSimulator teclado = new InputSimulator();
                string path = GetDownloadFolderPath();
                int userInput = 0;
                do
                {
                    userInput = DisplayMenu();
                } while (userInput > 5);

                if (userInput == 5)
                    Environment.Exit(0);

                Console.WriteLine("Set username :");
                string username = Console.ReadLine();
                Console.Clear();

                Console.WriteLine("Set password :");
                string password = GetPassword();
                Console.Clear();
                Console.WriteLine("\n Username and Password correctly obtained.");

                path = path + @"\Protonmail\";
                string InboxPath = path + "Inbox";
                string SentPath = path + "Sent";
                string reportPath = path + DateTime.Now.ToString("yyyy-MM-dd_HH-mm") + "report.txt";

                System.IO.Directory.CreateDirectory(InboxPath);
                System.IO.Directory.CreateDirectory(SentPath);

                string firefoxPath = GetFirefoxPathFromRegistry();
                new DriverManager().SetUpDriver(new FirefoxConfig());
                FirefoxOptions options = SetPreferences(path, firefoxPath);

                WebDriver driver = new FirefoxDriver(options);
                Console.Clear();
                Console.WriteLine("Firefox driver Launched\r\n");

                WebDriverWait wait30 = new WebDriverWait(driver, TimeSpan.FromSeconds(30));
                WebDriverWait wait120 = new WebDriverWait(driver, TimeSpan.FromSeconds(120));
                wait30.IgnoreExceptionTypes(typeof(StaleElementReferenceException));
                wait30.IgnoreExceptionTypes(typeof(NoSuchElementException));
                driver.Manage().Window.Maximize();

                //Login
                driver.Navigate().GoToUrl("https://mail.proton.me/u/0/sent");
                wait30.Until(ExpectedConditions.ElementExists(By.Id("username")));
                IWebElement inputUser = driver.FindElement(By.Id("username"));
                wait30.Until(ExpectedConditions.ElementToBeClickable(By.Id("username")));
                inputUser.SendKeys(username);
                wait30.Until(ExpectedConditions.ElementExists(By.Id("password")));
                IWebElement inputPassword = driver.FindElement(By.Id("password"));
                inputPassword.Click();
                inputPassword.Clear();
                inputPassword.SendKeys(password);
                password = "";
                driver.FindElement(By.CssSelector("button.button:nth-child(6)")).Click();

                //Set No reply grouping
                bool Grouping, Layout;
                string originalLanguage;
                SetSettings(false, true, "English", out Grouping, out Layout, out originalLanguage);
                var Lista_Correos = new List<Correo>();
                if (userInput == 3 || userInput == 1)
                {

                    Lista_Correos = DownloadMailInfo("Inbox", true);


                    MoveFiles(path, InboxPath);
                    if (userInput == 1)
                        DisplayandSaveReport(Lista_Correos);

                }
                if (userInput == 3 || userInput == 2)
                {
                    driver.Navigate().GoToUrl("https://mail.proton.me/u/0/sent");
                    var ListaRecibidos = DownloadMailInfo("Sent", true);

                    if (Lista_Correos.Count > 0)
                        Lista_Correos = Lista_Correos.Concat(ListaRecibidos).ToList();
                    else
                        Lista_Correos = ListaRecibidos;

                    MoveFiles(path, SentPath);
                    DisplayandSaveReport(Lista_Correos);


                }
                if (userInput == 4)
                {
                    var ListaCorreosRecibidos = DownloadMailInfo("Inbox", false);
                    driver.Navigate().GoToUrl("https://mail.proton.me/u/0/sent");

                    var ListaCorreosEnviados = DownloadMailInfo("Sent", false);

                    Lista_Correos = ListaCorreosEnviados.Concat(ListaCorreosRecibidos).ToList();
                    DisplayandSaveReport(Lista_Correos);

                }

                SetSettings(Grouping, Layout, originalLanguage, out Grouping, out Layout, out originalLanguage);
                driver.Quit();

                Console.WriteLine("Press any key to continue");
                Console.ReadLine();
                Console.Clear();
                Environment.Exit(0);





                //Metodos
                void MoveFiles(string OrigPath, string DestPath)
                {
                    string[] files = Directory.GetFiles(OrigPath, "*.eml");
                    foreach (var file in files)
                    {
                        string targetPath = DestPath + "\\" + file.Replace(OrigPath, "");
                        if (File.Exists(targetPath))
                            File.Delete(targetPath);

                        File.Move(file, targetPath);
                    }
                }




                void DisplayandSaveReport(List<Correo> ListaCorreos)
                {
                    Console.Clear();
                    string ReportText = ShowConsoleLineAndReturnText("Statistics");
                    ReportText += ShowConsoleLineAndReturnText("==================================================================================================================");

                    ReportText += ShowConsoleLineAndReturnText("Total mails received by sender");
                    var resumenPorRemitente = ListaCorreos.Where(j => j.location == "Inbox").GroupBy(c => c.remitente).Select(g => new Resumen { Remitente = g.Key, Total = g.Count(), Size = g.Sum(c => long.Parse(c.size)) }).OrderByDescending(z => z.Total).ToList();
                    ReportText = ReflejarConsola(resumenPorRemitente, true);

                    ReportText += ShowConsoleLineAndReturnText("==================================================================================================================");
                    ReportText += ShowConsoleLineAndReturnText("Size of mails received by sender");
                    var resumenPorRemitente2 = ListaCorreos.Where(j => j.location == "Inbox").GroupBy(c => c.remitente).Select(g => new Resumen { Remitente = g.Key, Total = g.Count(), Size = g.Sum(c => long.Parse(c.size)) }).OrderByDescending(z => z.Size).ToList();
                    ReportText += ReflejarConsola(resumenPorRemitente2, false);


                    ReportText += ShowConsoleLineAndReturnText("==================================================================================================================");
                    ReportText += ShowConsoleLineAndReturnText("Total mails sent  by recipient");
                    var resumenPorDestinatario = ListaCorreos.Where(j => j.location == "Sent").GroupBy(c => c.destinatario).Select(g => new Resumen { Remitente = g.Key, Total = g.Count() }).OrderByDescending(z => z.Total).ToList();
                    ReportText += ReflejarConsola(resumenPorDestinatario, true, false, true);

                    ReportText += ShowConsoleLineAndReturnText("==================================================================================================================");
                    ReportText += ShowConsoleLineAndReturnText("Size of mails received by recipient");
                    var resumenPorDestinatario2 = ListaCorreos.Where(j => j.location == "Sent").GroupBy(c => c.destinatario).Select(g => new Resumen { Remitente = g.Key, Total = g.Count(), Size = g.Sum(c => long.Parse(c.size)) }).OrderByDescending(z => z.Size).ToList();
                    ReportText += ReflejarConsola(resumenPorDestinatario2, false, false, true);

                    ReportText += ShowConsoleLineAndReturnText("==================================================================================================================");
                    ReportText += ShowConsoleLineAndReturnText("Size of mails by subject");
                    var resumenPorTamaño = ListaCorreos.GroupBy(c => c.size).Select(g => new Resumen { Size = long.Parse(g.Key), Total = g.Count(), asunto = g.Select(c => c.asunto).FirstOrDefault() }).OrderByDescending(z => z.Total).ToList();
                    ReportText += ReflejarConsola(resumenPorTamaño, true, true);
                    ReportText += ShowConsoleLineAndReturnText("Count of mails by subject");
                    ReportText += ShowConsoleLineAndReturnText("==================================================================================================================");

                    var resumenPorTamaño2 = ListaCorreos.GroupBy(c => long.Parse(c.size)).Select(g => new Resumen { Size = g.Key, Total = g.Count(), asunto = g.Select(c => c.asunto).FirstOrDefault() }).OrderByDescending(z => z.Size).ToList();
                    ReportText += ReflejarConsola(resumenPorTamaño2, false, true);

                    ReportText += ShowConsoleLineAndReturnText("==================================================================================================================");
                    ReportText += ShowConsoleLineAndReturnText("Total mails by date");

                    ReportText += ReflejarDatosFechaConsola(ListaCorreos, c => c.fecha.Date);
                    ReportText += ShowConsoleLineAndReturnText("Total mails by Month");
                    ReportText += ReflejarDatosFechaConsola(ListaCorreos, c => new DateTime(c.fecha.Year, c.fecha.Month, 1),true);
                    ReportText += ShowConsoleLineAndReturnText("Total mails by Year");
                    ReportText += ReflejarDatosFechaConsola(ListaCorreos, c => c.fecha.Year);

                    ReportText += ShowConsoleLineAndReturnText("==================================================================================================================");
                    ReportText += ShowConsoleLineAndReturnText("Summary");

                    List<Correo> listaCorreosEnviados = ListaCorreos.Where(p => p.location == "Sent").ToList();
                    List<Correo> ListaCorreosRecibidos = ListaCorreos.Where(p => p.location == "Inbox").ToList();

                    long totalSizeSend = listaCorreosEnviados.Sum(x => long.Parse(x.size));
                    long totalSizeRec = ListaCorreosRecibidos.Sum(x => long.Parse(x.size));
                    long totalSize = ListaCorreos.Sum(x => long.Parse(x.size));
                    long maxValorTotal = ListaCorreos.Max(x => long.Parse(x.size));

                    ReportText += ShowConsoleLineAndReturnText("Received: " + ListaCorreosRecibidos.Count + " Size: " + ConvertirBytesAFormatoLegible(totalSizeRec)).PadRight(20);
                    ReportText += ShowConsoleLineAndReturnText("Sent: " + listaCorreosEnviados.Count + "|| Size: " + ConvertirBytesAFormatoLegible(totalSizeSend)).PadRight(20);
                    ReportText += ShowConsoleLineAndReturnText("Total: " + ListaCorreos.Count + "|| Size: " + ConvertirBytesAFormatoLegible(totalSize)).PadRight(20);

                    File.WriteAllText(reportPath, ReportText);

                }



                string ReflejarDatosFechaConsola<T>(List<Correo> ListaCorreos, Func<Correo, T> agrupador,bool SetMes = false)
                {
                    var resumenPorFecha = ListaCorreos.GroupBy(agrupador).Select(g => new { fecha = g.Key, Total = g.Count() }).OrderBy(z => z.fecha).ToList();
                    int maxValor = resumenPorFecha.Max(r => r.Total);
                    int maxBarWidth = 50;
                    string text = "";

                    foreach (var item in resumenPorFecha)
                    {
                        int barLength = (int)((item.Total / (double)maxValor) * maxBarWidth);
                        string bar = new string('█', barLength);
                        string fecha = item.fecha.ToString().Replace(" 0:00:00", "");

                        if (SetMes)
                            fecha = DateTime.Parse(fecha).ToString("MMMM yyyy");

                        string consoleLine = $"{fecha} | {bar} ({item.Total})";
                        Console.WriteLine(consoleLine);
                        text = text + consoleLine + "\r\n";
                    }

                    return text;
                }

                List<Correo> DownloadMailInfo(string location, bool downloadMails)
                {
                    try
                    {
                        var ListaCorreos = new List<Correo>();
                        bool continuar = true;
                        int contador = 0;
                        //;
                        while (continuar)
                        {
         

                            Thread.Sleep(600); 

                            wait120.Until(ExpectedConditions.VisibilityOfAllElementsLocatedBy(By.CssSelector(".flex.flex-nowrap.item-container")));
                            var correosPagina = driver.FindElements(By.CssSelector(".flex.flex-nowrap.item-container"));

                            contador++;

                            for (int i = 0; i < correosPagina.Count; i++)
                            {
                                ((IJavaScriptExecutor)driver).ExecuteScript("window.focus();");


                                Correo objetoMail = new Correo("", "", "", "", DateTime.Now, "");
                                bool ReadStatus = false;
                                try
                                {
                                    ReadStatus = clicarCorreo(correosPagina[i], i);
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine(ex.Message);
                                }
                                try
                                {
                                    objetoMail = GetInfoFromMail(location, i);
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine(ex.Message);
                                }

                                ListaCorreos.Add(objetoMail);

                                if (downloadMails)
                                {
                                    try
                                    {
                                        DescargarMail();

                                    }
                                    catch (Exception ex)
                                    {
                                        Console.WriteLine(ex.Message);
                                    }
                                }

                                if (ReadStatus)
                                {
                                    Thread.Sleep(1000);
                                    WaitAndClickElement("button[data-testid='message-header-expanded:mark-as-unread']");
                                }
                            }
                            continuar = ClickNextButton(continuar);
                        }
                        return ListaCorreos;
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                        var st = new StackTrace(ex, true);
                        var frame = st.GetFrame(0);
                        var line = frame.GetFileLineNumber();
                        return Lista_Correos;
                    }
                }


                bool ClickNextButton(bool continuar)
                {
                    try
                    {
                        wait120.Until(ExpectedConditions.ElementToBeClickable(By.CssSelector("button[data-testid='pagination-row:go-to-next-page']")));
                        var NextButton = driver.FindElement(By.CssSelector("button[data-testid='pagination-row:go-to-next-page']"));
                        continuar = NextButton.Enabled;
                        try
                        {
                            var Element = driver.FindElement(By.CssSelector(".button-solid-danger"));
                            if (Element.Enabled)
                                Element.Click();
                        }
                        catch { }
                        if (continuar)
                            WaitAndClickElement("button[data-testid='pagination-row:go-to-next-page");

                        return continuar;
                    }
                    catch
                    {
                        return false;

                    }
                }


                void DescargarMail()
                {
                    bool ClickBoton = false;

                    try
                    {
                        wait120.Until(ExpectedConditions.ElementExists(By.ClassName("message-recipient-item-label")));
                        wait120.Until(ExpectedConditions.ElementIsVisible(By.ClassName("message-recipient-item-label")));
                        wait120.Until(ExpectedConditions.ElementToBeClickable(By.ClassName("message-recipient-item-label")));
                        wait120.Until(ExpectedConditions.ElementExists(By.CssSelector("button[data-testid='message-header-expanded:more-dropdown'")));


                        //More Dropdown
                        WaitAndClickElement("button[data-testid='message-header-expanded:more-dropdown']");

                        ClickBoton = true;
                        //Export
                        WaitAndClickElement("button[data-testid='message-view-more-dropdown:export']");

                    }
                    catch (Exception ex)
                    {
                        if (!ex.Message.Contains("obscure"))
                        {
                            Console.WriteLine(ex.Message.ToString());
                            wait120.Until(ExpectedConditions.InvisibilityOfElementLocated(By.CssSelector(".modal-two modal-two--out")));
                            if (ClickBoton)
                                WaitAndClickElement("button[data-testid='message-view-more-dropdown:export']");
                            else
                            {
                                WaitAndClickElement("button[data-testid='message-header-expanded:more-dropdown']");

                                WaitAndClickElement("button[data-testid='message-view-more-dropdown:export']");
                            }
                        }

                    }

                }


                Correo GetInfoFromMail(string location, int i)
                {
                    Correo objetoMail = null;
                    try
                    {
                        var asunto = driver.FindElements(By.CssSelector(".message-conversation-summary-header"))[0]
                            .Text.Replace("<", "").Replace(">", "").Replace("\r", "").Replace("\n", "");

                        wait120.Until(ExpectedConditions.ElementExists(By.CssSelector("[data-testid='recipients:sender']")));
                        var remitente = "";
                        try
                        {
                            remitente = driver.FindElements(By.CssSelector("[data-testid='recipients:sender']"))[0]
                           .FindElement(By.CssSelector(".message-recipient-item-label"))
                           .Text.Replace("<", "").Replace(">", "").Replace("\r", "").Replace("\n", "");
                        }
                        catch
                        {
                            remitente = driver.FindElements(By.CssSelector("[data-testid='recipients:sender']"))[0]
                                .Text.Replace("<", "").Replace(">", "").Replace("\r", "").Replace("\n", "");
                        }
                        wait120.Until(ExpectedConditions.ElementExists(By.CssSelector("[data-testid='recipients:partial-recipients-list']")));
                        var destinatario = "";
                        try
                        {
                            destinatario = driver.FindElements(By.CssSelector("[data-testid='recipients:partial-recipients-list']"))[0]
                           .FindElement(By.CssSelector(".message-recipient-item-label"))
                           .Text.Replace("<", "").Replace(">", "");
                        }
                        catch
                        {
                            destinatario = driver.FindElements(By.CssSelector("[data-testid='recipients:partial-recipients-list']"))[0]
                        .Text.Replace("<", "").Replace(">", "");

                        }
                        var fechaString = driver.FindElements(By.TagName("time"))[i].GetAttribute("datetime");
                        DateTime fecha = NormalizarFecha(fechaString);

                        WaitAndClickElement("button[data-testid='message-header-expanded:more-dropdown']");
                        WaitAndClickElement("button[data-testid='message-view-more-dropdown:view-message-details']");
                        wait120.Until(ExpectedConditions.ElementToBeClickable(By.CssSelector("[data-testid='message-details:size']")));

                        var size = driver.FindElements(By.CssSelector("[data-testid='message-details:size']"))[0].Text;
                        size = size.Replace("Size:\r\n", "");
                        if (size == "")
                            Thread.Sleep(1000);
                        size = ConvertirATamañoBytes(size).ToString();

                        objetoMail = new Correo(remitente, destinatario, location, size.ToString(), fecha, asunto);
                        WaitAndClickElement(".modal-two-footer > button:nth-child(1)");
                        return objetoMail;
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.ToString());
                        try
                        {
                            WaitAndClickElement(".modal-two-footer > button:nth-child(1)");
                        }
                        catch { }
                        return objetoMail;
                    }
                }


                bool clicarCorreo(IWebElement correo, int i)
                {
                    bool ReadStatus = false;
                    try
                    {
                        ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView(true);", correo);
                        Thread.Sleep(200);
                        ReadStatus = correo.GetAttribute("class").Contains("unread");
                        wait120.Until(ExpectedConditions.ElementToBeClickable(correo));
                        correo.Click();
                        return ReadStatus;
                    }
                    catch (Exception ex)
                    {
                        if (ex.Message.Contains("obscures it"))
                        {
                            ((IJavaScriptExecutor)driver).ExecuteScript("var modal = document.querySelector('.modal-two');if (modal) modal.hide();");
                            //((IJavaScriptExecutor)driver).ExecuteScript("var modal2 = document.querySelector('.modal-two-backdrop');if (modal2) modal2.style.display = 'none';");

                            Thread.Sleep(5000);
                        }

                        wait120.Until(ExpectedConditions.ElementExists(By.CssSelector(".flex.flex-nowrap.item-container")));
                        var correo2 = driver.FindElements(By.CssSelector(".flex.flex-nowrap.item-container"));
                        correo2[i].Click();

                        ReadStatus = correo2[i].GetAttribute("class").Contains("unread");
                        return ReadStatus;
                    }
                }


                DateTime NormalizarFecha(string fechaString)
                {
                    try
                    {
                        string[] partes = fechaString.Split(new[] { " at " }, StringSplitOptions.None);
                        string fechaParte = partes[0];

                        string[] elementos = fechaParte.Split(',');
                        if (elementos.Length < 3)
                            throw new ArgumentException("Invalid Date format");

                        string año = elementos[2].Trim();
                        string[] mesDia = elementos[1].Trim().Split(' ');

                        if (mesDia.Length < 2)
                            throw new ArgumentException("Error while obtaining date");

                        string mes = mesDia[0].Trim();
                        string dia = Regex.Replace(mesDia[1], @"\D", "");

                        string fechaFinal = $"{año}-{mes}-{dia}";

                        return DateTime.ParseExact(fechaFinal, "yyyy-MMMM-d", CultureInfo.InvariantCulture);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error: {ex.Message}");
                        throw;
                    }
                }


                string ConvertirBytesAFormatoLegible(long bytes)
                {
                    string[] unidades = { "B", "KB", "MB", "GB", "TB" };
                    int index = 0;
                    double tamaño = bytes;

                    while (tamaño >= 1024 && index < unidades.Length - 1)
                    {
                        tamaño /= 1024;
                        index++;
                    }
                    return $"{tamaño:0.##} {unidades[index]}";
                }




                string ShowConsoleLineAndReturnText(string p)
                {
                    Console.WriteLine(p);
                    return p + "\n\r";
                }


                long ConvertirATamañoBytes(string tamañoString)
                {
                    string[] partes = tamañoString.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                    if (partes.Length != 2)
                        throw new ArgumentException("Wrong format on size label");

                    decimal valor;
                    if (!decimal.TryParse(partes[0], NumberStyles.Float, CultureInfo.InvariantCulture, out valor))
                        throw new ArgumentException("Invalid size number");

                    string unidad = partes[1].ToUpper();
                    long factor;

                    if (unidad == "B" || unidad == "BYTES") factor = 1;
                    else if (unidad == "KB") factor = 1024;
                    else if (unidad == "MB") factor = 1024 * 1024;
                    else if (unidad == "GB") factor = 1024 * 1024 * 1024;
                    else throw new ArgumentException("Unexpected unit on size.");
                    return (long)(valor * factor);
                }


                string ReflejarConsola(List<Resumen> x, bool usarTotal, bool usarAsunto = false, bool AsRecipient = false)
                {
                    string text = "";
                    if (x.Count <= 0)
                        return text;

                    int maxBarWidth = 50;
                    long maxValor = (long)(usarTotal ? x.Max(r => r.Total) : x.Max(r => r.Size));
                    string textoRemitente = AsRecipient ? "Recipient" : "Sender";

                    foreach (var item in x)
                    {
                        long barLength = usarTotal ? (long)((item.Total / (double)maxValor) * maxBarWidth) : barLength = (long)((item.Size / (double)maxValor) * maxBarWidth);
                        string bar = new string('█', (int)barLength);

                        string consoleLine = $"{textoRemitente}: {item.Remitente}".PadRight(60) + $"|| Mails: {item.Total}".PadRight(20) + $"|| Size: {bar} ({item.Total})";
                        string asunto = "";

                        if (!usarTotal)
                        {
                            var tamaño = ConvertirBytesAFormatoLegible(item.Size);
                            consoleLine = $"{textoRemitente}: {item.Remitente}".PadRight(60) + $"|| Mails: {item.Total}".PadRight(20) + $"|| Size: {bar} ({tamaño})";
                        }

                        if (usarAsunto)
                        {
                            asunto = item.asunto.Length > 90 ? item.asunto.Substring(0, 86) + "..." : item.asunto;
                            asunto = "Subject :" + asunto;

                            if (usarTotal)
                                consoleLine = $"Size: " + ConvertirBytesAFormatoLegible(item.Size).PadRight(10) + "|| " + asunto.PadRight(100) + $"|| Count: " + $" {bar}(" + item.Total.ToString() + ")";
                            else if (!usarTotal)
                                consoleLine = $"Size: " + ConvertirBytesAFormatoLegible(item.Size).PadRight(10) + "|| " + asunto.PadRight(100) + "|| Size:  " + $" {bar} (" + ConvertirBytesAFormatoLegible(item.Size) + ")";
                        }
                        Console.WriteLine(consoleLine);
                        text = text + consoleLine + "\r\n";
                    }

                    return text;
                }
                void SetSettings(bool originalGrouping, bool originalLayout, string setLanguage, out bool GroupingReturn, out bool LayoutReturn, out string OriginalLanguage)
                {
                    //Click Ajustes
                    WaitAndClickElement(".drawer-sidebar-button.rounded.flex.interactive");
                    //All Settings Click
                    WaitAndClickElement("a[data-testid='drawer-quick-settings:all-settings-button']");

                    //SET-LANGUAGE
                    WaitAndClickElement("a[href = '/u/0/mail/language-time']");
                    OriginalLanguage = driver.FindElement(By.Id("languageSelect")).Text;
                    if (OriginalLanguage != setLanguage)
                    {
                        driver.FindElement(By.Id("languageSelect")).Click();
                        driver.FindElement(By.CssSelector("button.dropdown-item-button[title=\"" + setLanguage + "\"]")).Click();
                    }
                    //MessagesCompositions
                    WaitAndClickElement("a[href = '/u/0/mail/general']");

                    //ViewMode;
                    wait30.Until(ExpectedConditions.ElementExists(By.Id("viewMode")));
                    Thread.Sleep(1000);
                    GroupingReturn = driver.FindElement(By.Id("viewMode")).Selected;

                    if (GroupingReturn || originalGrouping)
                        driver.ExecuteScript("arguments[0].click();", driver.FindElement(By.Id("viewMode")));
                    Thread.Sleep(1000);

                    //layout Columns
                    LayoutReturn = Boolean.Parse(driver.FindElement(By.CssSelector("div.mb-4:nth-child(1) > ul:nth-child(2) > li:nth-child(1) > button:nth-child(1)")).GetAttribute("aria-pressed"));
                    if (!LayoutReturn)
                        WaitAndClickElement("button[data-testid='layout:Column'");

                    if (!originalLayout)
                        WaitAndClickElement("button[data-testid='layout:Row'");

                    driver.Navigate().GoToUrl("https://mail.protonmail.com/u/0");
                    //Close Side Panel
                    WaitAndClickElement("button[data-testid='drawer-app-header:close']");

                }

                int DisplayMenu()
                {
                    Console.WriteLine("Download all mails from your protonmail account");
                    Console.WriteLine();
                    Console.WriteLine("1. Download Inbox items");
                    Console.WriteLine("2. Download Sent items");
                    Console.WriteLine("3. Download all");
                    Console.WriteLine("4. Download only summary");
                    Console.WriteLine("5. Exit");
                    var result = Console.ReadLine();
                    try
                    {
                        return Convert.ToInt32(result);
                    }
                    catch
                    {
                        return 6;
                    }
                }
                string GetDownloadFolderPath()
                {
                    return Registry.GetValue(@"HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "{374DE290-123F-4565-9164-39C4925E467B}", String.Empty).ToString();
                }
                string GetFirefoxPathFromRegistry()
                {
                    string firefoxRuta = null;
                    string registryKey = @"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\firefox.exe";
                    using (RegistryKey clave = Registry.LocalMachine.OpenSubKey(registryKey))
                    {
                        if (clave != null)
                        {
                            firefoxRuta = clave.GetValue(null) as string;
                        }
                    }
                    return firefoxRuta;
                }

                void WaitAndClickElement(string CssSelector)
                {
                    //Scroll into view in events
                    wait120.Until(ExpectedConditions.ElementExists(By.CssSelector(CssSelector)));
                    IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
                    ;
                    js.ExecuteScript("arguments[0].scrollIntoView(true);", driver.FindElement(By.CssSelector(CssSelector)));
                    wait120.Until(ExpectedConditions.ElementToBeClickable(By.CssSelector(CssSelector)));
                    var elemento = driver.FindElement((By.CssSelector(CssSelector)));
                    try
                    {


                        elemento.Click();
                        //wait4.Until(ExpectedConditions.ElementIsVisible(By.CssSelector(".modal-two modal-two--out")));
                        //Element < button class="button button-for-icon button-group-item button-medium button-outline-weak text-nowrap" type="button"> is not clickable at point(1033,147) because another element<div class="modal-two modal-two--out"> obscures it
                    }
                    catch (Exception ex)
                    {
                        if (ex.Message.Contains("obscures it"))
                        {
                            js.ExecuteScript("var modal = document.querySelector('.modal-two');if (modal) modal.hide();");

                            //wait30.Until(ExpectedConditions.InvisibilityOfElementLocated(By.CssSelector(".modal-two modal-two--out")));

                            //Thread.Sleep(5000);
                            elemento.Click();
                        }
                        else
                            throw ex;

                    }
                }


                FirefoxOptions SetPreferences(string ruta, string firefox_Ruta)
                {
                    FirefoxOptions opcion = new FirefoxOptions();
                    opcion.SetPreference("browser.download.folderList", 2);
                    opcion.SetPreference("browser.download.manager.showWhenStarting", false);
                    opcion.SetPreference("browser.download.dir", ruta);
                    opcion.SetPreference("browser.download.improvements_to_download_panel", true);
                    opcion.SetPreference("browser.download.manager.showWhenStarting", false);
                    opcion.SetPreference("browser.download.useDownloadDir", true);
                    opcion.SetPreference("browser.download.viewableInternally.enabledTypes", "");
                    opcion.SetPreference("browser.helperApps.alwaysAsk.force", false);
                    opcion.SetPreference("browser.helperApps.neverAsk.saveToDisk", "Thunderbird Document, blob: ,application/vnd.protonmail.v1+json, application/json, json, media-src,blob,message, message/rfc6532,message/partial, message/external-body, message/rfc822, application/octet-stream, text/plain, application/download, application/octet-stream, binary/octet-stream, application/binary, application/x-unknown, texto/html");
                    opcion.SetPreference("pdfjs.disabled", true);
                    opcion.BrowserExecutableLocation = firefox_Ruta;
                    return opcion;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.WriteLine("Press any key to close the application \n");
            }
        }

        private static string GetPassword()
        {
            string password = "";
            string tecla = "";
            ConsoleKeyInfo key = Console.ReadKey(true);
            while (key.Key != ConsoleKey.Enter)
            {
                if (key.Key != ConsoleKey.Enter && key.Key != ConsoleKey.Backspace)
                {
                    tecla = key.KeyChar.ToString().ToLower();
                    if (key.Modifiers == ConsoleModifiers.Shift)
                        tecla = key.KeyChar.ToString().ToUpper();
                    password += tecla;
                }
                else if (key.Key == ConsoleKey.Backspace && password.Length > 0)
                {
                    password = password.Substring(0, password.Length - 1);
                }
                key = Console.ReadKey(true);
            }
            return password;
        }
    }
}
