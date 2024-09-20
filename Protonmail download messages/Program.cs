using Microsoft.Win32;
using OpenQA.Selenium;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.Support.UI;
using SeleniumExtras.WaitHelpers;
using System;
using System.IO;
using System.Threading;
using WebDriverManager;
using WebDriverManager.DriverConfigs.Impl;
using WindowsInput;

namespace Protonmail_download_messages
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                bool automaticDownload = true;
                InputSimulator teclado = new InputSimulator();
                string path = GetDownloadFolderPath();
                int userInput = 0;
                do
                {
                    userInput = DisplayMenu();
                } while (userInput > 4);

                if (userInput == 4)
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
                driver.Navigate().GoToUrl("https://mail.protonmail.com");
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
                SetSettings(false,true,out Grouping,out Layout);

                if (userInput == 3 || userInput == 1)
                {
                    DownloadMails();
                    MoveFiles(path, InboxPath);
                }
                if (userInput == 3 || userInput == 2)
                {
                    WaitAndClickElement("a[data-testid='navigation-link:sent']");
                    DownloadMails();
                    MoveFiles(path, SentPath);
                }

                SetSettings(Grouping, Layout,out Grouping,out Layout);
                Console.Clear();
                Console.WriteLine("Export completed!");
                driver.Quit();
                driver.Close();
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
                void DownloadMails()
                {
                    bool continuar = true;
                    while (continuar)
                    {
                        Thread.Sleep(600);
                        wait120.Until(ExpectedConditions.VisibilityOfAllElementsLocatedBy(By.CssSelector(".flex.flex-nowrap.item-container")));
                        var correosPagina = driver.FindElements(By.CssSelector(".flex.flex-nowrap.item-container"));

                        for (int i = 0; i < correosPagina.Count; i++)
                        {
                            bool ReadStatus = false;
                            try
                            {
                                ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView(true);", correosPagina[i]);
                                Thread.Sleep(200);
                                ReadStatus = correosPagina[i].GetAttribute("class").Contains("unread");
                                wait120.Until(ExpectedConditions.ElementToBeClickable(correosPagina[i]));
                                correosPagina[i].Click();
                            }
                            catch
                            {
                                Thread.Sleep(60000);
                                correosPagina = driver.FindElements(By.CssSelector(".flex.flex-nowrap.item-container"));
                                correosPagina[i].Click();
                            }
                            wait30.Until(ExpectedConditions.ElementExists(By.ClassName("message-recipient-item-label")));
                            wait30.Until(ExpectedConditions.ElementIsVisible(By.ClassName("message-recipient-item-label")));
                            wait30.Until(ExpectedConditions.ElementToBeClickable(By.ClassName("message-recipient-item-label")));
                            wait30.Until(ExpectedConditions.ElementExists(By.CssSelector("button[data-testid='message-header-expanded:more-dropdown'")));

                            //More Dropdown
                            WaitAndClickElement("button[data-testid='message-header-expanded:more-dropdown']");
                            //Export
                            WaitAndClickElement("button[data-testid='message-view-more-dropdown:export']");

                            if (ReadStatus)
                            {
                                Thread.Sleep(1000);
                                WaitAndClickElement("button[data-testid='message-header-expanded:mark-as-unread']");
                            }

                            if (!automaticDownload)
                                automaticDownload = SetAutoDownloads();
                        }
                        try
                        {
                            var NextButton = driver.FindElement(By.CssSelector("button[data-testid='pagination-row:go-to-next-page']"));
                            continuar = NextButton.Enabled;
                            if (continuar)
                                NextButton.Click();
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                            continuar = false;
                        }
                    }


                }
                void SetSettings(bool originalGrouping , bool originalLayout, out bool GroupingReturn, out bool LayoutReturn)
                {
                    //Set No reply grouping
                    //bool originalGrouping = false, bool originalLayout = true
                    //driver.FindElement(By.CssSelector("button.button:nth-child(6)")).Click();
                    //Click Ajustes
                    WaitAndClickElement(".drawer-sidebar-button.rounded.flex.interactive");
                    //All Settings Click
                    WaitAndClickElement("a[data-testid='drawer-quick-settings:all-settings-button']");
                    //MessagesCompositions
                    WaitAndClickElement("a[href = '/u/0/mail/general']");
                    //ViewMode;
                    wait30.Until(ExpectedConditions.ElementExists(By.Id("viewMode")));
                    Thread.Sleep(1000);
                     GroupingReturn = driver.FindElement(By.Id("viewMode")).Selected;
                    //GroupingReturn = bool.Parse(prueba);
                    if (GroupingReturn || originalGrouping)
                        driver.ExecuteScript("arguments[0].click();", driver.FindElement(By.Id("viewMode")));

                    //layout Columns
                    LayoutReturn = Boolean.Parse(driver.FindElement(By.CssSelector("button[data-testid='layout:Column'")).GetAttribute("aria-pressed"));
                    if (!LayoutReturn)
                        WaitAndClickElement("button[data-testid='layout:Column'");

                    if(!originalLayout)
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
                    Console.WriteLine("4. Exit");
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
                bool SetAutoDownloads()
                {
                    Thread.Sleep(5000);
                    teclado.Keyboard.KeyPress(WindowsInput.Native.VirtualKeyCode.UP);
                    Thread.Sleep(200);
                    teclado.Keyboard.KeyPress(WindowsInput.Native.VirtualKeyCode.TAB);
                    Thread.Sleep(200);
                    teclado.Keyboard.KeyPress(WindowsInput.Native.VirtualKeyCode.TAB);
                    Thread.Sleep(200);
                    teclado.Keyboard.KeyPress(WindowsInput.Native.VirtualKeyCode.SPACE);
                    Thread.Sleep(200);
                    teclado.Keyboard.KeyPress(WindowsInput.Native.VirtualKeyCode.TAB);
                    Thread.Sleep(200);
                    teclado.Keyboard.KeyPress(WindowsInput.Native.VirtualKeyCode.SPACE);
                    return true;
                }

                void WaitAndClickElement(string CssSelector)
                {
                    wait30.Until(ExpectedConditions.ElementExists(By.CssSelector(CssSelector)));
                    wait30.Until(ExpectedConditions.ElementIsVisible(By.CssSelector(CssSelector)));
                    wait30.Until(ExpectedConditions.ElementToBeClickable(By.CssSelector(CssSelector)));
                    var elemento = driver.FindElement((By.CssSelector(CssSelector)));
                    elemento.Click();
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
                Console.Clear();
                Console.WriteLine(ex.Message);
                Console.WriteLine("Press any key to close the application \n");
                Console.ReadKey();
                Environment.Exit(0);
            }
        }

        private static string GetPassword()
        {
            string password = "";
            string tecla = "";

            ConsoleKeyInfo key = Console.ReadKey(true);
            while (key.Key != ConsoleKey.Enter)
            {
                if (key.Key != ConsoleKey.Enter && key.Key!= ConsoleKey.Backspace)
                {
                    tecla = key.KeyChar.ToString().ToLower();
                    if (key.Modifiers == ConsoleModifiers.Shift)
                        tecla = key.KeyChar.ToString().ToUpper();
                    password += tecla;
                }
                else if (key.Key == ConsoleKey.Backspace && password.Length > 0)
                {
                    // Elimina el último carácter de la contraseña y borra un asterisco en la consola
                    password = password.Substring(0, password.Length - 1);
                }
                key = Console.ReadKey(true);
            }

            return password;
        }
    }
}
