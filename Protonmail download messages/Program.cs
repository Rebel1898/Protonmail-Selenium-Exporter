using Microsoft.Win32;
using OpenQA.Selenium;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.Support.UI;
using SeleniumExtras.WaitHelpers;
using System;
using System.IO;
using System.Threading;
using WindowsInput;
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
                } while (userInput > 4);

                if (userInput == 4)
                    Environment.Exit(0);

                Console.WriteLine("Set username :");
                string username = Console.ReadLine();
                Console.Clear();

                Console.WriteLine("Set password :");
                string password = "";
                string tecla = "";
                ConsoleKeyInfo key = Console.ReadKey(true);
                while (key.Key != ConsoleKey.Enter)
                {
                    if (key.Key != ConsoleKey.Enter)
                    {
                        tecla = key.KeyChar.ToString().ToLower();
                        if (key.Modifiers == ConsoleModifiers.Shift)
                            tecla = key.KeyChar.ToString().ToUpper();
                        password += tecla;
                    }
                    key = Console.ReadKey(true);
                }
                Console.Clear();

                path = path + @"\Protonmail\";
                string InboxPath = path + "Inbox";
                string SentPath = path + "Sent";

                System.IO.Directory.CreateDirectory(InboxPath);
                System.IO.Directory.CreateDirectory(SentPath);

                FirefoxOptions options = SetPreferences(path);
                WebDriver driver = new FirefoxDriver(options);
                WebDriverWait wait30 = new WebDriverWait(driver, TimeSpan.FromSeconds(30));
                wait30.IgnoreExceptionTypes(typeof(StaleElementReferenceException));
                wait30.IgnoreExceptionTypes(typeof(NoSuchElementException));
                driver.Manage().Window.Maximize();
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
                driver.FindElement(By.CssSelector("button.button-large.button-solid-norm.w100.mt1-75")).Click();
                wait30.Until(ExpectedConditions.ElementExists(By.ClassName("text-keep-space")));
                driver.Navigate().GoToUrl("https://account.protonmail.com/u/0/mail/appearance");

                wait30.Until(ExpectedConditions.ElementExists(By.Id("viewMode")));
                var Grouping = driver.FindElement(By.Id("viewMode")).GetAttribute("Checked");
                if (Grouping == "true")
                    driver.ExecuteScript("arguments[0].click();", driver.FindElement(By.Id("viewMode")));
                bool automaticDownload = true;

                if (userInput == 3 || userInput == 1)
                {
                    driver.Navigate().GoToUrl("https://mail.protonmail.com/u/0");
                    DownloadMails();
                    MoveFiles(path, InboxPath);
                }
                if (userInput == 3 || userInput == 2)
                {
                    driver.Navigate().GoToUrl("https://mail.protonmail.com/u/0/sent");
                    DownloadMails();
                    MoveFiles(path, SentPath);
                }

                driver.Quit();
                driver.Close();
                Environment.Exit(0);

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
                        wait30.Until(ExpectedConditions.VisibilityOfAllElementsLocatedBy(By.CssSelector(".flex.flex-nowrap.flex-align-items-center.cursor-pointer.item-container")));
                        var correosPagina = driver.FindElements(By.CssSelector(".flex.flex-nowrap.flex-align-items-center.cursor-pointer.item-container"));

                        for (int i = 0; i < correosPagina.Count; i++)
                        {
                            try
                            {
                                ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView(true);", correosPagina[i]);
                                Thread.Sleep(200);
                                wait30.Until(ExpectedConditions.ElementToBeClickable(correosPagina[i]));
                                correosPagina[i].Click();
                            }
                            catch
                            {
                                Thread.Sleep(60000);
                                correosPagina = driver.FindElements(By.CssSelector(".flex.flex-nowrap.flex-align-items-center.cursor-pointer.item-container"));
                                correosPagina[i].Click();

                            }
                            wait30.Until(ExpectedConditions.ElementExists(By.ClassName("message-recipient-item-label")));
                            wait30.Until(ExpectedConditions.ElementIsVisible(By.ClassName("message-recipient-item-label")));
                            wait30.Until(ExpectedConditions.ElementToBeClickable(By.ClassName("message-recipient-item-label")));

                            var correosVisibles = driver.FindElements(By.ClassName("message-recipient-item-label"));
                            var articlesinHilo = driver.FindElements(By.TagName("article"));

                            for (int j = correosVisibles.Count - 1; j > -1; j--)
                            {
                                var claseArticle = articlesinHilo[j].GetAttribute("class");
                                ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView(true);", correosVisibles[j]);
                                if (claseArticle.IndexOf("is-opened") == -1)
                                {
                                    driver.ExecuteScript("arguments[0].click();", correosVisibles[j]);
                                    claseArticle = articlesinHilo[j].GetAttribute("class");
                                    if (claseArticle.IndexOf("is-closed") != -1)
                                        driver.ExecuteScript("arguments[0].click();", correosVisibles[j]);
                                }
                            }
                            wait30.Until(ExpectedConditions.ElementExists(By.CssSelector(".icon-16p.caret-like")));
                            var botonesHilo = driver.FindElements(By.CssSelector(".icon-16p.caret-like"));

                            for (int r = 0; r < correosVisibles.Count; r++)
                            {

                                wait30.Until(ExpectedConditions.ElementIsVisible(By.CssSelector(".message-content.scroll-horizontal-if-needed.relative.bodyDecrypted.bg-norm.color-norm")));
                                wait30.Until(ExpectedConditions.ElementToBeClickable(By.CssSelector(".message-content.scroll-horizontal-if-needed.relative.bodyDecrypted.bg-norm.color-norm")));

                                IWebElement MailContainer = driver.FindElement(By.CssSelector(".message-content.scroll-horizontal-if-needed.relative.bodyDecrypted.bg-norm.color-norm"));

                                var Textocorreo = MailContainer.Text;
                                int contador = 60;
                                while (MailContainer.Text == "" && contador > 0)
                                {
                                    if (MailContainer.Text == "null")
                                        Console.WriteLine("NULL");
                                    Thread.Sleep(500);
                                    contador--;
                                }
                                driver.ExecuteScript("arguments[0].click();", botonesHilo[r].FindElement(By.XPath("./..")));
                                wait30.Until(ExpectedConditions.ElementExists(By.CssSelector(".flex-item-fluid.mtauto.mbauto")));
                                driver.FindElements(By.CssSelector(".flex-item-fluid.mtauto.mbauto"))[4].Click();
                                if (!automaticDownload)
                                    automaticDownload = SetAutoDownloads();
                            }
                        }
                        try
                        {
                            var NextButton = driver.FindElement(By.CssSelector(".icon-16p.block.rotateZ-270"));
                            continuar = !bool.Parse(driver.ExecuteScript("return document.getElementsByClassName('icon-16p block rotateZ-270')[0].parentElement.disabled").ToString());
                            driver.ExecuteScript("arguments[0].click();", NextButton.FindElement(By.XPath("./..")));
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                            continuar = false;
                        }
                    }


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
                FirefoxOptions SetPreferences(string ruta)
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
    }
}
