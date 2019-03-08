using ConsoleApp.Helper;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using WatiN.Core;

namespace ConsoleApp.Connectors
{
    class GoogleScraping
    {
        public string FindTerm(string term)
        {
            Exception exception = null;
            string output = "";
            System.Threading.Thread thread = new Thread(new ParameterizedThreadStart(s => Start(ref exception, ref output, term)));
            thread.SetApartmentState(ApartmentState.STA);

            thread.Start();
            thread.Join();
            
            if (exception != null)
            {
                throw exception;
            }

            return output;
        }

        private void Start(ref Exception exception, ref string output, string term)
        {
            IE browser = null;
            try
            {
                browser = IEBrowserHelper.GetBrowser();

                browser.GoTo("https://www.google.com.br/");
                browser.WaitForComplete();

                TextField txtSearch = browser.TextField(Find.ByName("q"));

                if (txtSearch.Exists)
                {
                    txtSearch.SetAttributeValue("value", term);
                }

                Element btnFind = browser.Element(Find.ByName("btnK"));

                if (btnFind.Exists)
                {
                    btnFind.Click();
                }
                
                var resultadosComplementares = browser.Div(Find.ByClass(p => p.Contains("kno-ecr-pt kno-fb-ctx")));

                if (resultadosComplementares.Exists)
                {
                    output = resultadosComplementares.OuterText;

                    var resultadosComplementaresDescricao = browser.Div(Find.ByClass(p => p.Contains("kno-rdesc")));

                    if (resultadosComplementaresDescricao.Exists)
                    {
                        output += ": " + resultadosComplementaresDescricao.Spans[0].OuterText.Replace("\r\n", string.Empty);
                    }

                }
                else
                {
                    output = "O termo de busca não resultou em algo com a area de Resultados complementares";
                }
 

            }
            catch (Exception e)
            {
                exception = e;
            }
finally
            {

                //Close the browser
                if (browser != null)
                {
                    browser.Close();
                    browser.Dispose();
                }
            }
        }
    }
}
