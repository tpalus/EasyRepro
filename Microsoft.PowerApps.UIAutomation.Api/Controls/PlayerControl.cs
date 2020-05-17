using Microsoft.Dynamics365.UIAutomation.Browser;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Microsoft.PowerApps.UIAutomation.Api
{
    public class PlayerControl
    {
        public string ControlName { get; set; }
        public string ControlValue { get; set; }

        public virtual void SetValue(InteractiveBrowser driver)
        {

        }

        public void ScrollIntoView(InteractiveBrowser browser, IWebElement element)
        {
            browser.Driver.ExecuteScript("arguments[0].scrollIntoView(true);", element);
        }
    }
}
