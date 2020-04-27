using Microsoft.Dynamics365.UIAutomation.Browser;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Microsoft.PowerApps.UIAutomation.Api.Controls
{
    public class TextArea : PlayerControl
    {
        public override void SetValue(InteractiveBrowser browser)
        {
            if (browser.Driver.HasElement(By.XPath($"//div[contains(@data-control-name,'{this.ControlName}')]")))
            {
                //Click the down arrow6
                var appControl = browser.Driver.FindElement(By.XPath($"//div[@data-control-name='{this.ControlName}']"));
                appControl.FindElement(By.TagName("textarea")).SendKeys(this.ControlValue);

                browser.ThinkTime(100);

            }
        }
    }
}
