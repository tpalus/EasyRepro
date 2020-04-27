using Microsoft.Dynamics365.UIAutomation.Browser;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Microsoft.PowerApps.UIAutomation.Api.Controls
{
    public class DropDownList : PlayerControl
    {
        public override void SetValue(InteractiveBrowser browser)
        {
            if (browser.Driver.HasElement(By.XPath($"//div[contains(@data-control-name,'{this.ControlName}')]")))
            {
                //Click the down arrow
                //var appControl = browser.Driver.FindElements(By.XPath($"//div[contains(@data-control-name,'{this.ControlName}')]"))[0];
                var appControl = browser.Driver.FindElement(By.XPath($"//div[@data-control-name='{this.ControlName}']"));

                var downArrow = appControl.FindElements(By.XPath(".//div[contains(@class, 'dropdownLabelArrow')]"));

                if (downArrow.Count > 0)
                    downArrow[0].Click(true);

                browser.ThinkTime(100);

                //Select the Value
                var listItems = browser.Driver.FindElements(By.XPath("//div[contains(@class,'dropdownListItem')]"));

                bool found = false;

                foreach (var listItem in listItems)
                {
                    if (listItem.Text.Equals(this.ControlValue, StringComparison.OrdinalIgnoreCase))
                    {
                        listItem.Click(true);
                        found = true;
                        break;
                    }
                }

                if (!found)
                    throw new NotFoundException($"Unable to find {this.ControlValue} in {this.ControlName} control");

                browser.ThinkTime(100);
            }
        }
    }
}
