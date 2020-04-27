// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

using Microsoft.Dynamics365.UIAutomation.Browser;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;

namespace Microsoft.PowerApps.UIAutomation.Api
{

    /// <summary>
    ///  The Home page.
    ///  </summary>
    public class Player
        : AppPage
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="Player"/> class.
        /// </summary>
        /// <param name="browser">The browser.</param>
        public Player(InteractiveBrowser browser)
            : base(browser)
        {
            if (browser.Driver.HasElement(By.XPath("//div[contains(@class,'publishedAppContainer')]")))
            {
                var appContainer = browser.Driver.FindElement(By.XPath("//div[contains(@class,'publishedAppContainer')]"));
                browser.Driver.SwitchTo().Frame(appContainer.FindElements(By.TagName("iframe"))[0]);
            }
        }
        
        /// <summary>
        /// Hide the Studio Welcome Dialog
        /// </summary>
        public BrowserCommandResult<bool> OpenApp(string id, Dictionary<string,string> parameters = null, bool openInNewTab = false)
        {
            return this.Execute(GetOptions("Open App By ID"), driver =>
            {
                var uri = new Uri(this.Browser.Driver.Url);
                var link = $"{uri.Scheme}://{uri.Authority}/webplayer/app?appId=/providers/Microsoft.PowerApps/apps/{id}";

                foreach (var param in parameters)
                {
                    link += $"&{param.Key}={param.Value}";
                }

                if (openInNewTab)
                {
                    driver.OpenNewTab();

                    Thread.Sleep(1000);
                    driver.LastWindow();

                    Thread.Sleep(1000);

                    driver.Navigate().GoToUrl(link);
                }
                else
                {
                    driver.Navigate().GoToUrl(link);
                }

                driver.WaitForPlayerToLoad();
                
                return true;
            });
        }

        public BrowserCommandResult<bool> SetValue(PlayerControl control)
        {
            return this.Execute($"Set Value for {control.ControlName} control", driver =>
            {
                control.SetValue(this.Browser);

                return true;
            });
        }

        public BrowserCommandResult<bool> ClickButton(string buttonName)
        {
            return this.Execute($"Click {buttonName} Button", driver =>
            {
                if (this.Browser.Driver.HasElement(By.XPath($"//div[contains(text(),'{buttonName}')]")))
                    this.Browser.Driver.FindElement(By.XPath($"//div[contains(text(),'{buttonName}')]")).Click(true);
                else
                    throw new NotFoundException($"Unable to find {buttonName} button");

                //Use WaitForPageToLoad if possible
                driver.WaitForPageToLoad();
                Browser.ThinkTime(2000);

                return true;
            });
            
        }

        public BrowserCommandResult<bool> ClickGalleryItem(string itemName,bool focusOnNewWindow = false)
        {
            return this.Execute($"Click {itemName} Gallery Item", driver =>
            {
                if (driver.HasElement(By.XPath($"//div[contains(text(),'{ itemName }')]")))
                    driver.FindElements(By.XPath($"//div[contains(text(),'{ itemName }')]"))[0].Click();
                else
                    throw new NotFoundException($"Unable to locate {itemName} item in the gallery");

                Browser.ThinkTime(5000);

                if(focusOnNewWindow)
                {
                    driver.SwitchTo().Window(driver.WindowHandles.Last());
                    driver.SwitchTo().DefaultContent();
                    driver.SwitchTo().Frame(0);
                }

                return true;

            });
        }

        public BrowserCommandResult<bool> ReturnToFacilityMenu(string facility)
        {

            return this.Execute($"Return to Facility Menu", driver =>
            {
                driver.SwitchTo().Window(driver.WindowHandles[1]);
                driver.SwitchTo().DefaultContent();
                driver.SwitchTo().Frame(0);

                var labels = driver.FindElements(By.XPath($"//div[contains(@data-control-name,'lblFacility')]"));

                labels.First(l => l.Text.StartsWith(facility, StringComparison.OrdinalIgnoreCase)).Click(true);

                Browser.ThinkTime(500);

                return true;

            });

        }

        public BrowserCommandResult<bool> ClickFeedbackIcon()
        {
            return this.Execute($"Click Feedback", driver =>
            {
                //driver.SwitchTo().Window(driver.WindowHandles.Last());
                driver.SwitchTo().DefaultContent();
                driver.SwitchTo().Frame(0);

                if (driver.HasElement(By.XPath($"//div[contains(@data-control-name,'Feedback_Icon')]")))
                    driver.FindElement(By.XPath($"//div[contains(@data-control-name,'Feedback_Icon')]")).Click(true);

                Browser.ThinkTime(1000);

                return true;

            });

        }
    }
}