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
    public class AdminApp
        : AppPage
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="Player"/> class.
        /// </summary>
        /// <param name="browser">The browser.</param>
        public AdminApp(InteractiveBrowser browser)
            : base(browser)
        {
        }
        
        public BrowserCommandResult<bool> OpenSubArea(string subArea)
        {
            return this.Execute($"Open {subArea} subarea", driver =>
            {

                return true;
            });
        }

    }
}