// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.Dynamics365.UIAutomation.Api.UCI;
using Microsoft.Dynamics365.UIAutomation.Browser;
using System;
using System.Security;

namespace Microsoft.Dynamics365.UIAutomation.Testing
{
    [TestClass]
    public class CreateOperations
    {
        private readonly SecureString _username = System.Configuration.ConfigurationManager.AppSettings["OnlineUsername"].ToSecureString();
        private readonly SecureString _password = System.Configuration.ConfigurationManager.AppSettings["OnlinePassword"].ToSecureString();
        private readonly Uri _xrmUri = new Uri(System.Configuration.ConfigurationManager.AppSettings["OnlineCrmUrl"].ToString());

        [TestCategory("Internal")]
        [TestMethod]
        public void CreateAccount()
        {
            var client = new WebClient(TestSettings.Options);
            using (var xrmApp = new XrmApp(client))
            {

                try
                {
                    xrmApp.OnlineLogin.Login(_xrmUri, _username, _password);
                    xrmApp.Navigation.OpenApp(UCIAppName.Sales);

                    xrmApp.Navigation.OpenSubArea("Sales", "Accounts");

                    xrmApp.CommandBar.ClickCommand("New");

                    xrmApp.Entity.SetValue("name", TestSettings.GetRandomString(5, 15));                    

                    xrmApp.Entity.Save();

                    xrmApp.Telemetry.TrackCommandEvents();
                }
                catch (Exception ex)
                {
                    xrmApp.Telemetry.TrackExceptions(ex);
                    throw;
                }
            }
           
        }

        [TestCategory("Internal")]
        [TestMethod]
        public void CreateContact()
        {
            var client = new WebClient(TestSettings.Options);
            using (var xrmApp = new XrmApp(client))
            {
                try
                { 
                    xrmApp.OnlineLogin.Login(_xrmUri, _username, _password);
                    xrmApp.Navigation.OpenApp(UCIAppName.Sales);

                    xrmApp.Navigation.OpenSubArea("Sales", "Contacts");

                    xrmApp.CommandBar.ClickCommand("New");

                    xrmApp.Entity.SetValue("firstname", TestSettings.GetRandomString(5, 10));
                    xrmApp.Entity.SetValue("lastname", TestSettings.GetRandomString(5, 10));

                    xrmApp.Entity.Save();

                    xrmApp.Telemetry.TrackCommandEvents();
                }
                catch (Exception ex)
                {
                    xrmApp.Telemetry.TrackExceptions(ex);
                    throw;
                }
            }
           
        }
        [TestCategory("Internal")]
        [TestMethod]
        public void CreateOpportunity()
        {
            var client = new WebClient(TestSettings.Options);
            using (var xrmApp = new XrmApp(client))
            {
                try
                {
                    xrmApp.OnlineLogin.Login(_xrmUri, _username, _password);
                    xrmApp.Navigation.OpenApp(UCIAppName.Sales);

                    xrmApp.Navigation.OpenSubArea("Sales", "Opportunities");

                    xrmApp.CommandBar.ClickCommand("New");

                    xrmApp.Entity.SetValue("name", TestSettings.GetRandomString(5, 10));

                    xrmApp.Entity.Save();
                    xrmApp.Telemetry.TrackCommandEvents();
                }
                catch (Exception ex)
                {
                    xrmApp.Telemetry.TrackExceptions(ex);
                    throw;
                }
            }
        }
        [TestCategory("Internal")]
        [TestMethod]
        public void CreateCase()
        {
            
            var client = new WebClient(TestSettings.Options);
            using (var xrmApp = new XrmApp(client))
            {
                try
                {
                    xrmApp.OnlineLogin.Login(_xrmUri, _username, _password);
                    xrmApp.Navigation.OpenApp(UCIAppName.Sales);

                    xrmApp.Navigation.OpenSubArea("Service", "Cases");

                    xrmApp.CommandBar.ClickCommand("New Case");

                    xrmApp.ThinkTime(5000);

                    xrmApp.Entity.SetValue("title", TestSettings.GetRandomString(5, 10));

                    LookupItem customer = new LookupItem { Name = "customerid", Value = "Test Account" };
                    xrmApp.Entity.SetValue(customer);

                    xrmApp.Entity.Save();
                    xrmApp.Telemetry.TrackCommandEvents();
                }
                catch (Exception ex)
                {
                    xrmApp.Telemetry.TrackExceptions(ex);
                    throw;
                }
            }
            
        }
        [TestCategory("Internal")]
        [TestMethod]
        public void CreateLead()
        {            
            var client = new WebClient(TestSettings.Options);
            using (var xrmApp = new XrmApp(client))
            {
                try
                {
                    xrmApp.OnlineLogin.Login(_xrmUri, _username, _password);
                    xrmApp.Navigation.OpenApp(UCIAppName.Sales);

                    xrmApp.Navigation.OpenSubArea("Sales", "Leads");

                    xrmApp.CommandBar.ClickCommand("New");

                    xrmApp.ThinkTime(5000);

                    xrmApp.Entity.SetValue("subject", TestSettings.GetRandomString(5, 15));
                    xrmApp.Entity.SetValue("firstname", TestSettings.GetRandomString(5, 10));
                    xrmApp.Entity.SetValue("lastname", TestSettings.GetRandomString(5, 10));

                    xrmApp.Entity.Save();
                    xrmApp.Telemetry.TrackCommandEvents();
                }
                catch (Exception ex)
                {
                    xrmApp.Telemetry.TrackExceptions(ex);
                    throw;
                }
            }
        }
    }
}