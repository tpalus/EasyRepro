using Microsoft.Dynamics365.UIAutomation.Browser;
using Microsoft.PowerApps.UIAutomation.Api;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Configuration;
using System.Collections.Generic;
using System;
using System.Diagnostics;
using OpenQA.Selenium;
using Microsoft.PowerApps.UIAutomation.Api.Controls;

namespace Microsoft.PowerApps.UIAutomation.Sample.EmergencyResponseApp
{
    [TestClass]
    public class EmergencyResponseApp
    {


        [TestMethod]
        public void UseEmergencyResponseApp()
        {
            var username = ConfigurationManager.AppSettings["OnlineUsername"];
            var password = ConfigurationManager.AppSettings["OnlinePassword"];
            var uri = new System.Uri("https://web.powerapps.com");
            var appId = ConfigurationManager.AppSettings["AppId"];

            using (var appBrowser = new PowerAppBrowser(TestSettings.Options))
            {
                Uri appsPage = new Uri(ConfigurationManager.AppSettings["OnlineUrl"]);

                appBrowser.OnlineLogin.PowerAppsLogin(appsPage, username.ToSecureString(), null);

                //Pick the DDL Lists
                COVIDLocation location = GetNewLocation();

                DropDownList systemDDL = new DropDownList() { ControlName = "SystemSelect_System_DD", ControlValue = location.System };
                DropDownList regionDDL = new DropDownList() { ControlName = "SystemSelect_Region_DD", ControlValue = String.Format("{0} {1}", location.System, location.Region)};
                DropDownList facilityDDL = new DropDownList() { ControlName = "SystemSelect_Facility_DD", ControlValue = String.Format("{0} {1} {2}", location.System, location.Region, location.Facility)};

                appBrowser.Player.SetValue(systemDDL);
                appBrowser.Player.SetValue(regionDDL);
                appBrowser.Player.SetValue(facilityDDL);


                //Click Next Button
                appBrowser.Player.ClickButton("Next");

                //Click icon/facility name
                appBrowser.Player.ClickGalleryItem("Supplies");

                //Select a different system/location/facility
                appBrowser.Player.ReturnToFacilityMenu(location.System);

                //Close the App
                appBrowser.Driver.Close();
                appBrowser.ThinkTime(500);

                //Reopen the App
                appBrowser.Driver.SwitchTo().Window(appBrowser.Driver.WindowHandles[0]);
                appBrowser.Player.ClickGalleryItem("Supplies");

                appBrowser.ThinkTime(1000);

            }
        }

        [TestMethod]
        public void UseSuppliesApp()
        {
            var username = ConfigurationManager.AppSettings["OnlineUsername"];
            var password = ConfigurationManager.AppSettings["OnlinePassword"];
            var uri = new System.Uri("https://web.powerapps.com");
            var appId = ConfigurationManager.AppSettings["AppId"];

            using (var appBrowser = new PowerAppBrowser(TestSettings.Options))
            {
                Uri appsPage = new Uri(ConfigurationManager.AppSettings["OnlineUrl"]);

                appBrowser.OnlineLogin.PowerAppsLogin(appsPage, username.ToSecureString(), null);

                //Pick the DDL Lists

                COVIDLocation location = GetNewLocation();

                DropDownList systemDDL = new DropDownList() { ControlName = "SystemSelect_System_DD", ControlValue = location.System };
                DropDownList regionDDL = new DropDownList() { ControlName = "SystemSelect_Region_DD", ControlValue = String.Format("{0} {1}", location.System, location.Region) };
                DropDownList facilityDDL = new DropDownList() { ControlName = "SystemSelect_Facility_DD", ControlValue = String.Format("{0} {1} {2}", location.System, location.Region, location.Facility) };

                appBrowser.Player.SetValue(systemDDL);
                appBrowser.Player.SetValue(regionDDL);
                appBrowser.Player.SetValue(facilityDDL);

                //Click Next Button
                appBrowser.Player.ClickButton("Next");

                //Click icon/facility name
                appBrowser.Player.ClickGalleryItem("Supplies", true);

                //Enter Numbers values in the fields and click Send
                var rows = appBrowser.Driver.FindElements(By.XPath("//*[contains(@data-control-part,'gallery-item')]"));

                int cnt = 1;

                foreach(var row in rows)
                {
                    var rowLabel = row.FindElement(By.XPath(".//div[contains(@data-control-name,'lblSupplyName')]")).Text;

                    row.FindElement(By.XPath(".//div[contains(@data-control-name,'txtSuppliesOnHand')]")).FindElement(By.TagName("input")).SendKeys(cnt++.ToString());
                    row.FindElement(By.XPath(".//div[contains(@data-control-name,'txtSupplyBurnRate')]")).FindElement(By.TagName("input")).SendKeys(cnt++.ToString());
                    
                }

                //Click Submit
                appBrowser.ThinkTime(100);
                appBrowser.Player.ClickButton("Submit");
                appBrowser.ThinkTime(2000);

                //Click Home
                appBrowser.Player.ClickButton("Home");
                appBrowser.ThinkTime(2000);

                //Check the inserted values are correctly opening the entity in the Admin App

                //Click on facility Icon
                appBrowser.Player.ReturnToFacilityMenu(location.System);

                //Change System/Region/Facility and click Next
                COVIDLocation location2 = GetNewLocation();

                DropDownList systemDDL2 = new DropDownList() { ControlName = "SystemSelect_System_DD", ControlValue = location2.System };
                DropDownList regionDDL2 = new DropDownList() { ControlName = "SystemSelect_Region_DD", ControlValue = String.Format("{0} {1}", location2.System, location2.Region) };
                DropDownList facilityDDL2 = new DropDownList() { ControlName = "SystemSelect_Facility_DD", ControlValue = String.Format("{0} {1} {2}", location2.System, location2.Region, location2.Facility) };

                appBrowser.Player.SetValue(systemDDL2);
                appBrowser.Player.SetValue(regionDDL2);
                appBrowser.Player.SetValue(facilityDDL2);

                //Click Next Button
                appBrowser.Player.ClickButton("Next");
                
                //Click Supplies App

                appBrowser.Player.ClickGalleryItem("Supplies", true);


                //Click the Feedback Icon
                appBrowser.Player.ClickFeedbackIcon();

                //Fill in the feedback and click send
                ComboBox topicCombo = new ComboBox() { ControlName = "DataCardValue21", ControlValue = "Idea" };
                appBrowser.Player.SetValue(topicCombo);

                string feedbackString = "This app is fantastic";
                appBrowser.Driver.FindElement(By.TagName("textarea")).SendKeys(feedbackString);

                appBrowser.Player.ClickButton("Submit");

                appBrowser.ThinkTime(1000);
            }
        }




        //Helpers
        private COVIDLocation GetNewLocation()
        {
            string[] systems = new string[] { "Alpine", "Blue Yonder", "Fabrikam", "Litware", "Southridge" };
            string[] regions = new string[] { "North", "South", "East", "West" };
            string[] facilities = new string[] { "Facility 1", "Facility 2" };

            Random systemRandom = new Random();
            Random regionRandom = new Random();
            Random facilityRandom = new Random();

            COVIDLocation location = new COVIDLocation();

            location.System = systems[systemRandom.Next(0, systems.Length)];
            location.Region = regions[regionRandom.Next(0, regions.Length)];
            location.Facility = facilities[facilityRandom.Next(0, facilities.Length)];

            return location;
        }
        
    }

    public class COVIDLocation
    {
        public string System { get; set; }
        public string Region { get; set; }
        public string Facility { get; set; }
    }


}