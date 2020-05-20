using Microsoft.Dynamics365.UIAutomation.Browser;
using Microsoft.PowerApps.UIAutomation.Api;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Configuration;
using System.Collections.Generic;
using System;
using System.Diagnostics;
using OpenQA.Selenium;
using Microsoft.PowerApps.UIAutomation.Api.Controls;
using System.Linq;
using Microsoft.Dynamics365.UIAutomation.Api.UCI;
using System.Security;

namespace Microsoft.PowerApps.UIAutomation.Sample.EmergencyResponseApp
{
    [TestClass]
    public class EmergencyResponseApp
    {
        // Test Scenarios
        //https://dev.azure.com/dynamicscrm/OneCRM/_queries/query/?tempQueryId=cbadfbdc-08dc-4474-8b2f-ea0554b38fe2

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
                COVIDLocation location = SetNewLocation(appBrowser);

                //Click Next Button
                appBrowser.Player.ClickButton("Next");

                //Click icon/facility name
                appBrowser.Player.ClickGalleryItem(EmergencyResponseAppNames.Supplies);

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
                COVIDLocation location = SetNewLocation(appBrowser);

                //Click Next Button
                appBrowser.Player.ClickButton("Next");

                //Click icon/facility name
                appBrowser.Player.ClickGalleryItem(EmergencyResponseAppNames.Supplies, true);

                //Enter Numbers values in the fields and click Send
                SetGalleryItemValuesForSupplies(appBrowser);

                //Click Submit
                appBrowser.Player.ClickButton("Submit");
                appBrowser.ThinkTime(2000);

                //Click Home
                appBrowser.Player.ClickButton("Home");
                appBrowser.ThinkTime(2000);

                //Check the inserted values are correctly opening the entity in the Admin App

                //Click on facility Icon
                appBrowser.Player.ReturnToFacilityMenu(location.System);

                //Change System/Region/Facility and click Next
                COVIDLocation location2 = SetNewLocation(appBrowser);

                //Click Next Button
                appBrowser.Player.ClickButton("Next");

                //Click Supplies App
                appBrowser.Player.ClickGalleryItem(EmergencyResponseAppNames.Supplies, true);

                //Click the Feedback Icon
                ProvideFeedback(appBrowser);

                appBrowser.Player.ClickButton("Submit");

                appBrowser.ThinkTime(1000);
            }
        }

        [TestMethod]
        public void UseStaffPlusEquipmentApp()
        {

            var username = ConfigurationManager.AppSettings["OnlineUsername"];
            var password = ConfigurationManager.AppSettings["OnlinePassword"];
            var uri = new System.Uri("https://web.powerapps.com");
            var appId = ConfigurationManager.AppSettings["AppId"];

            using (var appBrowser = new PowerAppBrowser(TestSettings.Options))
            {
                Uri appsPage = new Uri(ConfigurationManager.AppSettings["OnlineUrl"]);

                appBrowser.OnlineLogin.PowerAppsLogin(appsPage, username.ToSecureString(), null);

                //Launch Emergency Response App
                COVIDLocation covidLocation = SetNewLocation(appBrowser);

                //Click Next Button
                appBrowser.Player.ClickButton("Next");

                //Select Staff + equipment App
                appBrowser.Player.ClickGalleryItem(EmergencyResponseAppNames.Staff, true);

                //Enter values in the fields
                SetGalleryItemValuesForStaff(appBrowser);

                //Click Submit
                appBrowser.Player.ClickButton("Submit");
                appBrowser.ThinkTime(2000);

                //Check the inserted values are correct by opening the Entity in the Admin App

                //Click Home
                appBrowser.Player.ClickButton("Home");
                appBrowser.ThinkTime(2000);

                //Click the Facility Icon
                appBrowser.Player.ReturnToFacilityMenu(covidLocation.System);

                //Change the System/REgion/Facility and click Next
                COVIDLocation location2 = SetNewLocation(appBrowser);

                //Click Next Button
                appBrowser.Player.ClickButton("Next");

                //Select Staff + equipment App
                appBrowser.Player.ClickGalleryItem(EmergencyResponseAppNames.Staff, true);

                //Click the feedback Icon
                ProvideFeedback(appBrowser);

                appBrowser.Player.ClickButton("Submit");

                appBrowser.ThinkTime(1000);
            }
        }

        [TestMethod]
        public void UseDischargePlanningApp()
        {
            var username = ConfigurationManager.AppSettings["OnlineUsername"];
            var password = ConfigurationManager.AppSettings["OnlinePassword"];
            var uri = new System.Uri("https://web.powerapps.com");
            var appId = ConfigurationManager.AppSettings["AppId"];

            using (var appBrowser = new PowerAppBrowser(TestSettings.Options))
            {
                Uri appsPage = new Uri(ConfigurationManager.AppSettings["OnlineUrl"]);

                //Launch the App
                appBrowser.OnlineLogin.PowerAppsLogin(appsPage, username.ToSecureString(), null);

                //Pick the DDL Lists
                COVIDLocation location = SetNewLocation(appBrowser);

                //Click Next Button
                appBrowser.Player.ClickButton("Next");

                //Select Discharge Planning App
                appBrowser.Player.ClickGalleryItem(EmergencyResponseAppNames.DischargePlanning, true);

                //Fill out the forms and click Send
                // Write this Next
                //
                //
                //
                SetGalleryItemValuesForDischargePlanning(appBrowser);

                //Check the Inserted Values

                //Select a different system/location/facility
                appBrowser.Player.ReturnToFacilityMenu(location.System);

                //Choose New Location
                COVIDLocation location2 = SetNewLocation(appBrowser);

                //Click Next Button
                appBrowser.Player.ClickButton("Next");

                //Reopen the App
                appBrowser.Player.ClickGalleryItem(EmergencyResponseAppNames.DischargePlanning);

                //Click the feedback Icon
                ProvideFeedback(appBrowser);

                appBrowser.Player.ClickButton("Submit");

                appBrowser.ThinkTime(1000);

            }
        }

        [TestMethod]
        public void UseCovidStatsApp()
        {
            var username = ConfigurationManager.AppSettings["OnlineUsername"];
            var password = ConfigurationManager.AppSettings["OnlinePassword"];
            var uri = new System.Uri("https://web.powerapps.com");
            var appId = ConfigurationManager.AppSettings["AppId"];

            using (var appBrowser = new PowerAppBrowser(TestSettings.Options))
            {
                Uri appsPage = new Uri(ConfigurationManager.AppSettings["OnlineUrl"]);

                //Launch the App
                appBrowser.OnlineLogin.PowerAppsLogin(appsPage, username.ToSecureString(), null);

                //Pick the DDL Lists
                COVIDLocation location = SetNewLocation(appBrowser);

                //Click Next Button
                appBrowser.Player.ClickButton("Next");

                //Select Discharge Planning App
                appBrowser.Player.ClickGalleryItem(EmergencyResponseAppNames.COVIDStats, true);

                //Fill out the form and click Send
                SetValuesForCOVIDStats(appBrowser);

                appBrowser.Player.ClickButton("Submit");
                appBrowser.ThinkTime(2000);

                //Click the Home button
                appBrowser.Player.ClickButton("Home");

                //Check the Inserted Values

                //Click Home

                //Select a different system/location/facility
                appBrowser.Player.ReturnToFacilityMenu(location.System);

                //Choose New Location
                COVIDLocation location2 = SetNewLocation(appBrowser);

                //Click Next Button
                appBrowser.Player.ClickButton("Next");

                //Reopen the App
                appBrowser.Player.ClickGalleryItem(EmergencyResponseAppNames.COVIDStats, true);

                //Click the feedback Icon
                ProvideFeedback(appBrowser);

                appBrowser.Player.ClickButton("Submit");

                appBrowser.ThinkTime(1000);
            }

        }

        [TestMethod]
        public void UseStaffingNeedsApp()
        {
            var username = ConfigurationManager.AppSettings["OnlineUsername"];
            var password = ConfigurationManager.AppSettings["OnlinePassword"];
            var uri = new System.Uri("https://web.powerapps.com");
            var appId = ConfigurationManager.AppSettings["AppId"];

            using (var appBrowser = new PowerAppBrowser(TestSettings.Options))
            {
                Uri appsPage = new Uri(ConfigurationManager.AppSettings["OnlineUrl"]);

                //Launch the App
                appBrowser.OnlineLogin.PowerAppsLogin(appsPage, username.ToSecureString(), null);

                //Pick the DDL Lists
                COVIDLocation location = SetNewLocation(appBrowser);

                //Click Next Button
                appBrowser.Player.ClickButton("Next");

                //Select Staffing Needs App
                appBrowser.Player.ClickGalleryItem(EmergencyResponseAppNames.StaffingNeeds, true);

                //Fill out the form and click Send
                SetValuesForStaffingNeeds(appBrowser);

                appBrowser.Player.ClickButton("Submit");
                appBrowser.ThinkTime(2000);

                //Check the Inserted Values

                //Click the Home button
                appBrowser.Player.ClickButton("Home");

                //Select a different system/location/facility
                appBrowser.Player.ReturnToFacilityMenu(location.System);

                //Choose New Location
                COVIDLocation location2 = SetNewLocation(appBrowser);

                //Click Next Button
                appBrowser.Player.ClickButton("Next");

                //Reopen the App
                appBrowser.Player.ClickGalleryItem(EmergencyResponseAppNames.Staff, true);

                //Click the feedback Icon
                ProvideFeedback(appBrowser);

                appBrowser.Player.ClickButton("Submit");

                appBrowser.ThinkTime(1000);
            }

        }

        [TestMethod]
        public void UseBedCapacityApp()
        {
            var username = ConfigurationManager.AppSettings["OnlineUsername"];
            var password = ConfigurationManager.AppSettings["OnlinePassword"];
            var uri = new System.Uri("https://web.powerapps.com");
            var appId = ConfigurationManager.AppSettings["AppId"];

            using (var appBrowser = new PowerAppBrowser(TestSettings.Options))
            {
                Uri appsPage = new Uri(ConfigurationManager.AppSettings["OnlineUrl"]);

                //Launch the App
                appBrowser.OnlineLogin.PowerAppsLogin(appsPage, username.ToSecureString(), null);

                //Pick the DDL Lists
                COVIDLocation location = SetNewLocation(appBrowser);

                //Click Next Button
                appBrowser.Player.ClickButton("Next");

                //Select Staffing Needs App
                appBrowser.Player.ClickGalleryItem(EmergencyResponseAppNames.BedCapacity, true);

                //Fill out the form and click Send
                SetBedCapacityValues(appBrowser);

                appBrowser.Player.ClickButton("Submit");
                appBrowser.ThinkTime(2000);

                //Check the Inserted Values

                //Click the Home button
                appBrowser.Player.ClickButton("Home");

                //Select a different system/location/facility
                appBrowser.Player.ReturnToFacilityMenu(location.System);

                //Choose New Location
                COVIDLocation location2 = SetNewLocation(appBrowser);

                //Click Next Button
                appBrowser.Player.ClickButton("Next");

                //Reopen the App
                appBrowser.Player.ClickGalleryItem(EmergencyResponseAppNames.BedCapacity, true);

                //Click the feedback Icon
                ProvideFeedback(appBrowser);

                appBrowser.Player.ClickButton("Submit");

                appBrowser.ThinkTime(1000);

            }

        }

        [TestMethod]
        public void UseAdminApp()
        {
            var username = ConfigurationManager.AppSettings["OnlineUsername"];
            var password = ConfigurationManager.AppSettings["OnlinePassword"];
            var uri = new System.Uri("https://web.powerapps.com");
            var appId = ConfigurationManager.AppSettings["AppId"];

            WebClient client = new WebClient(TestSettings.Options);

            using (var appBrowser = new XrmApp(client))
            {
                Uri adminPage = new Uri("https://org322c1f7a.crm.dynamics.com/main.aspx?appid=3bb0a59f-4c76-ea11-a811-000d3a569e35&pagetype=entitylist&etn=msft_system&viewid=18461f05-4cb5-4ec7-b979-bc392ac0db79&viewType=1039");
                appBrowser.OnlineLogin.Login(adminPage, username.ToSecureString(), String.Empty.ToSecureString(), String.Empty.ToSecureString(), ADFSLoginAction);

                EmergencyResponseAppTestRecord record = new EmergencyResponseAppTestRecord();

                appBrowser.Navigation.OpenSubArea("Locations", "Systems");
                appBrowser.CommandBar.ClickCommand("New");
                appBrowser.Entity.SetValue("msft_systemname", record.SystemName);
                appBrowser.Entity.Save();

                appBrowser.Navigation.OpenSubArea("Regions");
                appBrowser.CommandBar.ClickCommand("New");
                appBrowser.Entity.SetValue(new LookupItem() { Name = "msft_system", Value = record.SystemName });
                appBrowser.Entity.SetValue("msft_regionname", record.RegionName);
                appBrowser.Entity.Save();

                appBrowser.Navigation.OpenSubArea("Facilities");
                appBrowser.CommandBar.ClickCommand("New");
                appBrowser.Entity.SetValue(new LookupItem() { Name = "msft_region", Value = record.RegionName });
                appBrowser.Entity.SetValue("msft_facilityname", record.FacilityName);
                appBrowser.Entity.Save();

                appBrowser.Navigation.OpenSubArea("Locations");
                appBrowser.CommandBar.ClickCommand("New");
                appBrowser.Entity.SetValue(new LookupItem() { Name = "msft_factility", Value = record.FacilityName });
                appBrowser.Entity.SetValue("msft_locationname", record.LocationName);
                appBrowser.Entity.Save();

                appBrowser.Navigation.OpenSubArea("Departments");
                appBrowser.CommandBar.ClickCommand("New");
                appBrowser.Entity.SetValue("msft_departmentname", record.DepartmentName);
                appBrowser.Entity.Save();

            }

        }

        [TestMethod]
        public void UseEquipmentApp()
        {
            var username = ConfigurationManager.AppSettings["OnlineUsername"];
            var password = ConfigurationManager.AppSettings["OnlinePassword"];
            var uri = new System.Uri("https://web.powerapps.com");
            var appId = ConfigurationManager.AppSettings["AppId"];

            using (var appBrowser = new PowerAppBrowser(TestSettings.Options))
            {
                Uri appsPage = new Uri(ConfigurationManager.AppSettings["OnlineUrl"]);

                //Launch the App
                appBrowser.OnlineLogin.PowerAppsLogin(appsPage, username.ToSecureString(), null);

                //Pick the DDL Lists
                COVIDLocation location = SetNewLocation(appBrowser);

                //Click Next Button
                appBrowser.Player.ClickButton("Next");

                //Select Staffing Needs App
                appBrowser.Player.ClickGalleryItem(EmergencyResponseAppNames.Equipment, true);

                //Fill out the form and click Send
                SetEquipmentValues(appBrowser);

                appBrowser.Player.ClickButton("Submit");
                appBrowser.ThinkTime(2000);

                //Check the Inserted Values

                //Click the Home button
                appBrowser.Player.ClickButton("Home");

                //Select a different system/location/facility
                appBrowser.Player.ReturnToFacilityMenu(location.System);

                //Choose New Location
                COVIDLocation location2 = SetNewLocation(appBrowser);

                //Click Next Button
                appBrowser.Player.ClickButton("Next");

                //Reopen the App
                appBrowser.Player.ClickGalleryItem(EmergencyResponseAppNames.Equipment, true);

                //Click the feedback Icon
                ProvideFeedback(appBrowser);

                appBrowser.Player.ClickButton("Submit");

                appBrowser.ThinkTime(1000);
            }

        }


        //Helpers
        private COVIDLocation GetNewLocation()
        {
            string[] systems = new string[] { "Alpine", "Blue Yonder", "Fabrikam", "Litware" };
            //string[] systems = new string[] { "Alpine", "Blue Yonder", "Fabrikam", "Litware", "Southridge" };
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
        private COVIDLocation SetNewLocation(PowerAppBrowser appBrowser)
        {
            COVIDLocation location = GetNewLocation();

            DropDownList systemDDL = new DropDownList() { ControlName = "SystemSelect_System_DD", ControlValue = location.System };
            DropDownList regionDDL = new DropDownList() { ControlName = "SystemSelect_Region_DD", ControlValue = String.Format("{0} {1}", location.System, location.Region) };
            DropDownList facilityDDL = new DropDownList() { ControlName = "SystemSelect_Facility_DD", ControlValue = String.Format("{0} {1} {2}", location.System, location.Region, location.Facility) };

            appBrowser.Player.SetValue(systemDDL);
            appBrowser.ThinkTime(100);
            appBrowser.Player.SetValue(regionDDL);
            appBrowser.ThinkTime(100);
            appBrowser.Player.SetValue(facilityDDL);
            appBrowser.ThinkTime(100);

            return location;
        }
        private void ProvideFeedback(PowerAppBrowser appBrowser)
        {
            //Click the feedback Icon
            appBrowser.Player.ClickFeedbackIcon();

            //Fill in the feedback and Click Send
            ComboBox topicCombo = new ComboBox() { ControlName = "DataCardValue21", ControlValue = "Idea" };
            appBrowser.Player.SetValue(topicCombo);

            string feedbackString = "This app is fantastic";
            TextArea comments = new TextArea() { ControlName = "DataCardValue22", ControlValue = feedbackString };
            appBrowser.Player.SetValue(comments);

            appBrowser.Player.ClickButton("Submit");
        }
        public void ADFSLoginAction(Dynamics365.UIAutomation.Api.UCI.LoginRedirectEventArgs args)
        {
            //Wait for CRM Page to load
            args.Driver.WaitUntilVisible(By.XPath(Dynamics365.UIAutomation.Api.UCI.Elements.Xpath[Dynamics365.UIAutomation.Api.UCI.Reference.Login.CrmUCIMainPage]),
                TimeSpan.FromSeconds(60),
                e =>
                {
                    args.Driver.WaitForPageToLoad();
                    args.Driver.SwitchTo().Frame(0);
                    args.Driver.WaitForPageToLoad();
                }
            );
        }

        private void SetGalleryItemValuesForSupplies(PowerAppBrowser appBrowser)
        {
            var rows = appBrowser.Driver.FindElements(By.XPath("//*[contains(@data-control-part,'gallery-item')]"));

            int cnt = 1;

            foreach (var row in rows)
            {
                var rowLabel = row.FindElement(By.XPath(".//div[contains(@data-control-name,'lblSupplyName')]")).Text;

                row.FindElement(By.XPath(".//div[contains(@data-control-name,'txtSuppliesOnHand')]")).FindElement(By.TagName("input")).SendKeys(cnt++.ToString());
                row.FindElement(By.XPath(".//div[contains(@data-control-name,'txtSupplyBurnRate')]")).FindElement(By.TagName("input")).SendKeys(cnt++.ToString());

            }

            appBrowser.ThinkTime(100);
        }

        private void SetGalleryItemValuesForStaff(PowerAppBrowser appBrowser)
        {
            int cnt = 1;

            SetGalleryValueForStaff(appBrowser, "Number of patients", "inputValue", cnt + 15.ToString());
            SetGalleryValueForStaff(appBrowser, "Partners", "txtValue", cnt++.ToString());
            SetGalleryValueForStaff(appBrowser, "Requested", "txtValue", cnt++.ToString());
            SetGalleryValueForStaff(appBrowser, "Assigned", "txtValue", cnt++.ToString());
            SetGalleryValueForStaff(appBrowser, "Unassigned", "txtValue", cnt++.ToString());

            appBrowser.ThinkTime(100);
        }
        private void SetGalleryValueForStaff(PowerAppBrowser appBrowser, string rowLabel, string controlName, string textBoxValue)
        {
            var rows = appBrowser.Driver.FindElements(By.XPath("//*[contains(@data-control-part,'gallery-item')]"));

            foreach (var row in rows)
            {
                //var thisRowLabel = row.FindElement(By.XPath(".//div[contains(@data-control-name,'txtlblInline') or contains(@data-control-name,'txtlblAbove')]")).Text;
                var rowLabels = row.FindElements(By.XPath(".//div[contains(@data-control-name,'txtlblInline') or contains(@data-control-name,'txtlblAbove')]"));

                foreach (var thisRowLabel in rowLabels)
                {
                    if (thisRowLabel.Text.Equals(rowLabel, StringComparison.OrdinalIgnoreCase))
                    {
                        var txtBoxContainer = row.FindElement(By.XPath($".//div[contains(@data-control-name,'{ controlName }')]"));
                        txtBoxContainer.FindElement(By.TagName("input")).SendKeys(textBoxValue);
                        break;
                    }
                }
            }
        }

        private void SetGalleryItemValuesForDischargePlanning(PowerAppBrowser appBrowser)
        {
            int cnt = 1;

            SetGalleryValueForDischarge(appBrowser, "Authorization", "txtValue", cnt++.ToString());
            SetGalleryValueForDischarge(appBrowser, "Durable medical equipment", "txtValue", cnt++.ToString());
            SetGalleryValueForDischarge(appBrowser, "Guardianship", "txtValue", cnt++.ToString());
            SetGalleryValueForDischarge(appBrowser, "Home + Community Services", "txtValue", cnt++.ToString());
            SetGalleryValueForDischarge(appBrowser, "Placement", "txtValue", cnt++.ToString());
            SetGalleryValueForDischarge(appBrowser, "Skilled nursing facility", "txtValue", cnt++.ToString());
            SetGalleryValueForDischarge(appBrowser, "Past 24 h", "txtValue", cnt++.ToString());
            SetGalleryValueForDischarge(appBrowser, "Likely next 24 h", "txtValue", cnt++.ToString());

            appBrowser.ThinkTime(100);
        }
        private void SetGalleryValueForDischarge(PowerAppBrowser appBrowser, string rowLabel, string controlName, string textBoxValue)
        {
            var rows = appBrowser.Driver.FindElements(By.XPath("//*[contains(@data-control-part,'gallery-item')]"));

            foreach (var row in rows)
            {
                //var thisRowLabel = row.FindElement(By.XPath(".//div[contains(@data-control-name,'txtlblInline') or contains(@data-control-name,'txtlblAbove')]")).Text;
                var rowLabels = row.FindElements(By.XPath(".//div[contains(@data-control-name,'txtlbl')]"));

                foreach (var thisRowLabel in rowLabels)
                {
                    if (thisRowLabel.Text.Equals(rowLabel, StringComparison.OrdinalIgnoreCase))
                    {
                        var txtBoxContainer = row.FindElement(By.XPath($".//div[contains(@data-control-name,'{ controlName }')]"));
                        txtBoxContainer.FindElement(By.TagName("input")).SendKeys(textBoxValue);
                        break;
                    }
                }
            }
        }

        private void SetValuesForCOVIDStats(PowerAppBrowser appBrowser)
        {
            //Click the arrow
            appBrowser.Driver.FindElement(By.XPath("//div[contains(@data-control-name,'drpLocation')]")).FindElement(By.XPath("//div[contains(@class,'dropdownLabelArrow')]")).Click(true);

            var dropDownContainer = appBrowser.Driver.FindElements(By.XPath("//div[contains(@role,'listbox')]"));

            appBrowser.ThinkTime(250);

            if (dropDownContainer.Count > 0)
            {
                var listItems = dropDownContainer[0].FindElements(By.XPath("//div[contains(@role,'option')]"));

                if (listItems.Count > 0)
                {
                    if (listItems.Count == 1)
                        listItems[0].Click(true);
                    else
                        listItems[listItems.Count - 1].Click(true);
                }
            }

            int cnt = 1;

            SetCOVIDStatsValue(appBrowser, "PUIs", "txtPUI", cnt++.ToString());
            SetCOVIDStatsValue(appBrowser, "Positive", "txtPOS", cnt++.ToString());
            SetCOVIDStatsValue(appBrowser, "Intubated", "txtIntubated", cnt++.ToString());
            SetCOVIDStatsValue(appBrowser, "Discharged", "txtDischarged", cnt++.ToString());

            appBrowser.ThinkTime(100);
        }
        private void SetCOVIDStatsValue(PowerAppBrowser appBrowser, string rowLabel, string controlName, string textBoxValue)
        {
            appBrowser.Driver.FindElement(By.XPath($"//div[contains(@data-control-name,'{ controlName }')]")).FindElement(By.TagName("input")).SendKeys(textBoxValue);
            appBrowser.ThinkTime(100);
        }

        private void SetValuesForStaffingNeeds(PowerAppBrowser appBrowser)
        {
            int cnt = 1;

            SetStaffingNeedsValue_ComboBox(appBrowser, "Department", "DataCardValue2", "Surgical ICU");
            SetStaffingNeedsValue_TextBox(appBrowser, "Department location", "DataCardValue6", "7th floor");

            SetStaffingNeedsValue_ComboBox(appBrowser, "Request type", "DataCardValue7", "Clinical");
            SetStaffingNeedsValue_ComboBox(appBrowser, "Role needed", "DataCardValue8", "RN");
            SetStaffingNeedsValue_ComboBox(appBrowser, "Needed now or next shift", "DataCardValue1", "Upcoming Shift");

            SetStaffingNeedsValue_TextBox(appBrowser, "How many", "DataCardValue3", cnt++.ToString());
            SetStaffingNeedsValue_TextArea(appBrowser, "Details", "DataCardValue10", "Need additional staff in this location");

            appBrowser.ThinkTime(100);
        }
        private void SetStaffingNeedsValue_TextBox(PowerAppBrowser appBrowser, string rowLabel, string controlName, string textBoxValue)
        {
            TextBox textBox = new TextBox() { ControlName = controlName, ControlValue = textBoxValue };
            appBrowser.Player.SetValue(textBox);
        }
        private void SetStaffingNeedsValue_ComboBox(PowerAppBrowser appBrowser, string rowLabel, string controlName, string comboBoxValue)
        {
            ComboBox combo = new ComboBox() { ControlName = controlName, ControlValue = comboBoxValue };
            appBrowser.Player.SetValue(combo);

            appBrowser.ThinkTime(100);
        }
        private void SetStaffingNeedsValue_TextArea(PowerAppBrowser appBrowser, string rowLabel, string controlName, string textBoxValue)
        {
            TextArea textArea = new TextArea() { ControlName = controlName, ControlValue = textBoxValue };
            appBrowser.Player.SetValue(textArea);
        }

        private void SetValue_RadioButton(PowerAppBrowser appBrowser, string controlName, string radioValueToClick)
        {
            RadioButton radio = new RadioButton() { ControlName = controlName, ControlValue = radioValueToClick };
            appBrowser.Player.SetValue(radio);
        }
        private void SetBedCapacityValues(PowerAppBrowser browser)
        {
            int i = 0;

            SetStaffingNeedsValue_TextBox(browser, "LicensedBeds", "DataCardValue8_1", i++.ToString());
            SetStaffingNeedsValue_TextBox(browser, "# of ICU Beds (AIIR)", "DataCardValue9_1", i++.ToString());
            SetStaffingNeedsValue_TextBox(browser, "# of ICU Beds (non-AIIR)", "DataCardValue10_1", i++.ToString());
            SetStaffingNeedsValue_TextBox(browser, "# Acute Beds (AIIR)", "DataCardValue11_1", i++.ToString());
            SetStaffingNeedsValue_TextBox(browser, "# of Acute Beds (non-AIIR)", "DataCardValue1", i++.ToString());
            SetValue_RadioButton(browser, "Radio4", "Yes");
            SetValue_RadioButton(browser, "Radio4_1", "Yes");
            SetStaffingNeedsValue_TextBox(browser, "# of Surge Beds", "DataCardValue16_1", i++.ToString());
            SetStaffingNeedsValue_TextBox(browser, "# of Pediatric ICU Beds", "DataCardValue13", i++.ToString());
            SetStaffingNeedsValue_TextBox(browser, "# of Pediatric Acute Beds", "DataCardValue7", i++.ToString());
        }

        private void SetEquipmentValues(PowerAppBrowser browser)
        {
            var rows = browser.Driver.FindElements(By.XPath("//*[contains(@data-control-part,'gallery-item')]"));

            int cnt = 1;

            foreach (var row in rows)
            {
                var rowLabel = row.FindElement(By.XPath(".//div[contains(@data-control-name,'txtlblInline')]")).Text;

                if (!String.IsNullOrEmpty(rowLabel))
                    row.FindElement(By.XPath(".//div[contains(@data-control-name,'txtValue')]")).FindElement(By.TagName("input")).SendKeys(cnt++.ToString());
            }

            browser.ThinkTime(100);
        }
    }

    public class COVIDLocation
    {
        public string System { get; set; }
        public string Region { get; set; }
        public string Facility { get; set; }
    }

    public static class EmergencyResponseAppNames
    {
        public static string Supplies = "Supplies";
        public static string Staff = "Staff";
        public static string DischargePlanning = "Discharge planning";
        public static string BedCapacity = "Bed capacity";
        public static string COVIDStats = "COVID-19 stats";
        public static string Equipment = "Equipment";
        public static string StaffingNeeds = "Staffing needs";

    }

    public class EmergencyResponseAppTestRecord
    {
        private string rootName = "z_EasyRepro";
        private static char[] Alphabet = ("ABCDEFGHIJKLMNOPQRSTUVWXYZabcefghijklmnopqrstuvwxyz0123456789").ToCharArray();
        private static Object m_randLock = new object();
        private static Random m_randomInstance = new Random();

        public EmergencyResponseAppTestRecord()
        {
            this.SystemName = String.Format("{0}_System_{1}", rootName, getRandomString(5, 7));
            this.RegionName = String.Format("{0}_Region_{1}", rootName, getRandomString(5, 7));
            this.FacilityName = String.Format("{0}_Facility_{1}", rootName, getRandomString(5, 7));
            this.LocationName = String.Format("{0}_Location_{1}", rootName, getRandomString(5, 7));
            this.DepartmentName = String.Format("{0}_Department_{1}", rootName, getRandomString(5, 7));
        }

        public string SystemName { get; set; }
        public string RegionName { get; set; }
        public string FacilityName { get; set; }
        public string LocationName { get; set; }
        public string DepartmentName { get; set; }


        private string getRandomString(int minLen, int maxLen)
        {
            int alphabetLength = Alphabet.Length;
            int stringLength;
            lock (m_randLock) { stringLength = m_randomInstance.Next(minLen, maxLen); }
            char[] str = new char[stringLength];

            // max length of the randomizer array is 5
            int randomizerLength = (stringLength > 5) ? 5 : stringLength;

            int[] rndInts = new int[randomizerLength];
            int[] rndIncrements = new int[randomizerLength];

            // Prepare a "randomizing" array
            for (int i = 0; i < randomizerLength; i++)
            {
                int rnd = m_randomInstance.Next(alphabetLength);
                rndInts[i] = rnd;
                rndIncrements[i] = rnd;
            }

            // Generate "random" string out of the alphabet used
            for (int i = 0; i < stringLength; i++)
            {
                int indexRnd = i % randomizerLength;
                int indexAlphabet = rndInts[indexRnd] % alphabetLength;
                str[i] = Alphabet[indexAlphabet];

                // Each rndInt "cycles" characters from the array, 
                // so we have more or less random string as a result
                rndInts[indexRnd] += rndIncrements[indexRnd];
            }
            return (new string(str));
        }
    }


}