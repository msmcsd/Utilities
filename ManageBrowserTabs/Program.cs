using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
using UIAutomationClient;
using TreeScope = UIAutomationClient.TreeScope;


namespace ManageBrowserTabs
{
    static class Program
    {
        const int UIA_ValuePatternId = 10002;               // https://learn.microsoft.com/en-us/windows/win32/winauto/uiauto-controlpattern-ids
        const int UIA_LegacyIAccessiblePatternId = 10018;   // https://learn.microsoft.com/en-us/windows/win32/winauto/uiauto-controlpattern-ids
        const int UIA_ControlTypePropertyId = 30003;        // https://learn.microsoft.com/en-us/windows/win32/winauto/uiauto-automation-element-propids
        const int UIA_EditControlTypeId = 50004;            // https://learn.microsoft.com/en-us/windows/win32/winauto/uiauto-controltype-ids
        const int UIA_TabItemControlTypeId = 50019;         // https://learn.microsoft.com/en-us/windows/win32/winauto/uiauto-controltype-ids

        private static string urlFile = Path.Combine(Environment.GetEnvironmentVariable("USERPROFILE") + "\\Desktop", "urls.bat");

        static void Main(string[] args)
        {
            if (args.Length == 0)
            {
                Console.WriteLine($"{Assembly.GetExecutingAssembly().GetName().Name} [--save] [--open]\n");
                Console.WriteLine("  --save -s\tSave urls of open tabs in Chrome.");
                Console.WriteLine("  --open -o\tOpen Run urls.bat to open Chrome with urls in the file.");
                return;
            }

            switch (args[0].ToLower())
            {
                case "--save":
                case "-s":
                    SaveOpenTabsUrls();
                    break;
                case "--open":
                case "-o":
                    OpenChromeWithTabs();
                    break;
            }
        }

        static void OpenChromeWithTabs()
        {
            if (!File.Exists(urlFile))
            {
                return;
            }

            Process p = new Process
            {
                StartInfo = new ProcessStartInfo
                {
                    FileName = urlFile,
                    CreateNoWindow = false
                }
            };
            
            p.Start();
        }

        static void SaveOpenTabsUrls()
        {
            var urls = GeteOpenTabsUrls();
            
            if (urls != null) SaveUrls(urls);
        }

        static void SaveUrls(Dictionary<string, string> urls)
        {
            if (urls == null)
            {
                Console.WriteLine("Url collection is null.");
                return;
            }

            StringBuilder sb = new StringBuilder();
            foreach (var url in urls)
            {
                if (!string.IsNullOrEmpty(url.Value))
                {
                    sb.Append($"REM {url.Key}\n");
                    sb.Append($"start https://{url.Value}\n\n");
                }
            }

            if (sb.Length > 0)
            {
                File.WriteAllText(urlFile, sb.ToString());
                Console.Write($"Urls saved to {urlFile}");
            }
        }

        static Dictionary<string, string> GeteOpenTabsUrls()
        {
            Process[] mainChromes = Process.GetProcessesByName("chrome").Where(p => !string.IsNullOrEmpty(p.MainWindowTitle)).ToArray();
            if (mainChromes.Length == 0)
            {
                Console.WriteLine("No Chrome instance found.");
                return null;
            }

            var uiAutomation = new CUIAutomation();
            IUIAutomationElement chromeMainUIAElement = uiAutomation.ElementFromHandle(mainChromes[0].MainWindowHandle);

            // Return a collection of Chrome tabs.
            IUIAutomationCondition chromeTabCondition = uiAutomation.CreatePropertyCondition(UIA_ControlTypePropertyId, UIA_TabItemControlTypeId);
            IUIAutomationElementArray chromeTabCollection = chromeMainUIAElement.FindAll(TreeScope.TreeScope_Descendants, chromeTabCondition);

            IUIAutomationElement addressBarEditControl = null;
            Dictionary<string, string> urls = new Dictionary<string, string>();

            // Loop thru each tab to selet tab and get its url.
            for (int i = 0; i < chromeTabCollection.Length; i++)
            {
                // Activate the tab.
                var lp = chromeTabCollection.GetElement(i).GetCurrentPattern(UIA_LegacyIAccessiblePatternId) as IUIAutomationLegacyIAccessiblePattern;
                lp.DoDefaultAction();

                string tabName = chromeTabCollection.GetElement(i).CurrentName;
                Console.WriteLine($"Activated tab '{tabName}'");

                // Sleep after tab switch so the address can be refreshed with new url.
                Thread.Sleep(1000);

                // Retrieve the URL
                if (addressBarEditControl == null)
                {
                    addressBarEditControl = GetAddressBarEditControl(uiAutomation, chromeMainUIAElement);
                }
                //IUIAutomationCondition getEditControl = uiAutomation.CreatePropertyCondition(UIA_ControlTypePropertyId, UIA_EditControlTypeId);
                //IUIAutomationElement addressBarEditControl = chromeMainUIAElement.FindFirst(TreeScope.TreeScope_Descendants, getEditControl);
                IUIAutomationValuePattern addressBar = (IUIAutomationValuePattern)addressBarEditControl.GetCurrentPattern(UIA_ValuePatternId);
                if (addressBar.CurrentValue != "")
                {
                    Console.WriteLine("URL found: " + addressBar.CurrentValue);
                    urls.Add(tabName, addressBar.CurrentValue);
                }
                else
                {
                    Console.WriteLine($"URL not found for '{chromeTabCollection.GetElement(i).CurrentName}'");
                    urls.Add(tabName, string.Empty);
                }
            }

            return urls;
        }

        static IUIAutomationElement GetAddressBarEditControl(CUIAutomation uiAutomation, IUIAutomationElement chromeMainUIAElement)
        {
            IUIAutomationCondition getEditControl = uiAutomation.CreatePropertyCondition(UIA_ControlTypePropertyId, UIA_EditControlTypeId);
            IUIAutomationElement addressBarEditControl = chromeMainUIAElement.FindFirst(TreeScope.TreeScope_Descendants, getEditControl);

            return addressBarEditControl;
        }
    }
}
