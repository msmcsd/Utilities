using System;
using System.Diagnostics;
using UIAutomationClient;
using TreeScope = UIAutomationClient.TreeScope;

namespace ManageBrowserTabs
{
    internal class GetChromeTabs
    {
        private readonly CUIAutomation _automation;

        public GetChromeTabs()
        {
            _automation = new CUIAutomation();
            _automation.AddFocusChangedEventHandler(null, new FocusChangeHandler(this));
        }

        public class FocusChangeHandler : IUIAutomationFocusChangedEventHandler
        {
            private readonly GetChromeTabs _listener;

            public FocusChangeHandler(GetChromeTabs listener)
            {
                _listener = listener;
            }

            public void HandleFocusChangedEvent(IUIAutomationElement element)
            {
                if (element != null)
                {
                    Process[] processes = Process.GetProcessesByName("chrome");
                    if (processes.Length <= 0)
                    {
                        Console.WriteLine("Chrome is not running");
                    }
                    else
                    {
                        foreach (Process process in processes)
                        {
                            try
                            {
                                IUIAutomationElement elm = this._listener._automation.ElementFromHandle(process.MainWindowHandle);
                                IUIAutomationCondition Cond = this._listener._automation.CreatePropertyCondition(30003, 50004);
                                IUIAutomationElementArray elm2 = elm.FindAll(TreeScope.TreeScope_Descendants, Cond);
                                for (int i = 0; i < elm2.Length; i++)
                                {
                                    IUIAutomationElement elm3 = elm2.GetElement(i);
                                    IUIAutomationValuePattern val = (IUIAutomationValuePattern)elm3.GetCurrentPattern(10002);
                                    if (val.CurrentValue != "")
                                    {
                                        Console.WriteLine("URL found: " + val.CurrentValue);
                                    }
                                }
                            }
                            catch { }
                        }

                    }
                }
            }
        }
    }
}
