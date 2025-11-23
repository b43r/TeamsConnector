/*
 * TeamsConnector
 * 
 * MIT License
 * 
 * Copyright (C) 2023 by Simon Baer
 * 
 * Permission is hereby granted, free of charge, to any person obtaining a copy
 * of this software and associated documentation files (the "Software"), to deal
 * in the Software without restriction, including without limitation the rights
 * to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 * copies of the Software, and to permit persons to whom the Software is
 * furnished to do so, subject to the following conditions:
 * 
 * The above copyright notice and this permission notice shall be included in all
 * copies or substantial portions of the Software.
 * 
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 * IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 * FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 * AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 * LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 * OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
 * SOFTWARE.
 */

using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;

using Interop.UIAutomationClient;

namespace TeamsConnector.Automation
{
    internal class Automation
    {
        private delegate bool EnumWindowsProc(IntPtr hWnd, int lParam);

        [DllImport("user32.dll")]
        private static extern int GetWindowText(IntPtr hWnd, StringBuilder text, int count);

        [DllImport("user32.dll")]
        private static extern bool EnumWindows(EnumWindowsProc enumFunc, int lParam);

        [DllImport("user32.dll")]
        private static extern int GetWindowTextLength(IntPtr hWnd);

        /// <summary>
        /// Event that is raised on incoming calls.
        /// </summary>
        public event EventHandler<IncomingCallEventArgs>? IncomingCall;

        private IUIAutomation automation;

        private const int UIA_Window_WindowOpenedEventId = 20016;
        private const string TeamsNotificationWindowTitle = "Microsoft Teams";

        public Automation()
        {
            automation = new CUIAutomationClass();
            SubscribeEvent();
        }

        public void Dispose()
        {
            UnsubsribeEvent();
        }

        /// <summary>
        /// Subscribe for the 'WindowOpened' UI Automation event.
        /// </summary>
        /// <returns>true if successful</returns>
        private bool SubscribeEvent()
        {
            try
            {
                IUIAutomationElement rootElement = automation.GetRootElement();
                CUIAutomationEventHandler cUIAutomationEventHandler = new CUIAutomationEventHandler();
                cUIAutomationEventHandler.UIAutomationEvent += CUIAutomationEventHandler_UIAutomationEvent;
                automation.AddAutomationEventHandler(UIA_Window_WindowOpenedEventId, rootElement, TreeScope.TreeScope_Subtree, null, cUIAutomationEventHandler);
                Logger.Log("Automation events subscribed.");
                return true;
            }
            catch (Exception arg)
            {
                Logger.Log($"Subscribe exception: {arg}");
                return false;
            }
        }

        /// <summary>
        /// Unsubscribe from all UI Automation events.
        /// </summary>
        /// <returns>true if successful</returns>
        private bool UnsubsribeEvent()
        {
            try
            {
                automation.RemoveAllEventHandlers();
                return true;
            }
            catch (Exception arg)
            {
                Logger.Log($"Unsubscribe exception: {arg}");
                return false;
            }
        }

        /// <summary>
        /// Handle the UI Automation event.
        /// </summary>
        /// <param name="src">object</param>
        /// <param name="args">UIAutomationEventArgs</param>
        private void CUIAutomationEventHandler_UIAutomationEvent(object? src, UIAutomationEventArgs args)
        {
            IUIAutomationElement? iUIAutomationElement = src as IUIAutomationElement;
            if (iUIAutomationElement != null)
            {
                Logger.Log($"UIAutomation event: {iUIAutomationElement.CurrentName}, EventId: {args.EventId}");
                if (args.EventId == UIA_Window_WindowOpenedEventId && iUIAutomationElement.CurrentName.Contains(TeamsNotificationWindowTitle))
                {
                    // on the old Teams client the phone number may not be available from the beginning, so try again for up to 1 second
                    int retryCount = 10;
                    string phoneNumberFromNotifyWindow;
                    while ((phoneNumberFromNotifyWindow = ExtractPhoneNumber(iUIAutomationElement, 0)) == string.Empty && retryCount > 0)
                    {
                        retryCount--;
                        Thread.Sleep(100);
                    }

                    if (retryCount < 10)
                    {
                        Logger.Log($"Delay: {(10 - retryCount) * 100} ms.");
                    }

                    if (!string.IsNullOrEmpty(phoneNumberFromNotifyWindow))
                    {
                        Logger.Log($"Incoming call from: {phoneNumberFromNotifyWindow}");
                        IncomingCall?.Invoke(this, new IncomingCallEventArgs(phoneNumberFromNotifyWindow));
                    }
                }
            }
        }

        /// <summary>
        /// Try to extract a phone number from a UI element and all its descendents.
        /// </summary>
        /// <param name="rootElement">IUIAutomationElement</param>
        /// <param name="level">recursion level</param>
        /// <returns>phone number or empty string</returns>
        private string ExtractPhoneNumber(IUIAutomationElement rootElement, int level)
        {
            string text = string.Empty;
            try
            {
                IUIAutomationTreeWalker controlViewWalker = automation.ControlViewWalker;
                for (IUIAutomationElement iUIAutomationElement = controlViewWalker.GetFirstChildElement(rootElement); iUIAutomationElement != null; iUIAutomationElement = controlViewWalker.GetNextSiblingElement(iUIAutomationElement))
                {
                    if (Regex.IsMatch(iUIAutomationElement.CurrentName, "^\\+[\\d\\s]+"))
                    {
                        Logger.Log($"{level}:{new string('.', level)}:{iUIAutomationElement.CurrentName}");
                        return Regex.Replace(iUIAutomationElement.CurrentName, "[\\s\\(\\)]", "");
                    }
                    text = ExtractPhoneNumber(iUIAutomationElement, level++);
                    if (!string.IsNullOrEmpty(text))
                    {
                        return text;
                    }
                }
                return text;
            }
            catch (Exception arg)
            {
                Logger.Log($"ExtractPhoneNumber exception: {arg}");
                return text;
            }
        }

        /// <summary>
        /// Returns the UIAutomationElement of the window with the given title.
        /// </summary>
        /// <param name="titleRegex">regular expression that matches the window title</param>
        /// <returns>IUIAutomationElement or null</returns>
        public IUIAutomationElement? GetWindowByTitle(string titleRegex)
        {
            var hwnd = GetWindowHandle(titleRegex);
            if (hwnd != IntPtr.Zero)
            {
                return automation.ElementFromHandle(hwnd);
            }

            return null;
        }

        /// <summary>
        /// Returns the window handle of the window with the given title.
        /// </summary>
        /// <param name="titleRegex">regular expression that matches the window title</param>
        /// <returns>window handle or IntPtr.Zero</returns>
        private static IntPtr GetWindowHandle(string titleRegex)
        {
            var re = new Regex(titleRegex);

            IntPtr result = IntPtr.Zero;
            EnumWindows(delegate (IntPtr hWnd, int lParam)
            {
                int windowTextLength = GetWindowTextLength(hWnd);
                if (windowTextLength == 0)
                {
                    return true;
                }

                StringBuilder text = new StringBuilder(windowTextLength);
                GetWindowText(hWnd, text, windowTextLength + 1);
                if (re.IsMatch(text.ToString()))
                {
                    result = hWnd;
                    return false;
                }

                return true;
            }, 0);

            return result;
        }
    }
}