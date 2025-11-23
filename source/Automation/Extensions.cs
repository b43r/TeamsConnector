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

using Interop.UIAutomationClient;

namespace TeamsConnector.Automation
{
    /// <summary>
    /// Extension methods for the IUIAutomationElement class.
    /// </summary>
    internal static class Extensions
    {
        private static CUIAutomationClass automation = new CUIAutomationClass();

        private const int WaitTimeoutDelay = 100;

        /// <summary>
        /// Expand the element if it supports the ExpandCollapse pattern.
        /// </summary>
        /// <param name="button">IUIAutomationElement</param>
        /// <returns>true if successful</returns>
        public static bool Expand(this IUIAutomationElement button)
        {
            var pat = button.GetCurrentPattern(UIA_PatternIds.UIA_ExpandCollapsePatternId) as IUIAutomationExpandCollapsePattern;
            if (pat != null && pat.CurrentExpandCollapseState != ExpandCollapseState.ExpandCollapseState_Expanded)
            {
                pat.Expand();
                return true;
            }

            return false;
        }

        /// <summary>
        /// Collapse the element if it supports the ExpandCollapse pattern.
        /// </summary>
        /// <param name="button">IUIAutomationElement</param>
        /// <returns>true if successful</returns>
        public static bool Collapse(this IUIAutomationElement button)
        {
            var pat = button.GetCurrentPattern(UIA_PatternIds.UIA_ExpandCollapsePatternId) as IUIAutomationExpandCollapsePattern;
            if (pat != null && pat.CurrentExpandCollapseState != ExpandCollapseState.ExpandCollapseState_Collapsed)
            {
                pat.Collapse();
                return true;
            }

            return false;
        }

        /// <summary>
        /// Invoke the element if it supports the Invoke pattern.
        /// </summary>
        /// <param name="button">IUIAutomationElement</param>
        /// <returns>true if successful</returns>
        public static bool Invoke(this IUIAutomationElement button)
        {
            var pat = button.GetCurrentPattern(UIA_PatternIds.UIA_InvokePatternId) as IUIAutomationInvokePattern;
            if (pat != null)
            {
                pat.Invoke();
                return true;
            }

            return false;
        }

        /// <summary>
        /// Dump all object in the array to the debug console.
        /// </summary>
        /// <param name="array">IUIAutomationElementArray</param>
        public static void DumpToConsole(this IUIAutomationElementArray array)
        {
            for (int i = 0; i < array.Length; i++)
            {
                array.GetElement(i).DumpToConsole();
            }
        }

        /// <summary>
        /// Dump the object to the debug console.
        /// </summary>
        /// <param name="element">IUIAutomationElement</param>
        public static void DumpToConsole(this IUIAutomationElement element)
        {
            System.Diagnostics.Debug.WriteLine($"Id={element.CurrentAutomationId}, Name={element.CurrentName}, Type={element.CurrentLocalizedControlType} ({element.CurrentControlType}), Class={element.CurrentClassName}");
        }

        /// <summary>
        /// Find a child element with the given id.
        /// </summary>
        /// <param name="parent">IUIAutomationElement</param>
        /// <param name="id">automation id</param>
        /// <param name="waitTimeout">number of milliseconds to wait for the element to appear</param>
        /// <returns>IUIAutomationElement or null</returns>
        public static IUIAutomationElement? FindById(this IUIAutomationElement parent, string id, int waitTimeout = 0)
        {
            int tries = (int)Math.Ceiling((decimal)waitTimeout / WaitTimeoutDelay) + 1;
            IUIAutomationElement? result = null;
            var condition = automation.CreatePropertyCondition(UIA_PropertyIds.UIA_AutomationIdPropertyId, id);
            while (result == null && tries > 0)
            {
                result = parent.FindFirst(TreeScope.TreeScope_Subtree, condition);
                tries--;

                if (result == null && tries > 0)
                {
                    System.Threading.Thread.Sleep(WaitTimeoutDelay);
                }
            }

            return result;
        }

        /// <summary>
        /// Find the first child element that matches the given search criteria.
        /// </summary>
        /// <param name="parent">IUIAutomationElement</param>
        /// <param name="controlType">control type</param>
        /// <param name="className">part of class name or null</param>
        /// <param name="name">part of name or null</param>
        /// <param name="waitTimeout">number of milliseconds to wait for the element to appear</param>
        /// <returns>IUIAutomationElement or null</returns>
        public static IUIAutomationElement? FindFirst(this IUIAutomationElement parent, ControlType controlType, string? className, string? name, int waitTimeout = 0)
        {
            int tries = (int)Math.Ceiling((decimal)waitTimeout / WaitTimeoutDelay) + 1;
            IUIAutomationElement? result = null;
            var condition = CreateCondition(controlType, className, name);
            while (result == null && tries > 0)
            {
                result = parent.FindFirst(TreeScope.TreeScope_Subtree, condition);
                tries--;

                if (result == null && tries > 0)
                {
                    System.Threading.Thread.Sleep(WaitTimeoutDelay);
                }
            }

            return result;
        }

        /// <summary>
        /// Find all child elements that matches the given search criteria.
        /// </summary>
        /// <param name="parent">IUIAutomationElement</param>
        /// <param name="controlType">control type</param>
        /// <param name="className">part of class name or null</param>
        /// <param name="name">part of name or null</param>
        /// <param name="waitTimeout">number of milliseconds to wait for the elements to appear</param>
        /// <returns>IUIAutomationElementArray or null</returns>
        public static IUIAutomationElementArray? FindAll(this IUIAutomationElement parent, ControlType controlType, string? className, string? name, int waitTimeout = 0)
        {
            int tries = (int)Math.Ceiling((decimal)waitTimeout / WaitTimeoutDelay) + 1;
            IUIAutomationElementArray? result = null;
            var condition = CreateCondition(controlType, className, name);
            while ((result == null || result.Length == 0) && tries > 0)
            {
                result = parent?.FindAll(TreeScope.TreeScope_Subtree, condition);
                tries--;

                if ((result == null || result.Length == 0) && tries > 0)
                {
                    System.Threading.Thread.Sleep(WaitTimeoutDelay);
                }
            }

            return result;
        }

        /// <summary>
        /// Create a condition based on the given search criteria.
        /// </summary>
        /// <param name="controlType">controlType</param>
        /// <param name="className">part of class name or null</param>
        /// <param name="name">part of name or null</param>
        /// <returns>IUIAutomationCondition</returns>
        private static IUIAutomationCondition CreateCondition(ControlType controlType, string? className, string? name)
        {
            List<IUIAutomationCondition> conditions = new List<IUIAutomationCondition>();
            if (controlType != ControlType.All)
            {
                conditions.Add(automation.CreatePropertyCondition(UIA_PropertyIds.UIA_ControlTypePropertyId, (int)controlType));
            }

            if (!string.IsNullOrEmpty(className))
            {
                conditions.Add(automation.CreatePropertyConditionEx(UIA_PropertyIds.UIA_ClassNamePropertyId, className, PropertyConditionFlags.PropertyConditionFlags_MatchSubstring));
            }

            if (!string.IsNullOrEmpty(name))
            {
                conditions.Add(automation.CreatePropertyConditionEx(UIA_PropertyIds.UIA_NamePropertyId, name, PropertyConditionFlags.PropertyConditionFlags_MatchSubstring));
            }

            return conditions.Any() ? automation.CreateAndConditionFromArray(conditions.ToArray()) : automation.CreateTrueCondition();
        }
    }
}
