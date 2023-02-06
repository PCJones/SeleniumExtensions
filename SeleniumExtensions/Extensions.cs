using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using System.Diagnostics;
using System.Runtime.CompilerServices;

namespace SeleniumExtensions
{
    public static class SeleniumExtensions
    {
        #region Constructor & Internal Methods (Logic)
        public static ClickType DefaultClickType { get; private set; }
        public static ScrollBehaviour DefaultScrollBehaviour { get; private set; }
        public static SimulateInputType DefaultSimulateInputType { get; private set; }
        public static SimulateInputBehaviour DefaultSimulateInputBehaviour { get; private set; }
        public static TimeSpan DefaultWaitForPageContainsStringTimeout { get; private set; }
        public static TimeSpan DefaultWaitForElementExistsTimeout { get; private set; }
        public static TimeSpan DefaultWaitForElementDisplayedTimeout { get; private set; }

        // TODO: implement method overloads that use defaultBy
        //public static ByMethods DefaultBy { get; private set; }

        static SeleniumExtensions()
        {
            DefaultClickType = ClickType.Simple;
            DefaultScrollBehaviour = ScrollBehaviour.None;
            //DefaultBy = ByMethods.Id;
            DefaultSimulateInputType = SimulateInputType.SendKeys;
            DefaultSimulateInputBehaviour = SimulateInputBehaviour.None;
            DefaultWaitForElementExistsTimeout = new TimeSpan(0, 0, 30);
            DefaultWaitForElementDisplayedTimeout = new TimeSpan(0, 0, 30);
            DefaultWaitForPageContainsStringTimeout = new TimeSpan(0, 0, 30);
        }
        public static void SetDefaultClickType(ClickType clickType)
        {
            DefaultClickType = clickType;
        }
        public static void SetDefaultScrollBehaviour(ScrollBehaviour scrollBehaviour)
        {
            DefaultScrollBehaviour = scrollBehaviour;
        }
        //public static void SetDefaultBy(ByMethods by)
        //{
        //    DefaultBy = by;
        //}
        public static void SetSimulateInputType(SimulateInputType simulateInputType)
        {
            DefaultSimulateInputType = simulateInputType;
        }
        public static void SetSimulateInputBehaviour(SimulateInputBehaviour simulateInputBehaviour)
        {
            DefaultSimulateInputBehaviour = simulateInputBehaviour;
        }
        public static void SetDefaultWaitForElementExistsTimeout(TimeSpan timeSpan)
        {
            DefaultWaitForElementExistsTimeout = timeSpan;
        }
        public static void SetDefaultWaitForElementExistsShownTimeout(TimeSpan timeSpan)
        {
            DefaultWaitForElementDisplayedTimeout = timeSpan;
        }
        public static void SetDefaultWaitForPageContainsStringTimeout(TimeSpan timeSpan)
        {
            DefaultWaitForPageContainsStringTimeout = timeSpan;
        }

        #endregion

        #region Public Selenium Extensions

        #region ScriptExecutor
        /// <summary>
        /// Provides quick access to an IJavaScriptExecutor
        /// </summary>
        /// <returns>Returns an IJavaScriptExecutor</returns>
        public static IJavaScriptExecutor ScriptExecutor(this IWebDriver driver)
        {
            return (IJavaScriptExecutor)driver;
        }
        #endregion
        #region ExecuteJavaSript
        // TOOD: add method overloading
        // TODO: Write return documentation; check if return is null if nothing get's returned by the script
        /// <summary>
        /// Executes JavaScript on the current page
        /// </summary>
        /// <param name="script">JavaScript to execute</param>
        /// <param name="arguments">Provide IWebElements accessible by using arguments[n] in the script</param>
        /// <returns></returns>
        public static object ExecuteJavaScript(this IWebDriver driver, string script, params IWebElement[] arguments)
        {
            return driver.ScriptExecutor().ExecuteScript(script, arguments);
        }

        #endregion
        #region ScrollToElement
        // TOOD: Add Documentation
        public static void ScrollToElement(this IWebDriver driver, IWebElement element, ScrollBehaviour scrollBehaviour)
        {
            switch (scrollBehaviour)
            {
                case ScrollBehaviour.None:
                    return;
                case ScrollBehaviour.SeleniumScrollHack:
                    // evaluating the LocationOnScreenOnceScrolledIntoView property forces selenium to scroll to the element.
                    // TODO: Verify if this still works, if not find an alternative and/or remove this
                    OpenQA.Selenium.Remote.RemoteWebElement remoteElement = (OpenQA.Selenium.Remote.RemoteWebElement)element;

                    System.Drawing.Point position = remoteElement.LocationOnScreenOnceScrolledIntoView;
                    break;
                case ScrollBehaviour.JavaScriptScrollIntoView:
                    driver.ExecuteJavaScript("arguments[0].scrollIntoView();", element);
                    break;
                default:
                    break;
            }
        }

        #region Method Overloading
        // TOOD: Add documentation
        public static void ScrollToElement(this IWebDriver driver, By by, ScrollBehaviour scrollBehaviour, int elementIndex = 0)
        {
            var element = driver.FindElements(by)[elementIndex];
            ScrollToElement(driver, element, scrollBehaviour);
        }

        // TOOD: Add documentation
        public static void ScrollToElement(this IWebDriver driver, By by, int elementIndex = 0)
        {
            var element = driver.FindElements(by)[elementIndex];
            ScrollToElement(driver, element, DefaultScrollBehaviour);
        }
        #endregion
        #endregion
        #region SimulateJavaScriptEvent
        /// <summary>
        /// Simulate an event on an object
        /// </summary>
        /// <param name="element">Affected element</param>
        /// <param name="eventName">Event to simulate (e.g. click, dbclick, mouse(over), mouse(down))</param>
        public static void SimulateJavaScriptEvent(this IWebDriver driver, IWebElement element, string eventName)
        {
            var script = $"{Constants.JavaScriptEventSimulationCode} simulate(arguments[0], '{eventName}');";
            driver.ExecuteJavaScript(script, element);
        }

        #region Method Overloading
        /// <summary>
        /// Simulate an event over an object
        /// </summary>
        /// <param name="by">By</param>
        /// <param name="element">Affected element</param>
        /// <param name="eventName">Event to simulate (e.g. click, dbclick, mouse(over), mouse(down))</param>
        /// <param name="elementIndex">Index of the element if multiple can be found with the by selector</param>
        public static void SimulateJavaScriptEvent(this IWebDriver driver, By by, string eventName, int elementIndex = 0)
        {
            var element = driver.FindElements(by)[elementIndex];
            SimulateJavaScriptEvent(driver, element, eventName);
        }
        #endregion

        #endregion
        #region ClickElement
        /// <summary>
        /// Clicks on an element in the active frame
        /// </summary>
        /// <param name="element">Element to click</param>
        /// <param name="clickType">Type of click simulation</param>
        /// <param name="scrollBehaviour">Scroll behaviour before the click is being performed</param>
        public static void ClickElement(this IWebDriver driver, IWebElement element, ClickType clickType, ScrollBehaviour scrollBehaviour)
        {
            ScrollToElement(driver, element, scrollBehaviour);
            switch (clickType)
            {
                case ClickType.Simple:
                    element.Click();
                    break;
                case ClickType.MouseAction:
                    driver.MouseActionClick(element);
                    break;
                case ClickType.JavaScriptSimple:
                    driver.ExecuteJavaScript("arguments[0].click();", element);
                    break;
                case ClickType.JavaScriptEventSimulation:
                    driver.SimulateJavaScriptEvent(element, "click");
                    break;
                default:
                    break;
            }
        }

        #region Method overloading
        /// <summary>
        /// Clicks on an element in the active frame
        /// </summary>
        /// <param name="by">By</param>
        /// <param name="clickType">Type of click simulation</param>
        /// <param name="scrollBehaviour">Scroll behaviour before the click is being performed</param>
        /// <param name="elementIndex">Index of the element if multiple can be found with the by selector</param>        
        public static void ClickElement(this IWebDriver driver, By by, ClickType clickType, ScrollBehaviour scrollBehaviour, int elementIndex = 0)
        {
            var element = driver.FindElements(by)[elementIndex];
            ClickElement(driver, element, clickType, scrollBehaviour);
        }
        /// <summary>
        /// Clicks on an element in the active frame
        /// </summary>
        /// <param name="by">By</param>
        /// <param name="clickType">Type of click simulation</param>
        /// <param name="elementIndex">Index of the element if multiple can be found with the by selector</param>
        public static void ClickElement(this IWebDriver driver, By by, ClickType clickType, int elementIndex = 0)
        {
            ClickElement(driver, by, clickType, DefaultScrollBehaviour, elementIndex);
        }
        /// <summary>
        /// Clicks on an element in the active frame
        /// </summary>
        /// <param name="by">By</param>
        /// <param name="elementIndex">Index of the element if multiple can be found with the by selector</param>
        public static void ClickElement(this IWebDriver driver, By by, int elementIndex = 0)
        {
            ClickElement(driver, by, DefaultClickType, DefaultScrollBehaviour, elementIndex);
        }

        /// <summary>
        /// Clicks on an element in the active frame
        /// </summary>
        /// <param name="element">Element to click</param>
        /// <param name="clickType">Type of click simulation</param>
        public static void ClickElement(this IWebDriver driver, IWebElement element, ClickType clickType)
        {
            ClickElement(driver, element, clickType, DefaultScrollBehaviour);
        }
        /// <summary>
        /// Clicks on an element in the active frame
        /// </summary>
        /// <param name="element">Element to click</param>
        public static void ClickElement(this IWebDriver driver, IWebElement element)
        {
            ClickElement(driver, element, DefaultClickType, DefaultScrollBehaviour);
        }
        #endregion
        #endregion
        #region ClickAllElements
        /// <summary>
        /// Clicks on an all supplied elements
        /// </summary>
        /// <param name="elements">Elements to click</param>
        /// <param name="clickType">Type of click simulation</param>
        /// <param name="scrollBehaviour">Scroll behaviour before the click is being performed</param>
        /// <param name="displayedElementsOnly">If set to true only displayed (non invisible) elements will be clicked</param>
        /// <param name="sleepAfterEachClick">Sleep time in milliseconds after every click</param>
        public static async Task ClickAllElementsAsync(this IWebDriver driver, IEnumerable<IWebElement> elements, ClickType clickType, ScrollBehaviour scrollBehaviour, bool displayedElementsOnly = false, int sleepAfterEachClick = 0)
        {
            foreach (var element in elements)
            {
                if (displayedElementsOnly && !element.Displayed)
                {
                    continue;
                }
                ClickElement(driver, element, clickType, scrollBehaviour);
                await Task.Delay(sleepAfterEachClick);
            }
        }

        /// <summary>
        /// Clicks on an all found elements
        /// </summary>
        /// <param name="by">By selector to find the elements</param>
        /// <param name="clickType">Type of click simulation</param>
        /// <param name="scrollBehaviour">Scroll behaviour before the click is being performed</param>
        /// <param name="displayedElementsOnly">If set to true only displayed (non invisible) elements will be clicked</param>
        /// <param name="sleepAfterEachClick">Sleep time in milliseconds after every click</param>
        public static async Task ClickAllElementsAsync(this IWebDriver driver, By by, ClickType clickType, ScrollBehaviour scrollBehaviour, bool displayedElementsOnly = false, int sleepAfterEachClick = 0)
        {
            var elements = driver.FindElements(by);
            foreach (var element in elements)
            {
                if (displayedElementsOnly && !element.Displayed)
                {
                    continue;
                }
                ClickElement(driver, element, clickType, scrollBehaviour);
                await Task.Delay(sleepAfterEachClick);
            }
        }

        #region Method Overloading
        /// <summary>
        /// Clicks on an all supplied elements
        /// </summary>
        /// <param name="elements">Elements to click</param>
        /// <param name="displayedElementsOnly">If set to true only displayed (non invisible) elements will be clicked</param>
        /// <param name="sleepAfterEachClick">Sleep time in milliseconds after every click</param>
        public static async Task ClickAllElementsAsync(this IWebDriver driver, IEnumerable<IWebElement> elements, bool displayedElementsOnly = false, int sleepAfterEachClick = 0)
        {
            await ClickAllElementsAsync(driver, elements, DefaultClickType, DefaultScrollBehaviour, displayedElementsOnly, sleepAfterEachClick);
        }

        /// <summary>
        /// Clicks on an all supplied elements
        /// </summary>
        /// <param name="elements">Elements to click</param>
        /// <param name="clickType">Type of click simulation</param>
        /// <param name="scrollBehaviour">Scroll behaviour before the click is being performed</param>
        /// <param name="displayedElementsOnly">If set to true only displayed (non invisible) elements will be clicked</param>
        /// <param name="sleepAfterEachClick">Sleep time in milliseconds after every click</param>
        public static void ClickAllElements(this IWebDriver driver, IEnumerable<IWebElement> elements, ClickType clickType, ScrollBehaviour scrollBehaviour, bool displayedElementsOnly = false, int sleepAfterEachClick = 0)
        {
            ClickAllElementsAsync(driver, elements, clickType, scrollBehaviour, displayedElementsOnly, sleepAfterEachClick).Wait();
        }

        /// <summary>
        /// Clicks on an all supplied elements
        /// </summary>
        /// <param name="elements">Elements to click</param>
        /// <param name="displayedElementsOnly">If set to true only displayed (non invisible) elements will be clicked</param>
        /// <param name="sleepAfterEachClick">Sleep time in milliseconds after every click</param>
        public static void ClickAllElements(this IWebDriver driver, IEnumerable<IWebElement> elements, bool displayedElementsOnly = false, int sleepAfterEachClick = 0)
        {
            ClickAllElementsAsync(driver, elements, displayedElementsOnly, sleepAfterEachClick).Wait();
        }

        /// <summary>
        /// Clicks on an all found elements
        /// </summary>
        /// <param name="by">By selector to find the elements</param>
        /// <param name="displayedElementsOnly">If set to true only displayed (non invisible) elements will be clicked</param>
        /// <param name="sleepAfterEachClick">Sleep time in milliseconds after every click</param>
        public static async Task ClickAllElementsAsync(this IWebDriver driver, By by, bool displayedElementsOnly = false, int sleepAfterEachClick = 0)
        {
            await ClickAllElementsAsync(driver, by, DefaultClickType, DefaultScrollBehaviour, displayedElementsOnly, sleepAfterEachClick);
        }

        /// <summary>
        /// Clicks on an all found elements
        /// </summary>
        /// <param name="by">By selector to find the elements</param>
        /// <param name="clickType">Type of click simulation</param>
        /// <param name="scrollBehaviour">Scroll behaviour before the click is being performed</param>
        /// <param name="displayedElementsOnly">If set to true only displayed (non invisible) elements will be clicked</param>
        /// <param name="sleepAfterEachClick">Sleep time in milliseconds after every click</param>
        public static void ClickAllElements(this IWebDriver driver, By by, ClickType clickType, ScrollBehaviour scrollBehaviour, bool displayedElementsOnly = false, int sleepAfterEachClick = 0)
        {
            ClickAllElementsAsync(driver, by, clickType, scrollBehaviour, displayedElementsOnly, sleepAfterEachClick).Wait();
        }

        /// <summary>
        /// Clicks on an all found elements
        /// </summary>
        /// <param name="by">By selector to find the elements</param>
        /// <param name="displayedElementsOnly">If set to true only displayed (non invisible) elements will be clicked</param>
        /// <param name="sleepAfterEachClick">Sleep time in milliseconds after every click</param>
        public static void ClickAllElements(this IWebDriver driver, By by, bool displayedElementsOnly = false, int sleepAfterEachClick = 0)
        {
            ClickAllElementsAsync(driver, by, DefaultClickType, DefaultScrollBehaviour, displayedElementsOnly, sleepAfterEachClick).Wait();
        }
        #endregion
        #endregion
        #region GetElementAboveElement
        // TODO: Add method overloading
        // TOOD: Add returns documentation
        /// <summary>
        /// Returns the top most element, which can be the same as the specified one
        /// </summary>
        /// <param name="driver"></param>
        /// <param name="element">The element which has an element above it</param>
        /// <returns></returns>
        public static IWebElement? GetElementAboveElement(this IWebDriver driver, IWebElement element)
        {
            driver.ExecuteJavaScript("arguments[0].scrollIntoView();", element);
            int elementX = element.Location.X - Convert.ToInt32(driver.ExecuteJavaScript("return window.pageXOffset;"));
            int elementY = element.Location.Y - Convert.ToInt32(driver.ExecuteJavaScript("return window.pageYOffset;"));
            var returnElement = (IWebElement)driver.ExecuteJavaScript("return document.elementFromPoint(" + elementX + ", " + elementY + ")");

            return returnElement;
        }
        #endregion
        #region HoverOverElement
        /// <summary>
        /// Hovers over an element via the Actions API
        /// </summary>
        /// <param name="by">By</param>
        /// <param name="elementIndex">Index of the element if more than one can be found with the By selector</param>
        public static void HoverOverElement(this IWebDriver driver, By by, int elementIndex = 0)
        {
            IWebElement element = driver.FindElements(by)[elementIndex];
            Actions actions = new(driver);
            actions.MoveToElement(element).Build().Perform();
        }

        /// <summary>
        /// Hovers over an element via the Actions API
        /// </summary>
        /// <param name="element">Element to hover over</param>
        public static void HoverOverElement(this IWebDriver driver, IWebElement element)
        {
            Actions actions = new(driver);
            actions.MoveToElement(element).Build().Perform();
        }
        #endregion
        #region SubmitForm
        /// <summary>
        /// Submits a form
        /// </summary>
        /// <param name="element">Element to perform the submit action on</param>
        public static void SubmitForm(this IWebDriver driver, IWebElement element)
        {
            element.Submit();
        }

        #region Method Overloading
        /// <summary>
        /// Submits a form
        /// </summary>
        /// <param name="by">By</param>
        /// <param name="elementIndex">Index of the element if more than one can be found with the By selector</param>
        public static void SubmitForm(this IWebDriver driver, By by, int elementIndex = 0)
        {
            var element = driver.FindElements(by)[elementIndex];
            SubmitForm(driver, element);
        }
        #endregion
        #endregion
        #region SimulateInput
        /// <summary>
        /// Sends keystrokes to the provided element. Useful to fill out forms.
        /// </summary>
        /// <param name="element">Element to click</param>
        /// <param name="text">Text/Keystrokes to simulate</param>
        /// <param name="simulateInputType">Type of input simulation (SendKeys = Default behaviour, text is sent instantly. HumanLike = short random breaks after every char.</param>
        /// <param name="simulateInputBehaviour">Specify if there should be an an action before and/or after the input.</param>
        public static async Task SimulateInputAsync(this IWebDriver driver, IWebElement element, string text, SimulateInputType simulateInputType, SimulateInputBehaviour simulateInputBehaviour)
        {
            switch (simulateInputBehaviour)
            {
                case SimulateInputBehaviour.None:
                    break;
                case SimulateInputBehaviour.ClearAtStart:
                    element.Clear();
                    break;
                case SimulateInputBehaviour.TabAtEnd:
                    text += Keys.Tab;
                    break;
                case SimulateInputBehaviour.ClearAtStart_TabAtEnd:
                    element.Clear();
                    text += Keys.Tab;
                    break;
                default:
                    break;
            }

            switch (simulateInputType)
            {
                case SimulateInputType.SendKeys:
                    element.SendKeys(text);
                    break;
                case SimulateInputType.HumanLike:
                    foreach (char character in text)
                    {
                        element.SendKeys(character.ToString());
                        var sleepTime = Constants.Random.Next(Constants.HUMANLIKE_INPUT_MIN_SLEEP, Constants.HUMANLIKE_INPUT_MAX_SLEEP);
                        await Task.Delay(sleepTime);
                    }
                    break;
                default:
                    break;
            }
        }

        #region Method Overloading
        /// <summary>
        /// Sends keystrokes to the provided element. Useful to fill out forms.
        /// </summary>
        /// <param name="element">Element to click</param>
        /// <param name="text">Text/Keystrokes to simulate</param>
        /// <param name="simulateInputType">Type of input simulation (SendKeys = Default behaviour, text is sent instantly. HumanLike = short random breaks after every char.</param>
        /// <param name="simulateInputBehaviour">Specify if there should be an an action before and/or after the input.</param>
        public static void SimulateInput(this IWebDriver driver, IWebElement element, string text, SimulateInputType simulateInputType, SimulateInputBehaviour simulateInputBehaviour)
        {
            SimulateInputAsync(driver, element, text, simulateInputType, simulateInputBehaviour).Wait();
        }

        /// <summary>
        /// Sends keystrokes to the provided element. Useful to fill out forms.
        /// </summary>
        /// <param name="element">Element to click</param>
        /// <param name="text">Text/Keystrokes to simulate</param>
        /// <param name="simulateInputType">Type of input simulation (SendKeys = Default behaviour, text is sent instantly. HumanLike = short random breaks after every char.</param>
        public static void SimulateInput(this IWebDriver driver, IWebElement element, string text, SimulateInputType simulateInputType)
        {
            SimulateInputAsync(driver, element, text, simulateInputType, DefaultSimulateInputBehaviour).Wait();
        }

        /// <summary>
        /// Sends keystrokes to the provided element. Useful to fill out forms.
        /// </summary>
        /// <param name="element">Element to click</param>
        /// <param name="text">Text/Keystrokes to simulate</param>
        /// <param name="simulateInputType">Type of input simulation (SendKeys = Default behaviour, text is sent instantly. HumanLike = short random breaks after every char.</param>
        public static async Task SimulateInputAsync(this IWebDriver driver, IWebElement element, string text, SimulateInputType simulateInputType)
        {
            await SimulateInputAsync(driver, element, text, simulateInputType, DefaultSimulateInputBehaviour);
        }

        /// <summary>
        /// Sends keystrokes to the provided element. Useful to fill out forms.
        /// </summary>
        /// <param name="by">By</param>
        /// <param name="text">Text/Keystrokes to simulate</param>
        /// <param name="simulateInputType">Type of input simulation (SendKeys = Default behaviour, text is sent instantly. HumanLike = short random breaks after every char.</param>
        /// <param name="simulateInputBehaviour">Specify if there should be an an action before and/or after the input.</param>
        /// <param name="elementIndex">Index of the element if multiple can be found with the by selector</param>
        public static void SimulateInput(this IWebDriver driver, By by, string text, SimulateInputType simulateInputType, SimulateInputBehaviour simulateInputBehaviour, int elementIndex = 0)
        {
            var element = driver.FindElements(by)[elementIndex];
            SimulateInputAsync(driver, element, text, simulateInputType, simulateInputBehaviour).Wait();
        }

        /// <summary>
        /// Sends keystrokes to the provided element. Useful to fill out forms.
        /// </summary>
        /// <param name="by">By</param>
        /// <param name="text">Text/Keystrokes to simulate</param>
        /// <param name="simulateInputType">Type of input simulation (SendKeys = Default behaviour, text is sent instantly. HumanLike = short random breaks after every char.</param>
        /// <param name="simulateInputBehaviour">Specify if there should be an an action before and/or after the input.</param>
        /// <param name="elementIndex">Index of the element if multiple can be found with the by selector</param>
        public static async Task SimulateInputAsync(this IWebDriver driver, By by, string text, SimulateInputType simulateInputType, SimulateInputBehaviour simulateInputBehaviour, int elementIndex = 0)
        {
            var element = driver.FindElements(by)[elementIndex];
            await SimulateInputAsync(driver, element, text, simulateInputType, simulateInputBehaviour);
        }

        /// <summary>
        /// Sends keystrokes to the provided element. Useful to fill out forms.
        /// </summary>
        /// <param name="by">By</param>
        /// <param name="text">Text/Keystrokes to simulate</param>
        /// <param name="simulateInputType">Type of input simulation (SendKeys = Default behaviour, text is sent instantly. HumanLike = short random breaks after every char.</param>
        /// <param name="elementIndex">Index of the element if multiple can be found with the by selector</param>
        public static void SimulateInput(this IWebDriver driver, By by, string text, SimulateInputType simulateInputType, int elementIndex = 0)
        {
            var element = driver.FindElements(by)[elementIndex];
            SimulateInputAsync(driver, element, text, simulateInputType, DefaultSimulateInputBehaviour).Wait();
        }

        /// <summary>
        /// Sends keystrokes to the provided element. Useful to fill out forms.
        /// </summary>
        /// <param name="by">By</param>
        /// <param name="text">Text/Keystrokes to simulate</param>
        /// <param name="simulateInputType">Type of input simulation (SendKeys = Default behaviour, text is sent instantly. HumanLike = short random breaks after every char.</param>
        /// <param name="elementIndex">Index of the element if multiple can be found with the by selector</param>
        public static async Task SimulateInputAsync(this IWebDriver driver, By by, string text, SimulateInputType simulateInputType, int elementIndex = 0)
        {
            var element = driver.FindElements(by)[elementIndex];
            await SimulateInputAsync(driver, element, text, simulateInputType, DefaultSimulateInputBehaviour);
        }

        /// <summary>
        /// Sends keystrokes to the provided element. Useful to fill out forms.
        /// </summary>
        /// <param name="by">By</param>
        /// <param name="text">Text/Keystrokes to simulate</param>
        /// <param name="elementIndex">Index of the element if multiple can be found with the by selector</param>
        public static void SimulateInput(this IWebDriver driver, By by, string text, int elementIndex = 0)
        {
            var element = driver.FindElements(by)[elementIndex];
            SimulateInputAsync(driver, element, text, DefaultSimulateInputType, DefaultSimulateInputBehaviour).Wait();
        }

        /// <summary>
        /// Sends keystrokes to the provided element. Useful to fill out forms.
        /// </summary>
        /// <param name="by">By</param>
        /// <param name="text">Text/Keystrokes to simulate</param>
        /// <param name="elementIndex">Index of the element if multiple can be found with the by selector</param>
        public static async Task SimulateInputAsync(this IWebDriver driver, By by, string text, int elementIndex = 0)
        {
            var element = driver.FindElements(by)[elementIndex];
            await SimulateInputAsync(driver, element, text, DefaultSimulateInputType, DefaultSimulateInputBehaviour);
        }
        #endregion
        #endregion
        #region WaitForPageContainsString
        /// <summary>
        /// Waits until the page source code is containing a specific string. 
        /// </summary>
        /// <param name="text">String the page source code needs to contain</param>
        /// <param name="timeout">Maximum wait time</param>
        /// <returns>Returns true if the text is found, returns false if the timeout is reached.</returns>
        public static async Task<bool> WaitForPageContainsStringAsync(this IWebDriver driver, string text, TimeSpan timeout)
        {
            var maxIterations = timeout.Seconds * 4;
            // Waite till the page source contains a given peace of text
            for (int i = 0; i < maxIterations; i++)
            {
                try
                {
                    if (driver.PageSource.Contains(text))
                    {
                        return true;
                    }
                }
                catch (Exception)
                {
                    continue;
                }
                finally
                {
                    await Task.Delay(250);
                }
            }
            return false;
        }

        /// <summary>
        /// Waits until the page source code is containing a specific string.
        /// </summary>
        /// <param name="text">String the page source code needs to contain</param>
        /// <returns>Returns true if the text is found, returns false if the timeout is reached.</returns>
        public static bool WaitForPageContainsString(this IWebDriver driver, string text)
        {
            return WaitForPageContainsStringAsync(driver, text, DefaultWaitForPageContainsStringTimeout).Result;
        }

        /// <summary>
        /// Waits until the page source code is containing a specific string.
        /// </summary>
        /// <param name="text">String the page source code needs to contain</param>
        /// <returns>Returns true if the text is found, returns false if the timeout is reached.</returns>
        public static async Task<bool> WaitForPageContainsStringAsync(this IWebDriver driver, string text)
        {
            return await WaitForPageContainsStringAsync(driver, text, DefaultWaitForPageContainsStringTimeout);
        }

        /// <summary>
        /// Waits until the page source code is containing a specific string.
        /// </summary>
        /// <param name="text">String the page source code needs to contain</param>
        /// <param name="timespan">Maximum wait time</param>
        /// <returns>Returns true if the text is found, returns false if the timeout is reached.</returns>
        public static bool WaitForPageContainsString(this IWebDriver driver, string text, TimeSpan timeSpan)
        {
            return WaitForPageContainsStringAsync(driver, text, timeSpan).Result;
        }
        #endregion
        #region WaitForElementExists
        /// <summary>
        /// Waits until the active frame is containing an element.
        /// </summary>
        /// <param name="by">By</param>
        /// <param name="timeout">Maximum wait time</param>
        /// <param name="elementIndex">Index of the element if more than one can be found with the By selector</param>
        /// <returns>Returns true when the element is found, returns false when the timeout is reached.</returns>
        public static async Task<bool> WaitForElementExistsAsync(this IWebDriver driver, By by, TimeSpan timeout, int elementIndex = 0)
        {
            var maxIterations = timeout.Seconds * 4;
            for (int i = 0; i < maxIterations; i++)
            {
                try
                {
                    IWebElement element = driver.FindElements(by)[elementIndex];

                    if (element != null)
                    {
                        return true;
                    }
                }
                catch (Exception)
                {
                    continue;
                }
                finally
                {
                    await Task.Delay(250);
                }
            }
            return false;
        }

        #region Method overloading
        /// <summary>
        /// Waits until the active frame is containing an element.
        /// </summary>
        /// <param name="elementIndex">Index of the element if more than one can be found with the By selector</param>
        /// <param name="by">By</param>
        /// <returns>Returns true when the element is found, returns false when the timeout is reached.</returns>
        public static async Task<bool> WaitForElementExistsAsync(this IWebDriver driver, By by, int elementIndex = 0)
        {
            return await WaitForElementExistsAsync(driver, by, DefaultWaitForElementExistsTimeout, elementIndex);
        }

        /// <summary>
        /// Waits until the active frame is containing an element.
        /// </summary>
        /// <param name="elementIndex">Index of the element if more than one can be found with the By selector</param>
        /// <param name="timeout">Maximum wait time</param>
        /// <param name="by">By</param>
        /// <returns>Returns true when the element is found, returns false when the timeout is reached.</returns>
        public static bool WaitForElementExists(this IWebDriver driver, By by, TimeSpan timeout, int elementIndex = 0)
        {
            return WaitForElementExistsAsync(driver, by, timeout, elementIndex).Result;
        }

        /// <summary>
        /// Waits until the active frame is containing an element.
        /// </summary>
        /// <param name="elementIndex">Index of the element if more than one can be found with the By selector</param>
        /// <param name="by">By</param>
        /// <returns>Returns true when the element is found, returns false when the timeout is reached.</returns>
        public static bool WaitForElementExists(this IWebDriver driver, By by, int elementIndex = 0)
        {
            return WaitForElementExistsAsync(driver, by, DefaultWaitForElementExistsTimeout, elementIndex).Result;
        }
        #endregion
        #endregion
        #region WaitForElementDisplayed
        /// <summary>
        /// Waits until the element is being displayed (not hidden).
        /// </summary>
        /// <param name="element">Element to wait for being displayed</param>
        /// <param name="timeout">Maximum wait time</param>
        /// <returns>Returns true when the element is not hidden, returns false when the timeout is reached or the element does not exist.</returns>
        public static async Task<bool> WaitForElementDisplayedAsync(this IWebDriver driver, IWebElement element, TimeSpan timeout)
        {
            var maxIterations = timeout.Seconds * 4;
            for (int i = 0; i < maxIterations; i++)
            {
                if (element.Displayed)
                {
                    return true;
                }
                await Task.Delay(250);
            }
            return false;
        }

        /// <summary>
        /// Waits until the element is being displayed (not hidden).
        /// </summary>
        /// <param name="by">By</param>
        /// <param name="timeout">Maximum wait time</param>
        /// <param name="elementIndex">Index of the element if more than one can be found with the By selector</param>
        /// <returns>Returns true when the element is not hidden, returns false when the timeout is reached or the element does not exist.</returns>
        public static async Task<bool> WaitForElementDisplayedAsync(this IWebDriver driver, By by, TimeSpan timeout, int elementIndex = 0)
        {
            var elements = driver.FindElements(by);
            if (elementIndex + 1 > elements.Count)
            {
                return false;
            }

            return await WaitForElementDisplayedAsync(driver, elements[elementIndex], timeout);
        }

        #region Method overloading
        /// <summary>
        /// Waits until the element is being displayed (not hidden).
        /// </summary>
        /// <param name="element">Element to wait for being displayed</param>
        /// <returns>Returns true when the element is not hidden, returns false when the timeout is reached or the element does not exist.</returns>
        public static async Task<bool> WaitForElementDisplayedAsync(this IWebDriver driver, IWebElement element)
        {
            return await WaitForElementDisplayedAsync(driver, element, DefaultWaitForElementDisplayedTimeout);
        }

        /// <summary>
        /// Waits until the element is being displayed (not hidden).
        /// </summary>
        /// <param name="by">By</param>
        /// <param name="elementIndex">Index of the element if more than one can be found with the By selector</param>
        /// <returns>Returns true when the element is not hidden, returns false when the timeout is reached or the element does not exist.</returns>
        public static async Task<bool> WaitForElementDisplayedAsync(this IWebDriver driver, By by, int elementIndex = 0)
        {
            return await WaitForElementDisplayedAsync(driver, by, DefaultWaitForElementDisplayedTimeout, elementIndex);
        }

        /// <summary>
        /// Waits until the element is being displayed (not hidden).
        /// </summary>
        /// <param name="element">Element to wait for being displayed</param>
        /// <param name="timeout">Maximum wait time</param>
        /// <returns>Returns true when the element is not hidden, returns false when the timeout is reached or the element does not exist.</returns>
        public static bool WaitForElementDisplayed(this IWebDriver driver, IWebElement element, TimeSpan timeout)
        {
            return WaitForElementDisplayedAsync(driver, element, timeout).Result;
        }

        /// <summary>
        /// Waits until the element is being displayed (not hidden).
        /// </summary>
        /// <param name="element">Element to wait for being displayed</param>
        /// <returns>Returns true when the element is not hidden, returns false when the timeout is reached or the element does not exist.</returns>
        public static bool WaitForElementDisplayed(this IWebDriver driver, IWebElement element)
        {
            return WaitForElementDisplayedAsync(driver, element).Result;
        }

        /// <summary>
        /// Waits until the element is being displayed (not hidden).
        /// </summary>
        /// <param name="by">By</param>
        /// <param name="timeout">Maximum wait time</param>
        /// <param name="elementIndex">Index of the element if more than one can be found with the By selector</param>
        /// <returns>Returns true when the element is not hidden, returns false when the timeout is reached or the element does not exist.</returns>
        public static bool WaitForElementDisplayed(this IWebDriver driver, By by, TimeSpan timeout, int elementIndex = 0)
        {
            return WaitForElementDisplayedAsync(driver, by, timeout, elementIndex).Result;
        }

        /// <summary>
        /// Waits until the element is being displayed (not hidden).
        /// </summary>
        /// <param name="by">By</param>
        /// <param name="elementIndex">Index of the element if more than one can be found with the By selector</param>
        /// <returns>Returns true when the element is not hidden, returns false when the timeout is reached or the element does not exist.</returns>
        public static bool WaitForElementDisplayed(this IWebDriver driver, By by, int elementIndex = 0)
        {
            return WaitForElementDisplayedAsync(driver, by, elementIndex).Result;
        }
        #endregion
        #endregion
        #region WaitForElementExistsAndDisplayed
        /// <summary>
        /// Waits until the element exists and is being displayed (not hidden).
        /// </summary>
        /// <param name="By">by</param>
        /// <param name="timeout">Maximum wait time</param>
        /// <param name="elementIndex">Index of the element if more than one can be found with the By selector</param>
        /// <returns>Returns true when the element exists and is not hidden, returns false when the timeout is reached or the element does not exist or is not dispayed.</returns>
        public static async Task<bool> WaitForElementExistsAndDisplayedAsync(this IWebDriver driver, By by, TimeSpan timeout, int elementIndex = 0)
        {
            Stopwatch stopwatch = new Stopwatch();
            stopwatch.Start();
            var elementExists = await WaitForElementExistsAsync(driver, by, timeout, elementIndex).ConfigureAwait(false);
            stopwatch.Stop();

            if (!elementExists)
            {
                return false;
            }

            var element = driver.FindElements(by)[elementIndex];

            var maxIterations = Math.Max(1, (timeout.Seconds - stopwatch.Elapsed.Seconds) * 4);
            for (int i = 0; i < maxIterations; i++)
            {
                if (element.Displayed)
                {
                    return true;
                }
                await Task.Delay(250);
            }
            return false;
        }

        #region Method Overloading
        /// <summary>
        /// Waits until the element exists and is being displayed (not hidden).
        /// </summary>
        /// <param name="By">by</param>
        /// <param name="elementIndex">Index of the element if more than one can be found with the By selector</param>
        /// <returns>Returns true when the element exists and is not hidden, returns false when the timeout is reached or the element does not exist or is not dispayed.</returns>
        public static async Task<bool> WaitForElementExistsAndDisplayedAsync(this IWebDriver driver, By by, int elementIndex = 0)
        {
            var timeout = DefaultWaitForElementExistsTimeout + DefaultWaitForElementDisplayedTimeout;
            return await WaitForElementExistsAndDisplayedAsync(driver, by, timeout, elementIndex).ConfigureAwait(false);
        }

        /// <summary>
        /// Waits until the element exists and is being displayed (not hidden).
        /// </summary>
        /// <param name="By">by</param>
        /// <param name="timeout">Maximum wait time</param>
        /// <param name="elementIndex">Index of the element if more than one can be found with the By selector</param>
        /// <returns>Returns true when the element exists and is not hidden, returns false when the timeout is reached or the element does not exist or is not dispayed.</returns>
        public static bool WaitForElementExistsAndDisplayed(this IWebDriver driver, By by, TimeSpan timeout, int elementIndex = 0)
        {
            return WaitForElementExistsAndDisplayedAsync(driver, by, timeout, elementIndex).Result;
        }

        /// <summary>
        /// Waits until the element exists and is being displayed (not hidden).
        /// </summary>
        /// <param name="By">by</param>
        /// <param name="elementIndex">Index of the element if more than one can be found with the By selector</param>
        /// <returns>Returns true when the element exists and is not hidden, returns false when the timeout is reached or the element does not exist or is not dispayed.</returns>
        public static bool WaitForElementExistsAndDisplayed(this IWebDriver driver, By by, int elementIndex = 0)
        {
            return WaitForElementExistsAndDisplayedAsync(driver, by, elementIndex).Result;
        }
        #endregion
        #endregion
        #region SwitchToFrame
        /// <summary>
        /// Switches the active context to a specified frame / iframe
        /// </summary>
        /// <param name="frameName">Frame name</param>
        public static void SwitchToFrame(this IWebDriver driver, string frameName)
        {
            driver.SwitchTo().Frame(frameName);
        }

        /// <summary>
        /// Switches the active context to a specified frame / iframe
        /// </summary>
        /// <param name="frameIndex">Index of the frame</param>
        public static void SwitchToFrame(this IWebDriver driver, int frameIndex)
        {
            driver.SwitchTo().Frame(frameIndex);
        }

        /// <summary>
        /// Switches the active context to a specified frame / iframe
        /// </summary>
        /// <param name="frameElement">Element of the frame</param>
        public static void SwitchToFrame(this IWebDriver driver, IWebElement frameElement)
        {
            driver.SwitchTo().Frame(frameElement);
        }

        /// <summary>
        /// Switches the active context to a specified frame / iframe
        /// </summary>
        /// <param name="by">Selector of the frame</param>
        /// <param name="elementIndex">Index of the element if more than one can be found with the By selector</param>
        public static void SwitchToFrame(this IWebDriver driver, By by, int elementIndex = 0)
        {
            driver.SwitchTo().Frame(driver.FindElements(by)[elementIndex]);
        }
        #endregion
        #region SwitchToDefaultFrame
        /// <summary>
        /// Switches the active context to the default frame (main page)
        /// </summary>
        public static void SwitchToDefaultFrame(this IWebDriver driver)
        {
            driver.SwitchTo().DefaultContent();
        }
        #endregion
        #region OpenTab
        /// <summary>
        /// Opens a new tab
        /// </summary>
        /// <returns>Returns the window handle of the new tab</returns>
        public static async Task<string> OpenNewTabAsync(this IWebDriver driver)
        {
            var windowCount = driver.WindowHandles.Count;
            driver.ExecuteJavaScript("window.open('');");
            for (int i = 0; i < 10; i++)
            {
                await Task.Delay(100);
                if (driver.WindowHandles.Count > windowCount)
                {
                    break;
                }
            }
            return driver.WindowHandles.Last();
        }

        /// <summary>
        /// Opens a new tab
        /// </summary>
        /// <returns>Returns the window handle of the new tab</returns>
        public static string OpenNewTab(this IWebDriver driver)
        {
            return OpenNewTabAsync(driver).Result;
        }
        #endregion
        #region MakeElementVisible
        /// <summary>
        /// Tries to make a hidden element visible (displayed)
        /// </summary>
        /// <param name="element">Hidden element to display</param>
        public static void MakeElementVisible(this IWebDriver driver, IWebElement element)
        {
            driver.ExecuteJavaScript("arguments[0].style.display='block';", element);
            if (element.GetAttribute("type") == "hidden")
            {
                driver.ExecuteJavaScript("arguments[0].setAttribute('type', '');");
            }
        }

        /// <summary>
        /// Tries to make a hidden element visible (displayed)
        /// </summary>
        /// <param name="by">Selector of the frame</param>
        /// <param name="elementIndex">Index of the element if more than one can be found with the By selector</param>
        public static void MakeElementVisible(this IWebDriver driver, By by, int elementIndex = 0)
        {
            var element = driver.FindElements(by)[elementIndex];
            MakeElementVisible(driver, element);
        }
        #endregion
        #endregion

        #region Internal methods (Selenium)
        /// <summary>
        /// Clicks on an element via Actions API
        /// </summary>
        /// <param name="Element">Element to click</param>
        private static void MouseActionClick(this IWebDriver driver, IWebElement Element)
        {
            driver.ExecuteJavaScript("arguments[0].scrollIntoView();", Element);
            OpenQA.Selenium.Interactions.Actions actions = new(driver);
            actions.MoveToElement(Element).ClickAndHold(Element).Release().Build().Perform();
        }
        #endregion


        ///// <summary>
        ///// Selects a dropdown value by value
        ///// </summary>
        ///// <param name="WebDriver">SeleniumWebDriver</param>
        ///// <param name="by">By</param>
        ///// <param name="Data">Data / value</param>
        ///// <param name="elementIndex">Index of the element if more than one can be found with the By selector</param>
        //public static void SelectDropDownByValue(IWebDriver WebDriver, By by, string Data, int elementIndex = 0, [CallerMemberName] string callerName = "")
        //{
        //    try
        //    {
        //        new SelectElement(WebDriver.FindElements(by)[elementIndex]).SelectByValue(Data);
        //    }
        //    catch (Exception ex)
        //    {
        //        System.Windows.Forms.MessageBox.Show("Error: " + ex.Message);
        //    }
        //}

        ///// <summary>
        ///// Selects a dropdown value by index
        ///// </summary>
        ///// <param name="WebDriver">SeleniumWebDriver</param>
        ///// <param name="by">By</param>
        ///// <param name="Index">Index</param>
        ///// <param name="elementIndex">Index of the element if more than one can be found with the By selector</param>
        //public static void SelectDropDownByIndex(this IWebDriver webDriver, By by, int index, int elementIndex = 0, [CallerMemberName] string callerName = "")
        //{
        //    try
        //    {
        //        IJavaScriptExecutor executor = (IJavaScriptExecutor)webDriver;
        //        executor.ExecuteScript("arguments[0].selectedIndex = " + index.ToString() + ";", webDriver.FindElements(by)[elementIndex]);
        //    }
        //    catch (Exception ex)
        //    {
        //        StackFrame frame = new StackFrame(1);
        //        string calledMethod = frame.GetMethod().DeclaringType.Name + "." + callerName;
        //        ReportError(calledMethod, ex, "By: " + by.ToString() + Environment.NewLine + "Index: " + index.ToString() + "elementIndex: " + elementIndex.ToString());
        //    }
        //}

        ///// <summary>
        ///// Selects a dropdown value by index
        ///// </summary>
        ///// <param name="webDriver">SeleniumWebDriver</param>
        ///// <param name="by">By</param>
        ///// <param name="value">Value</param>
        ///// <param name="elementIndex">Index of the element if more than one can be found with the By selector</param>
        //public static void SelectDropDownByValue(this IWebDriver webDriver, By by, string value, int elementIndex = 0, [CallerMemberName] string callerName = "")
        //{
        //    try
        //    {
        //        IJavaScriptExecutor executor = (IJavaScriptExecutor)webDriver;
        //        executor.ExecuteScript("for(var i=0;i<arguments[0].options.length;i++)arguments[0].options[i].value==" + value + "&&(arguments[0].options[i].selected=!0);", webDriver.FindElements(by)[elementIndex]);
        //    }
        //    catch (Exception ex)
        //    {
        //        StackFrame frame = new StackFrame(1);
        //        string calledMethod = frame.GetMethod().DeclaringType.Name + "." + callerName;
        //        ReportError(calledMethod, ex, "By: " + by.ToString() + Environment.NewLine + "Value: " + value + "elementIndex: " + elementIndex.ToString());
        //    }
        //}
    }
}
