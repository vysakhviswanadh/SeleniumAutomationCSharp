using System;
using System.Threading;
using OpenQA.Selenium;
using NUnit.Framework;
using AutoPOC.Utils;
//using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace AutoPOC
{   
    public class SampleTestSuite
    {               
        [Test]
        public void Test_Ultimate_QA()
        {
            var bf_ultimateQA = new BaseFunctions();
            string testCaseName = TestContext.CurrentContext.Test.Name;  
            
            IWebElement _link_automationExercises;
            IWebElement _header_Automation;

            bf_ultimateQA.getBrowser();            
            bf_ultimateQA.launchWebsite();

            _link_automationExercises = bf_ultimateQA.getElementsFromRepo("link_automationExercises");            
            bf_ultimateQA.seleniumExecuter(_link_automationExercises, "click", "link_automationExercises", "");        
            Thread.Sleep(TimeSpan.FromSeconds(1));

            _header_Automation = bf_ultimateQA.getElementsFromRepo("header_Automation");
            string headerAutomation = bf_ultimateQA.getTestData(testCaseName, "Header Automation Exercise");

            if(_header_Automation.Text == headerAutomation)
            {
                Console.WriteLine("Automation Exercises page loaded");
                bf_ultimateQA.Results("Automation Exercises page load", "Page should be loaded", "Page loaded", "Pass");                
            }
            else
            {
                Console.WriteLine("Automation Exercises page NOT loaded");
                bf_ultimateQA.Results("Automation Exercises page load", "Page should be loaded", "Page loaded", "Fail");                
                Assert.Fail("Automation Exercises page NOT loaded");
            }
            
            bf_ultimateQA.closeBrowser();
        }
    }
}
