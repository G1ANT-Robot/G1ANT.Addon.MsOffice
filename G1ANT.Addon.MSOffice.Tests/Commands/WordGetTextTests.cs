/**
*    Copyright(C) G1ANT Ltd, All rights reserved
*    Solution G1ANT.Addon, Project G1ANT.Addon.MSOffice
*    www.g1ant.com
*
*    Licensed under the G1ANT license.
*    See License.txt file in the project root for full license information.
*
*/



using G1ANT.Addon.MSOffice.Tests.Properties;
using G1ANT.Engine;
using G1ANT.Language;
using NUnit.Framework;
using System;
using System.Diagnostics;
using System.Reflection;
using System.Threading;

namespace G1ANT.Addon.MSOffice.Tests
{
    [TestFixture]
    [Apartment(ApartmentState.STA)]
    public class WordGetTextTests
    {
        Scripter scripter;
        static string wordPath;
        static string valueTested = "Test, test, test....test";
        static string expected = "Test, test, test....testTest, test, test....test";
        static string expectedEmptyString = string.Empty;

        private void KillProcesses()
        {
            foreach (Process p in Process.GetProcessesByName("word"))
            {
                try
                {
                    p.Kill();
                }
                catch { }
            }
        }

        [OneTimeSetUp]
        public void ClassInit()
        {
            Environment.CurrentDirectory = TestContext.CurrentContext.TestDirectory;
            scripter = new Scripter();
scripter.InitVariables.Clear();
            wordPath = Assembly.GetExecutingAssembly().UnpackResourceToFile("Resources." + nameof(Resources.TestDocument), "docx");
           scripter.InitVariables.Add("wordPath", new TextStructure(wordPath));
        }

        [SetUp]
        public void TestInit()
        {
            Language.Addon addon = Language.Addon.Load(@"G1ANT.Addon.MSOffice.dll");

        }

        [Test]
        [Timeout(MSOfficeTests.TestsTimeout)]
        public void WordGetTextTest()
        {
            scripter.Text = ($@"word.open {SpecialChars.Variable}wordPath
                                word.inserttext {SpecialChars.Text}{valueTested}{SpecialChars.Text}
                                word.inserttext {SpecialChars.Text}{valueTested}{SpecialChars.Text}
                                word.gettext result {SpecialChars.Variable}result1
                                word.inserttext {SpecialChars.Text}{valueTested}{SpecialChars.Text}
                                word.inserttext {SpecialChars.Text}{expectedEmptyString}{SpecialChars.Text} replacealltext true 
                                word.gettext result {SpecialChars.Variable}result2
                                word.close");
            scripter.Run();
            string trimmedValue = scripter.Variables.GetVariableValue<string>("result1").Trim();
            Assert.AreEqual(expected, trimmedValue);
            trimmedValue = scripter.Variables.GetVariableValue<string>("result2").Trim();
            Assert.AreEqual(expectedEmptyString, trimmedValue);
        }

        [TearDown]
        public void TestCleanUp()
        {
            
            Process[] proc = Process.GetProcessesByName("word");
            if (proc.Length != 0)
            {
                KillProcesses();
            }
        }
    }
}
