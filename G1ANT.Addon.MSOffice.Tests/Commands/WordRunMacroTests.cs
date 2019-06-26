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
    public class WordRunMacroTests
    {
        Scripter scripter;
        static string wordPath;
        static string macroName = "SortText";
        static string testedValue = $"Pawel\rPatryk\rMarcin\rZuza\rChris\rMichal\rDiana\rPrzemek\rJano\r";
        static string expectedValue = $"\rChris\rDiana\rJano\rMarcin\rMichal\rPatryk\rPawel\rPrzemek\rZuza\r";
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
            
        }

        [SetUp]
        public void TestInit()
        {
            scripter = new Scripter();
            scripter.InitVariables.Clear();
            Language.Addon addon = Language.Addon.Load(@"G1ANT.Addon.MSOffice.dll");
            wordPath = Assembly.GetExecutingAssembly().UnpackResourceToFile("Resources." + nameof(Resources.TestDocumentMacro), "docm");
           scripter.InitVariables.Add("wordPath", new TextStructure(wordPath));
           scripter.InitVariables.Add("macroName", new TextStructure(macroName));

        }

        [Test]
        [Timeout(MSOfficeTests.TestsTimeout)]
        public void WordRunMacroTest()
        {
            scripter.Text =($@"word.open {SpecialChars.Variable}wordPath
                               word.inserttext {SpecialChars.Text}{testedValue}{SpecialChars.Text}
                               word.runmacro {SpecialChars.Variable}macroName
                               word.gettext
                               word.close");
            scripter.Run();
            string trimmedValue = scripter.Variables.GetVariableValue<string>("result");
            Assert.AreEqual(expectedValue, trimmedValue);
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
