/**
*    Copyright(C) G1ANT Ltd, All rights reserved
*    Solution G1ANT.Addon, Project G1ANT.Addon.MSOffice
*    www.g1ant.com
*
*    Licensed under the G1ANT license.
*    See License.txt file in the project root for full license information.
*
*/


using G1ANT.Engine;
using G1ANT.Language;
using NUnit.Framework;
using System;
using System.Diagnostics;
using System.IO;
using System.Threading;

namespace G1ANT.Addon.MSOffice.Tests
{
    [TestFixture]
    [Apartment(ApartmentState.STA)]
    public class WordSaveTests
    {
        Scripter scripter;
        static string wordToBeTested = "TestG1ant";
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
        public static void ClassInit()
        {
            Environment.CurrentDirectory = TestContext.CurrentContext.TestDirectory;           
        }
        [SetUp]
        public void SetUp()
        {
            Language.Addon addon = Language.Addon.Load(@"G1ANT.Addon.MSOffice.dll");
            scripter = new Scripter();
scripter.InitVariables.Clear();
        }

        [Test]
        [Timeout(MSOfficeTests.TestsTimeout)]
        public void WordSaveTest()
        {

            string savePath = Environment.CurrentDirectory + @"\test.docx";
            scripter.Text =
                $@"word.open
                window {SpecialChars.Text}{SpecialChars.Search}word{SpecialChars.Search}{SpecialChars.Text}
                keyboard {SpecialChars.Text}{wordToBeTested}{SpecialChars.Text}
                word.save {SpecialChars.Variable}savePath
                word.close
                delay 4
                word.open {SpecialChars.Variable}savePath
                window {SpecialChars.Text}{SpecialChars.Search}word{SpecialChars.Search}{SpecialChars.Text}
                word.gettext
                word.close";
            
           scripter.InitVariables.Add("savePath", new TextStructure(savePath));
            scripter.Run();
            Assert.AreEqual(wordToBeTested, scripter.Variables.GetVariableValue<string>("result").Trim());

            FileInfo f = new FileInfo(savePath);
            Assert.IsTrue(f.Exists);
            f.Delete();
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
