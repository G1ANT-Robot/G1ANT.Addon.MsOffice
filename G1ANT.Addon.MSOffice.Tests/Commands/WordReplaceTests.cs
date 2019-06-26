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
using System.Threading;

namespace G1ANT.Addon.MSOffice.Tests
{
    [TestFixture]
    [Apartment(ApartmentState.STA)]
    public class WordReplaceTests
    {
        Scripter scripter;
        static String replaceFrom = "tro";
        static String replaceTo = "lo";
        static String restOfText = "lololololo";
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
           scripter.InitVariables.Add("text", new TextStructure(replaceFrom + restOfText));
        }

        [SetUp]
        public void TestInit()
        {
            Language.Addon addon = Language.Addon.Load(@"G1ANT.Addon.MSOffice.dll");
            
        }

        [Test]
        [Timeout(MSOfficeTests.TestsTimeout)]
        public void WordReplaceTest()
        {
            scripter.Text = ($@"word.open
                               word.inserttext {SpecialChars.Variable}text
                               word.replace from {replaceFrom} to {replaceTo}
                               word.gettext");
            scripter.Run();
            Assert.AreEqual(replaceTo + restOfText, scripter.Variables.GetVariableValue<string>("result").Trim());
        }


        [TearDown]
        public void TestCleanUp()
        {
            scripter.RunLine("word.close");
            Process[] proc = Process.GetProcessesByName("word");
            if (proc.Length != 0)
            {
                KillProcesses();
            }
        }
    }
}
