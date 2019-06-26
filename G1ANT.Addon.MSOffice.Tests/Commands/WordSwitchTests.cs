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
    public class WordSwitchTests
    {
        Scripter scripter;
        static String someText = "lololololo";
        static String someText2 = "trolololololo";

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
           scripter.InitVariables.Add("text", new TextStructure(someText));
           scripter.InitVariables.Add("text2", new TextStructure(someText2));
        }

        [SetUp]
        public void TestInit()
        {
            Language.Addon addon = Language.Addon.Load(@"G1ANT.Addon.MSOffice.dll");
           
        }

        [Test]
        [Timeout(MSOfficeTests.TestsTimeout)]
        public void WordSwitchTest()
        {
            scripter.Text = ($@"word.open result {SpecialChars.Variable}id
                                word.open result {SpecialChars.Variable}id2
                                word.switch {SpecialChars.Variable}id
                                word.inserttext {SpecialChars.Variable}text
                                word.switch {SpecialChars.Variable}id2
                                word.inserttext {SpecialChars.Variable}text2
                                word.switch {SpecialChars.Variable}id
                                word.gettext result {SpecialChars.Variable}result1
                                word.switch {SpecialChars.Variable}id2
                                word.gettext result {SpecialChars.Variable}result2
                                word.switch {SpecialChars.Variable}id
                                word.close
                                word.switch {SpecialChars.Variable}id2
                                word.close");
            scripter.Run();
            Assert.AreEqual(((string)scripter.Variables.GetVariable("result2").GetValue().Object).Trim(), someText2);
            Assert.AreEqual(scripter.Variables.GetVariableValue<string>("result1").Trim(), someText);
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
