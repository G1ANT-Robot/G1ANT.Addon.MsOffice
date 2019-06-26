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
    public class ExcelAddSheetTests
    {
        Scripter scripter;
        static string sheetName = "TestSheet";
        static string otherSheet = "otherSheet";

        private void KillProcesses()
        {
            foreach (Process p in Process.GetProcessesByName("excel"))
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
            scripter.InitVariables.Add("TestSheet", new TextStructure(sheetName));
            scripter.InitVariables.Add("otherSheet", new TextStructure(otherSheet));

        }

        [Test]
        [Timeout(MSOfficeTests.TestsTimeout)]
        public void ExcelAddSheetTest()
        {
            scripter.Text = ($@"excel.open
                                excel.addsheet {SpecialChars.Variable}TestSheet
                                excel.activatesheet {SpecialChars.Variable}TestSheet
                                excel.setvalue 1 colindex 1 row 1
                                excel.getvalue row 1 colindex 1 result {SpecialChars.Variable}valTest
                                excel.addsheet {SpecialChars.Variable}otherSheet
                                excel.activatesheet { SpecialChars.Variable}otherSheet
                                excel.setvalue 5 colindex 1 row 1
                                excel.getvalue row 1 colindex 1 result { SpecialChars.Variable}valOther
                                excel.activatesheet {SpecialChars.Variable}TestSheet
                                excel.getvalue row 1 colindex 1 result {SpecialChars.Variable}val");
            scripter.Run();

            Assert.AreEqual(1, int.Parse(scripter.Variables.GetVariableValue<string>("valTest")));
            Assert.AreEqual(5, int.Parse(scripter.Variables.GetVariableValue<string>("valOther")));
            Assert.AreEqual(1, int.Parse(scripter.Variables.GetVariableValue<string>("val")));
        }

        [Test]
        [Timeout(MSOfficeTests.TestsTimeout)]
        public void ExcelAddSheetFailTest()
        {
            scripter.Text = ($@"excel.open
                              excel.addsheet {SpecialChars.Variable}TestSheet
                              excel.addsheet {SpecialChars.Variable}TestSheet");
            Exception exception = Assert.Throws<ApplicationException>(delegate
            {
                scripter.Run();
            });
            Assert.IsInstanceOf<ArgumentException>(exception.GetBaseException());
        }

        [TearDown]
        public void TestCleanUp()
        {
            scripter.RunLine("excel.close");
            Process[] proc = Process.GetProcessesByName("excel");
            if (proc.Length != 0)
            {
                KillProcesses();
            }
        }
    }
}
