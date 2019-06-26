/**
*    Copyright(C) G1ANT Ltd, All rights reserved
*    Solution G1ANT.Addon, Project G1ANT.Addon.MSOffice
*    www.g1ant.com
*
*    Licensed under the G1ANT license.
*    See License.txt file in the project root for full license information.
*
*/

using G1ANT.Addon.MSOffice;
using G1ANT.Engine;
using System;
using System.IO;
using NUnit.Framework;
using System.Reflection;


using System.Threading;
using System.Diagnostics;
using G1ANT.Language;
using G1ANT.Addon.MSOffice.Tests.Properties;

namespace G1ANT.Addon.MSOffice.Tests
{
    [TestFixture]
    [Apartment(ApartmentState.STA)]
    public class ExcelGetFormulaTests
    {
        Scripter scripter;
        static string xlsPath;
        static string formula = "=A1+B1";

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
            xlsPath = Assembly.GetExecutingAssembly().UnpackResourceToFile("Resources." + nameof(Resources.TestWorkbook), "xlsm");
           scripter.InitVariables.Add("xlsPath", new TextStructure(xlsPath));
           
        }

        [Test]
        [Timeout(MSOfficeTests.TestsTimeout)]
        public void ExcelgetFormulaTest()
        {
            scripter.Text = ($@"excel.open {SpecialChars.Variable}xlsPath sheet Add
                                excel.getformula row 1 colindex 3");
            scripter.Run();
            Assert.AreEqual(formula, scripter.Variables.GetVariableValue<string>("result"));
        }

        [Test]
        [Timeout(MSOfficeTests.TestsTimeout)]
        public void ExcelgetFormula2Test()
        {
            scripter.Text =($@"excel.open {SpecialChars.Variable}xlsPath sheet Add
                               excel.getformula row 10 colindex 10");
            scripter.Run();
            Assert.AreEqual(string.Empty, scripter.Variables.GetVariableValue<string>("result"));
        }

        [Test]
        [Timeout(MSOfficeTests.TestsTimeout)]
        public void ExcelgetFormula3Test()
        {
            scripter.Text = ($@"excel.open {SpecialChars.Variable}xlsPath sheet Add
                                excel.getformula row -1 colindex 10");
            Exception exception = Assert.Throws<ApplicationException>(delegate
                {
                    scripter.Run();
                });
            Assert.IsInstanceOf<ArgumentException>(exception.GetBaseException());
        }

        [Test]
        [Timeout(MSOfficeTests.TestsTimeout)]
        public void ExcelgetFormula5Test()
        {
            scripter.Text = ($@"excel.open {SpecialChars.Variable}xlsPath sheet Add
                                excel.getformula row -1 colname żd2");
            Exception exception = Assert.Throws<ApplicationException>(delegate
            {
                scripter.Run();
            });
            Assert.IsInstanceOf<ArgumentException>(exception.GetBaseException());
        }

        [TearDown]
        public void TestCleanUp()
        {
            Process[] proc = Process.GetProcessesByName("excel");
            if (proc.Length != 0)
            {
                KillProcesses();
            }
        }
    }
}
