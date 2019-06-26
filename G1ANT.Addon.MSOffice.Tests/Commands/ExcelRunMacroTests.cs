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


using NUnit.Framework;
using System;
using System.Reflection;
using System.Threading;

using System.Diagnostics;
using G1ANT.Engine;
using G1ANT.Language;
using G1ANT.Addon.MSOffice.Tests.Properties;

namespace G1ANT.Addon.MSOffice.Tests
{
    [TestFixture]
    [Apartment(ApartmentState.STA)]
    public class ExcelRunMacroTests
    {

        Scripter scripter;
        static string xlsPath;

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

        static string sheetName = "Macro";
        static string macroName = "Calculate";
        static int calculationRow = 6;
        static int calculationValueToBeCountedcolindexumn = 2;
        static int calculationValueExpectedcolindexumn = 1;

        [OneTimeSetUp]
        public void ClassInit()
        {
            Environment.CurrentDirectory = TestContext.CurrentContext.TestDirectory;
            scripter = new Scripter();
scripter.InitVariables.Clear();
        }

        [SetUp]
        public void TestInit()
        {
            Language.Addon addon = Language.Addon.Load(@"G1ANT.Addon.MSOffice.dll");
            xlsPath = Assembly.GetExecutingAssembly().UnpackResourceToFile("Resources." + nameof(Resources.TestWorkbook), "xlsm");
           scripter.InitVariables.Add("xlsPath", new TextStructure(xlsPath));
           scripter.InitVariables.Add("sheet", new TextStructure(sheetName));
           scripter.InitVariables.Add("macroName", new TextStructure(macroName));
            
        }

        [Test]
        [Timeout(MSOfficeTests.TestsTimeout)]
        public void ExcelRunMacroCalculationTest()
        {
            scripter.Text =($@"excel.open {SpecialChars.Variable}xlsPath sheet {SpecialChars.Variable}sheet
                               excel.setvalue {SpecialChars.Text}{SpecialChars.Text} row {calculationRow} colindex {calculationValueToBeCountedcolindexumn}
                               excel.runmacro {SpecialChars.Variable}macroName
                               excel.getvalue row {calculationRow} colindex {calculationValueExpectedcolindexumn} result {SpecialChars.Variable}result1
                               excel.setvalue {SpecialChars.Text}4{SpecialChars.Text} row {calculationRow} colindex {calculationValueToBeCountedcolindexumn}
                               excel.runmacro {SpecialChars.Variable}macroName
                               excel.getvalue row {calculationRow} colindex {calculationValueExpectedcolindexumn} result {SpecialChars.Variable}result2");
            scripter.Run();
            int expectedValue = 0;
            Assert.AreEqual(expectedValue, int.Parse(scripter.Variables.GetVariableValue<string>("result1")));
            expectedValue = 40;
            Assert.AreEqual(expectedValue, int.Parse(scripter.Variables.GetVariableValue<string>("result2")));
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
