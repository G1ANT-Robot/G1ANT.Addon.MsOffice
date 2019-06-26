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
    public class ExcelGetValueTests
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
        public void ExcelGetValueTest()
        {
            scripter.Text =($@"excel.open {SpecialChars.Variable}xlsPath sheet Add
                               excel.getvalue row 1 colindex 1 result {SpecialChars.Variable}result1
                               excel.getvalue row 1 colindex 3 result {SpecialChars.Variable}result3
                               excel.getvalue row 1 colindex 2 result {SpecialChars.Variable}result2");
            scripter.Run();
            Assert.AreEqual(3, int.Parse(scripter.Variables.GetVariable("result1").GetValue().Object.ToString()));
            Assert.AreEqual(4, int.Parse(scripter.Variables.GetVariable("result2").GetValue().Object.ToString()));
            Assert.AreEqual(7, int.Parse(scripter.Variables.GetVariable("result3").GetValue().Object.ToString()));
        }

        [Test]
        [Timeout(MSOfficeTests.TestsTimeout)]
        public void ExcelGetValueFailTest()
        {
            scripter.Text =($@"excel.open {SpecialChars.Variable}xlsPath sheet Add
                               excel.getvalue row 0 colindex 1");
            Exception exception = Assert.Throws<ApplicationException>(delegate
            {
                scripter.Run();
            });
            Assert.IsInstanceOf<ArgumentException>(exception.GetBaseException());
        }

        [Test]
        [Timeout(MSOfficeTests.TestsTimeout)]
        public void ExcelGetValueFail2Test()
        {
            scripter.Text = ($@"excel.open {SpecialChars.Variable}xlsPath sheet Add
                                excel.getvalue row 1 colindex 0");
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
