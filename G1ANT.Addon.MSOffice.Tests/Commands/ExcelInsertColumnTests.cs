/**
*    Copyright(C) G1ANT Ltd, All rights reserved
*    Solution G1ANT.Addon, Project G1ANT.Addon.MSOffice
*    www.g1ant.com
*
*    Licensed under the G1ANT license.
*    See License.txt file in the project root for full license information.
*
*/

using NUnit.Framework;
using System;
using System.Diagnostics;
using System.Threading;
using G1ANT.Engine;
namespace G1ANT.Addon.MSOffice.Tests
{
    [TestFixture]
    [Apartment(ApartmentState.STA)]
    public class ExcelInsertcolindexumnTests
    {

        Scripter scripter;

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
            scripter = new Scripter();
scripter.InitVariables.Clear();
        }

        [SetUp]
        public void TestInit()
        {
            Language.Addon addon = Language.Addon.Load(@"G1ANT.Addon.MSOffice.dll");
            scripter.RunLine($"excel.open");
        }

        [Test]
        [Timeout(MSOfficeTests.TestsTimeout)]
        public void ExcelInsertcolindexumnTest()
        {
            scripter.RunLine("excel.setvalue aaa row 1 colindex 1");
            scripter.RunLine("excel.setvalue bbb row 1 colindex 2");
            scripter.RunLine("excel.insertcolumn colindex 1 where after");
            scripter.RunLine("excel.getvalue row 1 colindex 3");
            Assert.AreEqual("bbb", scripter.Variables.GetVariableValue<string>("result"));

            scripter.RunLine("excel.removecolumn colindex 2");
            scripter.RunLine("excel.insertcolumn colname b where before");
            Assert.AreEqual("bbb", scripter.Variables.GetVariableValue<string>("result"));
        }

        [Test]
        [Timeout(MSOfficeTests.TestsTimeout)]
        public void ExcelInsertRowFailTest()
        {

            scripter.Text = "excel.insertcolumn colindex 1 where haha";
            Exception exception = Assert.Throws<ApplicationException>(delegate
            {
                scripter.Run();
            });
            Assert.IsInstanceOf<ArgumentException>(exception.GetBaseException());
        }

        [Test]
        [Timeout(MSOfficeTests.TestsTimeout)]
        public void ExcelInsertRowFail2Test()
        {

            scripter.Text = "excel.insertcolumn colindex 0 where after";
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
