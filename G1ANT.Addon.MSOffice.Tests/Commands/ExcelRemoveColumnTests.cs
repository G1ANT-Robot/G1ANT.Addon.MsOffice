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

using System;
using NUnit.Framework;
using System.Reflection;

using System.Threading;
using System.Diagnostics;
using G1ANT.Engine;
using G1ANT.Language;

namespace G1ANT.Addon.MSOffice.Tests
{
    [TestFixture]
    [Apartment(ApartmentState.STA)]
    public class ExcelRemovecolindexumnTests
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
            
        }

        [SetUp]
        public void TestInit()
        {
            scripter = new Scripter();
            scripter.InitVariables.Clear();
            Language.Addon addon = Language.Addon.Load(@"G1ANT.Addon.MSOffice.dll");
           
        }

        [Test]
        [Timeout(MSOfficeTests.TestsTimeout)]
        public void ExcelRemovecolindexumnTest()
        {
            scripter.Text = ($@"excel.open
                                excel.setvalue aaa row 1 colindex 1
                                excel.setvalue bbb row 1 colindex 2
                                excel.insertcolumn colindex 1 where after
                                excel.removecolumn colindex 2
                                excel.getvalue row 1 colindex 2 result {SpecialChars.Variable}result1
                                excel.insertcolumn colindex 2 where before
                                excel.removecolumn colname b
                                excel.getvalue row 1 colindex 2
                                excel.close");
            scripter.Run();
            Assert.AreEqual("bbb", scripter.Variables.GetVariableValue<string>("result1"));
            Assert.AreEqual("bbb", scripter.Variables.GetVariableValue<string>("result"));
        }

        [Test]
        public void ExcelRemovecolindexumnFailTest()
        {
            scripter.Text =($@"excel.open
                               excel.removecolumn colindex hadhaad2radfa");
           
                Exception exception = Assert.Throws<AggregateException>(delegate
                {
                    scripter.Run();
                });
                Assert.IsInstanceOf<AggregateException>(exception);
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
