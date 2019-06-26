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
    public class ExcelSetValTests
    {
        static int intVal = 5;
        static float fVal = 3.3f;
        static string stringVal = "somng";
        static string formula = "=B1*C1";
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
           scripter.InitVariables.Add("intVal", new Language.IntegerStructure(intVal));
           scripter.InitVariables.Add("fVal", new FloatStructure(fVal));
           scripter.InitVariables.Add("strVal", new TextStructure(stringVal));
           scripter.InitVariables.Add("formula", new TextStructure(formula));
        }
        

        [SetUp]
        public void TestInit()
        {
            Language.Addon addon = Language.Addon.Load(@"G1ANT.Addon.MSOffice.dll");
            
        }

        [Test]
        [Timeout(MSOfficeTests.TestsTimeout)]
        public void ExcelSetValTest()
        {
            scripter.Text = ($@"excel.open
                               excel.setvalue {SpecialChars.Variable}strVal row 1 colindex 1
                               excel.setvalue {SpecialChars.Variable}intVal row 1 colindex 2
                               excel.setvalue {SpecialChars.Variable}fVal row 1 colindex 3
                               excel.setvalue {SpecialChars.Variable}formula row 1 colindex 4
                               excel.getvalue row 1 colindex 1 result {SpecialChars.Variable}result1
                               excel.getvalue row 1 colindex 2 result {SpecialChars.Variable}result2
                               excel.getvalue row 1 colindex 3 result {SpecialChars.Variable}result3
                               excel.getvalue row 1 colindex 4 result {SpecialChars.Variable}product
                               excel.close");

            scripter.Run();

            Assert.AreEqual(stringVal, scripter.Variables.GetVariableValue<string>("result1"));
            Assert.AreEqual(intVal, Int32.Parse(scripter.Variables.GetVariable("result2").GetValue().Object as String));
            Assert.AreEqual(fVal, float.Parse((scripter.Variables.GetVariable("result3").GetValue().Object as String).Replace(",", ".")));
            Assert.AreEqual(intVal * fVal, float.Parse(scripter.Variables.GetVariableValue<string>("product").Replace(",", ".")), 0.00001);
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
