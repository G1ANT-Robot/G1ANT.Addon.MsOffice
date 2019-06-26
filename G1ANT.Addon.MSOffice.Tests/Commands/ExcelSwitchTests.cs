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
    public class ExcelSwitchTests
	{
		static Engine.Scripter scripter;
		static string xlsPath;
		static int someVal = 10;

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
        public static void ClassInit()
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
           scripter.InitVariables.Add("val", new IntegerStructure(someVal));
           
        }

        [Test]
        [Timeout(MSOfficeTests.TestsTimeout)]
		public void ExcelSwitchTest()
		{
            scripter.Text = ($@"excel.open {SpecialChars.Text}{SpecialChars.Text} result {SpecialChars.Variable}id
                                excel.open {SpecialChars.Variable}xlsPath sheet Add result {SpecialChars.Variable}id2
                                excel.switch {SpecialChars.Variable}id
			                    excel.setvalue {SpecialChars.Variable}val row 1 colindex 1
			                    excel.switch {SpecialChars.Variable}id2
			                    excel.getvalue row 1 colindex 1 result {SpecialChars.Variable}result1
			                    
			                    excel.activatesheet Macro
			                    excel.getvalue row 6 colindex 2 result {SpecialChars.Variable}val2
			                    excel.switch {SpecialChars.Variable}id
			                    excel.getvalue row 1 colindex 1 result {SpecialChars.Variable}result2
			                    
			                    excel.switch {SpecialChars.Variable}id2
			                    excel.getvalue row 6 colindex 2 result {SpecialChars.Variable}result3
                                excel.switch {SpecialChars.Variable}id
                                excel.close
                                excel.switch {SpecialChars.Variable}id2
                                excel.close");
            scripter.Run();
            Assert.AreNotEqual(someVal, int.Parse(scripter.Variables.GetVariable("result1").GetValue().Object.ToString()));
            Assert.AreEqual(someVal, int.Parse(scripter.Variables.GetVariable("result2").GetValue().Object.ToString()));
            Assert.AreEqual(int.Parse(scripter.Variables.GetVariableValue<string>("val2")), int.Parse(scripter.Variables.GetVariable("result3").GetValue().Object.ToString()));

            
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
