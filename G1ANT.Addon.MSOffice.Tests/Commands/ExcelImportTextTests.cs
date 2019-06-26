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
    public class ExcelImportTextTests
	{
		static string csvPath;
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
            csvPath = Assembly.GetExecutingAssembly().UnpackResourceToFile("Resources." + nameof(Resources.TestData), "csv");
           scripter.InitVariables.Add("csvPath", new TextStructure(csvPath));
            
        }

        [Test]
        [Timeout(MSOfficeTests.TestsTimeout), Ignore("Passing point as argument to commands doesnt work")]
        public void ExcelImportTextTest()
		{
            string twentyOne = 21.ToString();
            string thirtyTwo = 32.ToString();
           scripter.RunLine($@"excel.open {SpecialChars.Variable}csvPath
                              excel.importtext path {SpecialChars.Variable}csvPath delimiter ,
		                      excel.getvalue row 2 colindex 1 result {SpecialChars.Variable}result1
			                  excel.getvalue row 3 colindex 2 result {SpecialChars.Variable}result2
                              excel.importtext path {SpecialChars.Variable}csvPath destination D1 delimiter ,
                              excel.getvalue row 2 colindex 4 result {SpecialChars.Variable}result3
                              excel.getvalue row 3 colindex 5 result {SpecialChars.Variable}result4
                              excel.importtext path {SpecialChars.Variable}csvPath destination (point)4,1 delimiter ,
                              excel.getvalue row 5 colindex 1 result {SpecialChars.Variable}result5
                              excel.getvalue row 6 colindex 2 result {SpecialChars.Variable}result6
                              excel.close");
            Assert.AreEqual(twentyOne, scripter.Variables.GetVariableValue<string>("result1"));
            Assert.AreEqual(thirtyTwo, scripter.Variables.GetVariableValue<string>("result2"));
            Assert.AreEqual(twentyOne, scripter.Variables.GetVariableValue<string>("result3"));
            Assert.AreEqual(thirtyTwo, scripter.Variables.GetVariableValue<string>("result4"));
            Assert.AreEqual(twentyOne, scripter.Variables.GetVariableValue<string>("result5"));
            Assert.AreEqual(thirtyTwo, scripter.Variables.GetVariableValue<string>("result6"));
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
