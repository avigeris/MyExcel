using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using MyExcel;

namespace UnitTestProjectExcel
{
    [TestClass]
    public class Based26SysTest
    {
        [TestMethod]
        public void TestTo26Sys()
        {
            int number = 0;
            string result = "A";

            Based26Sys name = new Based26Sys();
            string actual = name.To26Sys(number);

            Assert.AreEqual(result, actual);
        }

        [TestMethod]
        public void TestFrom26Sys()
        {
            string header = "AA";
            int result = 26;

            Based26Sys name = new Based26Sys();
            int actual = name.From26Sys(header);

            Assert.AreEqual(result, actual);
        }

        [TestMethod]
        public void TestNegativeValue()
        {
            int number = -1;
            string result = "-B";

            Based26Sys name = new Based26Sys();
            string actual = name.To26Sys(number);

            Assert.AreEqual(result, actual);

        }
    }

    [TestClass]
    public class ParserTest
    {
        [TestMethod]
        public void AddTest()
        {
            string a = "-3,25+-5,75";
            double result = -9;

            Parser p = new Parser();
            p.SetCurrentName("123");
            double actual = p.Evaluate(a);

            Assert.AreEqual(result, actual);
        }
        
    }
}
