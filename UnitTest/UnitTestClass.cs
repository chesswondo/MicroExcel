using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using MicroExcel;

namespace UnitTest
{
    [TestClass]
    public class UnitTestClass
    {
        static void ZeroUpZero()
        {
            MEParser p = new MEParser();
            p.Evaluate("0^0", null);
        }
        static void Wrong1()
        {
            MEParser p = new MEParser();
            p.Evaluate("5/0", null);
        }
        static void Wrong2()
        {
            MEParser p = new MEParser();
            p.Evaluate("abc+3", null);
        }

        [TestMethod]
        public void TestParserArithmetic()
        {
            MEParser p = new MEParser();
            Assert.AreEqual(-3, p.Evaluate("2+3-8", null));
            Assert.AreEqual(10, p.Evaluate("2*3+8/2", null));
        }

        [TestMethod]
        public void TestParserWrong()
        {
            MEParser p = new MEParser();
            Assert.ThrowsException<ArgumentException>(ZeroUpZero);
            Assert.ThrowsException<ArgumentException>(Wrong1);
            Assert.ThrowsException<ArgumentException>(Wrong2);
        }
    }
}
