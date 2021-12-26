using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;

namespace UnitTestProject1
{
    [TestClass]
    public class UnitTest1
    {
        [TestMethod]
        public void TestMethod1()
        {
            int x = 378678;
            int y = 51;
            int expected = 7426;
            трпо.Sred sred = new трпо.Sred();
            int actual = sred.Srd(x, y);
            Assert.AreEqual(expected, actual);
        }
    }
}
