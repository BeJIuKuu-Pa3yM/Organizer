using Xunit;

namespace TestProject1
{
    public class UnitTest1
    {
        public void Test1()
        {
            int x = 378678;
            int y = 51;
            int expected = 7425;
            ����.Sred sred = new ����.Sred();
            int actual = sred.Srd(x, y);
            Assert.Equal(expected, actual);

        }
    }
}
