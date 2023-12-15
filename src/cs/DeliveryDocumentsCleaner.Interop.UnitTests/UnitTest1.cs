using System.Drawing;

namespace DeliveryDocumentsCleaner.Interop.UnitTests;

public class UnitTest1
{
    [Fact]
    public void Test()
    {
        static long ToOleColorCode(byte red, byte green, byte blue)
        {
            var color = Color.FromArgb(0, red, green, blue);
            var value = ColorTranslator.ToOle(color);
            return value;
        }

        static string ToColorHex(long fullColorCode)
        {
            var rValue = fullColorCode % 256;
            var gValue = (fullColorCode / 256) % 256;
            var bValue = fullColorCode / 65536;
            return $"#{rValue:x2}{gValue:x2}{bValue:x2}";
        }

        var r = ToOleColorCode(0xff, 0, 0);
        var rHex = ToColorHex(r);
        var g = ToOleColorCode(0, 0xff, 0);
        var gHex = ToColorHex(g);
        var b = ToOleColorCode(0, 0, 0xff);
        var bHex = ToColorHex(b);

        Assert.Equal(0x0000ff, r);
        Assert.Equal(0x00ff00, g);
        Assert.Equal(0xff0000, b);
        Assert.Equal("#ff0000", rHex);
        Assert.Equal("#00ff00", gHex);
        Assert.Equal("#0000ff", bHex);
    }

}
