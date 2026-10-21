namespace OfficeEngine.Tests.Core;

public sealed class RgbColorTests
{
    [Fact]
    public void FromRgb_StoresTheComponents()
    {
        RgbColor color = RgbColor.FromRgb(0x12, 0x34, 0x56);

        Assert.Equal(0x12, color.Red);
        Assert.Equal(0x34, color.Green);
        Assert.Equal(0x56, color.Blue);
    }

    [Fact]
    public void ToHex_FormatsAsUpperCaseSixDigits()
    {
        Assert.Equal("112233", RgbColor.FromRgb(0x11, 0x22, 0x33).ToHex());
        Assert.Equal("000000", RgbColor.FromRgb(0, 0, 0).ToHex());
        Assert.Equal("FFFFFF", RgbColor.FromRgb(0xFF, 0xFF, 0xFF).ToHex());
    }

    [Fact]
    public void ToString_PrefixesAHash()
    {
        Assert.Equal("#112233", RgbColor.FromRgb(0x11, 0x22, 0x33).ToString());
    }

    [Theory]
    [InlineData("112233", 0x11, 0x22, 0x33)]
    [InlineData("#112233", 0x11, 0x22, 0x33)]
    [InlineData("  112233  ", 0x11, 0x22, 0x33)]
    [InlineData("AABBCC", 0xAA, 0xBB, 0xCC)]
    [InlineData("ffffff", 0xFF, 0xFF, 0xFF)]
    public void TryParseHex_AcceptsDocumentedForms(string hex, int red, int green, int blue)
    {
        bool parsed = RgbColor.TryParseHex(hex, out RgbColor color);

        Assert.True(parsed);
        Assert.Equal((byte)red, color.Red);
        Assert.Equal((byte)green, color.Green);
        Assert.Equal((byte)blue, color.Blue);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("auto")]
    [InlineData("11223")]
    [InlineData("1122334")]
    [InlineData("1122GG")]
    public void TryParseHex_RejectsValuesThatAreNotSixHexDigits(string? hex)
    {
        bool parsed = RgbColor.TryParseHex(hex, out RgbColor color);

        Assert.False(parsed);
        Assert.Equal(default, color);
    }

    [Fact]
    public void ToHex_And_TryParseHex_RoundTrip()
    {
        RgbColor original = RgbColor.FromRgb(0xDE, 0xAD, 0xBE);

        Assert.True(RgbColor.TryParseHex(original.ToHex(), out RgbColor roundTripped));
        Assert.Equal(original, roundTripped);
    }

    [Fact]
    public void Equality_ComparesEveryComponent()
    {
        RgbColor red = RgbColor.FromRgb(1, 2, 3);

        Assert.Equal(RgbColor.FromRgb(1, 2, 3), red);
        Assert.True(RgbColor.FromRgb(1, 2, 3) == red);
        Assert.False(RgbColor.FromRgb(1, 2, 4) == red);
        Assert.True(RgbColor.FromRgb(1, 2, 4) != red);
        Assert.Equal(red.GetHashCode(), RgbColor.FromRgb(1, 2, 3).GetHashCode());
        Assert.False(red.Equals(null));
    }
}
