using NUnit.Framework;
using Rubberduck.UI.Converters;
using System.Globalization;
using System.Windows.Media;

namespace RubberduckTests.Converters
{
    [TestFixture]
    [Category("Converters")]
    public class RgbShadeConverterTests
    {
        [Test]
        [TestCase("#762330", "#000000", -1)]
        [TestCase("#762330", "#0B0304", -0.9)]
        [TestCase("#762330", "#170709", -0.8)]
        [TestCase("#762330", "#230A0E", -0.7)]
        [TestCase("#762330", "#2F0E13", -0.6)]
        [TestCase("#762330", "#3B1118", -0.5)]
        [TestCase("#762330", "#46151C", -0.4)]
        [TestCase("#762330", "#521821", -0.3)]
        [TestCase("#762330", "#5E1C26", -0.2)]
        [TestCase("#762330", "#6A1F2B", -0.1)]
        [TestCase("#762330", "#762330", 0)]
        [TestCase("#762330", "#912B3B", 0.1)] //#833944
        [TestCase("#762330", "#AD3346", 0.2)] //#914F59
        [TestCase("#762330", "#C53E53", 0.3)] //#9F656E
        [TestCase("#762330", "#CD596C", 0.4)] //#AC7B82
        [TestCase("#762330", "#D67584", 0.5)] //#BA9197
        [TestCase("#762330", "#DE909C", 0.6)] //#C8A7AC
        [TestCase("#762330", "#E6ACB5", 0.7)] //#D5BDC0
        [TestCase("#762330", "#EEC7CE", 0.8)] //#E3D3D5
        [TestCase("#762330", "#F6E3E6", 0.9)] //#F1E9EA
        [TestCase("#762330", "#FEFFFF", 1)]   //#FFFFFF

        [TestCase("#787878", "#000000", -1)]
        [TestCase("#787878", "#0C0C0C", -0.9)]
        [TestCase("#787878", "#181818", -0.8)]
        [TestCase("#787878", "#242424", -0.7)]
        [TestCase("#787878", "#303030", -0.6)]
        [TestCase("#787878", "#3C3C3C", -0.5)]
        [TestCase("#787878", "#484848", -0.4)]
        [TestCase("#787878", "#545454", -0.3)]
        [TestCase("#787878", "#606060", -0.2)]
        [TestCase("#787878", "#6C6C6C", -0.1)]
        [TestCase("#787878", "#787878", 0)]
        [TestCase("#787878", "#858585", 0.1)] 
        [TestCase("#787878", "#939393", 0.2)] 
        [TestCase("#787878", "#A0A0A0", 0.3)] 
        [TestCase("#787878", "#AEAEAE", 0.4)] 
        [TestCase("#787878", "#BBBBBB", 0.5)] 
        [TestCase("#787878", "#C9C9C9", 0.6)] 
        [TestCase("#787878", "#D6D6D6", 0.7)] 
        [TestCase("#787878", "#E4E4E4", 0.8)] 
        [TestCase("#787878", "#F1F1F1", 0.9)] 
        [TestCase("#787878", "#FFFFFF", 1)]
        public void ColorAdjustments(string inputColor, string expectedColor, double adjustment)
        {
            var input = (Color)ColorConverter.ConvertFromString(inputColor);
            var expected = (Color)ColorConverter.ConvertFromString(expectedColor);
            var converter = new RgbShadeConverter();
            
            var actual = converter.Convert(input, typeof(Color), adjustment, CultureInfo.InvariantCulture);

            Assert.AreEqual(expected, actual);
        }
    }
}
