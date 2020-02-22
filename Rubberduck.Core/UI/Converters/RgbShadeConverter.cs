using System;
using System.Globalization;
using System.Windows.Data;
using System.Windows.Media;

namespace Rubberduck.UI.Converters
{
    /// <remarks>
    /// Adapted from https://www.codeproject.com/Articles/49650/WPF-Color-Conversions
    /// Adjusted to handle both color or brush input and return similar output.
    /// </remarks>
    public class RgbShadeConverter : IValueConverter
    {
        private struct HlsColor
        {
            public double A { get; set; }
            public double H { get; set; }
            public double L { get; set; }
            public double S { get; set; }
        }

        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value == null) throw new ArgumentNullException(nameof(value));
            if (parameter == null) throw new ArgumentNullException(nameof(parameter));

            // Adjust color by factor passed in
            var brightnessAdjustment = double.Parse(parameter.ToString());
            if (brightnessAdjustment < -1 || brightnessAdjustment > 1)
            {
                throw new ArgumentOutOfRangeException(nameof(parameter), parameter, "Must be between -1 and 1");
            }

            Color rgbColorIn;
            bool returnBrush = false;
            switch(value)
            {
                case SolidColorBrush brush:
                    rgbColorIn = brush.Color;
                    returnBrush = true;
                    break;
                case Color color:
                    rgbColorIn = color;
                    break;
                default:
                    throw new ArgumentOutOfRangeException("parameter", "Parameter must be either a brush or a color");
            }
            
            var hlsColor = RgbToHls(rgbColorIn);
            
            if (brightnessAdjustment > 0)
            {
                // Brighten the color by a % between the value of L to 1 (the maximum)
                var lbound = hlsColor.L;
                const double ubound = 1;

                hlsColor.L = ((ubound - lbound) * brightnessAdjustment) + lbound;
            }
            else if (brightnessAdjustment < 0)
            {
                // Darken the color by a % between 0 (the minimum) to the value of L
                // Because the lower bound is 0, the formula is simplified.
                // Also, -0.1 becomes 10% darker, -0.9 becomes 90% darker.
                hlsColor.L -= hlsColor.L * Math.Abs(brightnessAdjustment);
            }

            // Return result
            var rgbColorOut = HlsToRgb(hlsColor);
            
            if (returnBrush)
            {
                var brushOut = new SolidColorBrush
                {
                    Color = rgbColorOut
                };
                return brushOut;
            }

            return rgbColorOut;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if(value is Color rgb && parameter is HlsColor adj)
            {
                throw new NotImplementedException();
            }

            return null;
        }

        /// <summary>
        /// Converts a WPF RGB color to an HSL color
        /// </summary>
        /// <param name="rgbColor">The RGB color to convert.</param>
        /// <returns>An HSL color object equivalent to the RGB color object passed in.</returns>
        static HlsColor RgbToHls(Color rgbColor)
        {
            // Initialize result
            var hlsColor = new HlsColor();

            // Convert RGB values to percentages
            double r = (double)rgbColor.R / 255;
            var g = (double)rgbColor.G / 255;
            var b = (double)rgbColor.B / 255;
            hlsColor.A = (double)rgbColor.A / 255;

            // Find min and max RGB values
            var min = Math.Min(r, Math.Min(g, b));
            var max = Math.Max(r, Math.Max(g, b));
            var delta = max - min;

            /* If max and min are equal, that means we are dealing with 
             * a shade of gray. So we set H and S to zero, and L to either
             * max or min (it doesn't matter which), and  then we exit. */

            //Special case: Gray
            if (max == min)
            {
                hlsColor.H = 0;
                hlsColor.S = 0;
                hlsColor.L = max;
                return hlsColor;
            }

            /* If we get to this point, we know we don't have a shade of gray. */
            
            // Set L
            hlsColor.L = (min + max) / 2;

            // Set S
            if (hlsColor.L < 0.5)
            {
                hlsColor.S = delta / (max + min);
            }
            else
            {
                hlsColor.S = delta / (2.0 - max - min);
            }

            // Set H
            if (r == max) hlsColor.H = (g - b) / delta;
            if (g == max) hlsColor.H = 2.0 + (b - r) / delta;
            if (b == max) hlsColor.H = 4.0 + (r - g) / delta;
            hlsColor.H *= 60;
            if (hlsColor.H < 0) hlsColor.H += 360;

            // Set return value
            return hlsColor;
        }

        /// <summary>
        /// Converts a WPF HSL color to an RGB color
        /// </summary>
        /// <param name="hlsColor">The HSL color to convert.</param>
        /// <returns>An RGB color object equivalent to the HSL color object passed in.</returns>
        static Color HlsToRgb(HlsColor hlsColor)
        {
            // Initialize result
            var rgbColor = new Color();

            /* If S = 0, that means we are dealing with a shade 
             * of gray. So, we set R, G, and B to L and exit. */

            // Special case: Gray
            if (hlsColor.S == 0)
            {
                rgbColor.R = (byte)(hlsColor.L * 255);
                rgbColor.G = (byte)(hlsColor.L * 255);
                rgbColor.B = (byte)(hlsColor.L * 255);
                rgbColor.A = (byte)(hlsColor.A * 255);
                return rgbColor;
            }

            double t1;
            if (hlsColor.L < 0.5)
            {
                t1 = hlsColor.L * (1.0 + hlsColor.S);
            }
            else
            {
                t1 = hlsColor.L + hlsColor.S - (hlsColor.L * hlsColor.S);
            }

            var t2 = 2.0 * hlsColor.L - t1;

            // Convert H from degrees to a percentage
            var h = hlsColor.H / 360;

            // Set colors as percentage values
            var tR = h + (1.0 / 3.0);
            var r = SetColor(t1, t2, tR);

            var tG = h;
            var g = SetColor(t1, t2, tG);

            var tB = h - (1.0 / 3.0);
            var b = SetColor(t1, t2, tB);

            // Assign colors to Color object
            rgbColor.R = (byte)(r * 255);
            rgbColor.G = (byte)(g * 255);
            rgbColor.B = (byte)(b * 255);
            rgbColor.A = (byte)(hlsColor.A * 255);

            // Set return value
            return rgbColor;
        }

        /// <summary>
        /// Used by the HSL-to-RGB converter.
        /// </summary>
        /// <param name="t1">A temporary variable.</param>
        /// <param name="t2">A temporary variable.</param>
        /// <param name="t3">A temporary variable.</param>
        /// <returns>An RGB color value, in decimal format.</returns>
        private static double SetColor(double t1, double t2, double t3)
        {
            if (t3 < 0) t3 += 1.0;
            if (t3 > 1) t3 -= 1.0;

            double color;
            if (6.0 * t3 < 1)
            {
                color = t2 + (t1 - t2) * 6.0 * t3;
            }
            else if (2.0 * t3 < 1)
            {
                color = t1;
            }
            else if (3.0 * t3 < 2)
            {
                color = t2 + (t1 - t2) * ((2.0 / 3.0) - t3) * 6.0;
            }
            else
            {
                color = t2;
            }

            // Set return value
            return color;
        }
    }
}
