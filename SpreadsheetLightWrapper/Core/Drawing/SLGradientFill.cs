using System;
using System.Collections.Generic;
using System.Drawing;
using SpreadsheetLightWrapper.Core.misc;
using A = DocumentFormat.OpenXml.Drawing;

namespace SpreadsheetLightWrapper.Core.Drawing
{
    internal class SLGradientFill
    {
        private bool bRotateWithShape;
        private decimal decAngle;

        internal bool HasFlip;

        internal bool HasRotateWithShape;

        internal bool IsLinear = true;
        internal List<Color> listThemeColors;
        private A.TileFlipValues vFlip;

        internal SLGradientFill(List<Color> ThemeColors)
        {
            int i;
            listThemeColors = new List<Color>();
            for (i = 0; i < ThemeColors.Count; ++i)
                listThemeColors.Add(ThemeColors[i]);

            IsLinear = true;
            Angle = 0;
            PathType = A.PathShadeValues.Circle;
            Direction = SLGradientDirectionValues.Center;
            vFlip = A.TileFlipValues.None;
            HasFlip = false;
            bRotateWithShape = true;
            HasRotateWithShape = false;
            GradientStops = new List<SLGradientStop>();
        }

        /// <summary>
        ///     The interpolation angle ranging from 0 degrees to 359.9 degrees. 0 degrees mean from left to right, 90 degrees mean
        ///     from top to bottom, 180 degrees mean from right to left and 270 degrees mean from bottom to top. Accurate to
        ///     1/60000 of a degree.
        /// </summary>
        internal decimal Angle
        {
            get { return decAngle; }
            set
            {
                decAngle = value;
                if (decAngle < 0m) decAngle = 0m;
                if (decAngle >= 360m) decAngle = 359.9m;
            }
        }

        internal A.PathShadeValues PathType { get; set; }
        internal SLGradientDirectionValues Direction { get; set; }

        internal A.TileFlipValues Flip
        {
            get { return vFlip; }
            set
            {
                HasFlip = true;
                vFlip = value;
            }
        }

        internal bool RotateWithShape
        {
            get { return bRotateWithShape; }
            set
            {
                HasRotateWithShape = true;
                bRotateWithShape = value;
            }
        }

        internal List<SLGradientStop> GradientStops { get; set; }

        internal void SetLinearGradient(SLGradientPresetValues Preset, decimal Angle)
        {
            IsLinear = true;
            this.Angle = Angle;
            FillGradientStops(Preset);
        }

        internal void SetRadialGradient(SLGradientPresetValues Preset, SLGradientDirectionValues Direction)
        {
            IsLinear = false;
            PathType = A.PathShadeValues.Circle;
            this.Direction = Direction;
            FillGradientStops(Preset);
        }

        internal void SetRectangularGradient(SLGradientPresetValues Preset, SLGradientDirectionValues Direction)
        {
            IsLinear = false;
            PathType = A.PathShadeValues.Rectangle;
            this.Direction = Direction;
            FillGradientStops(Preset);
        }

        internal void SetPathGradient(SLGradientPresetValues Preset)
        {
            IsLinear = false;
            PathType = A.PathShadeValues.Shape;
            FillGradientStops(Preset);
        }

        internal void AppendGradientStop(Color Color, decimal Transparency, decimal Position)
        {
            var gs = new SLGradientStop(listThemeColors);
            gs.Color.SetColor(Color, Transparency);
            gs.Position = Position;
            GradientStops.Add(gs);
        }

        internal void AppendGradientStop(SLThemeColorIndexValues Color, double Tint, decimal Transparency,
            decimal Position)
        {
            var gs = new SLGradientStop(listThemeColors);
            gs.Color.SetColor(Color, Tint, Transparency);
            gs.Position = Position;
            GradientStops.Add(gs);
        }

        internal void ClearGradientStops()
        {
            GradientStops.Clear();
        }

        internal void FillGradientStops(SLGradientPresetValues PresetType)
        {
            GradientStops = new List<SLGradientStop>();
            switch (PresetType)
            {
                case SLGradientPresetValues.EarlySunset:
                    GradientStops.Add(new SLGradientStop(listThemeColors, "000082", 0));
                    GradientStops.Add(new SLGradientStop(listThemeColors, "66008F", 30));
                    GradientStops.Add(new SLGradientStop(listThemeColors, "BA0066", 64.999m));
                    GradientStops.Add(new SLGradientStop(listThemeColors, "FF0000", 89.999m));
                    GradientStops.Add(new SLGradientStop(listThemeColors, "FF8200", 100));
                    break;
                case SLGradientPresetValues.LateSunset:
                    GradientStops.Add(new SLGradientStop(listThemeColors, "000000", 0));
                    GradientStops.Add(new SLGradientStop(listThemeColors, "000040", 20));
                    GradientStops.Add(new SLGradientStop(listThemeColors, "400040", 50));
                    GradientStops.Add(new SLGradientStop(listThemeColors, "8F0040", 75));
                    GradientStops.Add(new SLGradientStop(listThemeColors, "F27300", 89.999m));
                    GradientStops.Add(new SLGradientStop(listThemeColors, "FFBF00", 100));
                    break;
                case SLGradientPresetValues.Nightfall:
                    GradientStops.Add(new SLGradientStop(listThemeColors, "000000", 0));
                    GradientStops.Add(new SLGradientStop(listThemeColors, "0A128C", 39.999m));
                    GradientStops.Add(new SLGradientStop(listThemeColors, "181CC7", 70));
                    GradientStops.Add(new SLGradientStop(listThemeColors, "7005D4", 88));
                    GradientStops.Add(new SLGradientStop(listThemeColors, "8C3D91", 100));
                    break;
                case SLGradientPresetValues.Daybreak:
                    GradientStops.Add(new SLGradientStop(listThemeColors, "5E9EFF", 0));
                    GradientStops.Add(new SLGradientStop(listThemeColors, "85C2FF", 39.999m));
                    GradientStops.Add(new SLGradientStop(listThemeColors, "C4D6EB", 70));
                    GradientStops.Add(new SLGradientStop(listThemeColors, "FFEBFA", 100));
                    break;
                case SLGradientPresetValues.Horizon:
                    GradientStops.Add(new SLGradientStop(listThemeColors, "DCEBF5", 0));
                    GradientStops.Add(new SLGradientStop(listThemeColors, "83A7C3", 8));
                    GradientStops.Add(new SLGradientStop(listThemeColors, "768FB9", 13));
                    GradientStops.Add(new SLGradientStop(listThemeColors, "83A7C3", 21.001m));
                    GradientStops.Add(new SLGradientStop(listThemeColors, "FFFFFF", 52));
                    GradientStops.Add(new SLGradientStop(listThemeColors, "9C6563", 56));
                    GradientStops.Add(new SLGradientStop(listThemeColors, "80302D", 58));
                    GradientStops.Add(new SLGradientStop(listThemeColors, "C0524E", 71.001m));
                    GradientStops.Add(new SLGradientStop(listThemeColors, "EBDAD4", 94));
                    GradientStops.Add(new SLGradientStop(listThemeColors, "55261C", 100));
                    break;
                case SLGradientPresetValues.Desert:
                    GradientStops.Add(new SLGradientStop(listThemeColors, "FC9FCB", 0));
                    GradientStops.Add(new SLGradientStop(listThemeColors, "F8B049", 13));
                    GradientStops.Add(new SLGradientStop(listThemeColors, "F8B049", 21.001m));
                    GradientStops.Add(new SLGradientStop(listThemeColors, "FEE7F2", 63));
                    GradientStops.Add(new SLGradientStop(listThemeColors, "F952A0", 67));
                    GradientStops.Add(new SLGradientStop(listThemeColors, "C50849", 69));
                    GradientStops.Add(new SLGradientStop(listThemeColors, "B43E85", 82.001m));
                    GradientStops.Add(new SLGradientStop(listThemeColors, "F8B049", 100));
                    break;
                case SLGradientPresetValues.Ocean:
                    GradientStops.Add(new SLGradientStop(listThemeColors, "03D4A8", 0));
                    GradientStops.Add(new SLGradientStop(listThemeColors, "21D6E0", 25));
                    GradientStops.Add(new SLGradientStop(listThemeColors, "0087E6", 75));
                    GradientStops.Add(new SLGradientStop(listThemeColors, "005CBF", 100));
                    break;
                case SLGradientPresetValues.CalmWater:
                    GradientStops.Add(new SLGradientStop(listThemeColors, "CCCCFF", 0));
                    GradientStops.Add(new SLGradientStop(listThemeColors, "99CCFF", 17.999m));
                    GradientStops.Add(new SLGradientStop(listThemeColors, "9966FF", 36));
                    GradientStops.Add(new SLGradientStop(listThemeColors, "CC99FF", 61));
                    GradientStops.Add(new SLGradientStop(listThemeColors, "99CCFF", 82.001m));
                    GradientStops.Add(new SLGradientStop(listThemeColors, "CCCCFF", 100));
                    break;
                case SLGradientPresetValues.Fire:
                    GradientStops.Add(new SLGradientStop(listThemeColors, "FFF200", 0));
                    GradientStops.Add(new SLGradientStop(listThemeColors, "FF7A00", 45));
                    GradientStops.Add(new SLGradientStop(listThemeColors, "FF0300", 70));
                    GradientStops.Add(new SLGradientStop(listThemeColors, "4D0808", 100));
                    break;
                case SLGradientPresetValues.Fog:
                    GradientStops.Add(new SLGradientStop(listThemeColors, "8488C4", 0));
                    GradientStops.Add(new SLGradientStop(listThemeColors, "D4DEFF", 53));
                    GradientStops.Add(new SLGradientStop(listThemeColors, "D4DEFF", 83));
                    GradientStops.Add(new SLGradientStop(listThemeColors, "96AB94", 100));
                    break;
                case SLGradientPresetValues.Moss:
                    GradientStops.Add(new SLGradientStop(listThemeColors, "DDEBCF", 0));
                    GradientStops.Add(new SLGradientStop(listThemeColors, "9CB86E", 50));
                    GradientStops.Add(new SLGradientStop(listThemeColors, "156B13", 100));
                    break;
                case SLGradientPresetValues.Peacock:
                    GradientStops.Add(new SLGradientStop(listThemeColors, "3399FF", 0));
                    GradientStops.Add(new SLGradientStop(listThemeColors, "00CCCC", 16));
                    GradientStops.Add(new SLGradientStop(listThemeColors, "9999FF", 47));
                    GradientStops.Add(new SLGradientStop(listThemeColors, "2E6792", 60.001m));
                    GradientStops.Add(new SLGradientStop(listThemeColors, "3333CC", 71.001m));
                    GradientStops.Add(new SLGradientStop(listThemeColors, "1170FF", 81));
                    GradientStops.Add(new SLGradientStop(listThemeColors, "006699", 100));
                    break;
                case SLGradientPresetValues.Wheat:
                    GradientStops.Add(new SLGradientStop(listThemeColors, "FBEAC7", 0));
                    GradientStops.Add(new SLGradientStop(listThemeColors, "FEE7F2", 17.999m));
                    GradientStops.Add(new SLGradientStop(listThemeColors, "FAC77D", 36));
                    GradientStops.Add(new SLGradientStop(listThemeColors, "FBA97D", 61));
                    GradientStops.Add(new SLGradientStop(listThemeColors, "FBD49C", 82.001m));
                    GradientStops.Add(new SLGradientStop(listThemeColors, "FEE7F2", 100));
                    break;
                case SLGradientPresetValues.Parchment:
                    GradientStops.Add(new SLGradientStop(listThemeColors, "FFEFD1", 0));
                    GradientStops.Add(new SLGradientStop(listThemeColors, "F0EBD5", 64.999m));
                    GradientStops.Add(new SLGradientStop(listThemeColors, "D1C39F", 100));
                    break;
                case SLGradientPresetValues.Mahogany:
                    GradientStops.Add(new SLGradientStop(listThemeColors, "D6B19C", 0));
                    GradientStops.Add(new SLGradientStop(listThemeColors, "D49E6C", 30));
                    GradientStops.Add(new SLGradientStop(listThemeColors, "A65528", 70));
                    GradientStops.Add(new SLGradientStop(listThemeColors, "663012", 100));
                    break;
                case SLGradientPresetValues.Rainbow:
                    GradientStops.Add(new SLGradientStop(listThemeColors, "A603AB", 0));
                    GradientStops.Add(new SLGradientStop(listThemeColors, "0819FB", 21.001m));
                    GradientStops.Add(new SLGradientStop(listThemeColors, "1A8D48", 35.001m));
                    GradientStops.Add(new SLGradientStop(listThemeColors, "FFFF00", 52));
                    GradientStops.Add(new SLGradientStop(listThemeColors, "EE3F17", 73));
                    GradientStops.Add(new SLGradientStop(listThemeColors, "E81766", 88));
                    GradientStops.Add(new SLGradientStop(listThemeColors, "A603AB", 100));
                    break;
                case SLGradientPresetValues.Rainbow2:
                    GradientStops.Add(new SLGradientStop(listThemeColors, "FF3399", 0));
                    GradientStops.Add(new SLGradientStop(listThemeColors, "FF6633", 25));
                    GradientStops.Add(new SLGradientStop(listThemeColors, "FFFF00", 50));
                    GradientStops.Add(new SLGradientStop(listThemeColors, "01A78F", 75));
                    GradientStops.Add(new SLGradientStop(listThemeColors, "3366FF", 100));
                    break;
                case SLGradientPresetValues.Gold:
                    GradientStops.Add(new SLGradientStop(listThemeColors, "E6DCAC", 0));
                    GradientStops.Add(new SLGradientStop(listThemeColors, "E6D78A", 12));
                    GradientStops.Add(new SLGradientStop(listThemeColors, "C7AC4C", 30));
                    GradientStops.Add(new SLGradientStop(listThemeColors, "E6D78A", 45));
                    GradientStops.Add(new SLGradientStop(listThemeColors, "C7AC4C", 77));
                    GradientStops.Add(new SLGradientStop(listThemeColors, "E6DCAC", 100));
                    break;
                case SLGradientPresetValues.Gold2:
                    GradientStops.Add(new SLGradientStop(listThemeColors, "FBE4AE", 0));
                    GradientStops.Add(new SLGradientStop(listThemeColors, "BD922A", 13));
                    GradientStops.Add(new SLGradientStop(listThemeColors, "BD922A", 21.001m));
                    GradientStops.Add(new SLGradientStop(listThemeColors, "FBE4AE", 63));
                    GradientStops.Add(new SLGradientStop(listThemeColors, "BD922A", 67));
                    GradientStops.Add(new SLGradientStop(listThemeColors, "835E17", 69));
                    GradientStops.Add(new SLGradientStop(listThemeColors, "A28949", 82.001m));
                    GradientStops.Add(new SLGradientStop(listThemeColors, "FAE3B7", 100));
                    break;
                case SLGradientPresetValues.Brass:
                    GradientStops.Add(new SLGradientStop(listThemeColors, "825600", 0));
                    GradientStops.Add(new SLGradientStop(listThemeColors, "FFA800", 13));
                    GradientStops.Add(new SLGradientStop(listThemeColors, "825600", 28));
                    GradientStops.Add(new SLGradientStop(listThemeColors, "FFA800", 42.999m));
                    GradientStops.Add(new SLGradientStop(listThemeColors, "825600", 58));
                    GradientStops.Add(new SLGradientStop(listThemeColors, "FFA800", 72));
                    GradientStops.Add(new SLGradientStop(listThemeColors, "825600", 87));
                    GradientStops.Add(new SLGradientStop(listThemeColors, "FFA800", 100));
                    break;
                case SLGradientPresetValues.Chrome:
                    GradientStops.Add(new SLGradientStop(listThemeColors, "FFFFFF", 0));
                    GradientStops.Add(new SLGradientStop(listThemeColors, "1F1F1F", 16));
                    GradientStops.Add(new SLGradientStop(listThemeColors, "FFFFFF", 17.999m));
                    GradientStops.Add(new SLGradientStop(listThemeColors, "636363", 42));
                    GradientStops.Add(new SLGradientStop(listThemeColors, "CFCFCF", 53));
                    GradientStops.Add(new SLGradientStop(listThemeColors, "CFCFCF", 66));
                    GradientStops.Add(new SLGradientStop(listThemeColors, "1F1F1F", 75.999m));
                    GradientStops.Add(new SLGradientStop(listThemeColors, "FFFFFF", 78.999m));
                    GradientStops.Add(new SLGradientStop(listThemeColors, "7F7F7F", 100));
                    break;
                case SLGradientPresetValues.Chrome2:
                    GradientStops.Add(new SLGradientStop(listThemeColors, "CBCBCB", 0));
                    GradientStops.Add(new SLGradientStop(listThemeColors, "5F5F5F", 13));
                    GradientStops.Add(new SLGradientStop(listThemeColors, "5F5F5F", 21.001m));
                    GradientStops.Add(new SLGradientStop(listThemeColors, "FFFFFF", 63));
                    GradientStops.Add(new SLGradientStop(listThemeColors, "B2B2B2", 67));
                    GradientStops.Add(new SLGradientStop(listThemeColors, "292929", 69));
                    GradientStops.Add(new SLGradientStop(listThemeColors, "777777", 82.001m));
                    GradientStops.Add(new SLGradientStop(listThemeColors, "EAEAEA", 100));
                    break;
                case SLGradientPresetValues.Silver:
                    GradientStops.Add(new SLGradientStop(listThemeColors, "FFFFFF", 0));
                    GradientStops.Add(new SLGradientStop(listThemeColors, "E6E6E6", 7.001m));
                    GradientStops.Add(new SLGradientStop(listThemeColors, "7D8496", 32.001m));
                    GradientStops.Add(new SLGradientStop(listThemeColors, "E6E6E6", 47));
                    GradientStops.Add(new SLGradientStop(listThemeColors, "7D8496", 85.001m));
                    GradientStops.Add(new SLGradientStop(listThemeColors, "E6E6E6", 100));
                    break;
                case SLGradientPresetValues.Sapphire:
                    GradientStops.Add(new SLGradientStop(listThemeColors, "000082", 0));
                    GradientStops.Add(new SLGradientStop(listThemeColors, "0047FF", 13));
                    GradientStops.Add(new SLGradientStop(listThemeColors, "000082", 28));
                    GradientStops.Add(new SLGradientStop(listThemeColors, "0047FF", 42.999m));
                    GradientStops.Add(new SLGradientStop(listThemeColors, "000082", 58));
                    GradientStops.Add(new SLGradientStop(listThemeColors, "0047FF", 72));
                    GradientStops.Add(new SLGradientStop(listThemeColors, "000082", 87));
                    GradientStops.Add(new SLGradientStop(listThemeColors, "0047FF", 100));
                    break;
            }
        }

        internal A.GradientFill ToGradientFill()
        {
            var gf = new A.GradientFill();

            var gsl = new A.GradientStopList();
            for (var i = 0; i < GradientStops.Count; ++i)
                gsl.Append(GradientStops[i].ToGradientStop());
            gf.Append(gsl);

            if (IsLinear)
            {
                var lgf = new A.LinearGradientFill();
                lgf.Angle = Convert.ToInt32(Angle*SLConstants.DegreeToAngleRepresentation);
                lgf.Scaled = false;
                gf.Append(lgf);
                gf.Append(new A.TileRectangle());
            }
            else
            {
                if (PathType == A.PathShadeValues.Shape)
                {
                    var pgf = new A.PathGradientFill();
                    pgf.Path = PathType;
                    pgf.FillToRectangle = new A.FillToRectangle
                    {
                        Left = 50000,
                        Top = 50000,
                        Right = 50000,
                        Bottom = 50000
                    };
                    gf.Append(pgf);
                    gf.Append(new A.TileRectangle());
                }
                else
                {
                    var pgf = new A.PathGradientFill();
                    pgf.Path = PathType;
                    switch (Direction)
                    {
                        case SLGradientDirectionValues.CenterToTopLeftCorner:
                            pgf.FillToRectangle = new A.FillToRectangle {Left = 100000, Top = 100000};
                            gf.Append(pgf);
                            gf.Append(new A.TileRectangle {Right = -100000, Bottom = -100000});
                            break;
                        case SLGradientDirectionValues.CenterToTopRightCorner:
                            pgf.FillToRectangle = new A.FillToRectangle {Top = 100000, Right = 100000};
                            gf.Append(pgf);
                            gf.Append(new A.TileRectangle {Left = -100000, Bottom = -100000});
                            break;
                        case SLGradientDirectionValues.Center:
                            pgf.FillToRectangle = new A.FillToRectangle
                            {
                                Left = 50000,
                                Top = 50000,
                                Right = 50000,
                                Bottom = 50000
                            };
                            gf.Append(pgf);
                            gf.Append(new A.TileRectangle());
                            break;
                        case SLGradientDirectionValues.CenterToBottomLeftCorner:
                            pgf.FillToRectangle = new A.FillToRectangle {Left = 100000, Bottom = 100000};
                            gf.Append(pgf);
                            gf.Append(new A.TileRectangle {Top = -100000, Right = -100000});
                            break;
                        case SLGradientDirectionValues.CenterToBottomRightCorner:
                            pgf.FillToRectangle = new A.FillToRectangle {Right = 100000, Bottom = 100000};
                            gf.Append(pgf);
                            gf.Append(new A.TileRectangle {Left = -100000, Top = -100000});
                            break;
                    }
                }
            }

            if (HasFlip) gf.Flip = Flip;
            if (HasRotateWithShape) gf.RotateWithShape = RotateWithShape;

            return gf;
        }

        internal SLGradientFill Clone()
        {
            var gf = new SLGradientFill(listThemeColors);
            gf.IsLinear = IsLinear;
            gf.decAngle = decAngle;
            gf.PathType = PathType;
            gf.Direction = Direction;
            gf.HasFlip = HasFlip;
            gf.vFlip = vFlip;
            gf.HasRotateWithShape = HasRotateWithShape;
            gf.bRotateWithShape = bRotateWithShape;
            for (var i = 0; i < GradientStops.Count; ++i)
                gf.GradientStops.Add(GradientStops[i].Clone());

            return gf;
        }
    }
}