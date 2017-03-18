using A = DocumentFormat.OpenXml.Drawing;

namespace Ups.Toolkit.SpreadsheetLight.Core.Drawing
{
    /// <summary>
    ///     Specifies camera preset settings.
    /// </summary>
    public enum SLCameraPresetValues
    {
        /// <summary>
        ///     None
        /// </summary>
        None = 0,

        /// <summary>
        ///     Isometric Left Down
        /// </summary>
        IsometricLeftDown,

        /// <summary>
        ///     Isometric Right Up
        /// </summary>
        IsometricRightUp,

        /// <summary>
        ///     Isometric Top Up
        /// </summary>
        IsometricTopUp,

        /// <summary>
        ///     Isometric Bottom Down
        /// </summary>
        IsometricBottomDown,

        /// <summary>
        ///     Off Axis 1 Left
        /// </summary>
        OffAxis1Left,

        /// <summary>
        ///     Off Axis 1 Right
        /// </summary>
        OffAxis1Right,

        /// <summary>
        ///     Off Axis 1 Top
        /// </summary>
        OffAxis1Top,

        /// <summary>
        ///     Off Axis 2 Left
        /// </summary>
        OffAxis2Left,

        /// <summary>
        ///     Off Axis 2 Right
        /// </summary>
        OffAxis2Right,

        /// <summary>
        ///     Off Axis 2 Top
        /// </summary>
        OffAxis2Top,

        /// <summary>
        ///     Perspective Front
        /// </summary>
        PerspectiveFront,

        /// <summary>
        ///     Perspective Left
        /// </summary>
        PerspectiveLeft,

        /// <summary>
        ///     Perspective Right
        /// </summary>
        PerspectiveRight,

        /// <summary>
        ///     Perspective Below
        /// </summary>
        PerspectiveBelow,

        /// <summary>
        ///     Perspective Above
        /// </summary>
        PerspectiveAbove,

        /// <summary>
        ///     Perspective Relaxed Moderately
        /// </summary>
        PerspectiveRelaxedModerately,

        /// <summary>
        ///     Perspective Relaxed
        /// </summary>
        PerspectiveRelaxed,

        /// <summary>
        ///     Perspective Contrasting Left
        /// </summary>
        PerspectiveContrastingLeft,

        /// <summary>
        ///     Perspective Contrasting Right
        /// </summary>
        PerspectiveContrastingRight,

        /// <summary>
        ///     Perspective Heroic Extreme Left
        /// </summary>
        PerspectiveHeroicExtremeLeft,

        /// <summary>
        ///     Perspective Heroic Extreme Right
        /// </summary>
        PerspectiveHeroicExtremeRight,

        /// <summary>
        ///     Oblique Top Left
        /// </summary>
        ObliqueTopLeft,

        /// <summary>
        ///     Oblique Top Right
        /// </summary>
        ObliqueTopRight,

        /// <summary>
        ///     Oblique Bottom Left
        /// </summary>
        ObliqueBottomLeft,

        /// <summary>
        ///     Oblique Bottom Right
        /// </summary>
        ObliqueBottomRight
    }

    /// <summary>
    ///     Encapsulates 3D rotation properties. Works together with SLFormat3D class.
    ///     This simulates some properties of DocumentFormat.OpenXml.Drawing.Scene3DType
    ///     and DocumentFormat.OpenXml.Drawing.Shape3DType classes. The reason for this mixing
    ///     is because Excel separates different properties from both classes into 2 separate sections
    ///     on the user interface (3-D Format and 3-D Rotation). Hence SLRotation3D and SLFormat3D
    ///     classes instead of straightforward mapping of the SDK Scene3DType and Shape3DType classes.
    /// </summary>
    public class SLRotation3D
    {
        internal decimal decDistanceZ;

        internal decimal decPerspective;

        internal decimal decX;

        internal decimal decY;

        internal decimal decZ;
        internal bool HasCamera;
        internal bool HasPerspectiveSet;

        internal bool HasXYZSet;

        /// <summary>
        ///     Initializes an instance of SLRotation3D.
        /// </summary>
        public SLRotation3D()
        {
            SetAllNull();
        }

        internal A.PresetCameraValues CameraPreset { get; set; }

        /// <summary>
        ///     Longitude angle ranging from 0 degrees to 359.9 degrees. Accurate to 1/60000 of a degree.
        /// </summary>
        public decimal X
        {
            get { return decX; }
            set
            {
                decX = value;
                if (decX < 0m) decX = 0m;
                if (decX >= 360m) decX = 359.9m;
                HasCamera = true;
                HasXYZSet = true;
            }
        }

        /// <summary>
        ///     Latitude angle ranging from 0 degrees to 359.9 degrees. Accurate to 1/60000 of a degree.
        /// </summary>
        public decimal Y
        {
            get { return decY; }
            set
            {
                decY = value;
                if (decY < 0m) decY = 0m;
                if (decY >= 360m) decY = 359.9m;
                HasCamera = true;
                HasXYZSet = true;
            }
        }

        /// <summary>
        ///     Revolution angle ranging from 0 degrees to 359.9 degrees. Accurate to 1/60000 of a degree.
        /// </summary>
        public decimal Z
        {
            get { return decZ; }
            set
            {
                decZ = value;
                if (decZ < 0m) decZ = 0m;
                if (decZ >= 360m) decZ = 359.9m;
                HasCamera = true;
                HasXYZSet = true;
            }
        }

        /// <summary>
        ///     Perspective angle ranging from 0 degrees to 180 degrees. However, a suggested maximum is 120 degrees.
        /// </summary>
        public decimal Perspective
        {
            get { return decPerspective; }
            set
            {
                if (IsPerspectiveView(CameraPreset))
                {
                    decPerspective = value;
                    if (decPerspective < 0m) decPerspective = 0m;
                    if (decPerspective > 180m) decPerspective = 180m;
                    HasCamera = true;
                    HasPerspectiveSet = true;
                }
            }
        }

        /// <summary>
        ///     Distance from the ground, ranging from -2147483648 pt to 2147483647 pt. However, a suggested range is -4000 pt to
        ///     4000 pt.
        /// </summary>
        public decimal DistanceZ
        {
            get { return decDistanceZ; }
            set
            {
                decDistanceZ = value;
                if (decDistanceZ < -2147483648m) decDistanceZ = -2147483648m;
                if (decDistanceZ > 2147483647m) decDistanceZ = 2147483647m;
            }
        }

        private void SetAllNull()
        {
            HasCamera = false;
            HasXYZSet = false;
            HasPerspectiveSet = false;
            CameraPreset = A.PresetCameraValues.OrthographicFront;
            decX = 0;
            decY = 0;
            decZ = 0;
            decPerspective = 0;
            decDistanceZ = 0;
        }

        /// <summary>
        ///     Set camera settings using a preset.
        /// </summary>
        /// <param name="Preset">The preset to be used.</param>
        public void SetCameraPreset(SLCameraPresetValues Preset)
        {
            switch (Preset)
            {
                case SLCameraPresetValues.None:
                    CameraPreset = A.PresetCameraValues.OrthographicFront;
                    decX = 0;
                    decY = 0;
                    decZ = 0;
                    decPerspective = 0;
                    HasCamera = false;
                    HasXYZSet = false;
                    HasPerspectiveSet = false;
                    break;
                case SLCameraPresetValues.IsometricLeftDown:
                    CameraPreset = A.PresetCameraValues.IsometricLeftDown;
                    decX = 45;
                    decY = 35;
                    decZ = 0;
                    decPerspective = 0;
                    HasCamera = true;
                    HasXYZSet = false;
                    HasPerspectiveSet = false;
                    break;
                case SLCameraPresetValues.IsometricRightUp:
                    CameraPreset = A.PresetCameraValues.IsometricRightUp;
                    decX = 315;
                    decY = 35;
                    decZ = 0;
                    decPerspective = 0;
                    HasCamera = true;
                    HasXYZSet = false;
                    HasPerspectiveSet = false;
                    break;
                case SLCameraPresetValues.IsometricTopUp:
                    CameraPreset = A.PresetCameraValues.IsometricTopUp;
                    decX = 314.7m;
                    decY = 324.6m;
                    decZ = 60.2m;
                    decPerspective = 0;
                    HasCamera = true;
                    HasXYZSet = false;
                    HasPerspectiveSet = false;
                    break;
                case SLCameraPresetValues.IsometricBottomDown:
                    CameraPreset = A.PresetCameraValues.IsometricBottomDown;
                    decX = 314.7m;
                    decY = 35.39999999999999m;
                    decZ = 299.8m;
                    decPerspective = 0;
                    HasCamera = true;
                    HasXYZSet = false;
                    HasPerspectiveSet = false;
                    break;
                case SLCameraPresetValues.OffAxis1Left:
                    CameraPreset = A.PresetCameraValues.IsometricOffAxis1Left;
                    decX = 64m;
                    decY = 18m;
                    decZ = 0;
                    decPerspective = 0;
                    HasCamera = true;
                    HasXYZSet = false;
                    HasPerspectiveSet = false;
                    break;
                case SLCameraPresetValues.OffAxis1Right:
                    CameraPreset = A.PresetCameraValues.IsometricOffAxis1Right;
                    decX = 334m;
                    decY = 18m;
                    decZ = 0;
                    decPerspective = 0;
                    HasCamera = true;
                    HasXYZSet = false;
                    HasPerspectiveSet = false;
                    break;
                case SLCameraPresetValues.OffAxis1Top:
                    CameraPreset = A.PresetCameraValues.IsometricOffAxis1Top;
                    decX = 306.5m;
                    decY = 301.3m;
                    decZ = 57.6m;
                    decPerspective = 0;
                    HasCamera = true;
                    HasXYZSet = false;
                    HasPerspectiveSet = false;
                    break;
                case SLCameraPresetValues.OffAxis2Left:
                    CameraPreset = A.PresetCameraValues.IsometricOffAxis2Left;
                    decX = 26m;
                    decY = 18m;
                    decZ = 0m;
                    decPerspective = 0;
                    HasCamera = true;
                    HasXYZSet = false;
                    HasPerspectiveSet = false;
                    break;
                case SLCameraPresetValues.OffAxis2Right:
                    CameraPreset = A.PresetCameraValues.IsometricOffAxis2Right;
                    decX = 296m;
                    decY = 18m;
                    decZ = 0m;
                    decPerspective = 0;
                    HasCamera = true;
                    HasXYZSet = false;
                    HasPerspectiveSet = false;
                    break;
                case SLCameraPresetValues.OffAxis2Top:
                    CameraPreset = A.PresetCameraValues.IsometricOffAxis2Top;
                    decX = 53.49999999999999m;
                    decY = 301.3m;
                    decZ = 302.4m;
                    decPerspective = 0;
                    HasCamera = true;
                    HasXYZSet = false;
                    HasPerspectiveSet = false;
                    break;
                case SLCameraPresetValues.PerspectiveFront:
                    CameraPreset = A.PresetCameraValues.PerspectiveFront;
                    decX = 0m;
                    decY = 0m;
                    decZ = 0m;
                    decPerspective = 45m;
                    HasCamera = true;
                    HasXYZSet = false;
                    HasPerspectiveSet = false;
                    break;
                case SLCameraPresetValues.PerspectiveLeft:
                    CameraPreset = A.PresetCameraValues.PerspectiveLeft;
                    decX = 20m;
                    decY = 0m;
                    decZ = 0m;
                    decPerspective = 45m;
                    HasCamera = true;
                    HasXYZSet = false;
                    HasPerspectiveSet = false;
                    break;
                case SLCameraPresetValues.PerspectiveRight:
                    CameraPreset = A.PresetCameraValues.PerspectiveRight;
                    decX = 340m;
                    decY = 0m;
                    decZ = 0m;
                    decPerspective = 45m;
                    HasCamera = true;
                    HasXYZSet = false;
                    HasPerspectiveSet = false;
                    break;
                case SLCameraPresetValues.PerspectiveBelow:
                    CameraPreset = A.PresetCameraValues.PerspectiveBelow;
                    decX = 0m;
                    decY = 20m;
                    decZ = 0m;
                    decPerspective = 45m;
                    HasCamera = true;
                    HasXYZSet = false;
                    HasPerspectiveSet = false;
                    break;
                case SLCameraPresetValues.PerspectiveAbove:
                    CameraPreset = A.PresetCameraValues.PerspectiveAbove;
                    decX = 0m;
                    decY = 340m;
                    decZ = 0m;
                    decPerspective = 45m;
                    HasCamera = true;
                    HasXYZSet = false;
                    HasPerspectiveSet = false;
                    break;
                case SLCameraPresetValues.PerspectiveRelaxedModerately:
                    CameraPreset = A.PresetCameraValues.PerspectiveRelaxedModerately;
                    decX = 0m;
                    decY = 324.8m;
                    decZ = 0m;
                    decPerspective = 45m;
                    HasCamera = true;
                    HasXYZSet = false;
                    HasPerspectiveSet = false;
                    break;
                case SLCameraPresetValues.PerspectiveRelaxed:
                    CameraPreset = A.PresetCameraValues.PerspectiveRelaxed;
                    decX = 0m;
                    decY = 309.6m;
                    decZ = 0m;
                    decPerspective = 45m;
                    HasCamera = true;
                    HasXYZSet = false;
                    HasPerspectiveSet = false;
                    break;
                case SLCameraPresetValues.PerspectiveContrastingLeft:
                    CameraPreset = A.PresetCameraValues.PerspectiveContrastingLeftFacing;
                    decX = 43.89999999999999m;
                    decY = 10.4m;
                    decZ = 356.4m;
                    decPerspective = 45m;
                    HasCamera = true;
                    HasXYZSet = false;
                    HasPerspectiveSet = false;
                    break;
                case SLCameraPresetValues.PerspectiveContrastingRight:
                    CameraPreset = A.PresetCameraValues.PerspectiveContrastingRightFacing;
                    decX = 316.1m;
                    decY = 10.4m;
                    decZ = 3.6m;
                    decPerspective = 45m;
                    HasCamera = true;
                    HasXYZSet = false;
                    HasPerspectiveSet = false;
                    break;
                case SLCameraPresetValues.PerspectiveHeroicExtremeLeft:
                    CameraPreset = A.PresetCameraValues.PerspectiveHeroicExtremeLeftFacing;
                    decX = 34.49999999999999m;
                    decY = 8.1m;
                    decZ = 357.1m;
                    decPerspective = 80m;
                    HasCamera = true;
                    HasXYZSet = false;
                    HasPerspectiveSet = false;
                    break;
                case SLCameraPresetValues.PerspectiveHeroicExtremeRight:
                    CameraPreset = A.PresetCameraValues.PerspectiveHeroicExtremeRightFacing;
                    decX = 325.5m;
                    decY = 8.1m;
                    decZ = 2.9m;
                    decPerspective = 80m;
                    HasCamera = true;
                    HasXYZSet = false;
                    HasPerspectiveSet = false;
                    break;
                case SLCameraPresetValues.ObliqueTopLeft:
                    CameraPreset = A.PresetCameraValues.ObliqueTopLeft;
                    decX = 0m;
                    decY = 0m;
                    decZ = 0m;
                    decPerspective = 0m;
                    HasCamera = true;
                    HasXYZSet = false;
                    HasPerspectiveSet = false;
                    break;
                case SLCameraPresetValues.ObliqueTopRight:
                    CameraPreset = A.PresetCameraValues.ObliqueTopRight;
                    decX = 0m;
                    decY = 0m;
                    decZ = 0m;
                    decPerspective = 0m;
                    HasCamera = true;
                    HasXYZSet = false;
                    HasPerspectiveSet = false;
                    break;
                case SLCameraPresetValues.ObliqueBottomLeft:
                    CameraPreset = A.PresetCameraValues.ObliqueBottomLeft;
                    decX = 0m;
                    decY = 0m;
                    decZ = 0m;
                    decPerspective = 0m;
                    HasCamera = true;
                    HasXYZSet = false;
                    HasPerspectiveSet = false;
                    break;
                case SLCameraPresetValues.ObliqueBottomRight:
                    CameraPreset = A.PresetCameraValues.ObliqueBottomRight;
                    decX = 0m;
                    decY = 0m;
                    decZ = 0m;
                    decPerspective = 0m;
                    HasCamera = true;
                    HasXYZSet = false;
                    HasPerspectiveSet = false;
                    break;
            }
        }

        private bool IsPerspectiveView(A.PresetCameraValues Preset)
        {
            var result = false;

            switch (Preset)
            {
                case A.PresetCameraValues.LegacyPerspectiveBottom:
                case A.PresetCameraValues.LegacyPerspectiveBottomLeft:
                case A.PresetCameraValues.LegacyPerspectiveBottomRight:
                case A.PresetCameraValues.LegacyPerspectiveFront:
                case A.PresetCameraValues.LegacyPerspectiveLeft:
                case A.PresetCameraValues.LegacyPerspectiveRight:
                case A.PresetCameraValues.LegacyPerspectiveTop:
                case A.PresetCameraValues.LegacyPerspectiveTopLeft:
                case A.PresetCameraValues.LegacyPerspectiveTopRight:
                case A.PresetCameraValues.PerspectiveAbove:
                case A.PresetCameraValues.PerspectiveAboveLeftFacing:
                case A.PresetCameraValues.PerspectiveAboveRightFacing:
                case A.PresetCameraValues.PerspectiveBelow:
                case A.PresetCameraValues.PerspectiveContrastingLeftFacing:
                case A.PresetCameraValues.PerspectiveContrastingRightFacing:
                case A.PresetCameraValues.PerspectiveFront:
                case A.PresetCameraValues.PerspectiveHeroicExtremeLeftFacing:
                case A.PresetCameraValues.PerspectiveHeroicExtremeRightFacing:
                case A.PresetCameraValues.PerspectiveHeroicLeftFacing:
                case A.PresetCameraValues.PerspectiveHeroicRightFacing:
                case A.PresetCameraValues.PerspectiveLeft:
                case A.PresetCameraValues.PerspectiveRelaxed:
                case A.PresetCameraValues.PerspectiveRelaxedModerately:
                case A.PresetCameraValues.PerspectiveRight:
                    result = true;
                    break;
            }

            return result;
        }

        internal SLRotation3D Clone()
        {
            var rot = new SLRotation3D();
            rot.HasCamera = HasCamera;
            rot.CameraPreset = CameraPreset;
            rot.HasXYZSet = HasXYZSet;
            rot.HasPerspectiveSet = HasPerspectiveSet;
            rot.decX = decX;
            rot.decY = decY;
            rot.decZ = decZ;
            rot.decPerspective = decPerspective;
            rot.decDistanceZ = decDistanceZ;

            return rot;
        }
    }
}