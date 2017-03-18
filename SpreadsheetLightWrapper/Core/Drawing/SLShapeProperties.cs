using System.Collections.Generic;
using System.Drawing;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace SpreadsheetLightWrapper.Core.Drawing
{
    internal class SLShapeProperties
    {
        internal bool HasBlackWhiteMode;

        internal bool HasPresetGeometry;

        internal bool HasTransform2D;
        internal List<Color> listThemeColors;
        internal A.BlackWhiteModeValues vBlackWhiteMode;
        internal A.ShapeTypeValues vPresetGeometry;

        internal SLShapeProperties(List<Color> ThemeColors)
        {
            int i;
            listThemeColors = new List<Color>();
            for (i = 0; i < ThemeColors.Count; ++i)
                listThemeColors.Add(ThemeColors[i]);

            SetAllNull();
        }

        internal bool HasShapeProperties
        {
            get
            {
                return HasBlackWhiteMode || HasTransform2D || HasPresetGeometry
                       || Fill.HasFill || Outline.HasLine
                       || EffectList.HasEffectList || Rotation3D.HasCamera || Format3D.HasLighting
                       || Format3D.HasBevelTop || Format3D.HasBevelBottom || Format3D.HasExtrusionColor
                       || Format3D.HasContourColor || (Format3D.ExtrusionHeight != 0)
                       || (Format3D.ContourWidth != 0) || (Format3D.Material != A.PresetMaterialTypeValues.WarmMatte)
                       || (Rotation3D.DistanceZ != 0);
            }
        }

        internal A.BlackWhiteModeValues BlackWhiteMode
        {
            get { return vBlackWhiteMode; }
            set
            {
                vBlackWhiteMode = value;
                HasBlackWhiteMode = true;
            }
        }

        internal SLTransform2D Transform2D { get; set; }

        internal A.ShapeTypeValues PresetGeometry
        {
            get { return vPresetGeometry; }
            set
            {
                vPresetGeometry = value;
                HasPresetGeometry = true;
            }
        }

        internal SLFill Fill { get; set; }
        internal SLLinePropertiesType Outline { get; set; }

        internal SLEffectList EffectList { get; set; }

        internal SLRotation3D Rotation3D { get; set; }
        internal SLFormat3D Format3D { get; set; }

        private void SetAllNull()
        {
            vBlackWhiteMode = A.BlackWhiteModeValues.Auto;
            HasBlackWhiteMode = false;

            Transform2D = new SLTransform2D();
            HasTransform2D = false;
            vPresetGeometry = A.ShapeTypeValues.Rectangle;
            HasPresetGeometry = false;

            Fill = new SLFill(listThemeColors);
            Outline = new SLLinePropertiesType(listThemeColors);
            EffectList = new SLEffectList(listThemeColors);

            Rotation3D = new SLRotation3D();
            Format3D = new SLFormat3D(listThemeColors);
        }

        // the logic is exactly the same for C.ChartShapeProperties, C.ShapeProperties, A.ShapeProperties,
        // Xdr.ShapeProperties and other ShapeProperties classes but we're duplicating it because the
        // base class is different

        internal Xdr.ShapeProperties ToXdrShapeProperties()
        {
            var sp = new Xdr.ShapeProperties();

            if (HasBlackWhiteMode) sp.BlackWhiteMode = BlackWhiteMode;

            if (HasTransform2D) sp.Transform2D = Transform2D.ToTransform2D();

            if (HasPresetGeometry)
                sp.Append(new A.PresetGeometry {Preset = PresetGeometry, AdjustValueList = new A.AdjustValueList()});

            if (Fill.HasFill) sp.Append(Fill.ToFill());

            if (Outline.HasLine) sp.Append(Outline.ToOutline());

            if (EffectList.HasEffectList) sp.Append(EffectList.ToEffectList());

            // the bevel top and bottom seems to require camera and lighting.
            // Not sure if that's all the relationship linking, so just leave as it is first...
            if (Rotation3D.HasCamera || Format3D.HasLighting
                || Format3D.HasBevelTop || Format3D.HasBevelBottom)
            {
                var scene3d = new A.Scene3DType();
                if (Rotation3D.HasCamera)
                {
                    scene3d.Camera = new A.Camera();
                    scene3d.Camera.Preset = Rotation3D.CameraPreset;
                    if (Rotation3D.HasPerspectiveSet)
                        scene3d.Camera.FieldOfView = SLDrawingTool.CalculateFovAngle(Rotation3D.Perspective);
                    if (Rotation3D.HasXYZSet)
                    {
                        scene3d.Camera.Rotation = new A.Rotation();
                        scene3d.Camera.Rotation.Latitude = SLDrawingTool.CalculatePositiveFixedAngle(Rotation3D.Y);
                        scene3d.Camera.Rotation.Longitude = SLDrawingTool.CalculatePositiveFixedAngle(Rotation3D.X);
                        scene3d.Camera.Rotation.Revolution = SLDrawingTool.CalculatePositiveFixedAngle(Rotation3D.Z);
                    }
                }
                else
                {
                    scene3d.Camera = new A.Camera {Preset = A.PresetCameraValues.OrthographicFront};
                }

                if (Format3D.HasLighting)
                {
                    scene3d.LightRig = new A.LightRig();
                    scene3d.LightRig.Rig = Format3D.Lighting;
                    scene3d.LightRig.Direction = A.LightRigDirectionValues.Top;
                    if (Format3D.Angle != 0)
                        scene3d.LightRig.Rotation = new A.Rotation
                        {
                            Latitude = 0,
                            Longitude = 0,
                            Revolution = SLDrawingTool.CalculatePositiveFixedAngle(Format3D.Angle)
                        };
                }
                else
                {
                    scene3d.LightRig = new A.LightRig();
                    scene3d.LightRig.Rig = A.LightRigValues.ThreePoints;
                    scene3d.LightRig.Direction = A.LightRigDirectionValues.Top;
                }

                sp.Append(scene3d);
            }

            if (Format3D.HasBevelTop || Format3D.HasBevelBottom || Format3D.HasExtrusionColor
                || Format3D.HasContourColor || (Format3D.ExtrusionHeight != 0)
                || (Format3D.ContourWidth != 0) || (Format3D.Material != A.PresetMaterialTypeValues.WarmMatte)
                || (Rotation3D.DistanceZ != 0))
            {
                var shape3d = new A.Shape3DType();

                if (Format3D.HasBevelTop)
                {
                    shape3d.BevelTop = new A.BevelTop();
                    if (Format3D.BevelTopWidth != 6m)
                        shape3d.BevelTop.Width = SLDrawingTool.CalculatePositiveCoordinate(Format3D.BevelTopWidth);
                    if (Format3D.BevelTopHeight != 6m)
                        shape3d.BevelTop.Height = SLDrawingTool.CalculatePositiveCoordinate(Format3D.BevelTopHeight);
                    if (Format3D.BevelTopPreset != A.BevelPresetValues.Circle)
                        shape3d.BevelTop.Preset = Format3D.BevelTopPreset;
                }

                if (Format3D.HasBevelBottom)
                {
                    shape3d.BevelBottom = new A.BevelBottom();
                    if (Format3D.BevelBottomWidth != 6m)
                        shape3d.BevelBottom.Width = SLDrawingTool.CalculatePositiveCoordinate(Format3D.BevelBottomWidth);
                    if (Format3D.BevelBottomHeight != 6m)
                        shape3d.BevelBottom.Height =
                            SLDrawingTool.CalculatePositiveCoordinate(Format3D.BevelBottomHeight);
                    if (Format3D.BevelBottomPreset != A.BevelPresetValues.Circle)
                        shape3d.BevelBottom.Preset = Format3D.BevelBottomPreset;
                }

                if (Format3D.HasExtrusionColor)
                {
                    shape3d.ExtrusionColor = new A.ExtrusionColor();
                    if (Format3D.clrExtrusionColor.IsRgbColorModelHex)
                        shape3d.ExtrusionColor.RgbColorModelHex = Format3D.clrExtrusionColor.ToRgbColorModelHex();
                    else
                        shape3d.ExtrusionColor.SchemeColor = Format3D.clrExtrusionColor.ToSchemeColor();
                }

                if (Format3D.HasContourColor)
                {
                    shape3d.ContourColor = new A.ContourColor();
                    if (Format3D.clrContourColor.IsRgbColorModelHex)
                        shape3d.ContourColor.RgbColorModelHex = Format3D.clrContourColor.ToRgbColorModelHex();
                    else
                        shape3d.ContourColor.SchemeColor = Format3D.clrContourColor.ToSchemeColor();
                }

                if (Rotation3D.DistanceZ != 0m)
                    shape3d.Z = SLDrawingTool.CalculateCoordinate(Rotation3D.DistanceZ);

                if (Format3D.ExtrusionHeight != 0m)
                    shape3d.ExtrusionHeight = SLDrawingTool.CalculatePositiveCoordinate(Format3D.ExtrusionHeight);

                if (Format3D.ContourWidth != 0m)
                    shape3d.ContourWidth = SLDrawingTool.CalculatePositiveCoordinate(Format3D.ContourWidth);

                if (Format3D.Material != A.PresetMaterialTypeValues.WarmMatte)
                    shape3d.PresetMaterial = Format3D.Material;

                sp.Append(shape3d);
            }

            return sp;
        }

        internal C.ChartShapeProperties ToChartShapeProperties(bool IsStylish = false)
        {
            var sp = new C.ChartShapeProperties();

            if (HasBlackWhiteMode) sp.BlackWhiteMode = BlackWhiteMode;

            if (HasTransform2D) sp.Transform2D = Transform2D.ToTransform2D();

            if (HasPresetGeometry)
                sp.Append(new A.PresetGeometry {Preset = PresetGeometry, AdjustValueList = new A.AdjustValueList()});

            if (Fill.HasFill) sp.Append(Fill.ToFill());

            if (Outline.HasLine) sp.Append(Outline.ToOutline());

            if (IsStylish || EffectList.HasEffectList) sp.Append(EffectList.ToEffectList());

            // the bevel top and bottom seems to require camera and lighting.
            // Not sure if that's all the relationship linking, so just leave as it is first...
            if (Rotation3D.HasCamera || Format3D.HasLighting
                || Format3D.HasBevelTop || Format3D.HasBevelBottom)
            {
                var scene3d = new A.Scene3DType();
                if (Rotation3D.HasCamera)
                {
                    scene3d.Camera = new A.Camera();
                    scene3d.Camera.Preset = Rotation3D.CameraPreset;
                    if (Rotation3D.HasPerspectiveSet)
                        scene3d.Camera.FieldOfView = SLDrawingTool.CalculateFovAngle(Rotation3D.Perspective);
                    if (Rotation3D.HasXYZSet)
                    {
                        scene3d.Camera.Rotation = new A.Rotation();
                        scene3d.Camera.Rotation.Latitude = SLDrawingTool.CalculatePositiveFixedAngle(Rotation3D.Y);
                        scene3d.Camera.Rotation.Longitude = SLDrawingTool.CalculatePositiveFixedAngle(Rotation3D.X);
                        scene3d.Camera.Rotation.Revolution = SLDrawingTool.CalculatePositiveFixedAngle(Rotation3D.Z);
                    }
                }
                else
                {
                    scene3d.Camera = new A.Camera {Preset = A.PresetCameraValues.OrthographicFront};
                }

                if (Format3D.HasLighting)
                {
                    scene3d.LightRig = new A.LightRig();
                    scene3d.LightRig.Rig = Format3D.Lighting;
                    scene3d.LightRig.Direction = A.LightRigDirectionValues.Top;
                    if (Format3D.Angle != 0)
                        scene3d.LightRig.Rotation = new A.Rotation
                        {
                            Latitude = 0,
                            Longitude = 0,
                            Revolution = SLDrawingTool.CalculatePositiveFixedAngle(Format3D.Angle)
                        };
                }
                else
                {
                    scene3d.LightRig = new A.LightRig();
                    scene3d.LightRig.Rig = A.LightRigValues.ThreePoints;
                    scene3d.LightRig.Direction = A.LightRigDirectionValues.Top;
                }

                sp.Append(scene3d);
            }

            if (Format3D.HasBevelTop || Format3D.HasBevelBottom || Format3D.HasExtrusionColor
                || Format3D.HasContourColor || (Format3D.ExtrusionHeight != 0)
                || (Format3D.ContourWidth != 0) || (Format3D.Material != A.PresetMaterialTypeValues.WarmMatte)
                || (Rotation3D.DistanceZ != 0))
            {
                var shape3d = new A.Shape3DType();

                if (Format3D.HasBevelTop)
                {
                    shape3d.BevelTop = new A.BevelTop();
                    if (Format3D.BevelTopWidth != 6m)
                        shape3d.BevelTop.Width = SLDrawingTool.CalculatePositiveCoordinate(Format3D.BevelTopWidth);
                    if (Format3D.BevelTopHeight != 6m)
                        shape3d.BevelTop.Height = SLDrawingTool.CalculatePositiveCoordinate(Format3D.BevelTopHeight);
                    if (Format3D.BevelTopPreset != A.BevelPresetValues.Circle)
                        shape3d.BevelTop.Preset = Format3D.BevelTopPreset;
                }

                if (Format3D.HasBevelBottom)
                {
                    shape3d.BevelBottom = new A.BevelBottom();
                    if (Format3D.BevelBottomWidth != 6m)
                        shape3d.BevelBottom.Width = SLDrawingTool.CalculatePositiveCoordinate(Format3D.BevelBottomWidth);
                    if (Format3D.BevelBottomHeight != 6m)
                        shape3d.BevelBottom.Height =
                            SLDrawingTool.CalculatePositiveCoordinate(Format3D.BevelBottomHeight);
                    if (Format3D.BevelBottomPreset != A.BevelPresetValues.Circle)
                        shape3d.BevelBottom.Preset = Format3D.BevelBottomPreset;
                }

                if (Format3D.HasExtrusionColor)
                {
                    shape3d.ExtrusionColor = new A.ExtrusionColor();
                    if (Format3D.clrExtrusionColor.IsRgbColorModelHex)
                        shape3d.ExtrusionColor.RgbColorModelHex = Format3D.clrExtrusionColor.ToRgbColorModelHex();
                    else
                        shape3d.ExtrusionColor.SchemeColor = Format3D.clrExtrusionColor.ToSchemeColor();
                }

                if (Format3D.HasContourColor)
                {
                    shape3d.ContourColor = new A.ContourColor();
                    if (Format3D.clrContourColor.IsRgbColorModelHex)
                        shape3d.ContourColor.RgbColorModelHex = Format3D.clrContourColor.ToRgbColorModelHex();
                    else
                        shape3d.ContourColor.SchemeColor = Format3D.clrContourColor.ToSchemeColor();
                }

                if (Rotation3D.DistanceZ != 0m)
                    shape3d.Z = SLDrawingTool.CalculateCoordinate(Rotation3D.DistanceZ);

                if (Format3D.ExtrusionHeight != 0m)
                    shape3d.ExtrusionHeight = SLDrawingTool.CalculatePositiveCoordinate(Format3D.ExtrusionHeight);

                if (Format3D.ContourWidth != 0m)
                    shape3d.ContourWidth = SLDrawingTool.CalculatePositiveCoordinate(Format3D.ContourWidth);

                if (Format3D.Material != A.PresetMaterialTypeValues.WarmMatte)
                    shape3d.PresetMaterial = Format3D.Material;

                sp.Append(shape3d);
            }

            return sp;
        }

        /// <summary>
        ///     This is for C.ShapeProperties
        /// </summary>
        /// <returns></returns>
        internal C.ShapeProperties ToCShapeProperties(bool IsStylish = false)
        {
            var sp = new C.ShapeProperties();

            if (HasBlackWhiteMode) sp.BlackWhiteMode = BlackWhiteMode;

            if (HasTransform2D) sp.Transform2D = Transform2D.ToTransform2D();

            if (HasPresetGeometry)
                sp.Append(new A.PresetGeometry {Preset = PresetGeometry, AdjustValueList = new A.AdjustValueList()});

            if (Fill.HasFill) sp.Append(Fill.ToFill());

            if (Outline.HasLine) sp.Append(Outline.ToOutline());

            if (IsStylish || EffectList.HasEffectList) sp.Append(EffectList.ToEffectList());

            // the bevel top and bottom seems to require camera and lighting.
            // Not sure if that's all the relationship linking, so just leave as it is first...
            if (Rotation3D.HasCamera || Format3D.HasLighting
                || Format3D.HasBevelTop || Format3D.HasBevelBottom)
            {
                var scene3d = new A.Scene3DType();
                if (Rotation3D.HasCamera)
                {
                    scene3d.Camera = new A.Camera();
                    scene3d.Camera.Preset = Rotation3D.CameraPreset;
                    if (Rotation3D.HasPerspectiveSet)
                        scene3d.Camera.FieldOfView = SLDrawingTool.CalculateFovAngle(Rotation3D.Perspective);
                    if (Rotation3D.HasXYZSet)
                    {
                        scene3d.Camera.Rotation = new A.Rotation();
                        scene3d.Camera.Rotation.Latitude = SLDrawingTool.CalculatePositiveFixedAngle(Rotation3D.Y);
                        scene3d.Camera.Rotation.Longitude = SLDrawingTool.CalculatePositiveFixedAngle(Rotation3D.X);
                        scene3d.Camera.Rotation.Revolution = SLDrawingTool.CalculatePositiveFixedAngle(Rotation3D.Z);
                    }
                }
                else
                {
                    scene3d.Camera = new A.Camera {Preset = A.PresetCameraValues.OrthographicFront};
                }

                if (Format3D.HasLighting)
                {
                    scene3d.LightRig = new A.LightRig();
                    scene3d.LightRig.Rig = Format3D.Lighting;
                    scene3d.LightRig.Direction = A.LightRigDirectionValues.Top;
                    if (Format3D.Angle != 0)
                        scene3d.LightRig.Rotation = new A.Rotation
                        {
                            Latitude = 0,
                            Longitude = 0,
                            Revolution = SLDrawingTool.CalculatePositiveFixedAngle(Format3D.Angle)
                        };
                }
                else
                {
                    scene3d.LightRig = new A.LightRig();
                    scene3d.LightRig.Rig = A.LightRigValues.ThreePoints;
                    scene3d.LightRig.Direction = A.LightRigDirectionValues.Top;
                }

                sp.Append(scene3d);
            }

            if (Format3D.HasBevelTop || Format3D.HasBevelBottom || Format3D.HasExtrusionColor
                || Format3D.HasContourColor || (Format3D.ExtrusionHeight != 0)
                || (Format3D.ContourWidth != 0) || (Format3D.Material != A.PresetMaterialTypeValues.WarmMatte)
                || (Rotation3D.DistanceZ != 0))
            {
                var shape3d = new A.Shape3DType();

                if (Format3D.HasBevelTop)
                {
                    shape3d.BevelTop = new A.BevelTop();
                    if (Format3D.BevelTopWidth != 6m)
                        shape3d.BevelTop.Width = SLDrawingTool.CalculatePositiveCoordinate(Format3D.BevelTopWidth);
                    if (Format3D.BevelTopHeight != 6m)
                        shape3d.BevelTop.Height = SLDrawingTool.CalculatePositiveCoordinate(Format3D.BevelTopHeight);
                    if (Format3D.BevelTopPreset != A.BevelPresetValues.Circle)
                        shape3d.BevelTop.Preset = Format3D.BevelTopPreset;
                }

                if (Format3D.HasBevelBottom)
                {
                    shape3d.BevelBottom = new A.BevelBottom();
                    if (Format3D.BevelBottomWidth != 6m)
                        shape3d.BevelBottom.Width = SLDrawingTool.CalculatePositiveCoordinate(Format3D.BevelBottomWidth);
                    if (Format3D.BevelBottomHeight != 6m)
                        shape3d.BevelBottom.Height =
                            SLDrawingTool.CalculatePositiveCoordinate(Format3D.BevelBottomHeight);
                    if (Format3D.BevelBottomPreset != A.BevelPresetValues.Circle)
                        shape3d.BevelBottom.Preset = Format3D.BevelBottomPreset;
                }

                if (Format3D.HasExtrusionColor)
                {
                    shape3d.ExtrusionColor = new A.ExtrusionColor();
                    if (Format3D.clrExtrusionColor.IsRgbColorModelHex)
                        shape3d.ExtrusionColor.RgbColorModelHex = Format3D.clrExtrusionColor.ToRgbColorModelHex();
                    else
                        shape3d.ExtrusionColor.SchemeColor = Format3D.clrExtrusionColor.ToSchemeColor();
                }

                if (Format3D.HasContourColor)
                {
                    shape3d.ContourColor = new A.ContourColor();
                    if (Format3D.clrContourColor.IsRgbColorModelHex)
                        shape3d.ContourColor.RgbColorModelHex = Format3D.clrContourColor.ToRgbColorModelHex();
                    else
                        shape3d.ContourColor.SchemeColor = Format3D.clrContourColor.ToSchemeColor();
                }

                if (Rotation3D.DistanceZ != 0m)
                    shape3d.Z = SLDrawingTool.CalculateCoordinate(Rotation3D.DistanceZ);

                if (Format3D.ExtrusionHeight != 0m)
                    shape3d.ExtrusionHeight = SLDrawingTool.CalculatePositiveCoordinate(Format3D.ExtrusionHeight);

                if (Format3D.ContourWidth != 0m)
                    shape3d.ContourWidth = SLDrawingTool.CalculatePositiveCoordinate(Format3D.ContourWidth);

                if (Format3D.Material != A.PresetMaterialTypeValues.WarmMatte)
                    shape3d.PresetMaterial = Format3D.Material;

                sp.Append(shape3d);
            }

            return sp;
        }

        internal SLShapeProperties Clone()
        {
            var sp = new SLShapeProperties(listThemeColors);
            sp.HasBlackWhiteMode = HasBlackWhiteMode;
            sp.vBlackWhiteMode = vBlackWhiteMode;
            sp.HasTransform2D = HasTransform2D;
            sp.Transform2D = Transform2D.Clone();
            sp.HasPresetGeometry = HasPresetGeometry;
            sp.vPresetGeometry = vPresetGeometry;
            sp.Fill = Fill.Clone();
            sp.Outline = Outline.Clone();
            sp.EffectList = EffectList.Clone();
            sp.Rotation3D = Rotation3D.Clone();
            sp.Format3D = Format3D.Clone();

            return sp;
        }
    }
}