using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using static ELibrary.Standard.Modules.basMain;

namespace ELibrary.Standard
{

    // 
    // 
    // Changing the fonts of a user control affects the controls embedded in it if
    // and on IFF the embedded control doesnt have its own font set
    // 
    // This has been set not to resize user control fonts at all and forms
    // 
    // 
    // 

    public class ControlResolutionBlender
    {


        #region Constructors


        static ControlResolutionBlender()
        {
            ComputerScreenBound = My.MyProject.Computer.Screen.Bounds;
            ComputerWorkingArea = My.MyProject.Computer.Screen.WorkingArea.Size;
        }

        private ControlResolutionBlender(Control[] pDoNotResizeFont = null, Control[] pDoNotResizeAtAll = null, byte pTerminateRecursiveAtLevel = 20, Control[] pdoNotResizeInnerContentsOfthisControl = null)
        {
            doNotResizeControlAtAll = new List<Control>();
            doNotResizeFontList = new List<Control>();
            doNotResizeInnerContentsOfthisControl = new List<Control>();
            if (pDoNotResizeAtAll is object)
                doNotResizeControlAtAll.AddRange(pDoNotResizeAtAll);
            if (pDoNotResizeFont is object)
                doNotResizeFontList.AddRange(pDoNotResizeFont);
            if (pdoNotResizeInnerContentsOfthisControl is object)
                doNotResizeInnerContentsOfthisControl.AddRange(pdoNotResizeInnerContentsOfthisControl);
            TerminateRecursiveAtLevel = pTerminateRecursiveAtLevel;
        }

        /// <summary>
    /// This is for fixed size windows
    /// </summary>
    /// <param name="pParentControl"></param>
    /// <param name="TargetResolution"></param>
    /// <param name="adjustMainControl"></param>
    /// <remarks></remarks>
        public ControlResolutionBlender(ref Control pParentControl, Point TargetResolution, bool adjustMainControl = true, bool isDebugMode = false, Control[] pDoNotResizeFont = null, Control[] pDoNotResizeAtAll = null, byte pTerminateRecursiveAtLevel = 20, Control[] pdoNotResizeInnerContentsOfthisControl = null) : this(pDoNotResizeFont, pDoNotResizeAtAll, pTerminateRecursiveAtLevel, pdoNotResizeInnerContentsOfthisControl: pdoNotResizeInnerContentsOfthisControl)
        {
            isDebugMode = isDebugMode;
            var currentResolution = ComputerScreenBound;

            // 'Dim hzUnitMeasurement As Double = currentResolution.Width / TargetResolution.X
            // 'Dim vtUnitMeasurement As Double = currentResolution.Height / TargetResolution.Y
            var lMeasurement = getRelativeDifference(new Point(currentResolution.Width, currentResolution.Height), TargetResolution);
            if (isDebugMode)
                MyLogFile.Print("Resolution Difference: " + lMeasurement.ToString());

            // 'Debug.Print("hzUnitMeasurement: " & hzUnitMeasurement)
            // 'Debug.Print("vtUnitMeasurement: " & vtUnitMeasurement)
            // 'Debug.Print("hzUnitMeasurement: " & lMeasurement.X)
            // 'Debug.Print("vtUnitMeasurement: " & lMeasurement.Y)

            var oldSize = pParentControl.Size;
            if (isDebugMode)
                MyLogFile.Print("Form OldSize: " + oldSize.ToString());
            if (adjustMainControl)
            {

                // 'With pParentControl
                // '    If .MaximumSize.Width <> 0 Then
                // '        .MaximumSize = New Size(CInt(.Width * hzUnitMeasurement),
                // '                                              CInt(.Height * vtUnitMeasurement)
                // '                                              )
                // '    End If
                // '    If .MinimumSize.Width <> 0 Then
                // '        .MinimumSize = New Size(CInt(.Width * hzUnitMeasurement),
                // '                               CInt(.Height * vtUnitMeasurement)
                // '                               )
                // '    End If

                // '    .Width = CInt(.Width * hzUnitMeasurement)
                // '    .Height = CInt(.Height * vtUnitMeasurement)

                // 'End With
                {
                    var withBlock = pParentControl;
                    if (withBlock.MaximumSize.Width != 0)
                    {
                        withBlock.MaximumSize = new Size((int)Math.Round(withBlock.Width * lMeasurement.X), (int)Math.Round(withBlock.Height * lMeasurement.Y));
                    }

                    if (withBlock.MinimumSize.Width != 0)
                    {
                        withBlock.MinimumSize = new Size((int)Math.Round(withBlock.Width * lMeasurement.X), (int)Math.Round(withBlock.Height * lMeasurement.Y));
                    }

                    withBlock.Width = (int)Math.Round(withBlock.Width * lMeasurement.X);
                    withBlock.Height = (int)Math.Round(withBlock.Height * lMeasurement.Y);
                    if (isDebugMode)
                        MyLogFile.Print("Form New Size: " + withBlock.Size.ToString());
                }
            }

            ResizeMyControls(ref pParentControl, lMeasurement.X, lMeasurement.Y, oldSize, true, 1);
        }


        /// <summary>
    /// This is for Maximize all to full screen.
    /// </summary>
    /// <param name="pParentControl"></param>
    /// <param name="restrictToWorkingArea">This only works if you do not set your windowstate to maximize. it does not cover the taskbar</param>
    /// <param name="pLockFormSize">Prevents user from setting the form size to lower than it is after it has been loaded</param>
    /// <remarks></remarks>
        public ControlResolutionBlender(ref Form pParentControl, bool isDebugMode = false, bool restrictToWorkingArea = true, bool pLockFormSize = true, Control[] pDoNotResizeFont = null, Control[] pDoNotResizeAtAll = null, byte pTerminateRecursiveAtLevel = 20, Control[] pdoNotResizeInnerContentsOfthisControl = null) : this(pDoNotResizeFont, pDoNotResizeAtAll, pTerminateRecursiveAtLevel, pdoNotResizeInnerContentsOfthisControl: pdoNotResizeInnerContentsOfthisControl)
        {
            isDebugMode = isDebugMode;
            var oldSize = pParentControl.Size;
            if (isDebugMode)
                MyLogFile.Print("Form OldSize: " + oldSize.ToString());
            pParentControl.WindowState = FormWindowState.Normal;
            if (restrictToWorkingArea)
            {
                pParentControl.Size = ComputerWorkingArea;
                pParentControl.Location = new Point(0, 0);
            }
            else
            {
                pParentControl.Size = new Size(ComputerScreenBound.Width, ComputerScreenBound.Height);
                pParentControl.WindowState = FormWindowState.Maximized;
                // REM Else the form will go behind taskbar
            }

            if (pLockFormSize)
                Objects.EForm.LockControlSize(ref pParentControl);
            if (isDebugMode)
                MyLogFile.Print("Form New Size: " + pParentControl.Size.ToString());
            Control argpParentControl = pParentControl;
            ResizeMyControls(ref argpParentControl, 0f, 0f, oldSize, true, 1);
        }



        /// <summary>
    /// This is to fit all child control to parent current size
    /// </summary>
    /// <param name="pParentControl">with currrent size</param>
    /// <param name="TargetResolution">Old Size</param>
    /// <remarks></remarks>
        public ControlResolutionBlender(Point TargetResolution, ref Control pParentControl, Control[] pDoNotResizeFont = null, Control[] pDoNotResizeAtAll = null, byte pTerminateRecursiveAtLevel = 20, bool isDebugMode = false, Control[] pdoNotResizeInnerContentsOfthisControl = null) : this(pDoNotResizeFont, pDoNotResizeAtAll, pTerminateRecursiveAtLevel, pdoNotResizeInnerContentsOfthisControl: pdoNotResizeInnerContentsOfthisControl)
        {
            isDebugMode = isDebugMode;
            var oldSize = new Size(TargetResolution);
            if (isDebugMode)
                MyLogFile.Print("Form OldSize: " + oldSize.ToString());
            if (isDebugMode)
                MyLogFile.Print("Form New Size: " + pParentControl.Size.ToString());
            ResizeMyControls(ref pParentControl, 0f, 0f, oldSize, true, 1);
        }



        #endregion




        #region Enum and Const
        public const int WINDOWS__TASKBAR__EXPECTED__HEIGHT__PX = 40;
        #endregion



        #region Properties

        public static bool IsDebugMode { get; set; }
        public static Rectangle ComputerScreenBound { get; set; }
        public static Size ComputerWorkingArea { get; set; }

        private byte TerminateRecursiveAtLevel;
        private List<Control> doNotResizeFontList;
        private List<Control> doNotResizeControlAtAll;
        private List<Control> doNotResizeInnerContentsOfthisControl;

        #endregion






        private PointF getRelativeDifference(Point newSize, Point oldSize)
        {
            return new PointF((float)(newSize.X / (double)oldSize.X), (float)(newSize.Y / (double)oldSize.Y));
        }

        private Font getNewFont(Font oOriginalFont, float wWidthMeasure)
        {
            float fSize = oOriginalFont.Size * wWidthMeasure;
            return new Font(oOriginalFont.FontFamily, fSize, oOriginalFont.Style);
        }

        private void ResizeMyControls(ref Control pParentControl, float hzUnitMeasurement, float vtUnitMeasurement, Size OldSizeParent, bool isRoot, int recursiveLevel)
        {
            try
            {
                if (IsDebugMode)
                    MyLogFile.Print("Recursive Level: " + recursiveLevel);
                if (recursiveLevel == TerminateRecursiveAtLevel)
                    return;
                if (doNotResizeControlAtAll.Contains(pParentControl))
                    return;
                if (!isRoot)
                {
                    // REM remember to check for autosize member

                    if (IsDebugMode)
                        MyLogFile.Print(string.Format("Control: {0}, OldSize: {1}, OldLocation: {2}", pParentControl.Name, pParentControl.Size, pParentControl.Location));
                    pParentControl.Width = (int)Math.Round(pParentControl.Width * hzUnitMeasurement);
                    pParentControl.Height = (int)Math.Round(pParentControl.Height * vtUnitMeasurement);
                    if (pParentControl.AutoSize)
                    {
                        if (IsDebugMode)
                            MyLogFile.Print(string.Format("Assembly : {0}, {1} has its AutoSize Property Set to TRUE!!!", pParentControl.GetType().FullName, pParentControl.Name));
                    }

                    pParentControl.Location = new Point((int)Math.Round(pParentControl.Location.X * hzUnitMeasurement), (int)Math.Round(pParentControl.Location.Y * vtUnitMeasurement));
                    if (!doNotResizeFontList.Contains(pParentControl) && !(pParentControl is UserControl) && !(pParentControl is Form))

                    {

                        // resizing fonts of forms and controls affects
                        // their inner controls in a crazy way

                        pParentControl.Font = getNewFont(pParentControl.Font, hzUnitMeasurement);
                    }

                    if (!doNotResizeFontList.Contains(pParentControl) && pParentControl is IControlResolutionPropertyExtender)
                    {
                        {
                            var withBlock = (IControlResolutionPropertyExtender)pParentControl;
                            if (withBlock.Xtended__Font1 is object)
                                withBlock.Xtended__Font1 = getNewFont(withBlock.Xtended__Font1, hzUnitMeasurement);
                            if (withBlock.Xtended__Font2 is object)
                                withBlock.Xtended__Font2 = getNewFont(withBlock.Xtended__Font2, hzUnitMeasurement);
                        }
                    }

                    if (IsDebugMode)
                        MyLogFile.Print(string.Format("Control: {0}, NewSize: {1}, NewLocation: {2}, Using Mesurements {3} x {4}", pParentControl.Name, pParentControl.Size, pParentControl.Location, hzUnitMeasurement, vtUnitMeasurement));
                }

                if (doNotResizeInnerContentsOfthisControl.Contains(pParentControl))
                    return;
                var lMeasure = getRelativeDifference((Point)pParentControl.Size, (Point)OldSizeParent);
                foreach (Control c in pParentControl.Controls)
                    ResizeMyControls(ref c, lMeasure.X, lMeasure.Y, c.Size, false, recursiveLevel + 1);
            }
            catch (Exception ex)
            {
                MyLogFile.Print(string.Format("Error adjusting ControlName: {0}, Error Message: {1}", pParentControl.Name, ex.Message));
            }
        }
    }
}