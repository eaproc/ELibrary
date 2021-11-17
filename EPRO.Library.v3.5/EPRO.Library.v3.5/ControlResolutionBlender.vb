Imports System.Drawing
Imports System.Windows.Forms

Imports EPRO.Library.v3._5.Modules.basMain

'
'
'   Changing the fonts of a user control affects the controls embedded in it if
'   and on IFF the embedded control doesnt have its own font set
'
'   This has been set not to resize user control fonts at all and forms
'
'
'

Public Class ControlResolutionBlender


#Region "Constructors"


    Shared Sub New()
        ComputerScreenBound = My.Computer.Screen.Bounds
        ComputerWorkingArea = My.Computer.Screen.WorkingArea.Size
    End Sub






    Private Sub New(
            Optional pDoNotResizeFont() As Control = Nothing,
            Optional pDoNotResizeAtAll() As Control = Nothing,
            Optional pTerminateRecursiveAtLevel As Byte = 20,
            Optional pdoNotResizeInnerContentsOfthisControl() As Control = Nothing
                                                          )

        Me.doNotResizeControlAtAll = New List(Of Control)()
        Me.doNotResizeFontList = New List(Of Control)()
        Me.doNotResizeInnerContentsOfthisControl = New List(Of Control)()

        If pDoNotResizeAtAll IsNot Nothing Then Me.doNotResizeControlAtAll.AddRange(pDoNotResizeAtAll)
        If pDoNotResizeFont IsNot Nothing Then Me.doNotResizeFontList.AddRange(pDoNotResizeFont)
        If pdoNotResizeInnerContentsOfthisControl IsNot Nothing Then Me.doNotResizeInnerContentsOfthisControl.AddRange(pdoNotResizeInnerContentsOfthisControl)

        Me.TerminateRecursiveAtLevel = pTerminateRecursiveAtLevel

      
    End Sub

    ''' <summary>
    ''' This is for fixed size windows
    ''' </summary>
    ''' <param name="pParentControl"></param>
    ''' <param name="TargetResolution"></param>
    ''' <param name="adjustMainControl"></param>
    ''' <remarks></remarks>
    Sub New(ByRef pParentControl As Control,
            ByVal TargetResolution As Point,
            Optional adjustMainControl As Boolean = True,
            Optional isDebugMode As Boolean = False,
            Optional pDoNotResizeFont() As Control = Nothing,
            Optional pDoNotResizeAtAll() As Control = Nothing,
            Optional pTerminateRecursiveAtLevel As Byte = 20,
            Optional pdoNotResizeInnerContentsOfthisControl() As Control = Nothing)

        Me.New(pDoNotResizeFont, pDoNotResizeAtAll, pTerminateRecursiveAtLevel, pdoNotResizeInnerContentsOfthisControl:=pdoNotResizeInnerContentsOfthisControl)
        IsDebugMode = isDebugMode


        Dim currentResolution As Rectangle = ComputerScreenBound

        ''Dim hzUnitMeasurement As Double = currentResolution.Width / TargetResolution.X
        ''Dim vtUnitMeasurement As Double = currentResolution.Height / TargetResolution.Y
        Dim lMeasurement As Drawing.PointF = Me.getRelativeDifference(
                            New Point(currentResolution.Width, currentResolution.Height),
                                                             TargetResolution)



        If IsDebugMode Then MyLogFile.Print("Resolution Difference: " & lMeasurement.ToString())

        ''Debug.Print("hzUnitMeasurement: " & hzUnitMeasurement)
        ''Debug.Print("vtUnitMeasurement: " & vtUnitMeasurement)
        ''Debug.Print("hzUnitMeasurement: " & lMeasurement.X)
        ''Debug.Print("vtUnitMeasurement: " & lMeasurement.Y)

        Dim oldSize As Size = pParentControl.Size

        If IsDebugMode Then MyLogFile.Print("Form OldSize: " & oldSize.ToString())

        If adjustMainControl Then

            ''With pParentControl
            ''    If .MaximumSize.Width <> 0 Then
            ''        .MaximumSize = New Size(CInt(.Width * hzUnitMeasurement),
            ''                                              CInt(.Height * vtUnitMeasurement)
            ''                                              )
            ''    End If
            ''    If .MinimumSize.Width <> 0 Then
            ''        .MinimumSize = New Size(CInt(.Width * hzUnitMeasurement),
            ''                               CInt(.Height * vtUnitMeasurement)
            ''                               )
            ''    End If

            ''    .Width = CInt(.Width * hzUnitMeasurement)
            ''    .Height = CInt(.Height * vtUnitMeasurement)

            ''End With
            With pParentControl
                If .MaximumSize.Width <> 0 Then
                    .MaximumSize = New Size(CInt(.Width * lMeasurement.X),
                                                          CInt(.Height * lMeasurement.Y)
                                                          )
                End If
                If .MinimumSize.Width <> 0 Then
                    .MinimumSize = New Size(CInt(.Width * lMeasurement.X),
                                           CInt(.Height * lMeasurement.Y)
                                           )
                End If

                .Width = CInt(.Width * lMeasurement.X)
                .Height = CInt(.Height * lMeasurement.Y)



                If IsDebugMode Then MyLogFile.Print("Form New Size: " & .Size.ToString())
            End With

        End If



        Me.ResizeMyControls(pParentControl,
                            lMeasurement.X, lMeasurement.Y,
                            oldSize,
                            True, 1)

    End Sub


    ''' <summary>
    ''' This is for Maximize all to full screen.
    ''' </summary>
    ''' <param name="pParentControl"></param>
    ''' <param name="restrictToWorkingArea">This only works if you do not set your windowstate to maximize. it does not cover the taskbar</param>
    ''' <param name="pLockFormSize">Prevents user from setting the form size to lower than it is after it has been loaded</param>
    ''' <remarks></remarks>
    Sub New(ByRef pParentControl As Form,
            Optional isDebugMode As Boolean = False,
            Optional restrictToWorkingArea As Boolean = True,
            Optional pLockFormSize As Boolean = True,
            Optional pDoNotResizeFont() As Control = Nothing,
            Optional pDoNotResizeAtAll() As Control = Nothing,
             Optional pTerminateRecursiveAtLevel As Byte = 20,
            Optional pdoNotResizeInnerContentsOfthisControl() As Control = Nothing
            )
        Me.New(pDoNotResizeFont, pDoNotResizeAtAll, pTerminateRecursiveAtLevel, pdoNotResizeInnerContentsOfthisControl:=pdoNotResizeInnerContentsOfthisControl)
        IsDebugMode = isDebugMode


        Dim oldSize As Size = pParentControl.Size
        If IsDebugMode Then MyLogFile.Print("Form OldSize: " & oldSize.ToString())

        pParentControl.WindowState = FormWindowState.Normal
        If restrictToWorkingArea Then
            pParentControl.Size = ComputerWorkingArea
            pParentControl.Location = New Point(0, 0)
        Else
            pParentControl.Size = New Size(ComputerScreenBound.Width, ComputerScreenBound.Height)
            pParentControl.WindowState = FormWindowState.Maximized
            REM Else the form will go behind taskbar
        End If

        If pLockFormSize Then Objects.EForm.LockControlSize(pParentControl)

        If IsDebugMode Then MyLogFile.Print("Form New Size: " & pParentControl.Size.ToString())

        Me.ResizeMyControls(CType(pParentControl, Control),
                            0, 0,
                            oldSize,
                            True, 1)

    End Sub



    ''' <summary>
    ''' This is to fit all child control to parent current size
    ''' </summary>
    ''' <param name="pParentControl">with currrent size</param>
    ''' <param name="TargetResolution">Old Size</param>
    ''' <remarks></remarks>
    Sub New(ByVal TargetResolution As Point,
            ByRef pParentControl As Control,
            Optional pDoNotResizeFont() As Control = Nothing,
            Optional pDoNotResizeAtAll() As Control = Nothing,
             Optional pTerminateRecursiveAtLevel As Byte = 20,
             Optional isDebugMode As Boolean = False,
            Optional pdoNotResizeInnerContentsOfthisControl() As Control = Nothing
            )
        Me.New(pDoNotResizeFont, pDoNotResizeAtAll, pTerminateRecursiveAtLevel, pdoNotResizeInnerContentsOfthisControl:=pdoNotResizeInnerContentsOfthisControl)
        IsDebugMode = isDebugMode


        Dim oldSize As Size = New Size(TargetResolution)
        If IsDebugMode Then MyLogFile.Print("Form OldSize: " & oldSize.ToString())

        If IsDebugMode Then MyLogFile.Print("Form New Size: " & pParentControl.Size.ToString())

        Me.ResizeMyControls(pParentControl,
                            0, 0,
                            oldSize,
                            True, 1)

    End Sub



#End Region




#Region "Enum and Const"""
    Public Const WINDOWS__TASKBAR__EXPECTED__HEIGHT__PX As Int32 = 40
#End Region



#Region "Properties"

    Public Shared Property IsDebugMode As Boolean

    Public Shared Property ComputerScreenBound As Rectangle
    Public Shared Property ComputerWorkingArea As Size



    Dim TerminateRecursiveAtLevel As Byte
    Private doNotResizeFontList As List(Of Control)
    Private doNotResizeControlAtAll As List(Of Control)
    Private doNotResizeInnerContentsOfthisControl As List(Of Control)

#End Region






    Private Function getRelativeDifference(ByVal newSize As Point, oldSize As Point) As PointF
        Return New PointF(CSng((newSize.X / oldSize.X)),
                                CSng(newSize.Y / oldSize.Y)
                                    )
    End Function


    Private Function getNewFont(ByVal oOriginalFont As Font, wWidthMeasure As Single) As Font

        Dim fSize As Single = oOriginalFont.Size * wWidthMeasure
        Return New Font(oOriginalFont.FontFamily, fSize, oOriginalFont.Style)
    End Function


    Private Sub ResizeMyControls(ByRef pParentControl As Control, hzUnitMeasurement As Single, vtUnitMeasurement As Single,
                                 OldSizeParent As Size,
                                 ByVal isRoot As Boolean,
                                 ByVal recursiveLevel As Int32
                                 )
        Try


            If IsDebugMode Then MyLogFile.Print("Recursive Level: " & recursiveLevel)
            If recursiveLevel = Me.TerminateRecursiveAtLevel Then Return

            If Me.doNotResizeControlAtAll.Contains(pParentControl) Then Return

            If Not isRoot Then
                REM remember to check for autosize member

                With pParentControl

                    If IsDebugMode Then MyLogFile.Print(String.Format("Control: {0}, OldSize: {1}, OldLocation: {2}", .Name, .Size, .Location))

                    .Width = CInt(.Width * hzUnitMeasurement)
                    .Height = CInt(.Height * vtUnitMeasurement)

                    If pParentControl.AutoSize Then
                        If IsDebugMode Then Modules.basMain.MyLogFile.Print(
                                        String.Format("Assembly : {0}, {1} has its AutoSize Property Set to TRUE!!!",
                                                        pParentControl.GetType().FullName, pParentControl.Name
                                                        )
                                                    )
                    End If

                    .Location = New Point(CInt(.Location.X * hzUnitMeasurement), CInt(.Location.Y * vtUnitMeasurement))


                    If Not Me.doNotResizeFontList.Contains(pParentControl) AndAlso _
                       Not TypeOf pParentControl Is UserControl AndAlso _
                       Not TypeOf pParentControl Is Form Then

                        ' resizing fonts of forms and controls affects
                        ' their inner controls in a crazy way

                        .Font = Me.getNewFont(.Font, hzUnitMeasurement)

                    End If

                    If Not Me.doNotResizeFontList.Contains(pParentControl) AndAlso _
                       TypeOf pParentControl Is IControlResolutionPropertyExtender Then

                        With CType(pParentControl, IControlResolutionPropertyExtender)
                            If .Xtended__Font1 IsNot Nothing Then .Xtended__Font1 = Me.getNewFont(.Xtended__Font1, hzUnitMeasurement)
                            If .Xtended__Font2 IsNot Nothing Then .Xtended__Font2 = Me.getNewFont(.Xtended__Font2, hzUnitMeasurement)

                        End With

                    End If

                    If IsDebugMode Then MyLogFile.Print(String.Format("Control: {0}, NewSize: {1}, NewLocation: {2}, Using Mesurements {3} x {4}",
                            .Name, .Size, .Location, hzUnitMeasurement, vtUnitMeasurement))

                End With
            End If




            If Me.doNotResizeInnerContentsOfthisControl.Contains(pParentControl) Then Return

            Dim lMeasure As PointF = Me.getRelativeDifference(CType(pParentControl.Size, Point), CType(OldSizeParent, Point))
            For Each c As Control In pParentControl.Controls
                Me.ResizeMyControls(c, lMeasure.X, lMeasure.Y, c.Size, False, recursiveLevel + 1)
            Next

        Catch ex As Exception
            Modules.basMain.MyLogFile.Print(String.Format("Error adjusting ControlName: {0}, Error Message: {1}", pParentControl.Name, ex.Message))
        End Try
    End Sub




End Class


