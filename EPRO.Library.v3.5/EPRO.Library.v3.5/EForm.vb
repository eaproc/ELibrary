Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports EPRO.Library.v3._5.Modules
Imports CODERiT.Logger.v._3._5.Exceptions

Namespace Objects

    Public Class EForm


#Region "New Form Display"

        ''' <summary>
        ''' Display a Modal Child Form and return the dialog result
        ''' </summary>
        ''' <param name="Parent"></param>
        ''' <param name="Child"></param>
        ''' <param name="FadeParentOpaq"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function ShowChildFormDialog(ByRef Parent As Form, ByVal Child As Form,
                                            Optional ByVal FadeParentOpaq As Double = 0.7
                                                                                      ) As DialogResult

            If Not isValid_Form(Parent) Then Return DialogResult.No
            Dim PrevOpaq As Double = Parent.Opacity
            Try


                Dim dlgRst As DialogResult
                Parent.Opacity = FadeParentOpaq
                With Child
                    .ShowInTaskbar = False
                    dlgRst = .ShowDialog(Parent)
                End With
                Parent.Opacity = PrevOpaq

                Return dlgRst

            Catch ex As Exception

                If isValid_Form(Parent) Then Parent.Opacity = PrevOpaq
                MyLogFile.Log(New EException(ex))

                Return DialogResult.No

            End Try

        End Function


        ''' <summary>
        ''' Display a Modal Child Form  on any control handle  it doesnt fade .. .and return the dialog result. User Control Doesn't support Opacity in .NET 3.5
        ''' </summary>
        ''' <param name="Parent"></param>
        ''' <param name="Child"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function ShowChildFormDialog(Of T As {UserControl}, T2 As Form)(ByRef Parent As T, ByVal Child As T2
                                                                                      ) As DialogResult

            Try


                Dim dlgRst As DialogResult
                With Child
                    .ShowInTaskbar = False
                    dlgRst = .ShowDialog(Parent)
                End With


                Return dlgRst

            Catch ex As Exception

                MyLogFile.Log(New EException(ex))

                Return DialogResult.No

            End Try

        End Function

        ''' <summary>
        ''' Display a Modal Child Form  on any control handle  it doesnt fade .. .and return the dialog result. User Control Doesn't support Opacity in .NET 3.5
        ''' </summary>
        ''' <param name="Parent"></param>
        ''' <param name="Child"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function ShowChildFormNonModal(Of T As {UserControl}, T2 As Form)(ByRef Parent As T, ByVal Child As T2,
                                                                                        Optional ByVal CloseParentForm As Boolean = True
                                                                                      ) As Boolean

            Try



                With Child
                    .Show()
                    
                End With

                If Parent IsNot Nothing AndAlso CloseParentForm AndAlso Not Parent.IsDisposed Then Parent.Dispose()
                Return True
            Catch ex As Exception

                MyLogFile.Log(New EException(ex))
                Return False
            End Try

        End Function

        ''' <summary>
        ''' Opens a non modal child form
        ''' </summary>
        ''' <param name="ParentForm"></param>
        ''' <param name="child"></param>
        ''' <param name="CloseParentForm"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function showChildFormNonModal(ByRef ParentForm As Form, ByVal child As Form,
                                                            Optional ByVal CloseParentForm As Boolean = True) As Boolean

            Try

                'Debug.Print("CloseParentForm " & CloseParentForm)
                'Debug.Print("ParentForm IsNot Nothing" & (ParentForm IsNot Nothing))

                child.Show()
                If CloseParentForm AndAlso isValid_Form(ParentForm) Then
                    'Debug.Print("Closing Parent Form")
                    ParentForm.Close()

                End If


                Return True

            Catch ex As Exception
                MyLogFile.Log(New EException(ex))

            End Try

            Return False

        End Function

        ''' <summary>
        ''' Opens a non modal child form and does not close the parent form
        ''' </summary>
        ''' <param name="child"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function showChildFormNonModal(ByVal child As Form) As Boolean

            Try

                child.Show()

                Return True

            Catch ex As Exception

            End Try

            Return False

        End Function

        '' ''' <summary>
        '' ''' Opens a non modal child form
        '' ''' </summary>
        '' ''' <param name="ParentForm"></param>
        '' ''' <param name="child"></param>
        '' ''' <param name="CloseParentForm"></param>
        '' ''' <remarks>This is running on same thread and the performance is poor</remarks>
        ''Public Shared Sub showChildFormNonModal(ByRef ParentForm As Form, ByVal child As Form,
        ''                                                    ByVal WaitTillOpenDescription As String,
        ''                                                    Optional ByVal CloseParentForm As Boolean = True)

        ''    Dim WaitForm As New prjWaitAsyncManager.clsWaitAsyncMgr(, WaitTillOpenDescription)
        ''    Application.DoEvents()
        ''    showChildFormNonModal(ParentForm, child, CloseParentForm)
        ''    WaitForm.Dispose()
        ''    WaitForm = Nothing

        ''End Sub


#End Region


#Region "Locating (Setting Location and Sizes) Control"



        Public Shared Sub centerControl(ByRef ControlToCenter As Control,
                                             ByVal InReferenceTo As Control,
                                             Optional ByVal Horizontally As Boolean = True)

            With ControlToCenter
                If Horizontally Then
                    .Left = CInt((InReferenceTo.Width - .Width) / 2)
                Else
                    .Top = CInt((InReferenceTo.Height - .Height) / 2)
                End If

            End With

        End Sub


        Public Shared Sub LockControlSize(Of T As {Control})(ByRef pControl As T)

            pControl.MinimumSize = pControl.Size
            pControl.MaximumSize = pControl.Size


        End Sub



        Public Enum FormLocationOnScreen
            TOP__LEFT__CORNER
            TOP__MIDDLE
            TOP__RIGHT__CORNER
            MIDDLE__LEFT__CORNER
            CENTER
            MIDDLE__RIGHT__CORNER
            BOTTOM__LEFT__CORNER
            BOTTOM__MIDDLE
            BOTTOM__RIGHT__CORNER
        End Enum


        Public Shared Sub placeFormOnScreenAt(Of T As Form)(ByRef pControl As T,
                                                            Optional ByVal pPosition As FormLocationOnScreen = FormLocationOnScreen.CENTER)
            placeFormOnScreenAt(pControl, New Padding(0), pPosition)
        End Sub
        Public Shared Sub placeFormOnScreenAt(Of T As Form)(ByRef pControl As T, ByVal pPadding As Padding,
                                                            Optional ByVal pPosition As FormLocationOnScreen = FormLocationOnScreen.CENTER)


            With pControl
                

                Select Case pPosition
                    Case FormLocationOnScreen.TOP__LEFT__CORNER
                        .Left = pPadding.Left
                        .Top = pPadding.Top

                    Case FormLocationOnScreen.TOP__MIDDLE
                        .Left = CInt((My.Computer.Screen.WorkingArea.Width - .Width) / 2)
                        .Top = pPadding.Top

                    Case FormLocationOnScreen.TOP__RIGHT__CORNER
                        .Left = My.Computer.Screen.WorkingArea.Width - .Width - pPadding.Right
                        .Top = pPadding.Top

                    Case FormLocationOnScreen.MIDDLE__LEFT__CORNER
                        .Left = pPadding.Left
                        .Top = CInt((My.Computer.Screen.WorkingArea.Height - .Height) / 2)

                    Case FormLocationOnScreen.CENTER
                        .Left = CInt((My.Computer.Screen.WorkingArea.Width - .Width) / 2)
                        .Top = CInt((My.Computer.Screen.WorkingArea.Height - .Height) / 2)

                    Case FormLocationOnScreen.MIDDLE__RIGHT__CORNER
                        .Left = My.Computer.Screen.WorkingArea.Width - .Width - pPadding.Right
                        .Top = CInt((My.Computer.Screen.WorkingArea.Height - .Height) / 2)

                    Case FormLocationOnScreen.BOTTOM__LEFT__CORNER
                        .Left = pPadding.Left
                        .Top = My.Computer.Screen.WorkingArea.Height - .Height - pPadding.Bottom

                    Case FormLocationOnScreen.BOTTOM__MIDDLE
                        .Left = CInt((My.Computer.Screen.WorkingArea.Width - .Width) / 2)
                        .Top = My.Computer.Screen.WorkingArea.Height - .Height - pPadding.Bottom

                    Case FormLocationOnScreen.BOTTOM__RIGHT__CORNER
                        .Left = My.Computer.Screen.WorkingArea.Width - .Width - pPadding.Right
                        .Top = My.Computer.Screen.WorkingArea.Height - .Height - pPadding.Bottom
                End Select


            End With

        End Sub




#End Region


        ''' <summary>
        ''' Checks if it is not nothing and not disposed
        ''' </summary>
        ''' <param name="frm"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function isValid_Form(ByVal frm As Form) As Boolean
            Return (frm IsNot Nothing AndAlso Not frm.IsDisposed)
        End Function

        ''' <summary>
        ''' Fetch GUID from an obj.getType
        ''' </summary>
        ''' <param name="objType"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function get_guid_from_ObjectType(ByVal objType As System.Type) As String
            Return New Guid(
                            CType(
                                objType.Assembly.GetCustomAttributes(GetType(GuidAttribute), False)(0), 
                                GuidAttribute).Value()
                               ).ToString

        End Function

        ''' <summary>
        ''' Fetch GUID from an obj.getType
        ''' </summary>
        ''' <param name="objType"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function get_AppTitle_from_ObjectType(ByVal objType As System.Type) As String
            Dim arr As Object() = objType.Assembly.GetCustomAttributes(GetType(System.Reflection.AssemblyTitleAttribute), True)
            Dim AppTitle As String = String.Empty

            If arr IsNot Nothing AndAlso arr.Length > 0 Then
                AppTitle = CType(arr(0), System.Reflection.AssemblyTitleAttribute).Title()
            End If

            Return AppTitle

        End Function


        ''' <summary>
        ''' Force close all application open forms and ignore error
        ''' </summary>
        ''' <remarks></remarks>
        Public Shared Sub ClosedAllApplicationForms()
            Try
                For intC As Int32 = 0 To Application.OpenForms.Count - 1
                    Application.OpenForms(0).Close()
                Next
            Catch ex As Exception
                MyLogFile.Print(ex)
            End Try
        End Sub


#Region "Variables Declarations"
        Public Const TimeFormatUsedWithoutSeconds As Microsoft.VisualBasic.DateFormat = DateFormat.ShortTime
        Public Const DateFormatUsed As String = "dd/MMM/yyyy"

        ' ''97-122 small letters [a - z]
        Public Shared SmallLetters() As Byte =
           New Byte() {
               CByte(97), CByte(98), CByte(99), CByte(100), CByte(101), CByte(102), CByte(103), CByte(104), CByte(105),
               CByte(106), CByte(107), CByte(108), CByte(109), CByte(110), CByte(111), CByte(112), CByte(113),
               CByte(114), CByte(115), CByte(116), CByte(117), CByte(118), CByte(119), CByte(120), CByte(121),
               CByte(122)
            }

        ' ''65-90  Big Letters [A - Z]
        Public Shared BigLetters() As Byte =
            New Byte() {
                CByte(65), CByte(66), CByte(67), CByte(68), CByte(69), CByte(70), CByte(71), CByte(72), CByte(73),
                CByte(74), CByte(75), CByte(76), CByte(77), CByte(78), CByte(79), CByte(80), CByte(81), CByte(82),
                CByte(83), CByte(84), CByte(85), CByte(86), CByte(87), CByte(88), CByte(89), CByte(90)
                    }



        ' ''48 - 57 = > Digits [0 - 9]
        Public Shared Numbers() As Byte =
            New Byte() {
                    CByte(48), CByte(49), CByte(50), CByte(51), CByte(52), CByte(53), CByte(54), CByte(55), CByte(56),
                    CByte(57)
                }


        ' ''Space=>32
        Public Const SpaceBarCode As Byte = CByte(Keys.Space)

#End Region

#Region "EncodingAndDecoding"
        Private Shared Function UnicodeStringToBytes(
        ByVal str As String) As Byte()

            Return System.Text.Encoding.Unicode.GetBytes(str)
        End Function

        Private Shared Function UnicodeBytesToString(
        ByVal bytes() As Byte) As String

            Return System.Text.Encoding.Unicode.GetString(bytes)
        End Function

        Private Shared Function UTF8BytesToString(
     ByVal bytes() As Byte) As String

            Return System.Text.Encoding.UTF8.GetString(bytes)
        End Function
#End Region

#Region "Validation Rules"

        Public Shared Function ValidateShortText(ByRef objInput As String,
                                          ByRef objErrMessage As String,
                                          Optional ByVal MinimumLength As Long = 3,
                                          Optional ByVal MaximumLength As Long = 50) As Boolean

            Try

                If Len(objInput) < MinimumLength Then _
                    objErrMessage = String.Format("Input is too small (Minimum of {0} Characters)", MinimumLength) : Return False

                If Len(objInput) > MaximumLength Then _
        objErrMessage = String.Format(
            "Input is too Large (It has been reduced to {0}. Maximum of {0} Characters)", MaximumLength) _
        : objInput = Left(objInput, CInt(MaximumLength)) : Return False

            Catch ex As Exception


                Return False
            End Try

            Return True
        End Function

        Public Shared Function ValidateLongText(ByRef objInput As String,
                                          ByRef objErrMessage As String,
                                          Optional ByVal MinimumLength As Long = 3,
                                          Optional ByVal MaximumLength As Long = 10000) As Boolean


            Return ValidateShortText(objInput, objErrMessage, MinimumLength, MaximumLength)

        End Function

        Public Shared Function ValidateEmail(ByRef objInput As String,
                                      ByRef objErrMessage As String) As Boolean
            'More Rules
            'If there is space within the text 
            'if there is double @
            'if there is up to 3 min letters before and after the @
            'min of 7 letters
            'Allowed Characters _ . @ [0 - 9][a - z]
            '. after @ must atleast 1 dot
            ' . must not be the last thing and must not be the first thing
            '
            If Not ValidateLongText(objInput, objErrMessage, 7) Then _
                 Return False

            If objInput.IndexOf(" ") > 0 Then _
                objErrMessage = "Spaces are NOT allowed within an email address." : Return False
            Dim intval As Integer = objInput.IndexOf("@")
            If intval >= 0 Then
                If Len(objInput) > intval + 1 Then
                    If objInput.IndexOf("@", intval + 1) >= 0 Then _
                        objErrMessage = "Maximum of 1 @ Symbol is allowed within an email address." : Return False
                End If
            Else
                GoTo InvalidEmailAddress
            End If

            If objInput.IndexOf(".@") >= 0 Then _
               GoTo InvalidEmailAddress

            If objInput.IndexOf("@.") >= 0 Then _
               GoTo InvalidEmailAddress

            If objInput.IndexOf("..") >= 0 Then _
       GoTo InvalidEmailAddress


            If Not ValidateShortText(objInput.Substring(0,
                                                        objInput.IndexOf("@")
                                                        ),
                                     objErrMessage, 4) Then _
         GoTo InvalidEmailAddress

            If Not ValidateShortText(objInput.Substring(objInput.IndexOf("@")),
                             objErrMessage, 4) Then _
                    GoTo InvalidEmailAddress

            If objInput.Substring(objInput.IndexOf("@")).IndexOf(".") < 0 Then _
                GoTo InvalidEmailAddress

            For intC As Integer = 33 To 126
                If intC = 46 Then intC = 47
                If intC = 48 Then intC = 58
                If intC = 64 Then intC = 91
                If intC = 95 Then intC = 96
                If intC = 97 Then intC = 123

                If objInput.IndexOf(Chr(intC)) >= 0 Then _
                    GoTo InvalidEmailAddress
                Application.DoEvents()
            Next

            If objInput.Substring(0, 1) = "." Then _
                    GoTo InvalidEmailAddress

            If objInput.Substring(Len(objInput) - 1, 1) = "." Then _
                   GoTo InvalidEmailAddress

            Return True
InvalidEmailAddress:
            objErrMessage = "Invalid Email address. Hint: Check for Invalid Symbols." : Return False
        End Function

        Public Shared Function ValidateMobileNumber(ByRef objInput As String,
                                             ByRef objErrMessage As String,
                                             Optional ByVal AllowPlus As Boolean = True,
                                             Optional ByVal MinimumLength As Long = 8,
                                          Optional ByVal MaximumLength As Long = 14) As Boolean


            If Not ValidateShortText(objInput, objErrMessage, MinimumLength, MaximumLength) Then Return False
            If AllowPlus Then If objInput.IndexOf("+") > 0 Then objErrMessage = "(+) Sign is wrongly placed! Invalid Number" : Return False
            If AllowPlus Then
                If Not IsNumeric(objInput.Substring(1)) Then GoTo InvalidMobileNumber
            Else
                If Not IsNumeric(objInput) Then GoTo InvalidMobileNumber
            End If


            Return True
            Exit Function
InvalidMobileNumber:
            objErrMessage = "Invalid Mobile Number. [Numbers Only!]"
            Return False
        End Function

        Public Shared Function ValidateNumeric(ByRef objInput As String,
                                        ByRef objErrMessage As String,
                                         Optional ByVal MinimumLength As Long = 1,
                                          Optional ByVal MaximumLength As Long = 20) As Boolean
            If Not ValidateMobileNumber(objInput, objErrMessage, False, MinimumLength, MaximumLength) Then objErrMessage = "Invalid Number(s). Only 0 - 9 are allowed!!!" : Return False
            Return True
        End Function

        Public Shared Function ValidateCurrency(ByRef objInput As String,
                                    ByRef objErrMessage As String,
                                     Optional ByVal MinimumLength As Long = 1,
                                      Optional ByVal MaximumLength As Long = 20) As Boolean

            Return ValidateNumeric(objInput, objErrMessage, MinimumLength, MaximumLength)

        End Function

        Public Shared Function ValidateDate(ByRef objInput As String,
                                ByRef objErrMessage As String,
                                 Optional ByVal DateFormatAccepted As String = DateFormatUsed
                                                                        ) As Boolean

            Try
                objInput = Format(CDate(objInput), DateFormatAccepted)
            Catch ex As Exception
                objErrMessage = String.Format("Invalid Date Entry. Format [{0}]", DateFormatAccepted)
                Return False
            End Try

            Return True
        End Function

        Public Shared Function ValidateDate(ByRef objInput As String,
                              ByRef objErrMessage As String,
                               ByVal DateFormatAccepted As Microsoft.VisualBasic.DateFormat
                                                                      ) As Boolean

            Try
                objInput = FormatDateTime(CDate(objInput), DateFormatAccepted)
            Catch ex As Exception
                objErrMessage = String.Format("Invalid Date/Time Entry. Format [{0}]", DateFormatAccepted.ToString)
                Return False
            End Try

            Return True
        End Function

        Public Shared Function ValidateTime(ByRef objInput As String,
                               ByRef objErrMessage As String,
                                Optional ByVal TimeFormatAccepted As Microsoft.VisualBasic.DateFormat = TimeFormatUsedWithoutSeconds
                                                                       ) As Boolean
            If Not ValidateDate(objInput, objErrMessage, TimeFormatAccepted) Then objErrMessage = "Invalid Time Format" : Return False

            Return True
        End Function




        ''' <summary>
        ''' Allows Leters [a - Z] and Space Only
        ''' </summary>
        ''' <param name="objInput"></param>
        ''' <param name="objErrMessage"></param>
        ''' <param name="MinimumLength"></param>
        ''' <param name="MaximumLength"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function ValidateLettersAndSpaceOnly(ByRef objInput As String,
                                      ByRef objErrMessage As String,
                                      Optional ByVal MinimumLength As Long = 1,
                                      Optional ByVal MaximumLength As Long = 50) As Boolean

            ' '' '' '' ''Dim smallLetters As String = ""
            ' '' '' '' ''For intC As Integer = 48 To 57
            ' '' '' '' ''    smallLetters &=
            ' '' '' '' ''                String.Format(
            ' '' '' '' ''                        "CByte({0})", intC
            ' '' '' '' ''                        )

            ' '' '' '' ''    If 57 > intC Then smallLetters &= ","
            ' '' '' '' ''Next

            ' '' '' '' ''Debug.Print(smallLetters)

            Try
                'Ok, Here We correct the length 

                If Len(objInput) < MinimumLength Then _
                    objErrMessage = String.Format("Input is too small (Minimum of {0} Characters)", MinimumLength) : Return False

                If Len(objInput) > MaximumLength Then _
        objErrMessage = String.Format(
            "Input is too Large (It has been reduced to {0}. Maximum of {0} Characters)", MaximumLength) _
        : objInput = Left(objInput, CInt(MaximumLength)) : Return False
                '************************************************

                ''We assume the objInput carries the readable format. Allows Big Letters, Small Letters and Space
                ''Now we verify the contents

                'Returns Strings with each character separated by vbNull
                Dim Contents() As Byte = UnicodeStringToBytes(objInput)

                Dim AllowedCharacters As List(Of Byte) = New List(Of Byte)
                AllowedCharacters.AddRange(
                            BigLetters
                            )
                AllowedCharacters.AddRange(
                            SmallLetters
                            )
                AllowedCharacters.AddRange(
                            {SpaceBarCode, Asc(vbNullChar)}
                            )
                Dim interceptContents() As Byte = AllowedCharacters.ToArray


                Dim CharsNotAllowed As IEnumerable(Of Byte) = Contents.Except(interceptContents)


                'If this statement is true then the validation is perfect
                If Not CharsNotAllowed.Count <= 0 Then
                    ''    'Valid
                    ''    Return True
                    ''Else
                    objErrMessage = String.Format(
                                                "Allows Leters [a - Z] and Space Only. {0}The following characters are not allowed: {0}{1}",
                                                vbCrLf, UTF8BytesToString(CharsNotAllowed.ToArray)
                                                )

                    ' ''Debug.Print(
                    ' ''    System.Text.Encoding.UTF8.GetString(CharsNotAllowed.ToArray)
                    ' ''    )
                    Return False
                End If

            Catch ex As Exception


                Return False
            End Try

            'Valid
            Return True
        End Function


#End Region


#Region "Enforcing Form to TopMost View"
        Private Shared Sub ActivateFormLater(ByVal sender As System.Object, ByVal e As System.EventArgs)
            Try
                Dim pSender As System.Windows.Forms.Timer = CType(sender, Timer)
                pSender.Stop()
                Dim pForm As Form = CType(pSender.Tag, Form)
                If pForm IsNot Nothing AndAlso Not pForm.IsDisposed Then
                    pForm.Activate()
                End If
            Catch ex As Exception
                basMain.MyLogFile.Print(ex)
            End Try

        End Sub


        Private Shared Sub RemoveFromTopmost(ByVal sender As System.Object, ByVal e As System.EventArgs)
            Try
                Dim pSender As System.Windows.Forms.Timer = CType(sender, Timer)
                pSender.Stop()
                ' pSender.Enabled = False
                Dim pForm As Form = CType(pSender.Tag, Form)
                If pForm IsNot Nothing AndAlso Not pForm.IsDisposed Then pForm.TopMost = False
            Catch ex As Exception
                basMain.MyLogFile.Print(ex)

            End Try

        End Sub

        Public Shared Sub RemoveFormFromTopmostIn(pMilliseconds As Int32, pForm As Form)

            Dim p As New System.Windows.Forms.Timer()
            With p
                .Interval = pMilliseconds
                .Tag = pForm
                AddHandler .Tick, AddressOf RemoveFromTopmost
                .Start()

            End With


        End Sub


        Public Shared Sub AddFormToTopmostIn(pMilliseconds As Int32, pForm As Form)

            Dim p As New System.Windows.Forms.Timer()
            With p
                .Interval = pMilliseconds
                .Tag = pForm
                AddHandler .Tick, AddressOf ActivateFormLater
                pForm.TopMost = True
                .Start()

            End With


        End Sub

#End Region


    End Class

End Namespace