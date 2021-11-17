Imports System.Windows.Forms
Imports System.Drawing

''' <summary>
''' Helps to move anther control relative to mouse position after clicking the invoke. Works on left click and dragging
''' </summary>
''' <remarks></remarks>
Public Class MoveControl___RelativeToMousePosition
    Implements IDisposable







#Region "Properties"


    Private ______Control_to_InvokeDrag As Control
    Private ______Control_to_Relocate As Control

    Public ReadOnly Property IsDisposed As Boolean
        Get
            Return Me.disposedValue
        End Get
    End Property




    Dim _____MouseLocation____OnBeginMove As Point
    Dim _____ControlToMove__Location__OnBeginMove As Point
    Dim _____KeepMovingControl As Boolean



#End Region




#Region "Constructors"


    ''' <summary>
    ''' Helps to drag control
    ''' </summary>
    ''' <param name="p______Control_to_InvokeDrag">Control that will be mouse clicked to move</param>
    ''' <param name="p______Control_to_Relocate">Control that will be moved</param>
    ''' <remarks></remarks>
    Public Sub New(
                  p______Control_to_InvokeDrag As Control,
                  p______Control_to_Relocate As Control
                  )

        If p______Control_to_InvokeDrag Is Nothing OrElse p______Control_to_Relocate Is Nothing OrElse
            p______Control_to_InvokeDrag.IsDisposed OrElse p______Control_to_Relocate.IsDisposed Then _
            Throw New InvalidOperationException("You are calling this class with wrong or disposed parameters")

        Me.______Control_to_InvokeDrag = p______Control_to_InvokeDrag
        Me.______Control_to_Relocate = p______Control_to_Relocate


        AddHandler Me.______Control_to_InvokeDrag.MouseDown, AddressOf Me.______Control_to_InvokeDrag_MouseDown
        AddHandler Me.______Control_to_InvokeDrag.MouseUp, AddressOf Me.______Control_to_InvokeDrag_MouseUp
        AddHandler Me.______Control_to_InvokeDrag.MouseMove, AddressOf Me.______Control_to_InvokeDrag_MouseMove


    End Sub




#End Region





#Region "Events"

    Public Event Moved(pSender As MoveControl___RelativeToMousePosition)

#End Region



















#Region "IDisposable Support"
    Private disposedValue As Boolean ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' TODO: dispose managed state (managed objects).

                If ______Control_to_InvokeDrag IsNot Nothing AndAlso Not Me.______Control_to_InvokeDrag.IsDisposed Then
                    RemoveHandler Me.______Control_to_InvokeDrag.MouseDown, AddressOf Me.______Control_to_InvokeDrag_MouseDown
                    RemoveHandler Me.______Control_to_InvokeDrag.MouseUp, AddressOf Me.______Control_to_InvokeDrag_MouseUp
                    RemoveHandler Me.______Control_to_InvokeDrag.MouseMove, AddressOf Me.______Control_to_InvokeDrag_MouseMove
                End If

            End If

            ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
            ' TODO: set large fields to null.
        End If
        Me.disposedValue = True
    End Sub

    ' TODO: override Finalize() only if Dispose(ByVal disposing As Boolean) above has code to free unmanaged resources.
    'Protected Overrides Sub Finalize()
    '    ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
    '    Dispose(False)
    '    MyBase.Finalize()
    'End Sub

    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
#End Region





#Region "Methods"




    Private Sub ______Control_to_InvokeDrag_MouseDown(sender As Object, e As MouseEventArgs)
        _____MouseLocation____OnBeginMove = Control.MousePosition
        _____ControlToMove__Location__OnBeginMove = ______Control_to_Relocate.Location
        _____KeepMovingControl = True

    End Sub

    Private Sub ______Control_to_InvokeDrag_MouseUp(sender As Object, e As MouseEventArgs)
        _____KeepMovingControl = False

    End Sub

    Private Sub ______Control_to_InvokeDrag_MouseMove(sender As Object, e As MouseEventArgs)
        Dim pNewMousePoint = Control.MousePosition

        If e.Button = Windows.Forms.MouseButtons.Left AndAlso _____KeepMovingControl Then
            ______Control_to_Relocate.Location = New Point(
                _____ControlToMove__Location__OnBeginMove.X + (pNewMousePoint.X - _____MouseLocation____OnBeginMove.X),
                _____ControlToMove__Location__OnBeginMove.Y + (pNewMousePoint.Y - _____MouseLocation____OnBeginMove.Y)
                 )


            RaiseEvent Moved(Me)

        End If
    End Sub




#End Region



End Class
