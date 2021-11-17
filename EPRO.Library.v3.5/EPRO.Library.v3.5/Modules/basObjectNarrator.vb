Imports System.Linq.Expressions

Namespace Modules

    Public Module basObjectNarrator

        ''' <summary>
        ''' Usage: VB: Function() variableName, C# ()=> variableName
        ''' </summary>
        ''' <typeparam name="T"></typeparam>
        ''' <param name="expr"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetVariableFullName(Of T)(ByVal expr As Expression(Of Func(Of T))) As String
            Dim body As MemberExpression = CType(expr.Body, MemberExpression)
            If body.ToString().IndexOf("+") >= 0 Then Return body.Member.Name REM it is a local variable
            Return body.ToString().Split(New String() {", "}, StringSplitOptions.RemoveEmptyEntries)(1).Split(New String() {": "}, StringSplitOptions.RemoveEmptyEntries)(1)
        End Function

        ''' <summary>
        ''' Usage: VB: Function() variableName, C# ()=> variableName
        ''' </summary>
        ''' <typeparam name="T"></typeparam>
        ''' <param name="expr"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetVariableName(Of T)(ByVal expr As Expression(Of Func(Of T))) As String
            Dim body As MemberExpression = CType(expr.Body, MemberExpression)
            Return body.Member.Name
        End Function

        ''' <summary>
        ''' Usage: Function() Me.MethodName(). if there is param, pass null
        ''' </summary>
        ''' <typeparam name="T"></typeparam>
        ''' <param name="expr"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetMethodFullName(Of T)(ByVal expr As Expression(Of Func(Of T))) As String
            Dim body As MethodCallExpression = CType(expr.Body, MethodCallExpression)
            Return body.ToString().Split(New String() {", "}, StringSplitOptions.RemoveEmptyEntries)(1).Split(New String() {": "}, StringSplitOptions.RemoveEmptyEntries)(1)
        End Function

        ''' <summary>
        ''' Usage: Function() Me.MethodName(). if there is param, pass null
        ''' </summary>
        ''' <typeparam name="T"></typeparam>
        ''' <param name="expr"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetMethodName(Of T)(ByVal expr As Expression(Of Func(Of T))) As String
            Dim body As MethodCallExpression = CType(expr.Body, MethodCallExpression)
            Return body.Method.Name
        End Function

        Public Function GetTypeFullName(Of T)(ByVal obj As T) As String
            Return obj.GetType().FullName
        End Function

        Public Function GetTypeName(Of T)(ByVal obj As T) As String
            Return obj.GetType().Name
        End Function



    End Module

End Namespace