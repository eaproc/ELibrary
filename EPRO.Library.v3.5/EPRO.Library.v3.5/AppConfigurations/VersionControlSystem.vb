Imports EPRO.Library.v3._5.Modules
Imports EPRO.Library.v3._5.Objects

Namespace AppConfigurations

    ''' <summary>
    ''' Manages Versions in Major.Minor.Revision.Build at Same Sizes
    ''' </summary>
    ''' <remarks></remarks>
    <Serializable()> _
    Public Class VersionControlSystem

#Region "Constructors"

        ''' <summary>
        ''' Parse in already created Version String
        ''' </summary>
        ''' <param name="vcs"></param>
        ''' <remarks></remarks>
        Sub New(ByVal vcs As String,
                ByVal vcsSize As Byte)
            Me.New(vcs:=vcs, vcsSize:=vcsSize, PadOutputWIthZeros:=False)
        End Sub


        ''' <summary>
        ''' Parse in already created Version String
        ''' </summary>
        ''' <param name="vcs"></param>
        ''' <remarks></remarks>
        Sub New(ByVal vcs As String,
                ByVal vcsSize As Byte,
                ByVal PadOutputWIthZeros As Boolean
                )

            Me.New(vcs:=vcs, vcsSize:=vcsSize, PadOutputWIthZeros:=PadOutputWIthZeros, ___VersionDelimiter:=DEFAULT_VERSION_DELIMITER)

        End Sub

        ''' <summary>
        ''' Create an empty version 0.0.0.0
        ''' </summary>
        ''' <remarks></remarks>
        Sub New(
                ByVal vcsSize As Byte,
                ByVal PadOutputWIthZeros As Boolean,
                ByVal ___VersionDelimiter As String)
            Me.New(vcs:=String.Empty, vcsSize:=vcsSize, PadOutputWIthZeros:=PadOutputWIthZeros, ___VersionDelimiter:=___VersionDelimiter)
        End Sub
        ''' <summary>
        ''' Create an empty version 0.0.0.0
        ''' </summary>
        ''' <remarks></remarks>
        Sub New(
                ByVal vcsSize As Byte,
                ByVal PadOutputWIthZeros As Boolean
                )
            Me.New(vcsSize:=vcsSize, PadOutputWIthZeros:=PadOutputWIthZeros, ___VersionDelimiter:=DEFAULT_VERSION_DELIMITER)
        End Sub

        ''' <summary>
        ''' Parse in already created Version String
        ''' </summary>
        ''' <param name="vcs"></param>
        ''' <remarks></remarks>
        Sub New(ByVal vcs As String,
                ByVal vcsSize As Byte,
                ByVal PadOutputWIthZeros As Boolean,
                ByVal ___VersionDelimiter As String)

            Dim mj As ULong, mi As ULong, r As ULong, b As ULong

            If vcs Is Nothing Then vcs = String.Empty

            Try
                Dim pVals As String() = vcs.Split(New String() {___VersionDelimiter}, StringSplitOptions.RemoveEmptyEntries)

                If pVals.Length > 0 Then mj = CULng(pVals(0).toLong())
                If pVals.Length > 1 Then mi = CULng(pVals(1).toLong())
                If pVals.Length > 2 Then r = CULng(pVals(2).toLong())
                If pVals.Length > 3 Then b = CULng(pVals(3).toLong())
            Catch ex As Exception
                Throw New Exception("Invalid Version String")
            End Try

            Me.Load(
       ___Major:=mj, ___Minor:=mi, ___Revision:=r, ___Build:=b,
           vcsSize:=vcsSize, PadOutputWIthZeros:=PadOutputWIthZeros, ___VersionDelimiter:=___VersionDelimiter
       )


        End Sub


        ''' <summary>
        ''' Parse in already created Version String
        ''' </summary>
        ''' <param name="vcs"></param>
        ''' <remarks></remarks>
        Sub New(ByVal vcs As String)
            Me.New(vcs:=vcs, vcsSize:=MAXIMUM_VCS_SIZE)
        End Sub

        Sub New(ByVal ___Major As ULong,
                ByVal ___Minor As ULong,
                ByVal ___Revision As ULong,
                ByVal ___Build As ULong)

            Me.New(___Major:=___Major, ___Minor:=___Minor, ___Revision:=___Revision, ___Build:=___Build, vcsSize:=MAXIMUM_VCS_SIZE)

        End Sub

        ''' <summary>
        ''' 
        ''' </summary>
        ''' <param name="___Major"></param>
        ''' <param name="___Minor"></param>
        ''' <param name="___Revision"></param>
        ''' <param name="___Build"></param>
        ''' <param name="vcsSize">Maximum is 12 and Minimum is 1</param>
        ''' <remarks></remarks>
        Sub New(ByVal ___Major As ULong,
               ByVal ___Minor As ULong,
               ByVal ___Revision As ULong,
               ByVal ___Build As ULong,
               ByVal vcsSize As Byte
               )

            Me.New(___Major:=___Major, ___Minor:=___Minor, ___Revision:=___Revision, ___Build:=___Build, vcsSize:=vcsSize, PadOutputWIthZeros:=False)

        End Sub


        ''' <summary>
        ''' 
        ''' </summary>
        ''' <param name="___Major"></param>
        ''' <param name="___Minor"></param>
        ''' <param name="___Revision"></param>
        ''' <param name="___Build"></param>
        ''' <param name="vcsSize">Maximum is 12 and Minimum is 1</param>
        ''' <remarks></remarks>
        Sub New(ByVal ___Major As ULong,
               ByVal ___Minor As ULong,
               ByVal ___Revision As ULong,
               ByVal ___Build As ULong,
               ByVal vcsSize As Byte,
               ByVal PadOutputWIthZeros As Boolean)

            Me.New(___Major:=___Major, ___Minor:=___Minor, ___Revision:=___Revision, ___Build:=___Build,
                   vcsSize:=vcsSize, PadOutputWIthZeros:=PadOutputWIthZeros, ___VersionDelimiter:=DEFAULT_VERSION_DELIMITER)

        End Sub

        ''' <summary>
        ''' 
        ''' </summary>
        ''' <param name="___Major"></param>
        ''' <param name="___Minor"></param>
        ''' <param name="___Revision"></param>
        ''' <param name="___Build"></param>
        ''' <param name="vcsSize">Maximum is 12 and Minimum is 1</param>
        ''' <remarks></remarks>
        Sub New(ByVal ___Major As ULong,
               ByVal ___Minor As ULong,
               ByVal ___Revision As ULong,
               ByVal ___Build As ULong,
               ByVal vcsSize As Byte,
               ByVal PadOutputWIthZeros As Boolean,
               ByVal ___VersionDelimiter As String)
            Me.Load(
               ___Major:=___Major, ___Minor:=___Minor, ___Revision:=___Revision, ___Build:=___Build,
                   vcsSize:=vcsSize, PadOutputWIthZeros:=PadOutputWIthZeros, ___VersionDelimiter:=___VersionDelimiter
               )
        End Sub


        ''' <summary>
        ''' Proxy Load
        ''' </summary>
        ''' <param name="___Major"></param>
        ''' <param name="___Minor"></param>
        ''' <param name="___Revision"></param>
        ''' <param name="___Build"></param>
        ''' <param name="vcsSize"></param>
        ''' <param name="PadOutputWIthZeros"></param>
        ''' <param name="___VersionDelimiter"></param>
        ''' <remarks></remarks>
        Private Sub Load(ByVal ___Major As ULong,
            ByVal ___Minor As ULong,
            ByVal ___Revision As ULong,
            ByVal ___Build As ULong,
            ByVal vcsSize As Byte,
            ByVal PadOutputWIthZeros As Boolean,
            ByVal ___VersionDelimiter As String)

            If vcsSize > MAXIMUM_VCS_SIZE OrElse vcsSize < MINIMUM_VCS_SIZE Then Throw New Exception("INVALID PARAMETER (vcsSize) VALUE: " & vcsSize)
            Me._Major = ___Major
            Me._Minor = ___Minor
            Me._Revision = ___Revision
            Me._Build = ___Build

            Me.PadOutputWithZeros = PadOutputWIthZeros
            Me.VersionDelimiter = ___VersionDelimiter
            Me.VersionSize = vcsSize

            REM Confirm all inputs
            If Major.ToString().Length > vcsSize OrElse
                Minor.ToString().Length > vcsSize OrElse
                Revision.ToString().Length > vcsSize OrElse
                Build.ToString().Length > vcsSize Then Throw New Exception("The parsed in Values are bigger the Maximum Size")




        End Sub



#End Region


#Region "Consts"

        ''' <summary>
        ''' Maximum Length of each component like Major Size 3 will be 999 maximum value
        ''' </summary>
        ''' <remarks></remarks>
        Public Const MAXIMUM_VCS_SIZE As Byte = 12
        Public Const MINIMUM_VCS_SIZE As Byte = 1
        Public Const DEFAULT_VERSION_DELIMITER As String = "."

#End Region



#Region "Properties"


        Private VersionSize As Byte
        Private PadOutputWithZeros As Boolean
        Private VersionDelimiter As String



        Private _Major As ULong
        Public ReadOnly Property Major As ULong
            Get
                Return Me._Major
            End Get
        End Property

        Private _Minor As ULong
        Public ReadOnly Property Minor As ULong
            Get
                Return Me._Minor
            End Get
        End Property


        Private _Revision As ULong
        Public ReadOnly Property Revision As ULong
            Get
                Return Me._Revision
            End Get
        End Property



        Private _Build As ULong
        Public ReadOnly Property Build As ULong
            Get
                Return Me._Build
            End Get
        End Property


        Public ReadOnly Property ComponentMaximumValue As ULong
            Get
                Return CULng(
                                StrDup(Me.VersionSize, "9"c).toLong()
                                )
            End Get
        End Property

        Public ReadOnly Property MaximumVersionValue As String
            Get
                Return String.Format("{1}{0}{1}{0}{1}{0}{1}", Me.VersionDelimiter, Me.ComponentMaximumValue)
            End Get
        End Property



#End Region


#Region "Methods"


        Private ThreadSafe__IncreaseMethod As New Object()
        ''' <summary>
        ''' Increase this Current Version. Thread Safe. Throws Exception(MaxReach)
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function Increase() As Boolean

            SyncLock ThreadSafe__IncreaseMethod
                If Me.ToString() = Me.MaximumVersionValue Then Throw New Exception("Maximum Value Reached!!!")

                If Me.IncreaseComponentPart(Me._Build) OrElse
                    Me.IncreaseComponentPart(Me._Revision) OrElse
                    Me.IncreaseComponentPart(Me._Minor) OrElse
                    Me.IncreaseComponentPart(Me._Major) Then

                    Return True

                Else
                    Throw New Exception("Maximum Value Reached!!!")

                End If

            End SyncLock

            Return True
        End Function



        ''' <summary>
        ''' returns false if it is at maximum already. So pass one up
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function IncreaseComponentPart(ByRef compPart As ULong) As Boolean

            If compPart < Me.ComponentMaximumValue Then

                compPart = CULng(compPart + 1)
                Return True

            ElseIf compPart = Me.ComponentMaximumValue Then

                compPart = 0
                Return False

            End If

            Return False REM It can't be greater than. So, this line will never run

        End Function



        ''' <summary>
        ''' Returns Version as string
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Overrides Function ToString() As String


            If Me.PadOutputWithZeros Then
                Return String.Format("{1}{0}{2}{0}{3}{0}{4}",
                                     Me.VersionDelimiter,
                                     Me.Major.ToString().PadLeft(Me.VersionSize, "0"c),
                                     Me.Minor.ToString().PadLeft(Me.VersionSize, "0"c),
                                     Me.Revision.ToString().PadLeft(Me.VersionSize, "0"c),
                                     Me.Build.ToString().PadLeft(Me.VersionSize, "0"c)
                                     )
            Else
                Return String.Format("{1}{0}{2}{0}{3}{0}{4}", Me.VersionDelimiter, Me.Major, Me.Minor, Me.Revision, Me.Build)
            End If



        End Function

        ''' <summary>
        ''' Confirms if it is same version regardless of it's version size
        ''' </summary>
        ''' <param name="obj"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Overrides Function Equals(obj As Object) As Boolean
            If TypeOf obj Is VersionControlSystem Then
                With CType(obj, VersionControlSystem)
                    Return .Major = Me.Major AndAlso .Minor = Me.Minor AndAlso
                        .Revision = Me.Revision AndAlso .Build = Me.Build

                End With

            End If

            Return MyBase.Equals(obj)
        End Function

        ''' <summary>
        ''' Confirms if this version is greater than the other
        ''' </summary>
        ''' <param name="obj"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function IsGreaterThan(obj As VersionControlSystem) As Boolean
            With Me
                Return .Major > obj.Major OrElse
                    (.Major = obj.Major AndAlso .Minor > obj.Minor) OrElse
                    (.Major = obj.Major AndAlso .Minor = obj.Minor AndAlso .Revision > obj.Revision) OrElse
                    (.Major = obj.Major AndAlso .Minor = obj.Minor AndAlso .Revision = obj.Revision AndAlso .Build > obj.Build)


            End With
        End Function


        ''' <summary>
        ''' Helps to reduce string to needed size. trimming version string. 
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function NormalizeStringVersion(ByVal vcs As String,
                ByVal vcsSize As Byte,
                ByVal PadOutputWIthZeros As Boolean,
                ByVal ___VersionDelimiter As String) As VersionControlSystem



            Dim mj As String = String.Empty, mi As String = String.Empty, r As String = String.Empty, b As String = String.Empty

            Try
                Dim pVals As String() = vcs.Split(New String() {___VersionDelimiter}, StringSplitOptions.RemoveEmptyEntries)

                If pVals.Length > 0 Then mj = Left(pVals(0), vcsSize)
                If pVals.Length > 1 Then mi = Left(pVals(1), vcsSize)
                If pVals.Length > 2 Then r = Left(pVals(2), vcsSize)
                If pVals.Length > 3 Then b = Left(pVals(3), vcsSize)

                Return New VersionControlSystem(
                    String.Format("{1}{0}{2}{0}{3}{0}{4}", ___VersionDelimiter, mj, mi, r, b),
                    vcsSize, PadOutputWIthZeros, ___VersionDelimiter
                    )
            Catch ex As Exception
                Throw New Exception("Invalid Version String")
            End Try


        End Function

        Public Shared Function NormalizeStringVersion(ByVal vcs As String,
              ByVal vcsSize As Byte,
              ByVal PadOutputWIthZeros As Boolean) As VersionControlSystem
            Return NormalizeStringVersion(vcs:=vcs, vcsSize:=vcsSize, PadOutputWIthZeros:=PadOutputWIthZeros, ___VersionDelimiter:=DEFAULT_VERSION_DELIMITER)
        End Function


#End Region



    End Class

End Namespace