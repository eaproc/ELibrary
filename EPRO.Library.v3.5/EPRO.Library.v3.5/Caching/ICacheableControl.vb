Namespace Caching

    ''' <summary>
    ''' Don't write code under form_Load events instead write it under LoadControls. Do not cache start up forms
    ''' </summary>
    ''' <remarks></remarks>
    Public Interface ICacheableControl
        Inherits IDisposable

        ''' <summary>
        ''' Use this to set your form to the initial state you want it to be when it is loaded as new
        ''' </summary>
        ''' <remarks></remarks>
        Sub ClearControls()

        ''' <summary>
        ''' Serves as a constructor for you to load the copy you are getting
        ''' </summary>
        ''' <param name="objParams"></param>
        ''' <remarks></remarks>
        Sub LoadControls(ParamArray objParams() As Object)



        '' ''' <summary>
        '' ''' Use this to return a new version of this component. That will be use to create next version
        '' ''' </summary>
        '' ''' <returns></returns>
        '' ''' <remarks></remarks>
        ''Function getNew() As ICacheableControl



        ''ReadOnly Property IsDisposed As Boolean


    End Interface


End Namespace