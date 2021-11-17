using System;

namespace ELibrary.Standard.Caching
{

    /// <summary>
    /// Don't write code under form_Load events instead write it under LoadControls. Do not cache start up forms
    /// </summary>
    /// <remarks></remarks>
    public interface ICacheableControl : IDisposable
    {

        /// <summary>
        /// Use this to set your form to the initial state you want it to be when it is loaded as new
        /// </summary>
        /// <remarks></remarks>
        void ClearControls();

        /// <summary>
        /// Serves as a constructor for you to load the copy you are getting
        /// </summary>
        /// <param name="objParams"></param>
        /// <remarks></remarks>
        void LoadControls(params object[] objParams);



        // ' ''' <summary>
        // ' ''' Use this to return a new version of this component. That will be use to create next version
        // ' ''' </summary>
        // ' ''' <returns></returns>
        // ' ''' <remarks></remarks>
        // 'Function getNew() As ICacheableControl



        // 'ReadOnly Property IsDisposed As Boolean


    }
}