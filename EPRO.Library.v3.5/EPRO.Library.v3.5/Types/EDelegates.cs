
namespace ELibrary.Standard.Types
{



    /// <summary>
    /// Contains Different Delegate Definitions. Delegates is like const .. They dont need shared specifier. Once public
    /// They are available
    /// </summary>
    /// <remarks></remarks>
    public class EDelegates
    {

        #region Delegates Sub

        #endregion

        #region Delegates Functions

        #endregion

        /// <summary>
        /// Takes a string Parameter and returns boolean
        /// </summary>
        /// <param name="strParam"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public delegate bool delegateBoolFunc(string strParam);



        /// <summary>
        /// Takes a string Parameter and returns Nothing
        /// </summary>
        /// <param name="strParam"></param>
        /// <remarks></remarks>
        public delegate void delegateSubString(string strParam);


        /// <summary>
        /// Takes a boolean Parameter and returns Nothing
        /// </summary>
        /// <param name="strParam"></param>
        /// <remarks></remarks>
        public delegate void delegateSubBool(bool strParam);

        /// <summary>
        /// Takes a boolean Parameter,Thread Parameter and returns Nothing
        /// </summary>
        /// <param name="strParam"></param>
        /// <remarks></remarks>
        public delegate void delegateSubBoolThread(bool strParam, System.Threading.Thread strParam);


        /// <summary>
        /// Takes a boolean Parameter and returns Nothing
        /// </summary>
        /// <param name="strParam"></param>
        /// <remarks></remarks>
        public delegate void delegateSubThread(System.Threading.Thread strParam);



        /// <summary>
        /// Collects No Parameters
        /// </summary>
        /// <remarks></remarks>
        public delegate void delegateNoParam();
    }
}