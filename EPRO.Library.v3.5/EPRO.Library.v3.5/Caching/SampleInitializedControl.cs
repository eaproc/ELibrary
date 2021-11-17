
namespace ELibrary.Standard.Caching
{
    public class SampleInitializedControl : System.Windows.Forms.Control
    {

        /// <summary>
        /// Just a simple control that force the creation of handle under this calling thread
        /// </summary>
        /// <remarks></remarks>
        public SampleInitializedControl() : base()
        {
            CreateHandle();
        }
    }
}