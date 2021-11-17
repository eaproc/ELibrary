
namespace ELibrary.Standard.SplashCoordinator
{
    public interface IProcessCoordinator
    {
        /// <summary>
        /// Runs when process coordinator has completely processed all processes passed in as parameters
        /// </summary>
        /// <remarks></remarks>
        void onExitAction();

        /// <summary>
        /// It tries to invoke this if user terminates the class before completion
        /// </summary>
        /// <remarks></remarks>
        void onTerminateAction();
    }
}