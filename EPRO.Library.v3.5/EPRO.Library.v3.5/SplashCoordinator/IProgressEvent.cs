
namespace ELibrary.Standard.SplashCoordinator
{
    public interface IProgressEvent
    {
        void ProgressChanged(int TotalProgress, int CurrentProgress);
    }
}