using CODERiT.Logger.v._3._5;

namespace ELibrary.Standard.Modules
{
    internal static class basMain
    {
        private static Log1 _LogFile = new Log1(typeof(basMain)).Load(Log1.Modes.FILE, true);

        public static Log1 MyLogFile
        {
            get
            {
                return _LogFile;
            }
        }
    }
}