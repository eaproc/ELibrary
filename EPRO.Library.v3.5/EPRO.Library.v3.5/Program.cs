using CODERiT.Logger.v._3._5;

namespace ELibrary.Standard
{
    internal class Program
    {
        static Program()
        {
            ____Logger = new Log1(typeof(Program)).Load(Log1.Modes.FILE, true);
        }

        private static Log1 ____Logger;

        public static Log1 Logger
        {
            get
            {
                return ____Logger;
            }
        }
    }
}