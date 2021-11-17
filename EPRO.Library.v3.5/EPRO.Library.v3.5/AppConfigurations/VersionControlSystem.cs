using System;
using ELibrary.Standard.Modules;
using Microsoft.VisualBasic;

namespace ELibrary.Standard.AppConfigurations
{

    /// <summary>
    /// Manages Versions in Major.Minor.Revision.Build at Same Sizes
    /// </summary>
    /// <remarks></remarks>
    [Serializable()]
    public class VersionControlSystem
    {

        #region Constructors

        /// <summary>
        /// Parse in already created Version String
        /// </summary>
        /// <param name="vcs"></param>
        /// <remarks></remarks>
        public VersionControlSystem(string vcs, byte vcsSize) : this(vcs: vcs, vcsSize: vcsSize, PadOutputWIthZeros: false)
        {
        }


        /// <summary>
        /// Parse in already created Version String
        /// </summary>
        /// <param name="vcs"></param>
        /// <remarks></remarks>
        public VersionControlSystem(string vcs, byte vcsSize, bool PadOutputWIthZeros) : this(vcs: vcs, vcsSize: vcsSize, PadOutputWIthZeros: PadOutputWIthZeros, ___VersionDelimiter: DEFAULT_VERSION_DELIMITER)
        {
        }

        /// <summary>
        /// Create an empty version 0.0.0.0
        /// </summary>
        /// <remarks></remarks>
        public VersionControlSystem(byte vcsSize, bool PadOutputWIthZeros, string ___VersionDelimiter) : this(vcs: string.Empty, vcsSize: vcsSize, PadOutputWIthZeros: PadOutputWIthZeros, ___VersionDelimiter: ___VersionDelimiter)
        {
        }
        /// <summary>
        /// Create an empty version 0.0.0.0
        /// </summary>
        /// <remarks></remarks>
        public VersionControlSystem(byte vcsSize, bool PadOutputWIthZeros) : this(vcsSize: vcsSize, PadOutputWIthZeros: PadOutputWIthZeros, ___VersionDelimiter: DEFAULT_VERSION_DELIMITER)
        {
        }

        /// <summary>
        /// Parse in already created Version String
        /// </summary>
        /// <param name="vcs"></param>
        /// <remarks></remarks>
        public VersionControlSystem(string vcs, byte vcsSize, bool PadOutputWIthZeros, string ___VersionDelimiter)
        {
            var mj = default(ulong);
            var mi = default(ulong);
            var r = default(ulong);
            var b = default(ulong);
            if (vcs is null)
                vcs = string.Empty;
            try
            {
                var pVals = vcs.Split(new string[] { ___VersionDelimiter }, StringSplitOptions.RemoveEmptyEntries);
                if (pVals.Length > 0)
                    mj = (ulong)pVals[0].toLong();
                if (pVals.Length > 1)
                    mi = (ulong)pVals[1].toLong();
                if (pVals.Length > 2)
                    r = (ulong)pVals[2].toLong();
                if (pVals.Length > 3)
                    b = (ulong)pVals[3].toLong();
            }
            catch (Exception ex)
            {
                throw new Exception("Invalid Version String");
            }

            Load(___Major: mj, ___Minor: mi, ___Revision: r, ___Build: b, vcsSize: vcsSize, PadOutputWIthZeros: PadOutputWIthZeros, ___VersionDelimiter: ___VersionDelimiter);
        }


        /// <summary>
        /// Parse in already created Version String
        /// </summary>
        /// <param name="vcs"></param>
        /// <remarks></remarks>
        public VersionControlSystem(string vcs) : this(vcs: vcs, vcsSize: MAXIMUM_VCS_SIZE)
        {
        }

        public VersionControlSystem(ulong ___Major, ulong ___Minor, ulong ___Revision, ulong ___Build) : this(___Major: ___Major, ___Minor: ___Minor, ___Revision: ___Revision, ___Build: ___Build, vcsSize: MAXIMUM_VCS_SIZE)
        {
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="___Major"></param>
        /// <param name="___Minor"></param>
        /// <param name="___Revision"></param>
        /// <param name="___Build"></param>
        /// <param name="vcsSize">Maximum is 12 and Minimum is 1</param>
        /// <remarks></remarks>
        public VersionControlSystem(ulong ___Major, ulong ___Minor, ulong ___Revision, ulong ___Build, byte vcsSize) : this(___Major: ___Major, ___Minor: ___Minor, ___Revision: ___Revision, ___Build: ___Build, vcsSize: vcsSize, PadOutputWIthZeros: false)
        {
        }


        /// <summary>
        /// 
        /// </summary>
        /// <param name="___Major"></param>
        /// <param name="___Minor"></param>
        /// <param name="___Revision"></param>
        /// <param name="___Build"></param>
        /// <param name="vcsSize">Maximum is 12 and Minimum is 1</param>
        /// <remarks></remarks>
        public VersionControlSystem(ulong ___Major, ulong ___Minor, ulong ___Revision, ulong ___Build, byte vcsSize, bool PadOutputWIthZeros) : this(___Major: ___Major, ___Minor: ___Minor, ___Revision: ___Revision, ___Build: ___Build, vcsSize: vcsSize, PadOutputWIthZeros: PadOutputWIthZeros, ___VersionDelimiter: DEFAULT_VERSION_DELIMITER)
        {
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="___Major"></param>
        /// <param name="___Minor"></param>
        /// <param name="___Revision"></param>
        /// <param name="___Build"></param>
        /// <param name="vcsSize">Maximum is 12 and Minimum is 1</param>
        /// <remarks></remarks>
        public VersionControlSystem(ulong ___Major, ulong ___Minor, ulong ___Revision, ulong ___Build, byte vcsSize, bool PadOutputWIthZeros, string ___VersionDelimiter)
        {
            Load(___Major: ___Major, ___Minor: ___Minor, ___Revision: ___Revision, ___Build: ___Build, vcsSize: vcsSize, PadOutputWIthZeros: PadOutputWIthZeros, ___VersionDelimiter: ___VersionDelimiter);
        }


        /// <summary>
        /// Proxy Load
        /// </summary>
        /// <param name="___Major"></param>
        /// <param name="___Minor"></param>
        /// <param name="___Revision"></param>
        /// <param name="___Build"></param>
        /// <param name="vcsSize"></param>
        /// <param name="PadOutputWIthZeros"></param>
        /// <param name="___VersionDelimiter"></param>
        /// <remarks></remarks>
        private void Load(ulong ___Major, ulong ___Minor, ulong ___Revision, ulong ___Build, byte vcsSize, bool PadOutputWIthZeros, string ___VersionDelimiter)
        {
            if (vcsSize > MAXIMUM_VCS_SIZE || vcsSize < MINIMUM_VCS_SIZE)
                throw new Exception("INVALID PARAMETER (vcsSize) VALUE: " + vcsSize);
            _Major = ___Major;
            _Minor = ___Minor;
            _Revision = ___Revision;
            _Build = ___Build;
            PadOutputWithZeros = PadOutputWIthZeros;
            VersionDelimiter = ___VersionDelimiter;
            VersionSize = vcsSize;

            // REM Confirm all inputs
            if (Major.ToString().Length > vcsSize || Minor.ToString().Length > vcsSize || Revision.ToString().Length > vcsSize || Build.ToString().Length > vcsSize)
                throw new Exception("The parsed in Values are bigger the Maximum Size");
        }



        #endregion


        #region Consts

        /// <summary>
        /// Maximum Length of each component like Major Size 3 will be 999 maximum value
        /// </summary>
        /// <remarks></remarks>
        public const byte MAXIMUM_VCS_SIZE = 12;
        public const byte MINIMUM_VCS_SIZE = 1;
        public const string DEFAULT_VERSION_DELIMITER = ".";

        #endregion



        #region Properties


        private byte VersionSize;
        private bool PadOutputWithZeros;
        private string VersionDelimiter;
        private ulong _Major;

        public ulong Major
        {
            get
            {
                return _Major;
            }
        }

        private ulong _Minor;

        public ulong Minor
        {
            get
            {
                return _Minor;
            }
        }

        private ulong _Revision;

        public ulong Revision
        {
            get
            {
                return _Revision;
            }
        }

        private ulong _Build;

        public ulong Build
        {
            get
            {
                return _Build;
            }
        }

        public ulong ComponentMaximumValue
        {
            get
            {
                return (ulong)Strings.StrDup(VersionSize, '9').toLong();
            }
        }

        public string MaximumVersionValue
        {
            get
            {
                return string.Format("{1}{0}{1}{0}{1}{0}{1}", VersionDelimiter, ComponentMaximumValue);
            }
        }



        #endregion


        #region Methods


        private object ThreadSafe__IncreaseMethod = new object();
        /// <summary>
        /// Increase this Current Version. Thread Safe. Throws Exception(MaxReach)
        /// </summary>
        /// <returns></returns>
        /// <remarks></remarks>
        public bool Increase()
        {
            lock (ThreadSafe__IncreaseMethod)
            {
                if ((ToString() ?? "") == (MaximumVersionValue ?? ""))
                    throw new Exception("Maximum Value Reached!!!");
                if (IncreaseComponentPart(ref _Build) || IncreaseComponentPart(ref _Revision) || IncreaseComponentPart(ref _Minor) || IncreaseComponentPart(ref _Major))
                {
                    return true;
                }
                else
                {
                    throw new Exception("Maximum Value Reached!!!");
                }
            }

            return true;
        }



        /// <summary>
        /// returns false if it is at maximum already. So pass one up
        /// </summary>
        /// <returns></returns>
        /// <remarks></remarks>
        private bool IncreaseComponentPart(ref ulong compPart)
        {
            if (compPart < ComponentMaximumValue)
            {
                compPart = (ulong)Math.Round(compPart + 1m);
                return true;
            }
            else if (compPart == ComponentMaximumValue)
            {
                compPart = 0UL;
                return false;
            }

            return false; // REM It can't be greater than. So, this line will never run
        }



        /// <summary>
        /// Returns Version as string
        /// </summary>
        /// <returns></returns>
        /// <remarks></remarks>
        public override string ToString()
        {
            if (PadOutputWithZeros)
            {
                return string.Format("{1}{0}{2}{0}{3}{0}{4}", VersionDelimiter, Major.ToString().PadLeft(VersionSize, '0'), Minor.ToString().PadLeft(VersionSize, '0'), Revision.ToString().PadLeft(VersionSize, '0'), Build.ToString().PadLeft(VersionSize, '0'));
            }
            else
            {
                return string.Format("{1}{0}{2}{0}{3}{0}{4}", VersionDelimiter, Major, Minor, Revision, Build);
            }
        }

        /// <summary>
        /// Confirms if it is same version regardless of it's version size
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public override bool Equals(object obj)
        {
            if (obj is VersionControlSystem)
            {
                {
                    var withBlock = (VersionControlSystem)obj;
                    return withBlock.Major == Major && withBlock.Minor == Minor && withBlock.Revision == Revision && withBlock.Build == Build;
                }
            }

            return base.Equals(obj);
        }

        /// <summary>
        /// Confirms if this version is greater than the other
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public bool IsGreaterThan(VersionControlSystem obj)
        {
            {
                var withBlock = this;
                return withBlock.Major > obj.Major || withBlock.Major == obj.Major && withBlock.Minor > obj.Minor || withBlock.Major == obj.Major && withBlock.Minor == obj.Minor && withBlock.Revision > obj.Revision || withBlock.Major == obj.Major && withBlock.Minor == obj.Minor && withBlock.Revision == obj.Revision && withBlock.Build > obj.Build;
            }
        }


        /// <summary>
        /// Helps to reduce string to needed size. trimming version string. 
        /// </summary>
        /// <returns></returns>
        /// <remarks></remarks>
        public static VersionControlSystem NormalizeStringVersion(string vcs, byte vcsSize, bool PadOutputWIthZeros, string ___VersionDelimiter)
        {
            string mj = string.Empty;
            string mi = string.Empty;
            string r = string.Empty;
            string b = string.Empty;
            try
            {
                var pVals = vcs.Split(new string[] { ___VersionDelimiter }, StringSplitOptions.RemoveEmptyEntries);
                if (pVals.Length > 0)
                    mj = Strings.Left(pVals[0], vcsSize);
                if (pVals.Length > 1)
                    mi = Strings.Left(pVals[1], vcsSize);
                if (pVals.Length > 2)
                    r = Strings.Left(pVals[2], vcsSize);
                if (pVals.Length > 3)
                    b = Strings.Left(pVals[3], vcsSize);
                return new VersionControlSystem(string.Format("{1}{0}{2}{0}{3}{0}{4}", ___VersionDelimiter, mj, mi, r, b), vcsSize, PadOutputWIthZeros, ___VersionDelimiter);
            }
            catch (Exception ex)
            {
                throw new Exception("Invalid Version String");
            }
        }

        public static VersionControlSystem NormalizeStringVersion(string vcs, byte vcsSize, bool PadOutputWIthZeros)
        {
            return NormalizeStringVersion(vcs: vcs, vcsSize: vcsSize, PadOutputWIthZeros: PadOutputWIthZeros, ___VersionDelimiter: DEFAULT_VERSION_DELIMITER);
        }


        #endregion



    }
}