﻿using System;
using System.IO;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Formatters.Binary;

namespace ELibrary.Standard.Modules
{
    public static class basObjectCloner
    {
        public static T Clone<T>(this T source)
        {
            if (!typeof(T).IsSerializable)
                throw new ArgumentException("The type must be serializable.", "source");
            if (ReferenceEquals(source, null))
                return default;
            IFormatter formatter = new BinaryFormatter();
            Stream stream = new MemoryStream();
            using (stream)
            {
                formatter.Serialize(stream, source);
                stream.Seek(0L, SeekOrigin.Begin);
                return (T)formatter.Deserialize(stream);
            }
        }
    }
}