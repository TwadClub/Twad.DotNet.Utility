using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Runtime.Serialization;
using System.Text;
using System.Xml;
using System.Xml.Serialization;

namespace Twad.DotNet.Utility.XML
{
    public class ObjectXmlSerializer
    {
        /// <summary>
        /// deserialize an object from a file.
        /// 
        /// Add loggingEnabled Flag
        /// 
        /// Modified by Colin He,  5/14/2007
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="m_FileNamePattern"></param>
        /// <returns>
        /// loggingEnabled==true: Null is returned if any error occurs.
        /// loggingEnabled==false: throw exception
        /// </returns>
        public static T LoadFromXml<T>(string fileName) where T : class
        {
            return LoadFromXml<T>(fileName, true);
        }

        public static T LoadFromXml<T>(string fileName, bool loggingEnabled) where T : class
        {
            FileStream fs = null;
            try
            {
                XmlSerializer serializer = new XmlSerializer(typeof(T));
                fs = new FileStream(fileName, FileMode.Open, FileAccess.Read);
                return (T)serializer.Deserialize(fs);
            }
            catch (Exception e)
            {
                if (loggingEnabled)
                {
                    LogLoadFileException(fileName, e);
                    return null;
                }
                else throw;
            }
            finally
            {
                if (fs != null)
                {
                    fs.Close();
                }
            }
        }

        /// <summary>
        /// serialize an object to a file.
        /// 
        /// Colin He,  5/14/2007
        /// 
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="m_FileNamePattern"></param>
        /// <returns>
        /// loggingEnabled==true: log exception
        /// loggingEnabled==false: throw exception
        /// </returns>
        public static void SaveToXml<T>(string fileName, T data) where T : class
        {
            SaveToXml<T>(fileName, data, true);
        }

        public static void SaveToXml<T>(string fileName, T data, bool loggingEnabled) where T : class
        {
            FileStream fs = null;
            try
            {

                XmlSerializer serializer = new XmlSerializer(typeof(T));
                fs = new FileStream(fileName, FileMode.Create, FileAccess.Write);
                serializer.Serialize(fs, data);
            }
            catch (Exception e)
            {
                if (loggingEnabled) LogSaveFileException(fileName, e);
                else throw;
            }
            finally
            {
                if (fs != null)
                {
                    fs.Close();
                }
            }
        }

        /// <summary>
        /// XML & Datacontract Serialize & Deserialize Helper
        /// 
        /// Cj Xie, Colin He, 5/14/2007
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="serialObject"></param>
        /// <returns></returns>
        public static string XmlSerializer<T>(T serialObject) where T : class
        {
            XmlSerializer ser = new XmlSerializer(typeof(T));
            System.IO.MemoryStream mem = new MemoryStream();
            XmlTextWriter writer = new XmlTextWriter(mem, Encoding.UTF8);
            ser.Serialize(writer, serialObject);
            writer.Close();

            return Encoding.UTF8.GetString(mem.ToArray());
        }

        public static T XmlDeserialize<T>(string str) where T : class
        {
            XmlSerializer mySerializer = new XmlSerializer(typeof(T));
            StreamReader mem2 = new StreamReader(
                    new MemoryStream(System.Text.Encoding.UTF8.GetBytes(str)),
                    System.Text.Encoding.UTF8);

            return (T)mySerializer.Deserialize(mem2);
        }


        public static T DataContractDeserializer<T>(string xmlData) where T : class
        {
            MemoryStream stream = new MemoryStream(Encoding.UTF8.GetBytes(xmlData));

            XmlDictionaryReader reader =
                XmlDictionaryReader.CreateTextReader(stream, new XmlDictionaryReaderQuotas());
            DataContractSerializer ser = new DataContractSerializer(typeof(T));
            T deserializedPerson = (T)ser.ReadObject(reader, true);
            reader.Close();
            stream.Close();

            return deserializedPerson;
        }

        public static string DataContractSerializer<T>(T myObject) where T : class
        {
            MemoryStream stream = new MemoryStream();
            DataContractSerializer ser = new DataContractSerializer(typeof(T));
            ser.WriteObject(stream, myObject);
            stream.Close();

            return System.Text.UnicodeEncoding.UTF8.GetString(stream.ToArray());
        }

        #region Logging
        private const string LogCategory = "Framework.ObjectXmlSerializer";
        private const int LogEventLoadFileException = 1;

        [Conditional("TRACE")]
        private static void LogLoadFileException(string fileName, Exception ex)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("Fail to load xml file: ");
            sb.Append(fileName + System.Environment.NewLine);
            sb.Append(ex.ToString());
            //Logger.LogEvent(LogCategory, LogEventLoadFileException, sb.ToString());
        }

        [Conditional("TRACE")]
        private static void LogSaveFileException(string fileName, Exception ex)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("Fail to save xml file: ");
            sb.Append(fileName + System.Environment.NewLine);
            sb.Append(ex.ToString());
            //Logger.LogEvent(LogCategory, LogEventLoadFileException, sb.ToString());
        }
        #endregion
    }
}
