using Newtonsoft.Json;
using Newtonsoft.Json.Converters;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text;

namespace Twad.DotNet.Utility.Json
{
    public class JsonNetHelper
    {
        public static string EnJson(object result)
        {
            return EnJson(0, null, result, GetToken(2));
        }

        public static string EnJson(int reason, object result)
        {
            return EnJson(reason, null, result, GetToken(2));
        }

        public static string EnJson(int reason, string message, object result)
        {
            return EnJson(reason, message, result, GetToken(2));
        }

        public static string EnJson(int reason, string message, object result, string token)
        {
            var timeConverter = new IsoDateTimeConverter();

            timeConverter.DateTimeFormat = "yyyy'-'MM'-'dd' 'HH':'mm':'ss";

            var obj = new { Reason = reason, Message = message, Result = result, Token = token };

            var jsonStr = JsonConvert.SerializeObject(obj, timeConverter);

            return jsonStr;
        }

        public JToken DeJson(string json)
        {
            var jt = JsonConvert.DeserializeObject<JToken>(json);
            return jt;
        }

        public static byte[] EnEncodingJson(object result)
        {
            return EnEncodingJson(0, null, result, GetToken(2));
        }

        public static byte[] EnEncodingJson(int reason, object result)
        {
            return EnEncodingJson(reason, null, result, GetToken(2));
        }

        public static byte[] EnEncodingJson(int reason, string message, object result)
        {
            return EnEncodingJson(reason, message, result, GetToken(2));
        }

        public static byte[] EnEncodingJson(int reason, string message, object result, string token)
        {
            var timeConverter = new IsoDateTimeConverter();

            timeConverter.DateTimeFormat = "yyyy'-'MM'-'dd' 'HH':'mm':'ss";

            var obj = new { Reason = reason, Message = message, Result = result, Token = token };

            var jsonByte = Encoding.UTF8.GetBytes(JsonConvert.SerializeObject(obj, timeConverter));

            return jsonByte;
        }

        public JToken DeEncodingJson(byte[] jsonByte)
        {
            var jsonStr = Encoding.UTF8.GetString(jsonByte);
            var jt = JsonConvert.DeserializeObject<JToken>(jsonStr);
            return jt;
        }

        public static byte[] EnGzipJson(object result)
        {
            return EnGzipJson(0, null, result, GetToken(2));
        }

        public static byte[] EnGzipJson(int reason, object result)
        {
            return EnGzipJson(reason, null, result, GetToken(2));
        }

        public static byte[] EnGzipJson(int reason, string message, object result)
        {
            return EnGzipJson(reason, message, result, GetToken(2));
        }

        public static byte[] EnGzipJson(int reason, string message, object result, string token)
        {
            var timeConverter = new IsoDateTimeConverter();

            timeConverter.DateTimeFormat = "yyyy'-'MM'-'dd' 'HH':'mm':'ss";

            var obj = new { Reason = reason, Message = message, Result = result, Token = token };

            var jsonByte = Gzip.EnCompress(JsonConvert.SerializeObject(obj, timeConverter));

            return jsonByte;
        }

        public JToken DeGzipJson(byte[] jsonByte)
        {
            var jsonStr = Gzip.DeCompressStr(jsonByte);
            var jt = JsonConvert.DeserializeObject<JToken>(jsonStr);
            return jt;
        }

        /// <summary>
        ///     获取Token(即方法名)
        /// </summary>
        /// <param name="index"></param>
        /// <returns></returns>
        private static string GetToken(int index)
        {
            try
            {
                var month = new StackTrace().GetFrame(index).GetMethod();

                return month.ReflectedType.Name + "_" + month.Name;
            }
            catch (Exception)
            {
                return "Auto generate token error!";
            }
        }

        #region Json序列化/反序列化

        public static T JsonDeserialize<T>(string Json)
        {
            try
            {
                return JsonConvert.DeserializeObject<T>(Json);
            }
            catch (Exception ex)
            {
                throw;
            }
        }

        public static string JsonSerialize<T>(T obj)
        {
            var settings = new JsonSerializerSettings();
            settings.NullValueHandling = NullValueHandling.Ignore;

            return JsonConvert.SerializeObject(obj, Formatting.None, settings);
        }

        public static List<T> JsonDeserializeList<T>(string Json)
        {
            var list = new List<T>();
            try
            {
                return JsonConvert.DeserializeObject<List<T>>(Json);
            }
            catch (Exception ex)
            {
                throw;
            }
        }

        public static string JsonSerializeList<T>(List<T> list)
        {
            var settings = new JsonSerializerSettings();
            settings.NullValueHandling = NullValueHandling.Ignore;

            return JsonConvert.SerializeObject(list, Formatting.None, settings);
        }

        #endregion Json序列化/反序列化
    }
}
