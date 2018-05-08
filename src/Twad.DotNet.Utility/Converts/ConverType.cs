using System;
using System.Collections.Generic;
using System.Globalization;
using System.Reflection;
using System.Text;

namespace Twad.DotNet.Utility.Converts
{
    public static class ConverType
    {
        #region 将字符串表示转换成等效的对象

        /// <summary>
        /// 将字符串表示转换成等效的对象。
        /// </summary>
        /// <typeparam name="T">将value值转换成T类型。</typeparam>
        /// <param name="value">要转换的值。</param>
        /// <returns>返回转换等效的对象；对于 <paramref name="T"/> 为 <see cref="System.DateTime"/> 的转换需做日期校验，日期值不能比 <see cref="DateTimeHelper.MinValue"/> 还要小</returns>
        public static T ToType<T>(string value)
        {
            object targetValue = default(T);
            if (string.IsNullOrEmpty(value))
            {
                return (T)targetValue;
            }

            Type conversionType = typeof(T);
            if (Validation.IsNullableType(conversionType))
            {
                conversionType = Nullable.GetUnderlyingType(conversionType);
            }

            targetValue = ToType(value, conversionType);

            // 检查类型转换的合法性
            CheckTargetValue(ref targetValue);

            return (T)targetValue;
        }

        /// <summary>
        /// 将字符串表示转换成等效的对象。
        /// </summary>
        /// <typeparam name="T">将value值转换成T类型。</typeparam>
        /// <param name="value">要转换的值。</param>
        /// <returns>返回转换等效的对象；对于 <paramref name="T"/> 为 <see cref="System.DateTime"/> 的转换需做日期校验，日期值不能比 <see cref="DateTimeHelper.MinValue"/> 还要小</returns>
        public static T TryType<T>(this object value)
        {
            object targetValue = default(T);
            if (value != null)
            {
                Type conversionType = typeof(T);
                if (Validation.IsNullableType(conversionType))
                {
                    conversionType = Nullable.GetUnderlyingType(conversionType);
                }

                if (conversionType == typeof(string))
                {
                    targetValue = value.ToString();
                }
                else if (typeof(Enum).IsAssignableFrom(conversionType))
                {
                    targetValue = value == null ? Enum.ToObject(conversionType, 0) : Enum.Parse(conversionType, value.ToString(), true);
                }
                else if (conversionType == typeof(Guid))
                {
                    targetValue = value == null ? Guid.Empty : new Guid(value.ToString());
                }
                else
                {
                    Type[] types = new Type[2] { typeof(string), Type.GetType(typeof(T).FullName + "&") };
                    MethodInfo mi = conversionType.GetMethod("TryParse", types);
                    if (mi != null)
                    {
                        object[] outObj = new object[2] { value.ToString(), default(T) };
                        object retValue = mi.Invoke(null, outObj);
                        if ((bool)retValue)
                        {
                            targetValue = outObj[1];
                        }
                    }
                    else
                    {
                        throw new NotSupportedException(string.Format("不支持对类型“{0}”的转换。", conversionType.FullName));
                    }
                }
            }
            return (T)targetValue;
        }

        /// <summary>
        /// 将字符串表示转换成等效的对象。转换错误不抛异常，返回默认值
        /// </summary>
        /// <typeparam name="T">将value值转换成T类型。</typeparam>
        /// <param name="value">要转换的值。</param>
        /// <param name="defaultValue">转换失败后返回的默认值。</param>
        /// <returns>返回转换等效的对象；对于 <paramref name="T"/> 为 <see cref="System.DateTime"/> 的转换需做日期校验，日期值不能比 <see cref="DateTimeHelper.MinValue"/> 还要小</returns>
        public static T TryType<T>(this object value, T defaultValue)
        {
            object targetValue = default(T);
            try
            {
                if (value != null)
                {
                    Type conversionType = typeof(T);
                    if (Validation.IsNullableType(conversionType))
                    {
                        conversionType = Nullable.GetUnderlyingType(conversionType);
                    }

                    if (conversionType == typeof(string))
                    {
                        targetValue = value.ToString();
                    }
                    else if (typeof(Enum).IsAssignableFrom(conversionType))
                    {
                        targetValue = value == null ? Enum.ToObject(conversionType, 0) : Enum.Parse(conversionType, value.ToString(), true);
                    }
                    else if (conversionType == typeof(Guid))
                    {
                        targetValue = value == null ? Guid.Empty : new Guid(value.ToString());
                    }
                    else
                    {
                        Type[] types = new Type[2] { typeof(string), Type.GetType(typeof(T).FullName + "&") };
                        MethodInfo mi = conversionType.GetMethod("TryParse", types);
                        if (mi != null)
                        {
                            object[] outObj = new object[2] { value.ToString(), default(T) };
                            object retValue = mi.Invoke(null, outObj);
                            if ((bool)retValue)
                            {
                                targetValue = outObj[1];
                            }
                            else
                            {
                                return defaultValue;//转换失败
                            }
                        }
                        else
                        {
                            //转换失败
                            return defaultValue;
                            //throw new NotSupportedException(string.Format("不支持对类型“{0}”的转换。", conversionType.FullName));
                        }
                    }
                    return (T)targetValue;
                }
                else
                {
                    return defaultValue;
                }
            }
            catch
            {
                return defaultValue;
            }
        }

        /// <summary>
        /// 将字符串表示转换成等效的对象。转换错误不抛异常，返回默认值
        /// </summary>
        /// <typeparam name="T">将value值转换成T类型。</typeparam>
        /// <param name="value">要转换的值。</param>
        /// <param name="defaultValue">转换失败后返回的默认值。</param>
        /// <returns>返回转换等效的对象；对于 <paramref name="T"/> 为 <see cref="System.DateTime"/> 的转换需做日期校验，日期值不能比 <see cref="DateTimeHelper.MinValue"/> 还要小</returns>
        public static T ToType<T>(string value, T defaultValue)
        {
            T targetValue = defaultValue;
            if (string.IsNullOrEmpty(value))
            {
                return targetValue;
            }
            try
            {
                return ToType<T>(value);
            }
            catch
            {
                return targetValue;
            }
        }

        #endregion

        #region 检查类型转换的合法性

        /// <summary>
        /// 检查类型转换的合法性
        /// </summary>
        /// <param name="targetValue">包含需要校验的转换类型值。</param>
        private static void CheckTargetValue(ref object targetValue)
        {
            if (targetValue == null)
            {
                return;
            }

            if (targetValue.GetType() == typeof(DateTime))
            {
                // 日期值不能比DateTimeHelper.MinValue还要小
                if (targetValue != null
                    && ((DateTime)targetValue).CompareTo(DateTimeConver.MinValue) == -1)
                {
                    targetValue = DateTimeConver.MinValue;
                }
            }
        }



        #endregion

        #region 通过使用指定的区域性特定格式设置信息，将字符串表示转换成等效的对象。

        /// <summary>
        /// 通过使用指定的区域性特定格式设置信息，将字符串表示转换成等效的对象。
        /// </summary>
        /// <param name="value">包含要转换的 <see cref="String"/>。</param>
        /// <param name="conversionType">要转换成的 <see cref="System.Type"/>。</param>
        /// <returns><paramref name="conversionType"/> 类型的对象，其值由 value 表示。</returns>
        private static object ToType(string value, Type conversionType)
        {
            if (conversionType == typeof(string))
            {
                return value;
            }

            if (conversionType == typeof(int))
            {
                return value == null ? 0 : int.Parse(value, NumberStyles.Any);
            }

            if (conversionType == typeof(bool))
            {
                return ToBoolean(value);
            }

            if (conversionType == typeof(float))
            {
                return value == null ? 0f : float.Parse(value, NumberStyles.Any);
            }

            if (conversionType == typeof(double))
            {
                return value == null ? 0.0 : double.Parse(value, NumberStyles.Any);
            }

            if (conversionType == typeof(decimal))
            {
                return value == null ? 0M : decimal.Parse(value, NumberStyles.Any);
            }

            if (conversionType == typeof(DateTime))
            {
                return value == null ? DateTimeConver.MinValue : DateTime.Parse(value, CultureInfo.CurrentCulture, DateTimeStyles.None);
            }

            if (conversionType == typeof(char))
            {
                return System.Convert.ToChar(value);
            }

            if (conversionType == typeof(sbyte))
            {
                return sbyte.Parse(value, NumberStyles.Any);
            }

            if (conversionType == typeof(byte))
            {
                return byte.Parse(value, NumberStyles.Any);
            }

            if (conversionType == typeof(short))
            {
                return value == null ? 0 : short.Parse(value, NumberStyles.Any);
            }

            if (conversionType == typeof(ushort))
            {
                return value == null ? 0 : ushort.Parse(value, NumberStyles.Any);
            }

            if (conversionType == typeof(uint))
            {
                return value == null ? 0 : uint.Parse(value, NumberStyles.Any);
            }

            if (conversionType == typeof(long))
            {
                return value == null ? 0L : long.Parse(value, NumberStyles.Any);
            }

            if (conversionType == typeof(ulong))
            {
                return value == null ? 0 : ulong.Parse(value, NumberStyles.Any);
            }

            if (typeof(Enum).IsAssignableFrom(conversionType))
            {
                return value == null ? Enum.ToObject(conversionType, 0) : Enum.Parse(conversionType, value, true);
            }

            if (conversionType == typeof(Guid))
            {
                return value == null ? Guid.Empty : new Guid(value);
            }

            throw new NotSupportedException(string.Format("不支持对类型“{0}”的转换。", conversionType.FullName));
        }

        #endregion

        #region 将逻辑值转换为它的等效布尔值。
        /// <summary>
        /// 将布尔值 true 表示为可转换的字符串。此字段为只读。 
        /// </summary>
        /// <remarks>该字段包含字符串“TRUE”，“YES”，“ON”，“1”。</remarks>
        public static readonly List<string> TrueStringList = new List<string>() { "TRUE", "YES", "ON", "1" };

        /// <summary>
        /// 将逻辑值转换为它的等效布尔值。
        /// </summary>
        /// <param name="value">处理值的字符串。</param>
        /// <returns>若成功执行转换，返回 true</returns>
        private static bool ToBoolean(string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                return false;
            }

            return TrueStringList.Contains(value.Trim().ToUpper()) ? true : false;
        }
        #endregion
    }
}
