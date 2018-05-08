using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;

namespace Twad.DotNet.Utility.Converts
{
    public static class Validation
    {
        private const string email = @"\w+([-+.]\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*";
        private const string mobile = @"^((\(\d{2,3}\))|(\d{3}\-))?(13|14|15|18)\d{9}$";
        private const string phone = @"\d{3}-\d{8}|\d{4}-\d{7}";
        private const string url = @"[a-zA-z]+://[^\s]*";

        /// <summary>
        /// 验证Email地址
        /// </summary>
        /// <param name="input"></param>
        /// <returns></returns>
        public static bool ValidEmail(this string input)
        {
            return ValidRegex(input, email);
        }

        /// <summary>
        /// 验证手机号码
        /// </summary>
        /// <param name="value"></param>
        public static bool ValidMobile(this string input)
        {
            return ValidRegex(input, mobile);
        }

        /// <summary>
        /// 验证电话号码 匹配形式如 0511-4405222 或 021-87888822
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        public static bool ValidPhone(this string input)
        {
            return ValidRegex(input, phone);
        }

        /// <summary>
        /// 验证url
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        public static bool ValidUrl(this string input)
        {
            return ValidRegex(input, url);
        }

        /// <summary>
        /// 根据正则验证字符串
        /// </summary>
        /// <param name="input"></param>
        /// <param name="pattern"></param>
        /// <returns></returns>
        public static bool ValidRegex(this string input, string pattern)
        {
            if (string.IsNullOrEmpty(input) || string.IsNullOrEmpty(pattern))
            {
                return false;
            }

            return Regex.Match(input, pattern).Success;
        }

        /// <summary>
        /// 验证字符串是否为空
        /// </summary>
        /// <param name="input"></param>
        /// <returns></returns>
        public static bool IsEmpty(this string input)
        {
            return string.IsNullOrEmpty(input) || string.IsNullOrEmpty(input.Trim());
        }

        /// <summary>
        /// 验证对象是否为null
        /// </summary>
        /// <param name="input"></param>
        /// <returns></returns>
        public static bool IsEmpty(this object input)
        {
            return ReferenceEquals(input, null);
        }
        /// <summary>
        /// 验证集合是否为null或lenth为0
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="list"></param>
        /// <returns></returns>
        public static bool IsEmpty<T>(this IList<T> list)
        {
            return list == null || list.Count <= 0;
        }

        /// <summary>
        /// 判断Type是否为空
        /// </summary>
        /// <param name="theType"></param>
        /// <returns></returns>
        public static bool IsNullableType(this Type theType)
        {
            return (theType.IsGenericType && !theType.IsGenericTypeDefinition)
                && (theType.GetGenericTypeDefinition() == typeof(Nullable<>));
        }
    }
}
