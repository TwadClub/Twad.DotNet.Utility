using System;
using System.Collections.Generic;
using System.Text;

namespace Twad.DotNet.Utility.Converts
{
    public static class DateTimeConver
    {
        #region Fields

        /// <summary>
        /// 最小日期值
        /// </summary>
        public static readonly DateTime MinValue = new DateTime(1900, 1, 1);

        /// <summary>
        /// 最大日期值
        /// </summary>
        public static readonly DateTime MaxValue = DateTime.MaxValue;

        /// <summary>
        /// 日期信息
        /// </summary>
        private static readonly string[] weekString = new string[] { "日", "一", "二", "三", "四", "五", "六" };

        private static object lockObj = new object();

        #endregion

        /// <summary>
        /// 获取此实例所表示的日期是星期几
        /// </summary>
        /// <param name="time">日期</param>
        /// <returns>星期</returns>
        public static string FormatWeekString(this DateTime time)
        {
            return "星期" + weekString[(int)time.DayOfWeek];
        }

        /// <summary>
        /// yyyy-MM-dd
        /// </summary>
        /// <param name="time"></param>
        /// <returns></returns>
        public static string FormatShortString(this DateTime time)
        {
            return time.ToString("yyyy-MM-dd");
        }

        /// <summary>
        /// yyyy年MM月dd日
        /// </summary>
        /// <param name="time"></param>
        /// <returns></returns>
        public static string FormatShortStringC(this DateTime time)
        {
            return time.ToString("yyyy年MM月dd日");
        }

        /// <summary>
        /// yyyy-MM-dd HH:mm:ss
        /// </summary>
        /// <param name="time"></param>
        /// <returns></returns>
        public static string FormatLongString(this DateTime time)
        {
            return time.ToString("yyyy-MM-dd HH:mm:ss");
        }

        /// <summary>
        /// yyyy年MM月dd日 HH:mm:ss
        /// </summary>
        /// <param name="time"></param>
        /// <returns></returns>
        public static string FormatLongStringC(this DateTime time)
        {
            return time.ToString("yyyy年MM月dd日 HH:mm:ss");
        }
    }
}
