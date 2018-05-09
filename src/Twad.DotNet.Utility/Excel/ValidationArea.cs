using NPOI.SS.UserModel;

namespace Twad.DotNet.Utility.Excel
{

    /// <summary>
    /// 单元格样式对象
    /// </summary>
    public class ExeclCellStyle
    {

        private int _width; //30 * 256
                            /// <summary>
                            /// 字段名称
                            /// </summary>
        public string ColumnsName { set; get; }
        /// <summary>
        /// 是否锁定
        /// </summary>
        public bool IsLock { get; set; }
        /// <summary>
        /// 单元格格宽度
        /// </summary>

        public int Width
        {
            get { return _width; }
            set
            {
                _width = value * 256;//单位是1/256个字符宽度，如果是30个字符长度就是 30*256。  
            }
        }
        /// <summary>
        /// 列计算公式
        /// </summary>
        public string SetCellFormula { set; get; }

        /// <summary>
        /// 背景色赋值
        /// </summary>
        public short BackGrandIndexed { get; set; }
        /// <summary>
        /// 是否隐藏列
        /// </summary>
        public bool IsHidden { set; get; }

        /// <summary>
        /// 表头的样式
        /// </summary>
        public ICellStyle TitleStyle { set; get; }

        /// <summary>
        /// 单元格式样式
        /// </summary>
        public ICellStyle CellStyle { set; get; }

    }
    /// <summary>
    /// 验证区域
    /// </summary>
    public class ValidationArea
    {
        /// <summary>
        /// 开始行
        /// </summary>
        public int FirstRow { set; get; }
        /// <summary>
        /// 结束行
        /// </summary>
        public int LastRow { set; get; }
        /// <summary>
        /// 开始列
        /// </summary>
        public int FirstCol { set; get; }
        /// <summary>
        /// 结束列
        /// </summary>
        public int LastCol { set; get; }
        /// <summary>
        /// 数据类型
        /// </summary>
        public int ValidationType { set; get; }
    }

}