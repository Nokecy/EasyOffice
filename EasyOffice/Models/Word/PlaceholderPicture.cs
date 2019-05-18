namespace EasyOffice.Models.Word
{
    /// <summary>
    /// 替换占位符的图片类
    /// </summary>
    public class PlaceholderPicture
    {
        /// <summary>
        /// 占位符
        /// </summary>
        public string PlaceHolder { get; set; }

        /// <summary>
        /// 图片名称
        /// </summary>
        public string PictureName { get; set; }

        /// <summary>
        /// 图片全地址
        /// </summary>
        public string PictureUrl { get; set; }

        /// <summary>
        /// 图片类型
        /// </summary>
        public string ImageType { get; set; }

        /// <summary>
        /// 后缀名
        /// </summary>
        public string Extension { get; set; }

        /// <summary>
        /// 图片ID
        /// </summary>
        public uint PicId { get; set; }
    }
}
