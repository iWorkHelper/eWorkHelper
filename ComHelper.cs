using System;
using System.Runtime.InteropServices;

namespace eWorkhelper
{
    /// <summary>
    /// Excel COM 对象释放助手。
    /// </summary>
    /// <remarks>
    /// 项目此前没有任何 <see cref="Marshal.ReleaseComObject"/> 调用，双点写法产生的大量中间 RCW
    /// 无法回收，会导致工作簿关闭后仍被引用、Excel 进程无法干净退出、批量操作时 RCW 数量暴涨。
    /// 此处集中提供“逐个释放、失败忽略”的基础能力；调用方必须自行保证：
    /// 1) 不释放由 VSTO 宿主拥有的 <c>Globals.ThisAddIn.Application</c>；
    /// 2) 不释放仍然会被调用方或缓存继续使用的对象；
    /// 3) 释放顺序与获取顺序相反。
    /// </remarks>
    internal static class ComHelper
    {
        /// <summary>
        /// 释放单个 COM 对象。非 COM 对象与 null 会被安全忽略，释放过程中的异常也会被吞掉，
        /// 以免 <c>finally</c> 中的释放动作掩盖在途异常。
        /// </summary>
        internal static void Release(object comObject)
        {
            if (comObject == null || !Marshal.IsComObject(comObject))
            {
                return;
            }

            try
            {
                Marshal.FinalReleaseComObject(comObject);
            }
            catch (Exception)
            {
                // 对象可能已被 Excel 回收或已释放；释放路径不得抛出。
            }
        }

        /// <summary>
        /// 释放调用方临时取得的一次宿主 COM 引用。与 Release 不同，这里不能使用
        /// FinalReleaseComObject，因为对象可能仍由 Excel/VSTO 或其他调用方持有。
        /// </summary>
        internal static void ReleaseBorrowed(object comObject)
        {
            if (comObject == null || !Marshal.IsComObject(comObject))
            {
                return;
            }

            try
            {
                Marshal.ReleaseComObject(comObject);
            }
            catch (Exception)
            {
                // 借用引用的清理不得掩盖原始操作异常。
            }
        }
    }
}
