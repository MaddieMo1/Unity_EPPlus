using System.Collections;
using System.Collections.Generic;
using UnityEngine;
using OfficeOpenXml;
using UnityEditor;
using System.IO;
using OfficeOpenXml.Style;

public class EPPlusExcel_ZH : MonoBehaviour
{
    //文件夹路径
    private string _FilePath;
    //文件夹路径存储
    private FileInfo _FileInfo;

    ExcelWorksheet _Worksheet;
    ExcelWorksheet _WorksheetNew;

    void Start()
    {
        //路径
        _FilePath = Application.streamingAssetsPath + "/Epplus.xlsx";
        //获取Excel 文件路径
        _FileInfo = new FileInfo(_FilePath);
        //开启协程
        StartCoroutine(EPPlus());
    }

    private IEnumerator EPPlus()
    {
        //Excel 数据赋予
        using (ExcelPackage _ExcelPackage = new ExcelPackage(_FileInfo))
        {
            if (_ExcelPackage.Workbook.Worksheets["Epplus_Test"] == null)
            {
                //表格创建
                _WorksheetNew = _ExcelPackage.Workbook.Worksheets.Add("Epplus_Test");
            }

            //获取数据表中第一张表格数据
            _Worksheet = _ExcelPackage.Workbook.Worksheets[1];

            //获取数据表中第一行第一列数据
            string _Message = _Worksheet.Cells[1, 1].Value.ToString();

            //数据表写入
            _Worksheet.Cells[1, 1].Value = "A1";
            _Worksheet.Cells["A2"].Value = "A2";
            _Worksheet.SetValue(2, 2, "EPPlus");

            //这是乘法的公式，意思是第三列乘以第四列的值赋值给第五列
            _Worksheet.Cells["E2"].Formula = "C2*D2";

            //这是求和公式，意思是第二行第三列的值到第四行第三例的值求和后赋给第五行第三列  并转移到 五行五列
            _Worksheet.Cells[5, 3, 5, 5].Formula = string.Format("SUBTOTAL(9,{0})", new ExcelAddress(2, 3, 4, 3).Address);


            //获取一个区域，区域范围是第七行第一列到第七行第五列
            using (var _Range = _WorksheetNew.Cells[7, 1, 7, 5])
            {
                //粗体
                _Range.Style.Font.Bold = true;
                //设置单元格样式为无样式
                _Range.Style.Fill.PatternType = ExcelFillStyle.None;
                //设置单元格底色为红色
                _Range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Red);
                //设置单元格字体颜色为黄色
                _Range.Style.Font.Color.SetColor(System.Drawing.Color.Yellow);
            }

            //表格删除
            _ExcelPackage.Workbook.Worksheets.Delete("Epplus_Test");

            //保存数据表
            _ExcelPackage.Save();
            
        }
#if UNITY_EDITOR
        //资源刷新
        AssetDatabase.Refresh();

        // 保存所有修改
        AssetDatabase.SaveAssets();
#endif
        yield return null;
    }
}
