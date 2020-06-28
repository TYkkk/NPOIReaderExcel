using System.Collections;
using System.Collections.Generic;
using UnityEngine;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;
using System.Data;
using System.IO;
using System.Linq;

public class Test : MonoBehaviour
{
    void Start()
    {
        string filePath = @"D:\TY\油管道阀室坐标.xlsx";

        IWorkbook workbook = AnalysisExcelManager.Instance.ImportExcel(filePath);

        List<ISheet> sheetList = AnalysisExcelManager.Instance.GetExcelData(workbook);

        AnalysisExcelManager.Instance.AnalysisSheet(sheetList);
    }
}
