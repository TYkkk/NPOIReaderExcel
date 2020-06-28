using Newtonsoft.Json;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using UnityEngine;

public class AnalysisExcelManager : Singleton<AnalysisExcelManager>
{
    public void AnalysisSheet(ISheet sheetData)
    {
        if (sheetData == null)
        {
            return;
        }

        List<ICell> cellList = new List<ICell>();

        for (int i = sheetData.FirstRowNum; i < sheetData.LastRowNum; i++)
        {
            IRow row = sheetData.GetRow(i);
            cellList.AddRange(row.Cells);
        }
    }

    public void AnalysisSheet(List<ISheet> sheetList)
    {
        if (sheetList == null)
        {
            return;
        }

        for (int i = 0; i < sheetList.Count; i++)
        {
            AnalysisSheet(sheetList[i]);
        }
    }

    public List<ISheet> GetExcelData(IWorkbook workbook)
    {
        if (workbook == null)
        {
            return null;
        }

        List<ISheet> sheetList = new List<ISheet>();

        foreach (var childSheet in workbook)
        {
            sheetList.Add(childSheet);
        }

        return sheetList;
    }

    public IWorkbook ImportExcel(string filePath)
    {
        if (!File.Exists(filePath))
        {
            return null;
        }

        IWorkbook workbook = null;
        try
        {
            using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                string extension = Path.GetExtension(filePath);

                if (extension == Const.ExtensionXLS)
                {
                    workbook = new HSSFWorkbook(fs);
                }
                else if (extension == Const.ExtensionXLSX)
                {
                    workbook = new XSSFWorkbook(fs);
                }
            }
        }
        catch (Exception ex)
        {
            Debug.Log(ex.Message);
        }

        return workbook;
    }
}
