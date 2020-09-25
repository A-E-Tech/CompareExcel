
using CompareExcelCore.Class;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;

namespace CompareExcelCore
{
    public class OfficeHelper
    {

        public static string ReadExcelToTagList(string fileName, string sheetName, out List<Tag> tagList)
        {
            ISheet sheet = null;
            string sErr = "";
            tagList = new List<Tag>();
            try
            {
                if (!File.Exists(fileName))
                {
                    sErr += "There is no file!";
                    goto Get_Out;
                }
                FileStream fs = new FileStream(fileName, FileMode.Open, FileAccess.Read);
                IWorkbook workbook = WorkbookFactory.Create(fs);
                if (!string.IsNullOrEmpty(sheetName))
                {
                    sheet = workbook.GetSheet(sheetName);
                    if (sheet == null)
                    {
                        sErr += "Sheet name is incorrect!";
                        goto Get_Out;
                    }
                }
                else
                {

                    sErr += "Sheet name cannot be empty!";
                }
                if (sheet != null)
                {
                    IRow firstRow = sheet.GetRow(1); //should be 0
                    List<IRow> rows = new List<IRow>();
                    for (int j = 5; j < sheet.PhysicalNumberOfRows; j++)
                    {
                        IRow eachRow = sheet.GetRow(j);
                        rows.Add(eachRow);
                    }
                    int rowLength = firstRow.Cells.Count - 2;
                    for (int i = 2; i < firstRow.Cells.Count; i++)
                    {
                        for (int k = 0; k < rows.Count; k++)
                        {
                            Tag tagMV = new Tag()
                            {
                                ColumnIndex = i,
                                MV = firstRow.Cells[i].ToString(),
                                RowIndex = k + 5,
                                CV = rows[k].Cells[0].ToString(), //should be 0
                                Value = rows[k].Cells[i + 1].ToString(),
                            };
                            tagList.Add(tagMV);
                        }

                    }


                }
            }
            catch (Exception ex)
            {
                sErr += ex.Message + "\r\n" + ex.StackTrace;
            }
        Get_Out:
            return sErr;
        }

        public string QueryTag(List<Tag> input, string xMV, string xCV, out Tag res)
        {
            string sErr = "";
            res = new Tag();
            try
            {
                List<Tag> xCVtags = new List<Tag>();
                xCVtags = input.Where(x => x.CV == xCV).ToList();
                res = xCVtags.FirstOrDefault(x => x.MV == xMV);
            }
            catch (Exception ex)
            {

                sErr += ex.Message + "\r\n" + ex.StackTrace;
            }


            return sErr;
        }
        public int Compare(string exp, string act) //if same, return 0
        {
            int res = 0;
            res = String.Compare(exp, act);
            return res;
        }

        public string CompareSet(List<Tag> childSet, List<Tag> superSet, out string xMessage)
        {
            string sErr = "";
            xMessage = "";
            try
            {
                for (int i = 0; i < childSet.Count; i++)
                {
                    int res = 0;
                    Tag tagInSuperSet = new Tag();
                    string xCV = childSet[i].CV;
                    string xMV = childSet[i].MV;
                    sErr = QueryTag(superSet, xMV, xCV, out tagInSuperSet);
                    if (sErr != "") goto Get_Out;
                    res = Compare(tagInSuperSet.Value, childSet[i].Value);
                    if (sErr != "") goto Get_Out;
                    if (res != 0)
                    {
                        xMessage += $"Tag MV is {xMV}; CV is {xCV}: Value 1 is {tagInSuperSet.Value}; Value 2 is {childSet[i].Value}" + "\r\n";
                    }
                }
            }
            catch (Exception ex)
            {

                sErr += ex.Message + "\r\n" + ex.StackTrace;
            }
        Get_Out:
            return sErr;
        }

        public string WriteLog(string xMessage)
        {
            string sErr = "";
            try
            {
                FileStream fs = new FileStream("Result.log", FileMode.Create, FileAccess.ReadWrite);
                byte[] array = Encoding.UTF8.GetBytes(xMessage);
                fs.Write(array, 0, array.Length);
                fs.Close();

            }
            catch (Exception ex)
            {

                sErr += ex.Message + "\r\n" + ex.StackTrace;
            }
            return sErr;
        }

        public string CompareFiles(string superSetName, string superSetSheetName, string childSetName, string childSetSheetName)
        {
            string sErr = "";
            string xMessage = "";
            try
            {
                List<Tag> superSet = new List<Tag>();
                List<Tag> childSet = new List<Tag>();
                sErr = ReadExcelToTagList(superSetName, superSetSheetName, out superSet);
                if (sErr != "") goto Get_Out;
                sErr = ReadExcelToTagList(childSetName, childSetSheetName, out childSet);
                if (sErr != "") goto Get_Out;
                sErr = CompareSet(childSet, superSet, out xMessage);
                if (sErr != "") goto Get_Out;
                if(xMessage=="")
                {
                    xMessage = "No difference Ada ^^";
                }
                sErr = WriteLog(xMessage);
                if (sErr != "") goto Get_Out;
            }
            catch (Exception ex)
            {

                sErr += ex.Message + "\r\n" + ex.StackTrace;
            }
        Get_Out:
            return sErr;
        }
    }
}
