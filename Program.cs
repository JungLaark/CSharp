using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Diagnostics;
//EXCEL.EXE 파일을 추가할 경우 생긴다 . 

namespace MakeExcelFile
{

    class Program
    {
        //매개변수로 스트링 값을 받아야 하는데 어떻게 받지 VB에서 C#프로그램에 매개변수를 던져줄 수 있나??? 

        static void Main(string[] args)
        {
            string carType = "test";
            string bodyNo = "H1W 123456";
            string value1 = "test";
            string value2 = "test";
            int lastRow = 0;

            if(args.Length == 4)
            {
                carType = args[0];
                bodyNo = args[1];
                value1 = args[2];
                value2 = args[3];
            }

            string folderPath = "D:\\TEST\\COMDATA\\" + DateTime.Now.ToString("yyyy"); //폴더 경로 -> 테스트로 D:\\에서 실행 
            string fileName = "SensorData_" + bodyNo.Substring(0, 3) + "_" + DateTime.Now.ToString("yyyyMM") + ".xls";//파일명


            DirectoryInfo di = new DirectoryInfo(folderPath);

            if (di.Exists == false)//디렉토리가 없다면 생성 
            {
                di.Create();
            }
            try
            {
                Excel.Application excelApp = null;
                Excel.Workbook excelWorkBook = null;
                Excel.Worksheet excelWorkSheet = null;
                Excel.Range excelRange = null;
                object missing = System.Reflection.Missing.Value;

                FileInfo fileInfo = new FileInfo(folderPath +"\\"+ fileName);
                if (fileInfo.Exists)//파일이 있는 경우 
                {
                    excelApp = new Excel.Application();
                    excelWorkBook = excelApp.Workbooks.Open(folderPath + "\\" + fileName, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing);
                    excelWorkSheet = excelApp.ActiveSheet;
                    excelRange = excelWorkSheet.UsedRange;

                    for(int i=1; i<= excelRange.Rows.Count; i++)
                    {
                        lastRow += 1;
                    }

                    //현재 덮어쓰고 있음 
                    excelWorkSheet.Cells[lastRow+1, 1] = bodyNo;
                    excelWorkSheet.Cells[lastRow+1, 2] = DateTime.Now.ToString("yyyyMMdd");
                    excelWorkSheet.Cells[lastRow+1, 3] = DateTime.Now.ToString("hh:mm:ss");
                    excelWorkSheet.Cells[lastRow+1, 4] = value1;
                    excelWorkSheet.Cells[lastRow+1, 5] = value2;

                    excelWorkBook.Close(true, Type.Missing, Type.Missing);

                    if (excelApp != null)
                    {
                        Process[] pProcess;
                        pProcess = System.Diagnostics.Process.GetProcessesByName("Excel");
                        pProcess[0].Kill();
                    }
                    excelApp = null;
                }
                else
                {

                    //파일이 없을 경우 
                    excelApp = new Excel.Application(); //엑셀 객체 생성 
                    excelApp.DisplayAlerts = false;
                    excelWorkBook = excelApp.Workbooks.Add();
                    //excelWorkSheet = (Excel.Worksheet)excelWorkBook.Worksheets["Sheet1"];//이건 왜 강제 형변환하는거지??/
                    excelWorkSheet = excelApp.ActiveSheet;


                    //스타일
                    excelWorkSheet.Range["A1:E1"].Borders.Color = System.Drawing.Color.Black;
                    //excelWorkSheet.Range["A1:E1"].Cells.

                    excelWorkSheet.Name = "TEST";
                    //칼럼명
                    excelWorkSheet.Cells[1, 1] = "Body_No";
                    excelWorkSheet.Cells[1, 2] = "날짜";
                    excelWorkSheet.Cells[1, 3] = "시간";
                    excelWorkSheet.Cells[1, 4] = "Value1";
                    excelWorkSheet.Cells[1, 5] = "Value2";
                    //데이터 넣기 
                    excelWorkSheet.Cells[2, 1] = bodyNo;
                    excelWorkSheet.Cells[2, 2] = DateTime.Now.ToString("yyyyMMdd");
                    excelWorkSheet.Cells[2, 3] = DateTime.Now.ToString("hh:mm:ss");
                    excelWorkSheet.Cells[2, 4] = value1;
                    excelWorkSheet.Cells[2, 5] = value2;


                    //저장
                    excelWorkBook.SaveAs(folderPath +"\\"+ fileName, Excel.XlFileFormat.xlWorkbookNormal, missing, missing, missing, missing, Excel.XlSaveAsAccessMode.xlShared, Excel.XlSaveConflictResolution.xlLocalSessionChanges, missing, missing, missing, missing);
                        
                    //releaseExcelObject(excelWorkSheet);
                    //releaseExcelObject(excelWorkBook);
                    //releaseExcelObject(excelApp);
                    excelWorkBook.Close(true, fileName, Type.Missing);

                    if(excelApp != null)
                    {
                        Process[] pProcess;
                        pProcess = System.Diagnostics.Process.GetProcessesByName("Excel");
                        pProcess[0].Kill();
                    }
                    excelApp = null;

                }
            }
            catch (Exception ex)
            {

            }

        }

    }
}
