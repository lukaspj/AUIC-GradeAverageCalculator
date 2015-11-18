using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;

namespace NationalitetsParser
{
   public partial class Form1 : Form
   {
      private Dictionary<string, Student> mStudents;

      public Form1()
      {
         InitializeComponent();
            //const string bssFile = @"C:\Users\Lukas\OneDrive\AUIC\Gennemsnit\VS\NationalitetsParser\NationalitetsParser\bin\Debug\BSS lille.xlsx";
            //const string artsFile = @"C:\Users\Lukas\OneDrive\AUIC\Gennemsnit\VS\NationalitetsParser\NationalitetsParser\bin\Debug\ARTS lille.xlsx";
            //const string healthFile = @"C:\Users\Lukas\OneDrive\AUIC\Gennemsnit\VS\NationalitetsParser\NationalitetsParser\bin\Debug\HEALTH lille.xlsx";
            //const string stFile = @"C:\Users\Lukas\OneDrive\AUIC\Gennemsnit\VS\NationalitetsParser\NationalitetsParser\bin\Debug\ST lille.xlsx";
            //const string stFile = @"C:\Users\Lukas\OneDrive\AUIC\Gennemsnit\VS\NationalitetsParser\NationalitetsParser\bin\Debug\a.xlsx";
            string stFile = Path.GetFullPath("a.xlsx");

         mStudents = new Dictionary<string, Student>();

         //ReadStudentsFromFile(bssFile);
         //ReadStudentsFromFile(artsFile);
         //ReadStudentsFromFile(healthFile);
         ReadStudentsFromFile(stFile);

         OutputExcel();
      }

      private void ReadStudentsFromFile(string pExcelFile)
      {
         ExcelWrapper excelSheet = new ExcelWrapper(pExcelFile);

         excelSheet.SetActiveWorksheetByName("Sheet1");

         Range rows = excelSheet.GetRows();
         string SortText = excelSheet.GetTextByColumnHeader(rows[2], "CPR ");
         for (int i = 2; SortText != null && SortText.Equals("") != true; i++)
         {
            string studieNmr = excelSheet.GetTextByColumnHeader(rows[i], "STUDIENR");
            string name = excelSheet.GetTextByColumnHeader(rows[i], "FORNAVNE");
            if (name != null && studieNmr != null && studieNmr.Equals("0") != true)
            {
               if (!mStudents.ContainsKey(studieNmr))
               {
                  string cpr = excelSheet.GetTextByColumnHeader(rows[i], "CPR ");
                  string surname = excelSheet.GetTextByColumnHeader(rows[i], "EFTERNAVN");
                  string lineofstudy = excelSheet.GetTextByColumnHeader(rows[i], "UDDANNELSENSNAVN");
                  mStudents[studieNmr] = new Student(cpr, studieNmr, name, surname, lineofstudy);
               }
               float ects = excelSheet.GetFloatByColumnHeader(rows[i], "BELASTNING");
               string passed = excelSheet.GetTextByColumnHeader(rows[i], "GRUPPE_RES");
               int grade = excelSheet.GetIntByColumnHeader(rows[i], "KARAKTER");
               string semester = excelSheet.GetTextByColumnHeader(rows[i], "TERM_FORKORT");
               string course = excelSheet.GetTextByColumnHeader(rows[i], "AKTIVITET");
               string split = excelSheet.GetTextByColumnHeader(rows[i], "OVERS_FORTOLKET_RES2");
               if (grade == -1)
                  continue;
               int parseTest;
               Grade theGrade;
               //if (split != "bestået" && int.TryParse(split, out parseTest))
               //{
               //   int ects2 = excelSheet.GetIntByColumnHeader(rows[i + 1], "BELASTNINGSENHED");
               //   int ects3 = excelSheet.GetIntByColumnHeader(rows[i + 2], "BELASTNINGSENHED");
               //   theGrade = new Grade(course, semester, grade, ects2 + ects3, passed);
               //   i += 2;
               //}
               //else
                  theGrade = new Grade(course, semester, grade, ects, passed);
               mStudents[studieNmr].AddGrade(theGrade);
            }

            SortText = excelSheet.GetTextByColumnHeader(rows[i + 1], "CPR ");
         }
      }

      public void OutputExcel()
      {
         // alle 4 samlet, sorteret efter fornavn
         ExcelWrapper wrapper = ExcelWrapper.CreateNewWorksheet(new[] { 
            "Total", "CPR", "StudieNmr", "Fornavn", "Efternavn", 
            "Uddannelse", "Fag", "Semester", "r/o", "Karakter",
            "ECTS", "Vægt", "b/i", "ECTS i alt", "Samlet vægt", 
            "Gennemsnit" 
         });
         ExcelWrapper wrapper2 = ExcelWrapper.CreateNewWorksheet(new[] { 
            "StudieNmr", "Fornavn", "Efternavn", "Gennemsnit", "ECTS i alt"
         });
         int curRow = 2;
         List<Student> students = mStudents.Values.ToList();
         students.Sort((current, other) => current.Fornavn.CompareTo(other.Fornavn));
         foreach (Student student in students)
         {
            foreach (Grade grade in student.GetGrades())
            {
               OutputStudent(wrapper, curRow, student);
               OutputGrade(wrapper, curRow, grade);
               curRow++;
            }
            // Output Total
            OutputTotal(wrapper, curRow, student);
            curRow++;
         }
         curRow = 2;
         foreach (Student student in students)
         {
            OutputShortlistTotal(wrapper2, curRow, student);
            curRow++;
         }
      }

      private void OutputShortlistTotal(ExcelWrapper pWrapper, int pCurRow, Student pStudent)
      {
         Range row = pWrapper.GetRows()[pCurRow];
         float totalECTS = pStudent.GetTotalEcts();
         float totalWeight = pStudent.GetTotalWeight();
         pWrapper.SetCellsByColumnHeader(row, "StudieNmr", pStudent.StudyNmr);
         pWrapper.SetCellsByColumnHeader(row, "Fornavn", pStudent.Fornavn);
         pWrapper.SetCellsByColumnHeader(row, "Efternavn", pStudent.Efternavn);
         if (pStudent.GetGrades().Any())
            pWrapper.SetCellsByColumnHeader(row, "Gennemsnit", totalWeight / totalECTS);
         else
            pWrapper.SetCellsByColumnHeader(row, "Gennemsnit", "0");
        pWrapper.SetCellsByColumnHeader(row, "ECTS i alt", totalECTS.ToString());
        }

      private void OutputTotal(ExcelWrapper pWrapper, int pCurRow, Student pStudent)
      {
         Range row = pWrapper.GetRows()[pCurRow];
         float totalECTS = pStudent.GetTotalEcts();
         float totalWeight = pStudent.GetTotalWeight();
         pWrapper.SetCellsByColumnHeader(row, "Total", pStudent.Cpr + " Count");
         if (pStudent.GetGrades().Any())
         {
            pWrapper.SetCellsByColumnHeader(row, "ECTS i alt", totalECTS.ToString());
            pWrapper.SetCellsByColumnHeader(row, "Samlet vægt", totalWeight.ToString());
            pWrapper.SetCellsByColumnHeader(row, "Gennemsnit", totalWeight / totalECTS);
         }
         else
         {
            pWrapper.SetCellsByColumnHeader(row, "ECTS i alt", "0");
            pWrapper.SetCellsByColumnHeader(row, "Samlet vægt", "0");
            pWrapper.SetCellsByColumnHeader(row, "Gennemsnit", 0);
         }
      }

      private void OutputStudent(ExcelWrapper pWrapper, int pCurRow, Student pStudent)
      {
         Range row = pWrapper.GetRows()[pCurRow];
         pWrapper.SetCellsByColumnHeader(row, "CPR", pStudent.Cpr);
         pWrapper.SetCellsByColumnHeader(row, "StudieNmr", pStudent.StudyNmr);
         pWrapper.SetCellsByColumnHeader(row, "Fornavn", pStudent.Fornavn);
         pWrapper.SetCellsByColumnHeader(row, "Efternavn", pStudent.Efternavn);
         pWrapper.SetCellsByColumnHeader(row, "Uddannelse", pStudent.LineOfStudy);
      }

      private void OutputGrade(ExcelWrapper pWrapper, int pCurRow, Grade pGrade)
      {
         Range row = pWrapper.GetRows()[pCurRow];

         pWrapper.SetCellsByColumnHeader(row, "Semester", pGrade.Semester);
         pWrapper.SetCellsByColumnHeader(row, "Fag", pGrade.Course);
         pWrapper.SetCellsByColumnHeader(row, "Karakter", pGrade.Result);
         pWrapper.SetCellsByColumnHeader(row, "ECTS", pGrade.Ects);
         pWrapper.SetCellsByColumnHeader(row, "Vægt", pGrade.Weight);
         pWrapper.SetCellsByColumnHeader(row, "b/i", pGrade.Passed);
      }

      internal class Student
      {
         private List<Grade> mGrades;
         public string Cpr { get; private set; }
         public string StudyNmr { get; private set; }
         public string Fornavn { get; private set; }
         public string Efternavn { get; private set; }
         public string LineOfStudy { get; private set; }

         public Student(string pCpr, string pStudyNmr, string pName, string pSurname, string pLineOfStudy)
         {
            Cpr = pCpr;
            StudyNmr = pStudyNmr;
            Fornavn = pName;
            Efternavn = pSurname;
            LineOfStudy = pLineOfStudy;

            mGrades = new List<Grade>();
         }

         public float GetTotalEcts()
         {
            return mGrades.Sum(pGrade => float.Parse(pGrade.Ects));
         }

         public void AddGrade(Grade pGrade)
         {
            if (pGrade.Result.Equals("B") || pGrade.Result.Equals("IB"))
               return;
            if (mGrades.Any(grade => grade.Course.Equals(pGrade.Course)))
               return;
            mGrades.Add(pGrade);
         }

         public List<Grade> GetGrades()
         {
            return mGrades;
         }

         public float GetTotalWeight()
         {
            return mGrades.Sum(pGrade => float.Parse(pGrade.Weight));
         }
      }

      internal class Grade
      {
         public string Semester { get; private set; }
         public string Result { get; private set; }
         public string Ects { get; private set; }
         public string Weight { get; private set; }
         public string Passed { get; private set; }
         public string Course { get; private set; }

         public Grade(string pCourse, string pSemester, int pGrade, float pEcts, string pPassed)
         {
            Semester = pSemester;
            Result = pGrade.ToString();
            Ects = pEcts.ToString();
            Weight = (pGrade * pEcts).ToString();
            Passed = pPassed;
            Course = pCourse;
         }
      }
   }
}
