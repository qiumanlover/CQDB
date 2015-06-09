namespace CQDB
{
    using NPOI.HSSF.UserModel;
    using NPOI.SS.UserModel;
    using NPOI.XSSF.UserModel;
    using System;
    using System.CodeDom.Compiler;
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.Diagnostics;
    using System.IO;
    using System.Windows;
    using System.Windows.Forms;
    using System.Windows.Input;
    using System.Windows.Markup;
    using System.Windows.Media;

    public class MainWindow : Window, IComponentConnector
    {
        private bool _contentLoaded;
        internal System.Windows.Controls.Button btnGenerate;
        private bool isFileError = false;
        private int maxCount = 30;
        private string replaceStr = "※";
        private string replaceStr1 = ",";
        private List<Student> studentList = null;
        private Dictionary<string, int> stuFieldIdDic = null;
        private string targetRootPath = string.Empty;
        private string templateFilePath = string.Empty;
        internal System.Windows.Controls.TextBox txtSavePath;
        internal System.Windows.Controls.TextBox txtStudent;
        internal System.Windows.Controls.TextBox txtTemplate;

        public MainWindow()
        {
            this.InitializeComponent();
        }

        private void BeginGenerate(string filePath)
        {
            this.GetStudentFromExcel(filePath);
            this.studentList.Sort(new Comparison<Student>(MainWindow.CompareStudent));
            List<Student> stuList = null;
            int index = 0;
            int count = 0;
            if ((this.studentList != null) && (this.studentList.Count > 0))
            {
                count = 1;
                for (int i = 1; i < this.studentList.Count; i++)
                {
                    if (((this.studentList[i].Origin.Equals(this.studentList[i - 1].Origin) && this.studentList[i].CourseSpeciality.Equals(this.studentList[i - 1].CourseSpeciality)) && this.studentList[i].Layer.Equals(this.studentList[i - 1].Layer)) && this.studentList[i].Course.Equals(this.studentList[i - 1].Course))
                    {
                        count++;
                    }
                    else
                    {
                        stuList = this.studentList.GetRange(index, count);
                        this.GeneratQDG(stuList);
                        index = i;
                        count = 1;
                    }
                }
                stuList = this.studentList.GetRange(index, count);
                this.GeneratQDG(stuList);
            }
            this.studentList.Clear();
        }

        private void btnGenerate_Click(object sender, RoutedEventArgs e)
        {
            this.CheckFileFormat(new System.Windows.Controls.TextBox[] { this.txtTemplate });
            if (this.isFileError)
            {
                Note("请放入Excel格式文件！");
            }
            else
            {
                DateTime now = DateTime.Now;
                try
                {
                    this.InitContainer();
                    this.InitData();
                    this.BeginGenerate(this.txtStudent.Text.Trim());
                }
                catch (Exception exception)
                {
                    Note(exception.ToString());
                }
                this.CollectGarbage();
                DateTime time2 = DateTime.Now;
                TimeSpan span = (TimeSpan) (time2 - now);
                System.Windows.MessageBox.Show(string.Format("开始时间：{0}\r\n结束时间：{1}\r\n总耗时：{2:f1}", now, time2, span.TotalSeconds));
            }
        }

        private void CheckFileFormat(params System.Windows.Controls.TextBox[] tbs)
        {
            foreach (System.Windows.Controls.TextBox box in tbs)
            {
                if (!(string.IsNullOrEmpty(box.Text) || (!Path.GetExtension(box.Text).Equals(".xls") && !Path.GetExtension(box.Text).Equals(".xlsx"))))
                {
                    this.isFileError = false;
                    box.BorderBrush = Brushes.Black;
                    box.BorderThickness = new Thickness(0.5);
                }
                else
                {
                    this.isFileError = true;
                    box.BorderBrush = Brushes.Red;
                    box.BorderThickness = new Thickness(2.0);
                }
            }
        }

        private void CollectGarbage()
        {
            this.studentList.Clear();
            this.studentList.TrimExcess();
            GC.Collect();
        }

        private static int CompareStudent(Student s1, Student s2)
        {
            if (s1 == null)
            {
                if (s2 == null)
                {
                    return 0;
                }
                return -1;
            }
            if (s2 == null)
            {
                return 1;
            }
            if (s1.Origin.CompareTo(s2.Origin) == 0)
            {
                if (s1.CourseSpeciality.CompareTo(s2.CourseSpeciality) != 0)
                {
                    return s1.CourseSpeciality.CompareTo(s2.CourseSpeciality);
                }
                if (s1.Layer.CompareTo(s2.Layer) != 0)
                {
                    return s1.Layer.CompareTo(s2.Layer);
                }
                if (s1.Course.CompareTo(s2.Course) == 0)
                {
                    if (s1.Speciality.CompareTo(s2.Speciality) == 0)
                    {
                        return s1.Id.CompareTo(s2.Id);
                    }
                    return s1.Speciality.CompareTo(s2.Speciality);
                }
                return s1.Course.CompareTo(s2.Course);
            }
            return s1.Origin.CompareTo(s2.Origin);
        }

        private void GeneratQDG(List<Student> stuList)
        {
            Student student = stuList[0];
            string path = Path.Combine(this.targetRootPath, student.Origin, student.CourseSpeciality, student.Layer);
            string course = student.Course;
            if (course.Contains("*"))
            {
                course = course.Replace("*", this.replaceStr);
            }
            if (course.Contains("/"))
            {
                course = course.Replace("/", this.replaceStr1);
            }
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
            List<Student> list2 = null;
            int count = 0;
            char ch = '1';
            bool flag = false;
            list2 = stuList;
            while (true)
            {
                count = (list2.Count <= this.maxCount) ? list2.Count : this.maxCount;
                List<Student> range = list2.GetRange(0, count);
                list2.RemoveRange(0, count);
                string filePath = string.Empty;
                if (!flag)
                {
                    filePath = Path.Combine(path, course + ".xls");
                }
                else
                {
                    filePath = Path.Combine(path, course + "___" + ((char) (ch + '\x0001')).ToString() + ".xls");
                }
                this.SaveToExcelFile(filePath, course, range);
                count = (list2.Count <= this.maxCount) ? list2.Count : this.maxCount;
                if (count > 0)
                {
                    flag = true;
                }
                else
                {
                    return;
                }
            }
        }

        private static IWorkbook GetFileWorkbook(string filePath)
        {
            FileStream stream = null;
            try
            {
                using (stream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                {
                    char ch = filePath[filePath.Length - 1];
                    if (ch.Equals('x'))
                    {
                        return new XSSFWorkbook(stream);
                    }
                    return new HSSFWorkbook(stream);
                }
            }
            catch (Exception exception)
            {
                throw new IOException(string.Format("文件{0}无法打开，请检查该文件是否被其他程序锁定！\r\n{1}", filePath, exception.ToString()));
            }
            finally
            {
                if (stream != null)
                {
                    stream.Dispose();
                }
            }
        }

        private void GetStudentFromExcel(string filePath)
        {
            int num2;
            IWorkbook fileWorkbook = GetFileWorkbook(filePath);
            ISheet sheetAt = null;
            IRow row = null;
            if ((fileWorkbook == null) || (fileWorkbook.NumberOfSheets <= 0))
            {
                throw new Exception("学籍表文件错误！");
            }
            sheetAt = fileWorkbook.GetSheetAt(0);
            IRow row2 = sheetAt.GetRow(0);
            int lastCellNum = row2.LastCellNum;
            for (num2 = row2.FirstCellNum; num2 < lastCellNum; num2++)
            {
                if ((row2.GetCell(num2) == null) | string.IsNullOrEmpty(row2.GetCell(num2).StringCellValue.Trim()))
                {
                    break;
                }
                this.stuFieldIdDic[row2.GetCell(num2).StringCellValue.Trim()] = num2;
            }
            int lastRowNum = sheetAt.LastRowNum;
            for (num2 = sheetAt.FirstRowNum + 1; num2 <= lastRowNum; num2++)
            {
                row = sheetAt.GetRow(num2);
                if (row != null)
                {
                    Student item = new Student {
                        Id = row.GetCell(this.stuFieldIdDic["学号"]).StringCellValue,
                        Name = row.GetCell(this.stuFieldIdDic["姓名"]).StringCellValue,
                        Course = row.GetCell(this.stuFieldIdDic["课程名称"]).StringCellValue,
                        Layer = row.GetCell(this.stuFieldIdDic["层次"]).StringCellValue,
                        Speciality = row.GetCell(this.stuFieldIdDic["专业"]).StringCellValue,
                        CourseSpeciality = row.GetCell(this.stuFieldIdDic["课程性质"]).StringCellValue,
                        Origin = row.GetCell(this.stuFieldIdDic["来源"]).StringCellValue
                    };
                    this.studentList.Add(item);
                }
            }
            this.stuFieldIdDic.Clear();
        }

        private IWorkbook GetTemplate(string filePath)
        {
            return GetFileWorkbook(filePath);
        }

        private void InitContainer()
        {
            this.stuFieldIdDic = new Dictionary<string, int>();
            this.studentList = new List<Student>();
        }

        private void InitData()
        {
            this.targetRootPath = this.txtSavePath.Text.Trim();
            this.templateFilePath = this.txtTemplate.Text.Trim();
        }

        [GeneratedCode("PresentationBuildTasks", "4.0.0.0"), DebuggerNonUserCode]
        public void InitializeComponent()
        {
            if (!this._contentLoaded)
            {
                this._contentLoaded = true;
                Uri resourceLocator = new Uri("/CQDB;component/mainwindow.xaml", UriKind.Relative);
                System.Windows.Application.LoadComponent(this, resourceLocator);
            }
        }

        private static void Note(string msg)
        {
            System.Windows.MessageBox.Show(msg, "提示");
        }

        private void SaveToExcelFile(string filePath, string course, List<Student> stuList)
        {
            IWorkbook template = this.GetTemplate(this.txtTemplate.Text);
            ISheet sheetAt = template.GetSheetAt(0);
            IRow row = null;
            string stringCellValue = string.Empty;
            row = sheetAt.GetRow(2);
            stringCellValue = row.Cells[0].StringCellValue;
            stringCellValue = string.Format("{0} {1}", stringCellValue, course);
            row.Cells[0].SetCellValue(stringCellValue);
            for (int i = 0; i < stuList.Count; i++)
            {
                row = sheetAt.GetRow(i + 4);
                int num2 = i + 1;
                row.Cells[0].SetCellValue(num2.ToString());
                row.Cells[1].SetCellValue(stuList[i].Speciality);
                row.Cells[2].SetCellValue(stuList[i].Layer);
                row.Cells[3].SetCellValue(stuList[i].Id);
                row.Cells[4].SetCellValue(stuList[i].Name);
            }
            this.WriteIO(filePath, template);
        }

        [EditorBrowsable(EditorBrowsableState.Never), DebuggerNonUserCode, GeneratedCode("PresentationBuildTasks", "4.0.0.0")]
        void IComponentConnector.Connect(int connectionId, object target)
        {
            switch (connectionId)
            {
                case 1:
                    this.txtStudent = (System.Windows.Controls.TextBox) target;
                    this.txtStudent.PreviewDragEnter += new System.Windows.DragEventHandler(this.txtFile_PreviewDragEnter);
                    this.txtStudent.PreviewDragOver += new System.Windows.DragEventHandler(this.txtFile_PreviewDragEnter);
                    this.txtStudent.PreviewDrop += new System.Windows.DragEventHandler(this.txtFile_PreviewDrop);
                    this.txtStudent.MouseDoubleClick += new MouseButtonEventHandler(this.txtFile_MouseDoubleClick);
                    break;

                case 2:
                    this.txtTemplate = (System.Windows.Controls.TextBox) target;
                    this.txtTemplate.PreviewDragEnter += new System.Windows.DragEventHandler(this.txtFile_PreviewDragEnter);
                    this.txtTemplate.PreviewDragOver += new System.Windows.DragEventHandler(this.txtFile_PreviewDragEnter);
                    this.txtTemplate.PreviewDrop += new System.Windows.DragEventHandler(this.txtFile_PreviewDrop);
                    this.txtTemplate.MouseDoubleClick += new MouseButtonEventHandler(this.txtFile_MouseDoubleClick);
                    break;

                case 3:
                    this.txtSavePath = (System.Windows.Controls.TextBox) target;
                    this.txtSavePath.PreviewDragEnter += new System.Windows.DragEventHandler(this.txtPath_PreviewDragOver);
                    this.txtSavePath.PreviewDragOver += new System.Windows.DragEventHandler(this.txtPath_PreviewDragOver);
                    this.txtSavePath.PreviewDrop += new System.Windows.DragEventHandler(this.txtPath_PreviewDrop);
                    this.txtSavePath.MouseDoubleClick += new MouseButtonEventHandler(this.txtPath_MouseDoubleClick);
                    break;

                case 4:
                    this.btnGenerate = (System.Windows.Controls.Button) target;
                    this.btnGenerate.Click += new RoutedEventHandler(this.btnGenerate_Click);
                    break;
                case 0:
                    ((MainWindow)target).KeyUp += new System.Windows.Input.KeyEventHandler(this.Window_KeyUp);
                    return;
                default:
                    this._contentLoaded = true;
                    break;
            }
        }

        private void txtFile_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog dialog = new Microsoft.Win32.OpenFileDialog {
                Filter = "Excel文件|*.xls;*.xlsx|所有文件|*.*"
            };
            if (dialog.ShowDialog() == true)
            {
                (sender as System.Windows.Controls.TextBox).Text = dialog.FileName;
            }
        }

        private void txtFile_PreviewDragEnter(object sender, System.Windows.DragEventArgs e)
        {
            e.Effects = System.Windows.DragDropEffects.Copy;
            e.Handled = true;
        }

        private void txtFile_PreviewDrop(object sender, System.Windows.DragEventArgs e)
        {
            object data = e.Data.GetData(System.Windows.DataFormats.FileDrop);
            System.Windows.Controls.TextBox box = sender as System.Windows.Controls.TextBox;
            if (box != null)
            {
                box.Text = string.Format("{0}", ((string[]) data)[0]);
            }
            this.CheckFileFormat(new System.Windows.Controls.TextBox[] { box });
        }

        private void txtPath_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            FolderBrowserDialog dialog = new FolderBrowserDialog();
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                (sender as System.Windows.Controls.TextBox).Text = dialog.SelectedPath;
            }
        }

        private void txtPath_PreviewDragOver(object sender, System.Windows.DragEventArgs e)
        {
            e.Effects = System.Windows.DragDropEffects.Copy;
            e.Handled = true;
        }

        private void txtPath_PreviewDrop(object sender, System.Windows.DragEventArgs e)
        {
            object data = e.Data.GetData(System.Windows.DataFormats.FileDrop);
            System.Windows.Controls.TextBox box = sender as System.Windows.Controls.TextBox;
            if (box != null)
            {
                box.Text = string.Format("{0}", ((string[]) data)[0]);
            }
        }

        private void Window_KeyUp(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key == Key.F1)
            {
                System.Windows.MessageBox.Show("注意事项：\r\n1、课程名称不能包含除*和/两个符号以外的无法成为文件名的字符\r\n2、课程名中所有*号，在签到表名称中将被替换为※，符号/将被替换为 ,\r\n3、源文件标题必须包含以下项：\r\n    学号、姓名、来源、层次、专业、课程名称、课程性质\r\n4、所有单元格不能是数字格式的样式");
            }
        }

        private void WriteIO(string filePath, IWorkbook workbook)
        {
            FileStream stream = null;
            try
            {
                using (stream = new FileStream(filePath, FileMode.Create, FileAccess.Write))
                {
                    workbook.Write(stream);
                }
            }
            catch (Exception exception)
            {
                throw new IOException("在路径" + Path.GetDirectoryName(filePath) + "无法创建新文件，请检查该路径的访问权限！\r\n" + exception.ToString());
            }
            finally
            {
                if (stream != null)
                {
                    stream.Dispose();
                }
            }
        }
    }
}

