using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
using ClosedXML.Excel;

namespace الحفلة
{
    public partial class Form1 : Form
    {   
        private void October_CheckedChanged(object sender, EventArgs e)
        {

        }
        private void button2_Click(object sender, EventArgs e)
        {

        }
        private void Form1_Load(object sender, EventArgs e)
        {

        }
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
        private void label13_Click(object sender, EventArgs e)
        {

        }
        private void label14_Click(object sender, EventArgs e)
        {

        }

        public Form1()
        {
            this.AutoScaleMode = AutoScaleMode.None;
            InitializeComponent();
            comboBox1.Items.Add("كلي");
            comboBox1.Items.Add("حلوان");
            comboBox6.Items.Add("حلون");
            comboBox1.Items.Add("مصدق");
            comboBox6.Items.Add("مصدق");
            comboBox1.Items.Add("مدينة نصر");
            comboBox6.Items.Add("مدينة نصر");
            comboBox1.Items.Add("مصر الجديدة");
            comboBox6.Items.Add("مصر الجديدة");
            comboBox1.Items.Add("فيصل");
            comboBox6.Items.Add("فيصل");
            comboBox1.Items.Add("المعادي");
            comboBox6.Items.Add("المعادي");
            comboBox1.Items.Add("اسكندرية");
            comboBox6.Items.Add("اسكندرية");
            comboBox1.Items.Add("اكتوبر");
            comboBox6.Items.Add("اكتوبر");

            string[] items = new string[]
            {
                "كانتين","مدير فرع","ادراة فروع","اتش ار متطوعين",
                "موظف","براعم","استكشاف","عيني",
                "اتش ار تدريب","دعايا وميديا","اداره مركزية","متدربين","مرافقين","مدير تنفيذي"
            };

            // إضافة العناصر بدون تكرار
            foreach (var item in items.Distinct())
            {
                comboBox2.Items.Add(item);
                comboBox7.Items.Add(item);
            }


            comboBox3.Items.Add("حلوان");
            comboBox3.Items.Add("مصدق");
            comboBox3.Items.Add("مدينة نصر");
            comboBox3.Items.Add("مصر الجديدة");
            comboBox3.Items.Add("فيصل");
            comboBox3.Items.Add("المعادي");
            comboBox3.Items.Add("اسكندرية");    
            comboBox3.Items.Add("اكتوبر");
            comboBox4.Items.Add("حضور");
            comboBox4.Items.Add("غياب");
            comboBox5.Items.Add("حضور");
            comboBox5.Items.Add("غياب");
        }

        private static readonly string basePath = AppDomain.CurrentDomain.BaseDirectory;
        private static readonly string excelFile = Path.Combine(basePath, "Attendance.xlsx");

        private void EnsureExcelFile()
        {
            if (!File.Exists(excelFile))
            {
                using (var wb = new XLWorkbook())
                {
                    var ws = wb.AddWorksheet("Attendance");
                    ws.Cell(1, 1).Value = "الاسم";
                    ws.Cell(1, 2).Value = "الفرع";
                    ws.Cell(1, 3).Value = "التصنيف";
                    ws.Cell(1, 4).Value = "الحالة";
                    ws.Cell(1, 5).Value = "الكود";
                    wb.SaveAs(excelFile);
                }
            }
        }

        private void AddToExcel(string name, string branch, string type, string status , string code)
        {
            EnsureExcelFile();

            using (var wb = new XLWorkbook(excelFile))
            {
                var ws = wb.Worksheet("Attendance");
                int lastRow = ws.LastRowUsed()?.RowNumber() ?? 1;
                ws.Cell(lastRow + 1, 1).Value = name;
                ws.Cell(lastRow + 1, 2).Value = branch;
                ws.Cell(lastRow + 1, 3).Value = type;
                ws.Cell(lastRow + 1, 4).Value = status;
                ws.Cell(lastRow + 1, 4).Value = code;

                wb.Save();
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            string branch = comboBox1.SelectedItem?.ToString()?.Trim();
            if (string.IsNullOrWhiteSpace(branch))
            {
                MessageBox.Show("من فضلك اختر الفرع");
                return;
            }

            listBox1.Items.Clear();

            string excelFile = Path.Combine(basePath, "Attendance.xlsx");
            if (!File.Exists(excelFile))
            {
                MessageBox.Show("ملف Excel غير موجود");
                return;
            }

            using (var wb = new XLWorkbook(excelFile))
            {
                var ws = wb.Worksheet("Attendance");

                // تحديد إذا كان هيعرض حضور فقط أو غياب فقط أو الاتنين
                bool getAttendance = checkBox1.Checked || (!checkBox1.Checked && !checkBox2.Checked);
                bool getAbsence = checkBox2.Checked || (!checkBox1.Checked && !checkBox2.Checked);

                foreach (var row in ws.RowsUsed().Skip(1))
                {
                    string name = row.Cell(1).GetString().Trim();
                    string rowBranch = row.Cell(2).GetString().Trim();
                    string type = row.Cell(3).GetString().Trim();
                    string status = row.Cell(4).GetString().Trim();

                    // تصفية حسب الفرع
                    if (branch != "كلي" && rowBranch != branch) continue;

                    // تصفية حسب الحضور/الغياب
                    if ((status == "حضور" && !getAttendance) || (status == "غياب" && !getAbsence)) continue;

                    listBox1.Items.Add($"{name}");
                }
            }

            if (listBox1.Items.Count == 0)
            {
                MessageBox.Show("لا توجد بيانات لعرضها");
            }

            textBox7.Text = listBox1.Items.Count.ToString();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (listBox1.SelectedItem == null)
            {
                MessageBox.Show("من فضلك اختر اسم من القائمة أولاً");
                return;
            }

            // اسم السطر المختار في الليست بوكس (بنخليها الاسم فقط)
            string selectedLine = listBox1.SelectedItem.ToString();
            // مسار ملف Excel
            string excelFile = Path.Combine(basePath, "Attendance.xlsx");
            if (!File.Exists(excelFile))
            {
                MessageBox.Show("ملف Excel غير موجود");
                return;
            }

            using (var wb = new XLWorkbook(excelFile))
            {
                var ws = wb.Worksheet("Attendance");

                // البحث عن الصف اللي فيه الاسم
                var row = ws.RowsUsed().FirstOrDefault(r => r.Cell(1).GetString().Trim() == selectedLine);

                if (row != null)
                {
                    string fullName = row.Cell(1).GetString();
                    string branch = row.Cell(2).GetString();
                    string personType = row.Cell(3).GetString();
                    string status = row.Cell(4).GetString();
                    string code = row.Cell(5).GetString();

                    // عرض البيانات في الأدوات
                    textBox1.Text = fullName;
                    comboBox2.SelectedItem = personType;
                    comboBox3.SelectedItem = branch;
                    comboBox4.SelectedItem = status;
                    textBox4.Text = code;
                }
                else
                {
                    MessageBox.Show("لم يتم العثور على البيانات في ملف Excel");
                }
            }
        }


        private void button2_Click_1(object sender, EventArgs e)
        {
            if (listBox1.SelectedItem == null)
            {
                MessageBox.Show("من فضلك اختر اسم من القائمة أولاً");
                return;
            }

            // الاسم المختار من الليست بوكس (من غير الحالة)
            string selectedLine = listBox1.SelectedItem.ToString();
            
            

            // القيم الجديدة من الأدوات
            string newName = textBox1.Text.Trim();
            string newType = comboBox2.SelectedItem?.ToString() ?? "";
            string newBranch = comboBox3.SelectedItem?.ToString() ?? "";
            string newStatus = comboBox4.SelectedItem?.ToString() ?? "";
            string code = textBox4.Text.Trim();

            if (string.IsNullOrWhiteSpace(newName) ||
                string.IsNullOrWhiteSpace(newType) ||
                string.IsNullOrWhiteSpace(newBranch) ||
                string.IsNullOrWhiteSpace(newStatus))
            {
                MessageBox.Show("من فضلك تأكد من إدخال جميع البيانات");
                return;
            }

            string excelFile = Path.Combine(basePath, "Attendance.xlsx");
            if (!File.Exists(excelFile))
            {
                MessageBox.Show("ملف Excel غير موجود");
                return;
            }

            using (var wb = new XLWorkbook(excelFile))
            {
                var ws = wb.Worksheet("Attendance");

                // البحث عن الصف اللي فيه الاسم القديم
                var row = ws.RowsUsed().FirstOrDefault(r => r.Cell(1).GetString().Trim() == selectedLine);

                if (row != null)
                {
                    // تعديل البيانات مباشرة في Excel
                    row.Cell(1).Value = newName;
                    row.Cell(2).Value = newBranch;
                    row.Cell(3).Value = newType;
                    row.Cell(4).Value = newStatus;
                    row.Cell(5).Value = code;

                    wb.Save();
                    MessageBox.Show("تم تعديل البيانات بنجاح");
                }
                else
                {
                    MessageBox.Show("لم يتم العثور على الاسم في ملف Excel");
                }
            }

            // إعادة تحميل البيانات في الليست بوكس
            LoadListFromExcel();
            button3_Click(sender, e);
        }
        private void LoadListFromExcel()
        {
            listBox1.Items.Clear();
            EnsureExcelFile();

            using (var wb = new XLWorkbook(excelFile))
            {
                var ws = wb.Worksheet("Attendance");
                foreach (var row in ws.RowsUsed().Skip(1))
                {
                    string name = row.Cell(1).GetString();
                    string branch = row.Cell(2).GetString();
                    string type = row.Cell(3).GetString();
                    string status = row.Cell(4).GetString();
                    string code = row.Cell(5).GetString();
                    listBox1.Items.Add($"{name}");
                }
            }
        }

        private void RemoveLine(string filePath, string name)
        {
            if (!File.Exists(filePath)) return;
            var lines = File.ReadAllLines(filePath).ToList();
            lines.RemoveAll(l => l.Trim().StartsWith(name));
            File.WriteAllLines(filePath, lines);
        }

        private void AppendLine(string filePath, string line)
        {
            File.AppendAllText(filePath, line + Environment.NewLine, Encoding.UTF8);
        }

        private void button4_Click(object sender, EventArgs e)
        {
            // ------------------- قراءة البيانات الجديدة -------------------
            string newName = textBox1.Text.Trim();
            string newType = comboBox2.SelectedItem?.ToString() ?? "";
            string newBranch = comboBox3.SelectedItem?.ToString() ?? "";
            string newStatus = comboBox4.SelectedItem?.ToString() ?? "";
            string code = textBox4.Text.Trim();

            if (string.IsNullOrWhiteSpace(newName) ||
                string.IsNullOrWhiteSpace(newType) ||
                string.IsNullOrWhiteSpace(newBranch) ||
                string.IsNullOrWhiteSpace(newStatus))
            {
                MessageBox.Show("من فضلك تأكد من إدخال جميع البيانات");
                return;
            }

            // ------------------- إضافة البيانات للـ Excel -------------------
            string excelFile = Path.Combine(basePath, "Attendance.xlsx");

            // لو الملف مش موجود، ننشئه ونضيف الأعمدة
            if (!File.Exists(excelFile))
            {
                using (var wb = new XLWorkbook())
                {
                    var ws = wb.AddWorksheet("Attendance");
                    ws.Cell(1, 1).Value = "الاسم";
                    ws.Cell(1, 2).Value = "الفرع";
                    ws.Cell(1, 3).Value = "التصنيف";
                    ws.Cell(1, 4).Value = "الحالة";
                    ws.Cell(1, 5).Value = "الكود";
                    wb.SaveAs(excelFile);
                }
            }

            // فتح الملف وإضافة البيانات
            using (var wb = new XLWorkbook(excelFile))
            {
                var ws = wb.Worksheet("Attendance");
                int lastRow = ws.LastRowUsed()?.RowNumber() ?? 1;
                ws.Cell(lastRow + 1, 1).Value = newName;
                ws.Cell(lastRow + 1, 2).Value = newBranch;
                ws.Cell(lastRow + 1, 3).Value = newType;
                ws.Cell(lastRow + 1, 4).Value = newStatus;
                ws.Cell(lastRow + 1, 5).Value = code;

                wb.Save();
            }

            MessageBox.Show("تم إضافة البيانات بنجاح");

            // إعادة تحميل البيانات في الليست بوكس
            LoadListFromExcel();
            button3_Click(sender, e);
        }
       

        private void button5_Click(object sender, EventArgs e)
        {
            if (listBox1.SelectedItem == null)
            {
                MessageBox.Show("من فضلك اختر اسم من القائمة أولاً");
                return;
            }

            // ------------------- قراءة البيانات القديمة -------------------
            string oldLine = listBox1.SelectedItem.ToString();

            // ------------------- حذف من شيت Excel -------------------
            string excelFile = Path.Combine(basePath, "Attendance.xlsx");
            if (!File.Exists(excelFile))
            {
                MessageBox.Show("ملف Excel غير موجود");
                return;
            }

            using (var wb = new XLWorkbook(excelFile))
            {
                var ws = wb.Worksheet("Attendance");

                // نبحث عن الصف اللي فيه الاسم
                var row = ws.RowsUsed().FirstOrDefault(r => r.Cell(1).GetString().Trim() == oldLine);
                if (row != null)
                {
                    row.Delete();
                    wb.Save();
                    MessageBox.Show("تم حذف البيانات بنجاح");
                }
                else
                {
                    MessageBox.Show("لم يتم العثور على الاسم في ملف Excel");
                }
            }

            // ------------------- إعادة تحميل البيانات في الليست بوكس -------------------
            LoadListFromExcel();
            button3_Click(sender, e);
        }

        private void button6_Click(object sender, EventArgs e)
        {
            string attendanceFile = Path.Combine(basePath, "Attendance.xlsx");
            string absenceFile = Path.Combine(basePath, "Absence.xlsx");

            // خريطة تربط اسم الفرع باسم الملف
            Dictionary<string, string> branchFiles = new Dictionary<string, string>()
            {
                { "اكتوبر", Path.Combine(basePath, "اكتوبر.xlsx") },
                { "حلوان", Path.Combine(basePath, "حلوان.xlsx") },
                { "المعادي", Path.Combine(basePath, "المعادي.xlsx") },
                { "اسكندرية", Path.Combine(basePath, "اسكندرية.xlsx") },
                { "فيصل", Path.Combine(basePath, "فيصل.xlsx") },
                { "مدينة نصر", Path.Combine(basePath, "مدينة نصر.xlsx") },
                { "مصر الجديدة", Path.Combine(basePath, "مصر الجديدة.xlsx") },
                { "مصدق", Path.Combine(basePath, "مصدق.xlsx") },
            };

            if (!File.Exists(attendanceFile))
            {
                MessageBox.Show("ملف Attendance.xlsx غير موجود");
                return;
            }

            using (var wb = new XLWorkbook(attendanceFile))
            {
                var ws = wb.Worksheet("Attendance");

                var allRows = ws.RowsUsed().Skip(1) // تخطي العنوان
                                .Select(r => new
                                {
                                    Name = r.Cell(1).GetString(),
                                    Branch = r.Cell(2).GetString(),
                                    Type = r.Cell(3).GetString(),
                                    Status = r.Cell(4).GetString(),
                                    Code = r.Cell(5).GetString()
                                }).ToList();

                // ===== حفظ الغياب =====
                var absences = allRows.Where(r => r.Status.Trim() == "غياب").ToList();
                SaveToExcel(absenceFile, absences);

                // ===== حفظ حسب الفروع =====
                foreach (var branch in branchFiles.Keys)
                {
                    var branchRows = allRows.Where(r => r.Branch.Trim() == branch).ToList();
                    SaveToExcel(branchFiles[branch], branchRows);
                }
            }

            MessageBox.Show("تم استخراج الملفات بنجاح ✅");
        }

        private void SaveToExcel(string filePath, IEnumerable<dynamic> rows)
        {
            using (var newWb = new XLWorkbook())
            {
                var newWs = newWb.AddWorksheet("Sheet1");
                // عنوان الأعمدة
                newWs.Cell(1, 1).Value = "الاسم";
                newWs.Cell(1, 2).Value = "الفرع";
                newWs.Cell(1, 3).Value = "النوع";
                newWs.Cell(1, 4).Value = "الحالة";
                newWs.Cell(1, 5).Value = "الكود";

                int row = 2;
                foreach (var r in rows)
                {
                    newWs.Cell(row, 1).Value = r.Name;
                    newWs.Cell(row, 2).Value = r.Branch;
                    newWs.Cell(row, 3).Value = r.Type;
                    newWs.Cell(row, 4).Value = r.Status;
                    newWs.Cell(row, 5).Value = r.Code;
                    row++;
                }

                newWb.SaveAs(filePath);
            }
        }


        private void button7_Click(object sender, EventArgs e)
        {
            string searchName = textBox2.Text.Trim();
            if (string.IsNullOrWhiteSpace(searchName))
            {
                MessageBox.Show("من فضلك أدخل الاسم");
                return;
            }

            string excelFile = Path.Combine(basePath, "Attendance.xlsx");
            if (!File.Exists(excelFile))
            {
                MessageBox.Show("ملف Attendance.xlsx غير موجود");
                return;
            }

            using (var wb = new XLWorkbook(excelFile))
            {
                var ws = wb.Worksheet("Attendance");
                var row = ws.RowsUsed()
                            .Skip(1)
                            .FirstOrDefault(r => r.Cell(1).GetString().Trim() == searchName);

                if (row != null)
                {
                    textBox6.Text = row.Cell(1).GetString();      // الاسم
                    comboBox6.SelectedItem = row.Cell(2).GetString(); // الفرع
                    comboBox7.SelectedItem = row.Cell(3).GetString(); // التصنيف
                    comboBox5.SelectedItem = row.Cell(4).GetString(); // الحالة
                    textBox5.Text = row.Cell(5).GetString();
                    textBox3.Text = row.Cell(5).GetString();
                }
                else
                {
                    MessageBox.Show("الاسم غير موجود");
                }
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            string codeText = textBox3.Text.Trim();

            string excelFile = Path.Combine(basePath, "Attendance.xlsx");
            if (!File.Exists(excelFile))
            {
                MessageBox.Show("ملف Attendance.xlsx غير موجود");
                return;
            }

            using (var wb = new XLWorkbook(excelFile))
            {
                var ws = wb.Worksheet("Attendance");
                var row = ws.RowsUsed()
                            .Skip(1)
                            .Skip(2)
                            .Skip(3)
                            .Skip(4)
                            .FirstOrDefault(r => r.Cell(5).GetString() == codeText);

                if (row != null)
                {
                    textBox6.Text = row.Cell(1).GetString();      // الاسم
                    textBox2.Text = row.Cell(1).GetString();
                    comboBox6.SelectedItem = row.Cell(2).GetString(); // الفرع
                    comboBox7.SelectedItem = row.Cell(3).GetString(); // التصنيف
                    comboBox5.SelectedItem = row.Cell(4).GetString(); // الحالة
                    textBox5.Text = row.Cell(5).GetString();
                }
                else
                {
                    MessageBox.Show("الكود غير موجود");
                }
            }
        }

        private void button9_Click(object sender, EventArgs e)
        {
            string oldName = textBox2.Text.Trim();
            // القيم الجديدة من الأدوات
            string newName = textBox6.Text.Trim();
            string newType = comboBox7.SelectedItem?.ToString() ?? "";
            string newBranch = comboBox6.SelectedItem?.ToString() ?? "";
            string newStatus = comboBox5.SelectedItem?.ToString() ?? "";
            string code = textBox5.Text.Trim();

            if (string.IsNullOrWhiteSpace(newName) ||
                string.IsNullOrWhiteSpace(newType) ||
                string.IsNullOrWhiteSpace(newBranch) ||
                string.IsNullOrWhiteSpace(newStatus))
            {
                MessageBox.Show("من فضلك تأكد من إدخال جميع البيانات");
                return;
            }

            string excelFile = Path.Combine(basePath, "Attendance.xlsx");
            if (!File.Exists(excelFile))
            {
                MessageBox.Show("ملف Excel غير موجود");
                return;
            }

            using (var wb = new XLWorkbook(excelFile))
            {
                var ws = wb.Worksheet("Attendance");
                // البحث عن الصف اللي فيه الاسم القديم
                var row = ws.RowsUsed().FirstOrDefault(r => r.Cell(1).GetString().Trim() == oldName);
                if (row != null)
                {
                    // تعديل البيانات مباشرة في Excel
                    row.Cell(1).Value = newName;
                    row.Cell(2).Value = newBranch;
                    row.Cell(3).Value = newType;
                    row.Cell(4).Value = newStatus;
                    row.Cell(5).Value = code;

                    wb.Save();
                    MessageBox.Show("تم تعديل البيانات بنجاح");
                }
                else
                {
                    MessageBox.Show("لم يتم العثور على الاسم في ملف Excel");
                }
            }
            // إعادة تحميل البيانات في الليست بوكس
            LoadListFromExcel();
            button3_Click(sender, e);
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {

        }
    }
}