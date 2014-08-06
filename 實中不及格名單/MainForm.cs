using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml;

using Aspose.Words;
using Aspose.Words.Drawing;
using K12.Data;
using Campus.Report;

namespace JH.IBSH.FailPeriodicalExamList
{
    public partial class MainForm : FISCA.Presentation.Controls.BaseForm
    {
        private BackgroundWorker _bgw = new BackgroundWorker();
        private const string configList = "JH.IBSH.FailPeriodicalExamList.Config.1.List";
        private const string configSingle = "JH.IBSH.FailPeriodicalExamList.Config.1.Single";
        private string reportName = "";
        public static ReportConfiguration ReportConfigurationList = new Campus.Report.ReportConfiguration(configList);
        public static ReportConfiguration ReportConfigurationSingle = new Campus.Report.ReportConfiguration(configSingle);
        private mainFormConfig config;
        public class mainFormConfig {
            public mainFormFrom from;
            public bool hasList;
            public bool hasSingle;
            public string formName;
        }
        public enum mainFormFrom { 
            Student ,
            Class 
        }
        class filter
        {
            public SchoolYearSemester sys;
            public int exam;
            public string examText;
        }
        public static CourseGradeB.Tool.SubjectType StringToSubjectType(string type)
        {
            switch (type)
            {
                case "Honor":
                    return CourseGradeB.Tool.SubjectType.Honor;
                case "Regular":
                    return CourseGradeB.Tool.SubjectType.Regular;
                default:
                    return CourseGradeB.Tool.SubjectType.Regular;
            }
        }
        public MainForm(mainFormConfig config)
        {
            InitializeComponent();
            #region 設定comboBox選單
            foreach (int item in Enumerable.Range(int.Parse(School.DefaultSchoolYear) - 11, 13))
            {
                comboBoxEx2.Items.Add(item);
            }
            foreach (string item in new string[] { "1", "2" })
            {
                comboBoxEx3.Items.Add(item);
            }
            foreach (string item in new string[] { "Midterm", "Final" })
            {
                comboBoxEx4.Items.Add(item);
            }
            #endregion
            this.reportName = config.formName;
            this.Text = config.formName;
            comboBoxEx2.Text = School.DefaultSchoolYear;
            comboBoxEx3.Text = School.DefaultSemester;
            comboBoxEx4.SelectedIndex = 0;
            this.config = config;
            _bgw.DoWork += new DoWorkEventHandler(_bgw_DoWork);
            _bgw.RunWorkerCompleted += new RunWorkerCompletedEventHandler(_bgw_RunWorkerCompleted);
        }
        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            new Config().ShowDialog();
        }
        private void btnPrint_Click(object sender, EventArgs e)
        {
            if (config.from == mainFormFrom.Class && K12.Presentation.NLDPanels.Class.SelectedSource.Count < 1)
            {
                FISCA.Presentation.Controls.MsgBox.Show("請先選擇班級");
                return;
            }
            else if (config.from == mainFormFrom.Student && K12.Presentation.NLDPanels.Student.SelectedSource.Count < 1)
            {
                FISCA.Presentation.Controls.MsgBox.Show("請先選擇學生");
                return;
            }
            btnPrint.Enabled = false;
            _bgw.RunWorkerAsync(new filter
                {
                    sys = new SchoolYearSemester(int.Parse(comboBoxEx2.Text), int.Parse(comboBoxEx3.Text)),
                    exam = comboBoxEx4.Text == "Midterm" ? 1 : 2,
                    examText = comboBoxEx4.Text,
                }
            );
        }
        void _bgw_DoWork(object sender, DoWorkEventArgs e)
        {
            filter f = (filter)e.Argument;
            Document document = new Document();
            Document templateList = (ReportConfigurationList.Template != null)
                 ? ReportConfigurationList.Template.ToDocument()
                 : new Campus.Report.ReportTemplate(Properties.Resources.名單樣版, Campus.Report.TemplateType.Word).ToDocument();
            Document templateSingle = (ReportConfigurationSingle.Template != null)
                ? ReportConfigurationSingle.Template.ToDocument()
                : new Campus.Report.ReportTemplate(Properties.Resources.通知單樣版, Campus.Report.TemplateType.Word).ToDocument();
            //List<string> cids = isAll ?
            //      K12.Data.Class.SelectAll().Select(x => x.ID).ToList()
            //    : K12.Presentation.NLDPanels.Class.SelectedSource;
            List<StudentRecord> srl ;
            switch (config.from)
            {
                case mainFormFrom.Class:
                    srl = K12.Data.Student.SelectByClassIDs(K12.Presentation.NLDPanels.Class.SelectedSource);
                    break;
                case mainFormFrom.Student:
                   srl = K12.Data.Student.SelectByIDs(K12.Presentation.NLDPanels.Student.SelectedSource);
                   break;
                default :
                   return;
            }
            List<string> sids  = srl.Select(x => x.ID).ToList();
            int SchoolYear = f.sys.SchoolYear;
            int Semester = f.sys.Semester;
            int Exam = f.exam;
            string ExamText = f.examText;

            DataTable dt = tool._Q.Select(@"select student.id,student.english_name,student.name,student.student_number,student.seat_no,class.class_name,teacher.teacher_name,teacher.contact_phone as teacher_contact_phone,class.grade_year,course.id as course_id,course.period as period,course.credit as credit,course.subject as subject,$ischool.subject.list.group as group,$ischool.subject.list.type as type,$ischool.subject.list.english_name as subject_english_name,sc_attend.ref_student_id as student_id,sce_take.ref_sc_attend_id as sc_attend_id,sce_take.ref_exam_id as exam_id,xpath_string(sce_take.extension,'//Score') as score 
from sc_attend
join sce_take on sce_take.ref_sc_attend_id=sc_attend.id
join course on course.id=sc_attend.ref_course_id
join $ischool.course.extend on $ischool.course.extend.ref_course_id=course.id
left join student on student.id=sc_attend.ref_student_id
left join class on student.ref_class_id=class.id
left join $ischool.subject.list on course.subject=$ischool.subject.list.name 
left join teacher on teacher.id=class.ref_teacher_id
where student.id in (" + string.Join(",", sids) + ") and course.school_year=" + SchoolYear + " and course.semester=" + Semester + " and sce_take.ref_exam_id = " + Exam + " and cast('0'||xpath_string(sce_take.extension,'//Score') as numeric ) < 60");
            Dictionary<string, List<CustomSCETakeRecord>> dscetr = new Dictionary<string, List<CustomSCETakeRecord>>();
            foreach (DataRow row in dt.Rows)
            {
                string id = "" + row["id"];
                if (!dscetr.ContainsKey(id))
                    dscetr.Add(id, new List<CustomSCETakeRecord>());
                int tmp_period, tmp_credit;
                decimal tmp_score;
                decimal.TryParse("" + row["score"], out tmp_score);
                int.TryParse("" + row["period"], out tmp_period);
                int.TryParse("" + row["credit"], out tmp_credit);
                dscetr[id].Add(new CustomSCETakeRecord()
                {
                    RefStudentID = id,
                    Name = "" + row["name"],
                    EnglishName = "" + row["english_name"],
                    StudentNumber = "" + row["student_number"],
                    SeatNo = "" + row["seat_no"],
                    ClassName = "" + row["class_name"],
                    TeacherName = "" + row["teacher_name"],
                    GradeYear = "" + row["grade_year"],
                    Subject = "" + row["subject"],
                    Score = tmp_score,
                    CourseId = "" + row["course_id"],
                    CoursePeriod = tmp_period,
                    CourseCredit = tmp_credit,
                    SubjectEnglishName = "" + row["subject_english_name"],
                    CourseGroup = "" + row["group"],
                    CourseType = StringToSubjectType("" + row["type"]),
                    ExamId = "" + row["exam_id"],
                    TeacherContactPhone = "" + row["teacher_contact_phone"]
                });
            }
            Dictionary<string, object> mailmerge = new Dictionary<string, object>();
            
            Dictionary<string, ParentRecord> dpr = K12.Data.Parent.SelectByStudents(srl).ToDictionary(x => x.RefStudentID, x => x);
            srl.Sort(delegate(StudentRecord a, StudentRecord b)
            {
                StudentRecord aStudent = a;
                ClassRecord aClass = a.Class;
                StudentRecord bStudent = b;
                ClassRecord bClass = b.Class;

                string aa = aClass == null ? (string.Empty).PadLeft(10, '0') : (aClass.Name).PadLeft(10, '0');
                aa += aStudent == null ? (string.Empty).PadLeft(3, '0') : (aStudent.SeatNo + "").PadLeft(3, '0');
                aa += aStudent == null ? (string.Empty).PadLeft(10, '0') : (aStudent.StudentNumber).PadLeft(10, '0');

                string bb = bClass == null ? (string.Empty).PadLeft(10, '0') : (bClass.Name).PadLeft(10, '0');
                bb += bStudent == null ? (string.Empty).PadLeft(3, '0') : (bStudent.SeatNo + "").PadLeft(3, '0');
                bb += bStudent == null ? (string.Empty).PadLeft(10, '0') : (bStudent.StudentNumber).PadLeft(10, '0');

                return aa.CompareTo(bb);
            });
            Document list = (Document)templateList.Clone(true);
            DocumentBuilder db = new DocumentBuilder(list);
            Table table = (Table)list.GetChild(NodeType.Table, 0, true);
            foreach (StudentRecord sr in srl)
            {
                //名單
                if (!dscetr.ContainsKey(sr.ID))
                    continue;
                Document each = (Document)templateSingle.Clone(true);
                Table table2 = (Table)each.GetChild(NodeType.Table, 0, true);
                DocumentBuilder db2 = new DocumentBuilder(each);
                mailmerge.Clear();
                foreach (CustomSCETakeRecord cscetrl in dscetr[sr.ID])
                {
                    //名單
                    table.Rows.Add(table.Rows[table.Rows.Count-1].Clone(true));
                    db.MoveToMergeField("班級");
                    db.Write(cscetrl.ClassName);
                    db.MoveToMergeField("座號");
                    db.Write("" + sr.SeatNo);
                    db.MoveToMergeField("學號");
                    db.Write("" + sr.StudentNumber);
                    db.MoveToMergeField("科目");
                    db.Write(cscetrl.Subject);
                    db.MoveToMergeField("姓名");
                    db.Write(cscetrl.Name);
                    db.MoveToMergeField("英文名");
                    db.Write(cscetrl.EnglishName);
                    db.MoveToMergeField("教師姓名");
                    db.Write(cscetrl.TeacherName);
                    db.MoveToMergeField("分數");
                    db.Write("" + cscetrl.Score);

                    //通知單
                    table2.Rows.Add(table2.Rows[table2.Rows.Count-1].Clone(true));
                    db2.MoveToMergeField("科目");
                    db2.Write(cscetrl.Subject);
                    db2.MoveToMergeField("教師姓名");
                    db2.Write(cscetrl.TeacherName);
                    db2.MoveToMergeField("教師聯絡電話");
                    db2.Write(cscetrl.TeacherContactPhone);
                    db2.MoveToMergeField("分數");
                    db2.Write("" + cscetrl.Score);
                }
                mailmerge.Add("家長姓名", "");
                if (dpr.ContainsKey(sr.ID))
                {
                    mailmerge["家長姓名"] = dpr[sr.ID].CustodianName;
                }
                mailmerge.Add("學年", SchoolYear);
                mailmerge.Add("學期", Semester);
                mailmerge.Add("學段", ExamText);
                mailmerge.Add("班級", sr.Class.Name);
                mailmerge.Add("座號", sr.SeatNo);
                mailmerge.Add("姓名", sr.Name);
                mailmerge.Add("英文名", sr.EnglishName);
                mailmerge.Add("班導師", sr.Class.Teacher.Name);
                mailmerge.Add("學校名稱", School.ChineseName);
                mailmerge.Add("學校英文名稱", School.EnglishName);
                mailmerge.Add("列印日期", DateTime.Now.AddYears(-1911).ToString("yyy.M.d"));

                table2.Rows[table2.Rows.Count - 1].Remove();
                each.MailMerge.Execute(mailmerge.Keys.ToArray(), mailmerge.Values.ToArray());
                if ( config.hasSingle )
                    document.Sections.Add(document.ImportNode(each.FirstSection, true));
            }
            table.Rows[table.Rows.Count - 1].Remove();
            Dictionary<string, object> listmailmerge = new Dictionary<string, object>();
            listmailmerge.Add("列印日期", DateTime.Now.AddYears(-1911).ToString("yyy年M月d日"));
            listmailmerge.Add("學年度", SchoolYear);
            listmailmerge.Add("學期", Semester);
            listmailmerge.Add("學段", ExamText);
            string 學段_長 = "";
            switch (Exam)
            {
                case 1:
                    學段_長 = "第一次段考";
                    break;
                case 2:
                    學段_長 = "第二次段考";
                    break;
            }
            listmailmerge.Add("學段_長", 學段_長);
            list.MailMerge.Execute(listmailmerge.Keys.ToArray(), listmailmerge.Values.ToArray());
            if ( config.hasList )
                document.Sections.Insert(1, document.ImportNode(list.FirstSection, true));

            document.Sections.RemoveAt(0);
            e.Result = document;
        }
        void _bgw_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            Document inResult = (Document)e.Result;
            btnPrint.Enabled = true;
            try
            {
                SaveFileDialog SaveFileDialog1 = new SaveFileDialog();

                SaveFileDialog1.Filter = "Word (*.doc)|*.doc|所有檔案 (*.*)|*.*";
                SaveFileDialog1.FileName = reportName;

                if (SaveFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    inResult.Save(SaveFileDialog1.FileName);
                    Process.Start(SaveFileDialog1.FileName);
                    FISCA.Presentation.MotherForm.SetStatusBarMessage(SaveFileDialog1.FileName + ",列印完成!!");
                    //Update_ePaper ue = new Update_ePaper(new List<Document> { inResult }, current, PrefixStudent.學號);
                    //ue.ShowDialog();
                }
                else
                {
                    FISCA.Presentation.Controls.MsgBox.Show("檔案未儲存");
                    return;
                }
            }
            catch (Exception exp)
            {
                string msg = "檔案儲存錯誤,請檢查檔案是否開啟中!!";
                FISCA.Presentation.Controls.MsgBox.Show(msg + "\n" + exp.Message);
                FISCA.Presentation.MotherForm.SetStatusBarMessage(msg + "\n" + exp.Message);
            }
        }
        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }

}
