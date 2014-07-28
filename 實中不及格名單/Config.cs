using FISCA.Presentation.Controls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace JH.IBSH.FailPeriodicalExamList
{
    public partial class Config : BaseForm
    {
        public Config()
        {
            InitializeComponent();
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Campus.Report.TemplateSettingForm TemplateForm;
            if (JH.IBSH.FailPeriodicalExamList.MainForm.ReportConfigurationList.Template == null)
            {
                JH.IBSH.FailPeriodicalExamList.MainForm.ReportConfigurationList.Template = new Campus.Report.ReportTemplate(Properties.Resources.名單樣版, Campus.Report.TemplateType.Word);
            }
            TemplateForm = new Campus.Report.TemplateSettingForm(JH.IBSH.FailPeriodicalExamList.MainForm.ReportConfigurationList.Template, new Campus.Report.ReportTemplate(Properties.Resources.名單樣版, Campus.Report.TemplateType.Word));
            //預設名稱
            TemplateForm.DefaultFileName = "名單樣版";
            if (TemplateForm.ShowDialog() == DialogResult.OK)
            {
                JH.IBSH.FailPeriodicalExamList.MainForm.ReportConfigurationList.Template = TemplateForm.Template;
                JH.IBSH.FailPeriodicalExamList.MainForm.ReportConfigurationList.Save();
            }
        }
        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Campus.Report.TemplateSettingForm TemplateForm;
            if (JH.IBSH.FailPeriodicalExamList.MainForm.ReportConfigurationSingle.Template == null)
            {
                JH.IBSH.FailPeriodicalExamList.MainForm.ReportConfigurationSingle.Template = new Campus.Report.ReportTemplate(Properties.Resources.通知單樣版, Campus.Report.TemplateType.Word);
            }
            TemplateForm = new Campus.Report.TemplateSettingForm(JH.IBSH.FailPeriodicalExamList.MainForm.ReportConfigurationSingle.Template, new Campus.Report.ReportTemplate(Properties.Resources.通知單樣版, Campus.Report.TemplateType.Word));
            //預設名稱
            TemplateForm.DefaultFileName = "通知單樣版";
            if (TemplateForm.ShowDialog() == DialogResult.OK)
            {
                JH.IBSH.FailPeriodicalExamList.MainForm.ReportConfigurationSingle.Template = TemplateForm.Template;
                JH.IBSH.FailPeriodicalExamList.MainForm.ReportConfigurationSingle.Save();
            }
        }
    }
}
