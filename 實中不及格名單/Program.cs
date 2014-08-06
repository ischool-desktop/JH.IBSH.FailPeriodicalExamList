using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FISCA;
using FISCA.Presentation;
using FISCA.Permission;

namespace JH.IBSH.FailPeriodicalExamList
{
    public class Program
    {
        [MainMethod()]
        static public void Main()
        {
            MenuButton item = K12.Presentation.NLDPanels.Class.RibbonBarItems["資料統計"]["報表"]["成績相關報表"];
            item["不及格名單"].Enable = Permissions.實中班級不及格名單權限;
            item["不及格名單"].Click += delegate
            {
                new MainForm(new MainForm.mainFormConfig()
                {
                    formName = "班級不及格名單",
                    from = MainForm.mainFormFrom.Class,
                    hasList = true,
                    hasSingle = false
                }).ShowDialog();
            };
            item = K12.Presentation.NLDPanels.Class.RibbonBarItems["資料統計"]["報表"]["成績相關報表"];
            item["不及格通知單"].Enable = Permissions.實中班級不及格通知單權限;
            item["不及格通知單"].Click += delegate
            {
                new MainForm(new MainForm.mainFormConfig()
                {
                    formName = "班級不及格通知單",
                    from = MainForm.mainFormFrom.Class,
                    hasList = false,
                    hasSingle = true 
                }).ShowDialog();
            };
            item = K12.Presentation.NLDPanels.Student.RibbonBarItems["資料統計"]["報表"]["成績相關報表"];
            item["不及格通知單"].Enable = Permissions.實中學生不及格通知單權限;
            item["不及格通知單"].Click += delegate
            {
                new MainForm(new MainForm.mainFormConfig()
                {
                    formName = "學生不及格通知單",
                    from = MainForm.mainFormFrom.Student,
                    hasList = false,
                    hasSingle = true
                }).ShowDialog();
            };
            //FISCA.Presentation.RibbonBarItem item2 = FISCA.Presentation.MotherForm.RibbonBarItems["教務作業", "批次作業/檢視"];
            //item2["不及格名單"].Enable = Permissions.實中全部不及格名單權限;
            //item2["不及格名單"].Click += delegate
            //{
            //    new MainForm(true).ShowDialog();
            //};
            Catalog detail1 = RoleAclSource.Instance["班級"]["報表"];
            detail1.Add(new RibbonFeature(Permissions.實中班級不及格名單, "不及格名單"));
            detail1 = RoleAclSource.Instance["班級"]["報表"];
            detail1.Add(new RibbonFeature(Permissions.實中班級不及格通知單, "不及格通知單"));
            detail1 = RoleAclSource.Instance["學生"]["報表"];
            detail1.Add(new RibbonFeature(Permissions.實中學生不及格通知單, "不及格通知單"));
            //Catalog detail2 = RoleAclSource.Instance["教務作業"];
            //detail1.Add(new RibbonFeature(Permissions.實中全部不及格名單, "不及格名單"));
        }
    }
}
