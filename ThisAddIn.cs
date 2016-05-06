using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace SystemwebMonthlyReport
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new Ribbon();
        }

        public void CreateTask(string subject)
        {
            Outlook.TaskItem task = Application.CreateItem(
            Outlook.OlItemType.olTaskItem) as Outlook.TaskItem;

            task.Subject = subject;
            int year = DateTime.Now.Year;
            int month = DateTime.Now.Month;
            int day = DateTime.Now.Day;
            task.StartDate = new DateTime(year, month, 1);
            if (day < 20)
            {
                task.StartDate = task.StartDate.AddMonths(-1);
            }
            task.DueDate = task.StartDate.AddMonths(1).AddDays(-1);
            task.Complete = true;
            task.Mileage = task.StartDate.ToString("yyyy/MM/dd");
            task.DateCompleted = task.DueDate;
            task.Save();
        }

        #region VSTO 產生的程式碼

        /// <summary>
        /// 此為設計工具支援所需的方法 - 請勿使用程式碼編輯器
        /// 修改這個方法的內容。
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion VSTO 產生的程式碼
    }
}