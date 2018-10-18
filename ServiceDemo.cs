using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics; 
using System.ServiceProcess;
using System.Text; 

namespace FileToImgService
{
    public partial class ServiceDemo : ServiceBase
    {
        private System.Timers.Timer objPollTimer;   //定时器
        private static int intPollTimerDuration;    //服务的执行时间间隔
        private bool isDoJob = true;//是否执行任务

        public ServiceDemo()
        {
            InitializeComponent();
            //当前服务的名称
            this.ServiceName = ServiceConfig.ServiceName;
            //服务的执行时间间隔(单位毫秒)
            intPollTimerDuration = Convert.ToInt32(System.Configuration.ConfigurationSettings.AppSettings.Get("PollInterval"));

            objPollTimer = new System.Timers.Timer();
            //设置间隔时间
            objPollTimer.Interval = intPollTimerDuration;
            //要触发的方法
            objPollTimer.Elapsed += new System.Timers.ElapsedEventHandler(doJob);
            new Aspose.Words.License().SetLicense(LicenseHelper.License.LStream);//去除水印
        }
        /// <summary>
        /// 启动服务执行方法
        /// </summary>
        /// <param name="args"></param>
        protected override void OnStart(string[] args)
        {
            try
            {
                LogControl.LogInfo("服务开启" + DateTime.Now.ToString("yyyyMMddHHmmss"));
                //开启
                objPollTimer.Start();
                System.Diagnostics.EventLog.WriteEntry(ServiceConfig.ServiceName, ServiceConfig.ServiceName + " 在" + DateTime.Now.ToString() + " 时启动");
            }
            catch (Exception e)
            {
                System.Diagnostics.EventLog.WriteEntry(ServiceConfig.ServiceName, ServiceConfig.ServiceName + " 在" + DateTime.Now.ToString() + " 时启动失败,详细信息:" + e.Message, EventLogEntryType.Error);
                OnStop();
            }
        }
        /// <summary>
        /// 关闭服务执行方法
        /// </summary>
        protected override void OnStop()
        {

            //关闭定时器
            objPollTimer.Stop();
            System.Diagnostics.EventLog.WriteEntry(ServiceConfig.ServiceName, ServiceConfig.ServiceName + " 在" + DateTime.Now.ToString() + " 时停止运行");
        }

        #region 执行的方法
        /// <summary>
        /// 执行作业
        /// </summary>
        public void doJob(object sender, System.Timers.ElapsedEventArgs e)
        {
            if (isDoJob)
            {
                LogControl.LogInfo("****************************************服务开始执行任务" + DateTime.Now.ToString("yyyyMMddHHmmss") + "**************************************");
                ServerDemoJob job = new ServerDemoJob();
                isDoJob = false;
                job.DoJob();
                //job.test();
                isDoJob = true;
            }
        }
        #endregion
    }
}
