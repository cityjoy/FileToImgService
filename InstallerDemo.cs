using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration.Install; 
using System.ServiceProcess; 
namespace FileToImgService
{
    [RunInstaller(true)]
    public partial class InstallerDemo : Installer
    {
        #region 新增代码
        private ServiceInstaller HostInstaller;
        private ServiceProcessInstaller HostProcessInstaller;
        #endregion

        public InstallerDemo()
        {
            InitializeComponent();

            #region 新增代码
            HostInstaller = new ServiceInstaller();
            HostInstaller.StartType = System.ServiceProcess.ServiceStartMode.Manual;
            HostInstaller.ServiceName = ServiceConfig.ServiceName;//服务名字
            HostInstaller.DisplayName = ServiceConfig.ServiceName;//服务列表显示名字

            HostInstaller.Description = "站点监控服务";
            Installers.Add(HostInstaller);
            HostProcessInstaller = new ServiceProcessInstaller();
            HostProcessInstaller.Account = ServiceAccount.LocalSystem;
            Installers.Add(HostProcessInstaller);
            #endregion
        }
    }
}
