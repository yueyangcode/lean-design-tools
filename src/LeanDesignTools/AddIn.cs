using System;
using System.Runtime.InteropServices;
using Extensibility;
using Microsoft.Win32;
using Office = Microsoft.Office.Core;

namespace LeanDesignTools
{
    [ComVisible(true)]
    [Guid("5B4630A0-E087-4E24-93B5-8A79B0E1B842")]
    [ProgId("LeanDesignTools.AddIn")]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    public sealed class AddIn : IDTExtensibility2, Office.IRibbonExtensibility, Office.ICustomTaskPaneConsumer
    {
        private Office.CustomTaskPane _taskPane;
        private Office.ICTPFactory _ctpFactory;
        private Office.IRibbonUI _ribbon;
        private const string SettingsPath = @"Software\LeanDesignTools";

        public void OnConnection(object application, ext_ConnectMode connectMode, object addInInst, ref Array custom)
        {
            DesignPane.PowerPoint = application;
        }

        public void CTPFactoryAvailable(Office.ICTPFactory ctpFactoryInst)
        {
            _ctpFactory = ctpFactoryInst;
            if (ReadPaneVisible())
            {
                try { SetPaneVisible(true); }
                catch (Exception ex) { DesignPane.ShowError("创建设计工具窗格失败", ex); }
            }
        }

        public string GetCustomUI(string ribbonId)
        {
            return @"<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui' onLoad='RibbonLoaded'>
  <ribbon><tabs><tab id='LeanDesignToolsTab' label='设计工具'>
    <group id='LeanDesignToolsGroup' label='设计'>
      <toggleButton id='LeanDesignToolsToggle' label='设计工具' size='large' imageMso='ObjectAlignDialog'
                    onAction='TogglePane' getPressed='GetPanePressed'/>
    </group>
  </tab></tabs></ribbon>
</customUI>";
        }

        public void RibbonLoaded(Office.IRibbonUI ribbonUi) => _ribbon = ribbonUi;

        public void TogglePane(Office.IRibbonControl control, bool pressed)
        {
            try
            {
                SetPaneVisible(pressed);
                SavePaneVisible(pressed);
            }
            catch (Exception ex) { DesignPane.ShowError("创建设计工具窗格失败", ex); }
        }

        public bool GetPanePressed(Office.IRibbonControl control)
        {
            try { return _taskPane != null && _taskPane.Visible; }
            catch { return false; }
        }

        public void OnDisconnection(ext_DisconnectMode removeMode, ref Array custom)
        {
            try { if (_taskPane != null) SavePaneVisible(_taskPane.Visible); } catch { }
            _taskPane = null;
            _ctpFactory = null;
            _ribbon = null;
            DesignPane.PowerPoint = null;
        }

        public void OnAddInsUpdate(ref Array custom) { }
        public void OnStartupComplete(ref Array custom) { }
        public void OnBeginShutdown(ref Array custom) { }

        private void SetPaneVisible(bool visible)
        {
            if (_taskPane == null && visible)
            {
                if (_ctpFactory == null) throw new InvalidOperationException("PowerPoint 尚未提供任务窗格服务。");
                // Keep the visible title while avoiding a task-pane identity collision with iSlide.
                _taskPane = _ctpFactory.CreateCTP("LeanDesignTools.Pane", "设计工具\u200B", Type.Missing);
                _taskPane.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionRight;
                _taskPane.Width = 205;
            }
            if (_taskPane != null) _taskPane.Visible = visible;
            _ribbon?.InvalidateControl("LeanDesignToolsToggle");
        }

        private static bool ReadPaneVisible()
        {
            try
            {
                using (var key = Registry.CurrentUser.OpenSubKey(SettingsPath))
                    return key?.GetValue("PaneVisible") is int value ? value != 0 : true;
            }
            catch { return true; }
        }

        private static void SavePaneVisible(bool visible)
        {
            try
            {
                using (var key = Registry.CurrentUser.CreateSubKey(SettingsPath))
                    key?.SetValue("PaneVisible", visible ? 1 : 0, RegistryValueKind.DWord);
            }
            catch { }
        }
    }
}
