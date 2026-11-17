using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Extensibility;
using Microsoft.Win32;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace LeanDesignTools
{
    [ComVisible(true)]
    [Guid("5B4630A0-E087-4E24-93B5-8A79B0E1B842")]
    [ProgId("LeanDesignTools.AddIn")]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    public sealed class AddIn : IDTExtensibility2, Office.IRibbonExtensibility, Office.ICustomTaskPaneConsumer
    {
        private readonly Dictionary<string, Office.CustomTaskPane> _taskPanes = new Dictionary<string, Office.CustomTaskPane>(StringComparer.OrdinalIgnoreCase);
        private PowerPoint.Application _powerPoint;
        private Office.ICTPFactory _ctpFactory;
        private Office.IRibbonUI _ribbon;
        private const string SettingsPath = @"Software\LeanDesignTools";

        public void OnConnection(object application, ext_ConnectMode connectMode, object addInInst, ref Array custom)
        {
            _powerPoint = (PowerPoint.Application)application;
            DesignPane.PowerPoint = _powerPoint;
            _powerPoint.WindowActivate += OnWindowActivate;
        }

        public void CTPFactoryAvailable(Office.ICTPFactory ctpFactoryInst)
        {
            _ctpFactory = ctpFactoryInst;
            if (ReadPaneVisible())
            {
                try { SetPaneVisible(true); }
                catch { }
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
            try
            {
                var key = ActiveWindowKey();
                var pane = GetPane(key, false);
                if (pane == null && ReadPaneVisible())
                {
                    SetPaneVisible(true);
                    pane = GetPane(key, false);
                }
                return pane != null && pane.Visible;
            }
            catch { return false; }
        }

        public void OnDisconnection(ext_DisconnectMode removeMode, ref Array custom)
        {
            try
            {
                var pane = GetPane(ActiveWindowKey(), false);
                if (pane != null) SavePaneVisible(pane.Visible);
            }
            catch { }
            try { if (_powerPoint != null) _powerPoint.WindowActivate -= OnWindowActivate; }
            catch { }
            _taskPanes.Clear();
            _powerPoint = null;
            _ctpFactory = null;
            _ribbon = null;
            DesignPane.PowerPoint = null;
        }

        public void OnAddInsUpdate(ref Array custom) { }
        public void OnStartupComplete(ref Array custom) { }
        public void OnBeginShutdown(ref Array custom) { }

        private void OnWindowActivate(PowerPoint.Presentation presentation, PowerPoint.DocumentWindow window)
        {
            if (_ctpFactory == null || !ReadPaneVisible()) return;
            try { SetPaneVisible(true); }
            catch { }
        }

        private void SetPaneVisible(bool visible)
        {
            var key = ActiveWindowKey();
            var pane = GetPane(key, visible);
            if (pane != null)
            {
                try { pane.Visible = visible; }
                catch
                {
                    _taskPanes.Remove(key);
                    if (!visible) return;
                    pane = GetPane(key, true);
                    pane.Visible = true;
                }
            }
            _ribbon?.InvalidateControl("LeanDesignToolsToggle");
        }

        private Office.CustomTaskPane GetPane(string key, bool create)
        {
            if (_taskPanes.TryGetValue(key, out var pane))
            {
                try { var visible = pane.Visible; return pane; }
                catch { _taskPanes.Remove(key); }
            }
            if (!create) return null;
            if (_ctpFactory == null) throw new InvalidOperationException("PowerPoint 尚未提供任务窗格服务。");

            // Keep the visible title while avoiding a task-pane identity collision with iSlide.
            pane = _ctpFactory.CreateCTP("LeanDesignTools.Pane", "设计工具\u200B", DesignPane.PowerPoint.ActiveWindow);
            pane.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionRight;
            pane.Width = 205;
            _taskPanes[key] = pane;
            return pane;
        }

        private static string ActiveWindowKey()
        {
            var window = DesignPane.PowerPoint?.ActiveWindow as PowerPoint.DocumentWindow;
            if (window == null) throw new InvalidOperationException("PowerPoint 当前没有活动窗口。");
            // shortcut: one PowerPoint window per presentation, use a stable native window ID if multi-view support is added.
            return window.Presentation.FullName;
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
