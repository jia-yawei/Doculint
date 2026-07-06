using System;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Win32;

#pragma warning disable CA1416

namespace DocuLint.WpsAddin
{
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [Guid("7E4037DF-0FDB-454A-B438-8B9235CC4D44")]
    [ProgId("DocuLint.WpsAddin")]
    public sealed class WpsAddinBootstrap : IDTExtensibility2, IRibbonExtensibility
    {
        private static readonly string[] CommonStyles =
        {
            "1级标题-105模板",
            "2级标题-105模板",
            "3级标题-105模板",
            "正文-105",
            "题注"
        };

        private QuickLauncherForm launcherForm;
        private object ribbonUi;

        public object Application { get; private set; }

        public void OnConnection(object application, ext_ConnectMode connectMode, object addInInst, ref Array custom)
        {
            Application = application;
        }

        public void OnDisconnection(ext_DisconnectMode removeMode, ref Array custom)
        {
            if (launcherForm != null)
            {
                try
                {
                    launcherForm.Close();
                    launcherForm.Dispose();
                }
                catch
                {
                }

                launcherForm = null;
            }

            ribbonUi = null;
            Application = null;
        }

        public void OnAddInsUpdate(ref Array custom)
        {
        }

        public void OnStartupComplete(ref Array custom)
        {
        }

        public void OnBeginShutdown(ref Array custom)
        {
        }

        public string GetCustomUI(string ribbonId)
        {
            return @"<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui' onLoad='OnRibbonLoad'>
    <ribbon>
        <tabs>
            <tab idMso='TabAddIns'>
                <group id='grpDocuLintStyles' label='常用样式库'>
                    <toggleButton id='btnStyle1'
                                  size='large'
                                  imageMso='QuickStylesGallery'
                                  getLabel='GetCommonStyleLabel'
                                  onAction='OnCommonStyleClick'
                                  getPressed='GetCommonStylePressed'
                                  screentip='应用常用样式'
                                  supertip='将当前选区应用为预设常用样式。' />
                    <toggleButton id='btnStyle2'
                                  size='large'
                                  imageMso='QuickStylesGallery'
                                  getLabel='GetCommonStyleLabel'
                                  onAction='OnCommonStyleClick'
                                  getPressed='GetCommonStylePressed'
                                  screentip='应用常用样式'
                                  supertip='将当前选区应用为预设常用样式。' />
                    <toggleButton id='btnStyle3'
                                  size='large'
                                  imageMso='QuickStylesGallery'
                                  getLabel='GetCommonStyleLabel'
                                  onAction='OnCommonStyleClick'
                                  getPressed='GetCommonStylePressed'
                                  screentip='应用常用样式'
                                  supertip='将当前选区应用为预设常用样式。' />
                    <toggleButton id='btnStyle4'
                                  size='large'
                                  imageMso='QuickStylesGallery'
                                  getLabel='GetCommonStyleLabel'
                                  onAction='OnCommonStyleClick'
                                  getPressed='GetCommonStylePressed'
                                  screentip='应用常用样式'
                                  supertip='将当前选区应用为预设常用样式。' />
                    <toggleButton id='btnStyle5'
                                  size='large'
                                  imageMso='QuickStylesGallery'
                                  getLabel='GetCommonStyleLabel'
                                  onAction='OnCommonStyleClick'
                                  getPressed='GetCommonStylePressed'
                                  screentip='应用常用样式'
                                  supertip='将当前选区应用为预设常用样式。' />
                </group>
                <group id='grpDocuLintWps' label='文档不加班'>
                    <button id='btnDocuLintLauncher'
                            label='文档不加班面板'
                            size='large'
                            imageMso='ReviewNewComment'
                            onAction='OnDocuLintLauncherClick'
                            screentip='打开文档不加班面板'
                            supertip='打开文档不加班的快捷启动窗口。' />
                </group>
            </tab>
        </tabs>
    </ribbon>
</customUI>";
        }

        public void OnRibbonLoad(object ribbon)
        {
            ribbonUi = ribbon;
        }

        public string GetCommonStyleLabel(object control)
        {
            return GetStyleNameByControl(control);
        }

        public bool GetCommonStylePressed(object control)
        {
            string expectedStyle = GetStyleNameByControl(control);
            string currentStyle = GetCurrentSelectionStyleName();
            return !string.IsNullOrWhiteSpace(expectedStyle)
                && string.Equals(currentStyle, expectedStyle, StringComparison.OrdinalIgnoreCase);
        }

        public void OnCommonStyleClick(object control, bool isPressed)
        {
            string styleName = GetStyleNameByControl(control);
            ApplyCommonStyle(styleName);
        }

        public void OnDocuLintLauncherClick(object control)
        {
            ShowLauncher();
        }

        public void ShowLauncher()
        {
            if (launcherForm == null || launcherForm.IsDisposed)
            {
                launcherForm = new QuickLauncherForm();
            }

            launcherForm.SetHostStatus(Application == null ? "Host not connected" : "Host connected");

            if (!launcherForm.Visible)
            {
                launcherForm.Show();
                return;
            }

            launcherForm.BringToFront();
        }

        public void RefreshRibbon()
        {
            if (ribbonUi == null)
            {
                return;
            }

            try
            {
                ribbonUi.GetType().InvokeMember("Invalidate", BindingFlags.InvokeMethod, null, ribbonUi, null);
            }
            catch
            {
            }
        }

        private string GetStyleNameByControl(object control)
        {
            if (control == null)
            {
                return string.Empty;
            }

            string controlId = TryGetStringProperty(control, "Id");
            if (string.IsNullOrWhiteSpace(controlId) || !controlId.StartsWith("btnStyle", StringComparison.OrdinalIgnoreCase))
            {
                return string.Empty;
            }

            string suffix = controlId.Substring("btnStyle".Length);
            if (!int.TryParse(suffix, out int styleIndex))
            {
                return string.Empty;
            }

            styleIndex -= 1;
            if (styleIndex < 0 || styleIndex >= CommonStyles.Length)
            {
                return string.Empty;
            }

            return CommonStyles[styleIndex] ?? string.Empty;
        }

        private string GetCurrentSelectionStyleName()
        {
            try
            {
                object selection = TryGetPropertyValue(Application, "Selection");
                if (selection == null)
                {
                    return string.Empty;
                }

                object styleObj = TryInvoke(selection, "get_Style");
                if (styleObj == null)
                {
                    styleObj = TryGetPropertyValue(selection, "Style");
                }

                string nameLocal = TryGetStringProperty(styleObj, "NameLocal");
                if (!string.IsNullOrWhiteSpace(nameLocal))
                {
                    return nameLocal;
                }

                return Convert.ToString(styleObj) ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        private void ApplyCommonStyle(string styleName)
        {
            if (string.IsNullOrWhiteSpace(styleName))
            {
                return;
            }

            try
            {
                object selection = TryGetPropertyValue(Application, "Selection");
                object range = TryGetPropertyValue(selection, "Range");
                if (range == null)
                {
                    throw new InvalidOperationException("未获取到当前选区。");
                }

                bool applied = TryInvoke(range, "set_Style", styleName) != null;
                if (!applied)
                {
                    applied = TrySetPropertyValue(range, "Style", styleName);
                }

                if (!applied && selection != null)
                {
                    applied = TryInvoke(selection, "set_Style", styleName) != null
                        || TrySetPropertyValue(selection, "Style", styleName);
                }

                if (!applied)
                {
                    throw new InvalidOperationException("WPS 未接受该样式设置请求。");
                }

                RefreshRibbon();
            }
            catch (Exception ex)
            {
                MessageBox.Show("应用样式失败: " + ex.Message, "DocuLint for WPS");
            }
        }

        private static object TryGetPropertyValue(object target, string propertyName)
        {
            if (target == null || string.IsNullOrWhiteSpace(propertyName))
            {
                return null;
            }

            try
            {
                return target.GetType().InvokeMember(propertyName, BindingFlags.GetProperty, null, target, null);
            }
            catch
            {
                return null;
            }
        }

        private static string TryGetStringProperty(object target, string propertyName)
        {
            object value = TryGetPropertyValue(target, propertyName);
            return Convert.ToString(value) ?? string.Empty;
        }

        private static object TryInvoke(object target, string methodName, params object[] args)
        {
            if (target == null || string.IsNullOrWhiteSpace(methodName))
            {
                return null;
            }

            try
            {
                return target.GetType().InvokeMember(methodName, BindingFlags.InvokeMethod, null, target, args);
            }
            catch
            {
                return null;
            }
        }

        private static bool TrySetPropertyValue(object target, string propertyName, object value)
        {
            if (target == null || string.IsNullOrWhiteSpace(propertyName))
            {
                return false;
            }

            try
            {
                target.GetType().InvokeMember(propertyName, BindingFlags.SetProperty, null, target, new[] { value });
                return true;
            }
            catch
            {
                return false;
            }
        }

        [ComRegisterFunction]
        public static void Register(Type type)
        {
            if (type == null)
            {
                return;
            }

            string addinPath = @"Software\Kingsoft\Office\Addins\" + type.FullName;
            using (RegistryKey key = Registry.CurrentUser.CreateSubKey(addinPath))
            {
                if (key == null)
                {
                    return;
                }

                key.SetValue("FriendlyName", "文档不加班", RegistryValueKind.String);
                key.SetValue("Description", "文档不加班 WPS Writer 加载项", RegistryValueKind.String);
                key.SetValue("LoadBehavior", 3, RegistryValueKind.DWord);
                key.SetValue("CommandLineSafe", 0, RegistryValueKind.DWord);
            }
        }

        [ComUnregisterFunction]
        public static void Unregister(Type type)
        {
            if (type == null)
            {
                return;
            }

            string addinPath = @"Software\Kingsoft\Office\Addins\" + type.FullName;
            Registry.CurrentUser.DeleteSubKeyTree(addinPath, false);
            Registry.CurrentUser.DeleteSubKeyTree(@"Software\Microsoft\Office\Word\AddinsData\" + type.FullName, false);
        }
    }

    [ComImport]
    [Guid("B65AD801-ABAF-11D0-BB8B-00A0C90F2744")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IDTExtensibility2
    {
        void OnConnection([MarshalAs(UnmanagedType.IDispatch)] object application, ext_ConnectMode connectMode, [MarshalAs(UnmanagedType.IDispatch)] object addInInst, ref Array custom);

        void OnDisconnection(ext_DisconnectMode removeMode, ref Array custom);

        void OnAddInsUpdate(ref Array custom);

        void OnStartupComplete(ref Array custom);

        void OnBeginShutdown(ref Array custom);
    }

    public enum ext_ConnectMode
    {
        ext_cm_AfterStartup = 0,
        ext_cm_Startup = 1,
        ext_cm_External = 2,
        ext_cm_CommandLine = 3,
        ext_cm_Solution = 4,
        ext_cm_UISetup = 5
    }

    public enum ext_DisconnectMode
    {
        ext_dm_HostShutdown = 0,
        ext_dm_UserClosed = 1,
        ext_dm_UISetupComplete = 2,
        ext_dm_SolutionClosed = 3
    }

    [ComImport]
    [Guid("000C0396-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IRibbonExtensibility
    {
        [DispId(1)]
        string GetCustomUI([MarshalAs(UnmanagedType.BStr)] string ribbonId);
    }
}

#pragma warning restore CA1416
