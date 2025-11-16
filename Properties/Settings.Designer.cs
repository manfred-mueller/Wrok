// Auto-generated Settings-Designer (manuell hinzugef gt, einfach gehalten)
using System.Configuration;

namespace Wrok.Properties
{
    [global::System.Runtime.CompilerServices.CompilerGeneratedAttribute()]
    [global::System.CodeDom.Compiler.GeneratedCodeAttribute("SettingsSingleFileGenerator", "1.0.0.0")]
    internal sealed partial class Settings : ApplicationSettingsBase
    {
        private static Settings defaultInstance = ((Settings)(Synchronized(new Settings())));

        public static Settings Default
        {
            get { return defaultInstance; }
        }

        [UserScopedSetting()]
        [DefaultSettingValue("30")]
        public int InactivityTimeoutSeconds
        {
            get { return ((int)(this["InactivityTimeoutSeconds"])); }
            set { this["InactivityTimeoutSeconds"] = value; }
        }

        // Fenster-Position/Gr  e speichern
        [UserScopedSetting()]
        [DefaultSettingValue("0")]
        public int WindowTop
        {
            get { return ((int)(this["WindowTop"])); }
            set { this["WindowTop"] = value; }
        }

        [UserScopedSetting()]
        [DefaultSettingValue("0")]
        public int WindowLeft
        {
            get { return ((int)(this["WindowLeft"])); }
            set { this["WindowLeft"] = value; }
        }

        [UserScopedSetting()]
        [DefaultSettingValue("1000")]
        public int WindowWidth
        {
            get { return ((int)(this["WindowWidth"])); }
            set { this["WindowWidth"] = value; }
        }

        [UserScopedSetting()]
        [DefaultSettingValue("700")]
        public int WindowHeight
        {
            get { return ((int)(this["WindowHeight"])); }
            set { this["WindowHeight"] = value; }
        }

        [UserScopedSetting()]
        [DefaultSettingValue("false")]
        public bool IsMaximized
        {
            get { return ((bool)(this["IsMaximized"])); }
            set { this["IsMaximized"] = value; }
        }
    }
}