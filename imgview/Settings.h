#pragma once

using namespace System;
using namespace System::ComponentModel;
using namespace System::Configuration;
using namespace System::Windows;

namespace imgview {

	public ref class Settings sealed: public System::Configuration::ApplicationSettingsBase {
        public:
            [UserScopedSettingAttribute()]
            [DefaultSettingValueAttribute("False")]
            property bool SettingsUpgraded
            {
                bool get() { return (bool)this["SettingsUpgraded"]; }
                void set(bool t) { this["SettingsUpgraded"] = t; }
            }
        public:
            [UserScopedSettingAttribute()]
            [DefaultSettingValueAttribute("0")] //Normal
            property System::Windows::Forms::FormWindowState WindowState
            {
                Forms::FormWindowState get() {return (Forms::FormWindowState)this["WindowState"];}
                void set(Forms::FormWindowState state) {this["WindowState"] = state;}
            }
        public:
            [UserScopedSettingAttribute()]
            [DefaultSettingValueAttribute("640, 480")]
            property Drawing::Size FormClientSize
            {
                Drawing::Size get() { return (Drawing::Size)this["FormClientSize"]; }
                void set(Drawing::Size size) { this["FormClientSize"] = size; }
            }
        public:
            [UserScopedSettingAttribute()]
            [DefaultSettingValueAttribute("Control")]
            property Drawing::Color BackgroundColor
            {
                Drawing::Color get() { return (Drawing::Color)this["BackgroundColor"]; }
                void set(Drawing::Color color) { this["BackgroundColor"] = color; }
            }

        public:
            Settings() : ApplicationSettingsBase()
            {
                //バージョン(ビルド)が更新された際に古いuser.configを引き継ぐ
                if(!this->SettingsUpgraded) {
                    this->Upgrade();
                    this->SettingsUpgraded = true;
                    this->Save();
                    System::Diagnostics::Debug::WriteLine("Configuration file upgraded");
                }
            }
	};

}
