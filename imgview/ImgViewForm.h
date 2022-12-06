#pragma once

#include <algorithm>

#include <shlobj.h>
#include <exdisp.h>
#include <msclr\com\ptr.h>
#include <propsys.h>

#include <winuser.h>

#include "resource.h"

#include "Settings.h"
#include "FileSortInfo.h"

namespace imgview {
	using namespace System;
	using namespace System::ComponentModel;
	using namespace System::Collections;
	using namespace System::Windows::Forms;
	using namespace System::Data;
	using namespace System::Drawing;

	/// <summary>
	/// ImgViewForm の概要
	/// </summary>
	public ref class ImgViewForm : public System::Windows::Forms::Form
	{
	public:
		ImgViewForm(array<String^>^ filenames)
		{
			this->appSettings = gcnew imgview::Settings;

			InitializeComponent();
		
			this->InitializeExtensions();
			if(filenames && filenames->Length > 0) {
				this->ShowSelectedImage(filenames, true);
			}
		}

	public:
		/// <returns>ダイアログで選択された単一ファイルパス、キャンセル時はnullptr</returns>
		array<String^>^ FromOpenFileDialog() {
			String^ exts = String::Join(";", this->listImgExtensions);
			
			System::Windows::Forms::OpenFileDialog^ ofd = gcnew System::Windows::Forms::OpenFileDialog();
			ofd->Title = "ファイルを開く";
			ofd->Filter = String::Format("画像ファイル({0})|{0}|すべてのファイル(*.*)|*.*", exts);
			ofd->FilterIndex = 1;
			ofd->Multiselect = true;
			ofd->InitialDirectory = System::IO::Directory::GetCurrentDirectory();
			ofd->RestoreDirectory = false; //カレントディレクトリを移動する
			if (ofd->ShowDialog() == System::Windows::Forms::DialogResult::OK) {
				System::IO::Directory::SetCurrentDirectory(System::IO::Path::GetDirectoryName(ofd->FileNames[0])); //MEMO: 必要?
				return ofd->FileNames;
			} else {
				return nullptr;
			}
		}
	protected:
		imgview::Settings^ appSettings;

	protected:
		array<String^>^ foundFileNames = nullptr;
		int currentPos = 0;
	private: System::Windows::Forms::ToolStripStatusLabel^ statusFileIndex;
	private: System::Windows::Forms::ToolStripMenuItem^ 表示VToolStripMenuItem;
	private: System::Windows::Forms::ToolStripMenuItem^ 背景色BToolStripMenuItem;
	private: System::Windows::Forms::ToolStripMenuItem^ menuBackgroundColorWhite;
	private: System::Windows::Forms::ToolStripMenuItem^ menuBackgroundColorDefault;

	private: System::Windows::Forms::ToolStripMenuItem^ menuBackgroundColorBlack;




	protected:
		array<String^>^ listImgExtensions;
		System::Void InitializeExtensions() {
			this->listImgExtensions = gcnew array<String^> {"*.bmp", "*.jpg", "*.jpe", "*.jpeg", "*.png", "*.gif", "*.tiff", "*.exif"};
		}

	protected:
		/// <summary>
		/// 使用中のリソースをすべてクリーンアップします。
		/// </summary>
		~ImgViewForm()
		{
			if (components)
			{
				delete components;
			}
		}
	private: System::Windows::Forms::MenuStrip^ MainMenu;
	protected:
	private: System::Windows::Forms::ToolStripMenuItem^ ファイルFToolStripMenuItem;
	private: System::Windows::Forms::ToolStripMenuItem^ menuExit;
	private: System::Windows::Forms::StatusStrip^ statusStrip;
	private: System::Windows::Forms::ToolStripStatusLabel^ statusDisplayRatio;
	private: System::Windows::Forms::PictureBox^ pictureImage;
	private: System::Windows::Forms::ToolStripStatusLabel^ statusFilePath;
	private: System::Windows::Forms::ToolStripMenuItem^ menuOpenFile;


	private:
		/// <summary>
		/// 必要なデザイナー変数です。
		/// </summary>
		System::ComponentModel::Container ^components;

#pragma region Windows Form Designer generated code
		/// <summary>
		/// デザイナー サポートに必要なメソッドです。このメソッドの内容を
		/// コード エディターで変更しないでください。
		/// </summary>
		void InitializeComponent(void)
		{
			this->MainMenu = (gcnew System::Windows::Forms::MenuStrip());
			this->ファイルFToolStripMenuItem = (gcnew System::Windows::Forms::ToolStripMenuItem());
			this->menuOpenFile = (gcnew System::Windows::Forms::ToolStripMenuItem());
			this->menuExit = (gcnew System::Windows::Forms::ToolStripMenuItem());
			this->表示VToolStripMenuItem = (gcnew System::Windows::Forms::ToolStripMenuItem());
			this->背景色BToolStripMenuItem = (gcnew System::Windows::Forms::ToolStripMenuItem());
			this->menuBackgroundColorDefault = (gcnew System::Windows::Forms::ToolStripMenuItem());
			this->menuBackgroundColorWhite = (gcnew System::Windows::Forms::ToolStripMenuItem());
			this->menuBackgroundColorBlack = (gcnew System::Windows::Forms::ToolStripMenuItem());
			this->statusStrip = (gcnew System::Windows::Forms::StatusStrip());
			this->statusDisplayRatio = (gcnew System::Windows::Forms::ToolStripStatusLabel());
			this->statusFilePath = (gcnew System::Windows::Forms::ToolStripStatusLabel());
			this->statusFileIndex = (gcnew System::Windows::Forms::ToolStripStatusLabel());
			this->pictureImage = (gcnew System::Windows::Forms::PictureBox());
			this->MainMenu->SuspendLayout();
			this->statusStrip->SuspendLayout();
			(cli::safe_cast<System::ComponentModel::ISupportInitialize^>(this->pictureImage))->BeginInit();
			this->SuspendLayout();
			// 
			// MainMenu
			// 
			this->MainMenu->Items->AddRange(gcnew cli::array< System::Windows::Forms::ToolStripItem^  >(2) {
				this->ファイルFToolStripMenuItem,
					this->表示VToolStripMenuItem
			});
			this->MainMenu->Location = System::Drawing::Point(0, 0);
			this->MainMenu->Name = L"MainMenu";
			this->MainMenu->Size = System::Drawing::Size(624, 24);
			this->MainMenu->TabIndex = 0;
			this->MainMenu->Text = L"MainMenu";
			// 
			// ファイルFToolStripMenuItem
			// 
			this->ファイルFToolStripMenuItem->DropDownItems->AddRange(gcnew cli::array< System::Windows::Forms::ToolStripItem^  >(2) {
				this->menuOpenFile,
					this->menuExit
			});
			this->ファイルFToolStripMenuItem->Name = L"ファイルFToolStripMenuItem";
			this->ファイルFToolStripMenuItem->Size = System::Drawing::Size(67, 20);
			this->ファイルFToolStripMenuItem->Text = L"ファイル(&F)";
			// 
			// menuOpenFile
			// 
			this->menuOpenFile->Name = L"menuOpenFile";
			this->menuOpenFile->Size = System::Drawing::Size(113, 22);
			this->menuOpenFile->Text = L"開く(&O)";
			this->menuOpenFile->Click += gcnew System::EventHandler(this, &ImgViewForm::menuOpenFile_Click);
			// 
			// menuExit
			// 
			this->menuExit->Name = L"menuExit";
			this->menuExit->Size = System::Drawing::Size(113, 22);
			this->menuExit->Text = L"終了(&X)";
			this->menuExit->Click += gcnew System::EventHandler(this, &ImgViewForm::menuExit_Click);
			// 
			// 表示VToolStripMenuItem
			// 
			this->表示VToolStripMenuItem->DropDownItems->AddRange(gcnew cli::array< System::Windows::Forms::ToolStripItem^  >(1) { this->背景色BToolStripMenuItem });
			this->表示VToolStripMenuItem->Name = L"表示VToolStripMenuItem";
			this->表示VToolStripMenuItem->Size = System::Drawing::Size(58, 20);
			this->表示VToolStripMenuItem->Text = L"表示(&V)";
			// 
			// 背景色BToolStripMenuItem
			// 
			this->背景色BToolStripMenuItem->DropDownItems->AddRange(gcnew cli::array< System::Windows::Forms::ToolStripItem^  >(3) {
				this->menuBackgroundColorDefault,
					this->menuBackgroundColorWhite, this->menuBackgroundColorBlack
			});
			this->背景色BToolStripMenuItem->Name = L"背景色BToolStripMenuItem";
			this->背景色BToolStripMenuItem->Size = System::Drawing::Size(180, 22);
			this->背景色BToolStripMenuItem->Text = L"背景色(&B)";
			// 
			// menuBackgroundColorDefault
			// 
			this->menuBackgroundColorDefault->BackColor = System::Drawing::SystemColors::Control;
			this->menuBackgroundColorDefault->Name = L"menuBackgroundColorDefault";
			this->menuBackgroundColorDefault->Size = System::Drawing::Size(117, 22);
			this->menuBackgroundColorDefault->Text = L"デフォルト";
			this->menuBackgroundColorDefault->Click += gcnew System::EventHandler(this, &ImgViewForm::menuBackgroundColorDefault_Click);
			// 
			// menuBackgroundColorWhite
			// 
			this->menuBackgroundColorWhite->BackColor = System::Drawing::Color::White;
			this->menuBackgroundColorWhite->ForeColor = System::Drawing::Color::Black;
			this->menuBackgroundColorWhite->Name = L"menuBackgroundColorWhite";
			this->menuBackgroundColorWhite->Size = System::Drawing::Size(117, 22);
			this->menuBackgroundColorWhite->Text = L"白";
			this->menuBackgroundColorWhite->Click += gcnew System::EventHandler(this, &ImgViewForm::menuBackgroundColorWhite_Click);
			// 
			// menuBackgroundColorBlack
			// 
			this->menuBackgroundColorBlack->BackColor = System::Drawing::Color::Black;
			this->menuBackgroundColorBlack->ForeColor = System::Drawing::Color::White;
			this->menuBackgroundColorBlack->Name = L"menuBackgroundColorBlack";
			this->menuBackgroundColorBlack->Size = System::Drawing::Size(117, 22);
			this->menuBackgroundColorBlack->Text = L"黒";
			this->menuBackgroundColorBlack->Click += gcnew System::EventHandler(this, &ImgViewForm::menuBackgroundColorBlack_Click);
			// 
			// statusStrip
			// 
			this->statusStrip->Items->AddRange(gcnew cli::array< System::Windows::Forms::ToolStripItem^  >(3) {
				this->statusDisplayRatio,
					this->statusFilePath, this->statusFileIndex
			});
			this->statusStrip->Location = System::Drawing::Point(0, 419);
			this->statusStrip->Name = L"statusStrip";
			this->statusStrip->Size = System::Drawing::Size(624, 22);
			this->statusStrip->TabIndex = 1;
			this->statusStrip->Text = L"statusStrip";
			// 
			// statusDisplayRatio
			// 
			this->statusDisplayRatio->Name = L"statusDisplayRatio";
			this->statusDisplayRatio->Size = System::Drawing::Size(35, 17);
			this->statusDisplayRatio->Text = L"100%";
			// 
			// statusFilePath
			// 
			this->statusFilePath->Name = L"statusFilePath";
			this->statusFilePath->Size = System::Drawing::Size(10, 17);
			this->statusFilePath->Text = L" ";
			// 
			// statusFileIndex
			// 
			this->statusFileIndex->Name = L"statusFileIndex";
			this->statusFileIndex->Size = System::Drawing::Size(32, 17);
			this->statusFileIndex->Text = L"(0/0)";
			// 
			// pictureImage
			// 
			this->pictureImage->Dock = System::Windows::Forms::DockStyle::Fill;
			this->pictureImage->ImageLocation = L"";
			this->pictureImage->Location = System::Drawing::Point(0, 24);
			this->pictureImage->Name = L"pictureImage";
			this->pictureImage->Size = System::Drawing::Size(624, 395);
			this->pictureImage->SizeMode = System::Windows::Forms::PictureBoxSizeMode::Zoom;
			this->pictureImage->TabIndex = 2;
			this->pictureImage->TabStop = false;
			// 
			// ImgViewForm
			// 
			this->AutoScaleDimensions = System::Drawing::SizeF(6, 12);
			this->AutoScaleMode = System::Windows::Forms::AutoScaleMode::Font;
			this->ClientSize = System::Drawing::Size(624, 441);
			this->Controls->Add(this->pictureImage);
			this->Controls->Add(this->statusStrip);
			this->Controls->Add(this->MainMenu);
			this->MainMenuStrip = this->MainMenu;
			this->Name = L"ImgViewForm";
			this->Text = L"ImgView";
			this->FormClosing += gcnew System::Windows::Forms::FormClosingEventHandler(this, &ImgViewForm::ImgViewForm_FormClosing);
			this->Load += gcnew System::EventHandler(this, &ImgViewForm::ImgViewForm_Load);
			this->SizeChanged += gcnew System::EventHandler(this, &ImgViewForm::ImgViewForm_SizeChanged);
			this->KeyDown += gcnew System::Windows::Forms::KeyEventHandler(this, &ImgViewForm::ImgViewForm_KeyDown);
			this->MainMenu->ResumeLayout(false);
			this->MainMenu->PerformLayout();
			this->statusStrip->ResumeLayout(false);
			this->statusStrip->PerformLayout();
			(cli::safe_cast<System::ComponentModel::ISupportInitialize^>(this->pictureImage))->EndInit();
			this->ResumeLayout(false);
			this->PerformLayout();

		}
#pragma endregion
	/// <summary>
	/// 指定された画像を読み込み、各種コンポーネントを更新する
	/// </summary>
	private: System::Void LoadImage(String^ location) {
		try {
			this->pictureImage->Load(location); //NOTE: パフォーマンスが気になるなら非同期処理にするか検討(前後一つずつ裏でロードする等)
			ImgViewForm_SizeChanged(nullptr, nullptr);
		}
		catch(Exception^) {
			this->pictureImage->Image = this->pictureImage->ErrorImage; //NOTE: 拡大されるが結局エラーなので気にしない
		}
		this->statusFileIndex->Text = String::Format("({0:d}/{1:d})", this->currentPos + 1, this->foundFileNames->Length);

		this->statusFilePath->Text = location;
	}

	/// <summary>
	/// メニューから終了する
	/// </summary>
	private: System::Void menuExit_Click(System::Object^ sender, System::EventArgs^ e) {
		this->Close();
	}

	/// <summary>
	/// Form::OnLoad
	///   Settingsを適用する
	/// </summary>
	private: System::Void ImgViewForm_Load(System::Object^ sender, System::EventArgs^ e) {
		this->Size = this->appSettings->FormClientSize;
		this->WindowState = this->appSettings->WindowState;
		
		this->pictureImage->BackColor = this->appSettings->BackgroundColor;
	}
	/// <summary>
	/// Form::OnClosing
	///   Settingsを保存する
	/// </summary>
	private: System::Void ImgViewForm_FormClosing(System::Object^ sender, System::Windows::Forms::FormClosingEventArgs^ e) {
		switch(this->WindowState) {
		case FormWindowState::Normal:
			this->appSettings->WindowState = FormWindowState::Normal;
			this->appSettings->FormClientSize = this->Size;
			break;
		case FormWindowState::Maximized:
			this->appSettings->WindowState = FormWindowState::Maximized;
			break;
		}

		this->appSettings->BackgroundColor = this->pictureImage->BackColor;

		this->appSettings->Save();
	}

	/// <summary>
	/// ウィンドウサイズ変更で画像表示を更新
	/// INFO: 現状PictureBoxのSizeModeがZoom固定なので細かい処理はない
	/// </summary>
	private: System::Void ImgViewForm_SizeChanged(System::Object^ sender, System::EventArgs^ e) {
		int dr = 0;
		if(this->pictureImage->Image && this->pictureImage->Image->Size.Width != 0 && this->pictureImage->Image->Size.Height != 0) {
			dr = (std::min)(100 * this->pictureImage->Size.Width / this->pictureImage->Image->Size.Width, 100 * this->pictureImage->Size.Height / this->pictureImage->Image->Size.Height);
		}
		this->statusDisplayRatio->Text = String::Format("{0:d}%", dr);
	}

	//TODO: FileDialogからソートが取得できないためメニューから開く機能は削除する(テスト用に残してある)
	private: System::Void menuOpenFile_Click(System::Object^ sender, System::EventArgs^ e) {
		array<String^>^ files = this->FromOpenFileDialog();
		if (files == nullptr || files->Length == 0) return;
		
		this->ShowSelectedImage(files, true);
	}

	/// <summary>
	///  画像ファイルをPictureBoxにロードする
	///  単一ファイルが指定された場合は同ディレクトリにある他のファイルも列挙する
	/// </summary>
	private: System::Void ShowSelectedImage(array<String^>^ filenames, bool seekdirectory) {
		if (filenames == nullptr || filenames->Length == 0) return;

		if(filenames->Length > 1 || !seekdirectory) {
			this->foundFileNames = (array<String^>^)filenames->Clone();
			this->currentPos = 0;
			this->LoadImage(this->foundFileNames[this->currentPos]);
		}
		else
		{
			String^ cf = filenames[0];

			array<String^>^ founds = this->SeekDirectory(cf);

#ifdef _DEBUG
			for(int i = 0; i < (std::min)(10, founds->Length); i++) System::Diagnostics::Debug::WriteLine(founds[i]);
#endif
			this->foundFileNames = founds;
#ifdef _DEBUG
			this->currentPos = System::Array::IndexOf(founds, cf);
#else
			this->currentPos = (std::max)(0, System::Array::IndexOf(founds, cf));
#endif
			this->LoadImage(cf);
		}
	}
	private: array<String^>^ SeekDirectory(String^ filepath) {
		String^ dir = System::IO::Path::GetDirectoryName(filepath);
		int orderby = this->AcquireSortingOrder(dir, 0);

		Array::Sort(this->listImgExtensions);
		
		Generic::List<String^>^ fls = gcnew Generic::List<String^>();
		for each (String ^ ex in this->listImgExtensions) {
			array<String^>^ fs = System::IO::Directory::GetFiles(dir, ex); //NOTE: 単一ディレクトリで固定ならばファイル名部分のみでよい
			if(fs->Length > 0) fls->AddRange(fs);
		}

		array<String^>^ ret = fls->ToArray();
		//Array::Sort(ret);
		//TODO: Array::Sort()が安定ソートではないので同一値の中で並べ替えが不定になってしまう。データ型を作った方がよさそう
		switch(orderby) {
			case 1: {
				array<DateTime>^ dates = gcnew array<DateTime>(ret->Length);
				for (int i = 0; i < ret->Length; i++) {
					dates[i] = System::IO::File::GetLastWriteTime(ret[i]);
				}
				Array::Sort(dates, ret);
				break;
			}
			case 2: {
				array<Int64>^ sizes = gcnew array<Int64>(ret->Length);
				for (int i = 0; i < ret->Length; i++) {
					sizes[i] = System::IO::FileInfo(ret[i]).Length;
				}
				Array::Sort(sizes, ret);
				break;
			}
			case 3: {
				array<String^>^ exts = gcnew array<String^>(ret->Length);
				for (int i = 0; i < ret->Length; i++) {
					exts[i] = System::IO::Path::GetExtension(ret[i]);
				}
				Array::Sort(exts, ret);
				break;
			}
			default:
				Array::Sort(ret, gcnew LogicalStringComparer());
				break;
		}

		return ret;
	}

	/// <summary>
	/// 操作用キー入力処理
	/// </summary>
	private: System::Void ImgViewForm_KeyDown(System::Object^ sender, System::Windows::Forms::KeyEventArgs^ e) {
		switch(e->KeyCode) {
		case Keys::Left: //NOTE: 逆順にするならばここの処理を入れ替えるのが楽?
			if(this->foundFileNames && this->currentPos > 0) {
				this->currentPos--;
				this->LoadImage(this->foundFileNames[this->currentPos]);
			}
			break;
		case Keys::Right:
			if (this->foundFileNames && this->currentPos < this->foundFileNames->Length - 1) {
				this->currentPos++;
				this->LoadImage(this->foundFileNames[this->currentPos]);
			}
			break;
		}
	}

	/// <summary>
	/// 対象ディレクトリを開いているExplorerのウィンドウを探し、ソート方法を取得する
	/// TODO: メンバ関数にする必要がないのでFileSortInfoへ移動する
	/// </summary>
	private: int AcquireSortingOrder(String^ directory, int def) {
		int ret = def;
		
		msclr::com::ptr<IShellWindows> wnds;
		wnds.CreateInstance(CLSID_ShellWindows);

		long count;
		if(FAILED(wnds->get_Count(&count))) return ret;

		TCHAR* title = new TCHAR[directory->Length + 2]; //for GetWindowText(), comparing up to the length of the target directory + 1 char
		title[directory->Length + 1] = '\0';
		
		VARIANT vi;
		V_VT(&vi) = VT_I4;
		for(V_I4(&vi) = 0; V_I4(&vi) < count; V_I4(&vi)++) {
			msclr::com::ptr<IDispatch> disp;
			IDispatch* d;
			if(FAILED(wnds->Item(vi, &d)) || d == NULL) {
				System::Diagnostics::Debug::WriteLine("Failed: IShellWindows->Item");
				continue;
			}
			disp = d;

			msclr::com::ptr<IWebBrowserApp> ba;
			HWND hwnd;
			disp.QueryInterface(ba);
			if(!ba) continue;
			ba->get_HWND(reinterpret_cast<SHANDLE_PTR*>(&hwnd));
			
			// Compare to the title bar
			int rc;
			rc = ::GetWindowText(hwnd, title, directory->Length + 2);
			String^ ttl = gcnew String(title);
			if(rc == 0 || !ttl->Equals(directory)) continue;
#ifdef _DEBUG
			System::Diagnostics::Debug::WriteLine("shell window found.");
#endif

			msclr::com::ptr<::IServiceProvider> sp; //NOTE: is not "System::IServiceProvider"
			ba.QueryInterface(sp);
			if(!sp) continue;

			
			IShellBrowser *b;
			if(FAILED(sp->QueryService(SID_STopLevelBrowser, &b)) || b == NULL) continue;
			msclr::com::ptr<IShellBrowser> sb(b);

			IShellView *v;
			if(FAILED(sb->QueryActiveShellView(&v)) || v == NULL) continue;
			msclr::com::ptr<IShellView> sv(v);

			msclr::com::ptr<IFolderView2> view;
			sv.QueryInterface(view);
			if(!view){
			} else {
				SORTCOLUMN sort;
				PWSTR key;
				view->GetSortColumns(&sort, 1);
				PSGetNameFromPropertyKey(sort.propkey, &key);
				String^ ks = gcnew String(key);
#ifdef _DEBUG
				System::Diagnostics::Debug::WriteLine(String::Format("Sort={0}", ks));
#endif
				if (ks->EndsWith("DateModified")) {
					ret = 1;
				} else if (ks->EndsWith("Size")) {
					ret = 2;
				}else if (ks->EndsWith("ItemTypeText")) {
					ret = 3;
				}

				break;
			}
		}

		delete title;

		return ret;
	}

	/// <summary>
	/// 背景色変更
	/// </summary>
	private: System::Void menuBackgroundColorDefault_Click(System::Object^ sender, System::EventArgs^ e) {
		this->pictureImage->BackColor = System::Drawing::SystemColors::Control;
	}
	private: System::Void menuBackgroundColorWhite_Click(System::Object^ sender, System::EventArgs^ e) {
		this->pictureImage->BackColor = System::Drawing::Color::White;
	}
	private: System::Void menuBackgroundColorBlack_Click(System::Object^ sender, System::EventArgs^ e) {
		this->pictureImage->BackColor = System::Drawing::Color::Black;
	}

};


}
