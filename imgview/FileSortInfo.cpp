#include "FileSortInfo.h"

#include <vcclr.h>
#include <shlwapi.h>

#include <shlobj.h>
#include <exdisp.h>
#include <msclr\com\ptr.h>
#include <combaseapi.h>
#include <propsys.h>

#include <winuser.h>

namespace imgview {
namespace filesort {
		
	using namespace System;
	using namespace System::Collections;

	/// <summary>
	/// IComparer implementation using the same method as Explorer
	/// </summary>
	int LogicalStringComparer::Compare(String^ x, String^ y)
	{
		pin_ptr<const wchar_t> px = PtrToStringChars(x);
		pin_ptr<const wchar_t> py = PtrToStringChars(y);
		return ::StrCmpLogicalW(px, py);
	}

	/// <summary>
	/// 対象ディレクトリを開いているExplorerのウィンドウを探し、ソート方法を取得する
	/// </summary>
	SortingOrder AcquireSortingOrder(String^ directory, int def)
	{
		SortingOrder ret = SortingOrder::Name;

		msclr::com::ptr<IShellWindows> wnds;
		wnds.CreateInstance(CLSID_ShellWindows);

		long count;
		if (FAILED(wnds->get_Count(&count))) return ret;

		TCHAR* fdpathbuf = new TCHAR[directory->Length + 2]; //for current folder path, compared up to the length of the target directory + 1
		fdpathbuf[directory->Length + 1] = '\0';

		VARIANT vi;
		V_VT(&vi) = VT_I4;
		for (V_I4(&vi) = 0; V_I4(&vi) < count; V_I4(&vi)++) {
			msclr::com::ptr<IDispatch> disp;
			IDispatch* d;
			if (FAILED(wnds->Item(vi, &d)) || d == NULL) {
#ifdef _DEBUG
				System::Diagnostics::Debug::WriteLine("Failed: IShellWindows->Item");
#endif
				continue;
			}
			disp = d;

			msclr::com::ptr<IWebBrowserApp> ba;
			disp.QueryInterface(ba);
			if (!ba) continue;

			msclr::com::ptr<::IServiceProvider> sp; //NOTE: is not "System::IServiceProvider"
			ba.QueryInterface(sp);
			if (!sp) continue;

			IShellBrowser* b;
			if (FAILED(sp->QueryService(SID_STopLevelBrowser, &b)) || b == NULL) continue;
			msclr::com::ptr<IShellBrowser> sb(b);

			IShellView* v;
			if (FAILED(sb->QueryActiveShellView(&v)) || v == NULL) continue;
			msclr::com::ptr<IShellView> sv(v);

			msclr::com::ptr<IFolderView2> view;
			sv.QueryInterface(view);
			if (!view) continue;
			
			IPersistFolder2* f;
			if (view->GetFolder(IID_PPV_ARGS(&f)) != S_OK) {
				continue;
			}
			msclr::com::ptr<IPersistFolder2> folder(f);

			{ //MEMO: CComHeapPtrを使いたい
				LPITEMIDLIST fpid;
				folder->GetCurFolder(&fpid);
				if (!::SHGetPathFromIDListEx(fpid, fdpathbuf, directory->Length + 2, GPFIDL_DEFAULT)) {
					//TODO: continue;
				}
				CoTaskMemFree(fpid);
			}
			String^ cfpath = gcnew String(fdpathbuf);
			if (!cfpath->Equals(directory)) continue;
#ifdef _DEBUG
			System::Diagnostics::Debug::WriteLine("shell window found.");
#endif

			SORTCOLUMN sort;
			PWSTR key;
			view->GetSortColumns(&sort, 1);
			PSGetNameFromPropertyKey(sort.propkey, &key);
			String^ ks = gcnew String(key);
#ifdef _DEBUG
			System::Diagnostics::Debug::WriteLine(String::Format("Sort={0}", ks));
#endif
			if (ks->EndsWith("DateModified")) {
				ret = SortingOrder::Modified;
			}
			else if (ks->EndsWith("Size")) {
				ret = SortingOrder::Size;
			}
			else if (ks->EndsWith("ItemTypeText")) {
				ret = SortingOrder::Type;
			}

			break;
		}

		delete fdpathbuf;

		return ret;
	}

}}