#include "FileSortInfo.h"

#include<vcclr.h>
#include <shlwapi.h>

namespace imgview {
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

}