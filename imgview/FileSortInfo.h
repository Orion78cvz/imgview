#pragma once

namespace imgview {
	using namespace System;
	using namespace System::Collections;

	/// <summary>
	/// IComparer implementation using the same method as Explorer
	/// </summary>
	public ref class LogicalStringComparer : public Generic::IComparer<String^>
	{
	public:
		virtual int Compare(String^ x, String^ y);
	};

}