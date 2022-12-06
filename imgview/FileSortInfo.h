#pragma once

namespace imgview {
namespace filesort {
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

	enum class SortingOrder {
		Name, Modified, Size, Type
	};
	/// <summary>
	///   find Explorer window opening target directory, and get sorting order
	/// </summary>
	SortingOrder AcquireSortingOrder(String^ directory, int def);
}}