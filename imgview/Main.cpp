#include "ImgViewForm.h"

using namespace System;
using namespace System::Reflection;
using namespace System::Windows::Forms;
using namespace imgview;

[assembly: AssemblyVersionAttribute("1.1.0.1")];

[STAThreadAttribute]
int main(array<String^>^ args) {
	Application::EnableVisualStyles();
	Application::SetCompatibleTextRenderingDefault(false);
	
	//TODO: sanitize args
	Application::Run(gcnew ImgViewForm(args));

	return 0;
}