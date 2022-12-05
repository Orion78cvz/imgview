#include "ImgViewForm.h"

using namespace System;
using namespace System::Windows::Forms;
using namespace imgview;

[STAThreadAttribute]
int main(array<String^>^ args) {
	Application::EnableVisualStyles();
	Application::SetCompatibleTextRenderingDefault(false);
	
	//TODO: sanitize args
	Application::Run(gcnew ImgViewForm(args));

	return 0;
}