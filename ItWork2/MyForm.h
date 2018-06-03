#pragma once


#include <xlnt/xlnt.hpp>
#include "iostream"
#include "windows.h"
#include "ShellAPI.h"
#include <msclr\marshal_cppstd.h>  

namespace ItWork2 {

	using namespace System;
	using namespace System::IO;
	using namespace System::ComponentModel;
	using namespace System::Collections;
	using namespace System::Windows::Forms;
	using namespace System::Data;
	using namespace System::Drawing;
	using namespace Runtime::InteropServices;
	using namespace msclr::interop;

	/// <summary>
	/// Сводка для MyForm
	/// </summary>
	public ref class MyForm : public System::Windows::Forms::Form
	{
	public:
		MyForm(void)
		{
			InitializeComponent();
			//
			//TODO: добавьте код конструктора
			//
		}

	protected:
		/// <summary>
		/// Освободить все используемые ресурсы.
		/// </summary>
		~MyForm()
		{
			if (components)
			{
				delete components;
			}
		}
	private: System::Windows::Forms::Button^  button1;
	private: System::Windows::Forms::TextBox^  textBox1;
	private: System::Windows::Forms::FolderBrowserDialog^  folderBrowserDialog1;
	private: System::Windows::Forms::Button^  button2;
	private: System::Windows::Forms::Button^  button3;
	private: System::Windows::Forms::NumericUpDown^  numericUpDown1;
	private: System::Windows::Forms::ListView^  listView1;
	private: System::Windows::Forms::ColumnHeader^  columnHeader1;
	private: System::Windows::Forms::ColumnHeader^  columnHeader2;
	private: System::Windows::Forms::ColumnHeader^  columnHeader3;
	private: System::Windows::Forms::Button^  button4;
	private: System::Windows::Forms::OpenFileDialog^  openFileDialog1;
	private: System::Windows::Forms::ColumnHeader^  columnHeader4;
	private: System::Windows::Forms::ColumnHeader^  columnHeader5;
	private: System::Windows::Forms::ComboBox^  comboBox1;

	protected:

	private:
		/// <summary>
		/// Обязательная переменная конструктора.
		/// </summary>
		System::ComponentModel::Container ^components;

#pragma region Windows Form Designer generated code
		/// <summary>
		/// Требуемый метод для поддержки конструктора — не изменяйте 
		/// содержимое этого метода с помощью редактора кода.
		/// </summary>
		void InitializeComponent(void)
		{
			this->button1 = (gcnew System::Windows::Forms::Button());
			this->textBox1 = (gcnew System::Windows::Forms::TextBox());
			this->folderBrowserDialog1 = (gcnew System::Windows::Forms::FolderBrowserDialog());
			this->button2 = (gcnew System::Windows::Forms::Button());
			this->button3 = (gcnew System::Windows::Forms::Button());
			this->numericUpDown1 = (gcnew System::Windows::Forms::NumericUpDown());
			this->listView1 = (gcnew System::Windows::Forms::ListView());
			this->columnHeader1 = (gcnew System::Windows::Forms::ColumnHeader());
			this->columnHeader2 = (gcnew System::Windows::Forms::ColumnHeader());
			this->columnHeader3 = (gcnew System::Windows::Forms::ColumnHeader());
			this->columnHeader4 = (gcnew System::Windows::Forms::ColumnHeader());
			this->columnHeader5 = (gcnew System::Windows::Forms::ColumnHeader());
			this->button4 = (gcnew System::Windows::Forms::Button());
			this->openFileDialog1 = (gcnew System::Windows::Forms::OpenFileDialog());
			this->comboBox1 = (gcnew System::Windows::Forms::ComboBox());
			(cli::safe_cast<System::ComponentModel::ISupportInitialize^>(this->numericUpDown1))->BeginInit();
			this->SuspendLayout();
			// 
			// button1
			// 
			this->button1->Location = System::Drawing::Point(12, 12);
			this->button1->Name = L"button1";
			this->button1->Size = System::Drawing::Size(361, 64);
			this->button1->TabIndex = 0;
			this->button1->Text = L"Выберите каталог";
			this->button1->UseVisualStyleBackColor = true;
			this->button1->Click += gcnew System::EventHandler(this, &MyForm::button1_Click);
			// 
			// textBox1
			// 
			this->textBox1->Location = System::Drawing::Point(12, 315);
			this->textBox1->Multiline = true;
			this->textBox1->Name = L"textBox1";
			this->textBox1->ScrollBars = System::Windows::Forms::ScrollBars::Vertical;
			this->textBox1->Size = System::Drawing::Size(788, 179);
			this->textBox1->TabIndex = 1;
			// 
			// button2
			// 
			this->button2->Location = System::Drawing::Point(379, 12);
			this->button2->Name = L"button2";
			this->button2->Size = System::Drawing::Size(55, 64);
			this->button2->TabIndex = 2;
			this->button2->Text = L"output.xlxs";
			this->button2->UseVisualStyleBackColor = true;
			this->button2->Click += gcnew System::EventHandler(this, &MyForm::button2_Click);
			// 
			// button3
			// 
			this->button3->Location = System::Drawing::Point(440, 12);
			this->button3->Name = L"button3";
			this->button3->Size = System::Drawing::Size(141, 64);
			this->button3->TabIndex = 3;
			this->button3->Text = L"Открыть файл с номером, который указан -->";
			this->button3->UseVisualStyleBackColor = true;
			this->button3->Click += gcnew System::EventHandler(this, &MyForm::button3_Click);
			// 
			// numericUpDown1
			// 
			this->numericUpDown1->Location = System::Drawing::Point(587, 36);
			this->numericUpDown1->Name = L"numericUpDown1";
			this->numericUpDown1->Size = System::Drawing::Size(44, 20);
			this->numericUpDown1->TabIndex = 4;
			// 
			// listView1
			// 
			this->listView1->AllowColumnReorder = true;
			this->listView1->Columns->AddRange(gcnew cli::array< System::Windows::Forms::ColumnHeader^  >(5) {
				this->columnHeader1, this->columnHeader2,
					this->columnHeader3, this->columnHeader4, this->columnHeader5
			});
			this->listView1->FullRowSelect = true;
			this->listView1->GridLines = true;
			this->listView1->LabelEdit = true;
			this->listView1->Location = System::Drawing::Point(12, 82);
			this->listView1->MultiSelect = false;
			this->listView1->Name = L"listView1";
			this->listView1->Size = System::Drawing::Size(788, 227);
			this->listView1->Sorting = System::Windows::Forms::SortOrder::Descending;
			this->listView1->TabIndex = 5;
			this->listView1->UseCompatibleStateImageBehavior = false;
			this->listView1->View = System::Windows::Forms::View::Details;
			this->listView1->DoubleClick += gcnew System::EventHandler(this, &MyForm::listView1_DoubleClick);
			// 
			// columnHeader1
			// 
			this->columnHeader1->Text = L"Имя файла";
			this->columnHeader1->Width = 84;
			// 
			// columnHeader2
			// 
			this->columnHeader2->Text = L"Размер";
			this->columnHeader2->Width = 82;
			// 
			// columnHeader3
			// 
			this->columnHeader3->Text = L"Дата изменения";
			this->columnHeader3->Width = 115;
			// 
			// columnHeader4
			// 
			this->columnHeader4->Text = L"Адрес файла";
			this->columnHeader4->Width = 80;
			// 
			// columnHeader5
			// 
			this->columnHeader5->Text = L"Тип файла";
			this->columnHeader5->Width = 74;
			// 
			// button4
			// 
			this->button4->Location = System::Drawing::Point(587, 9);
			this->button4->Name = L"button4";
			this->button4->Size = System::Drawing::Size(53, 23);
			this->button4->TabIndex = 6;
			this->button4->Text = L"button4";
			this->button4->UseVisualStyleBackColor = true;
			this->button4->Click += gcnew System::EventHandler(this, &MyForm::button4_Click);
			// 
			// openFileDialog1
			// 
			this->openFileDialog1->FileName = L"openFileDialog1";
			// 
			// comboBox1
			// 
			this->comboBox1->FormattingEnabled = true;
			this->comboBox1->Items->AddRange(gcnew cli::array< System::Object^  >(5) {
				L"Имя файла", L"Размер", L"Дата изменения", L"Адрес файла",
					L"Тип файла"
			});
			this->comboBox1->Location = System::Drawing::Point(668, 35);
			this->comboBox1->Name = L"comboBox1";
			this->comboBox1->Size = System::Drawing::Size(121, 21);
			this->comboBox1->TabIndex = 7;
			this->comboBox1->SelectedValueChanged += gcnew System::EventHandler(this, &MyForm::comboBox1_SelectedValueChanged);
			// 
			// MyForm
			// 
			this->AutoScaleDimensions = System::Drawing::SizeF(6, 13);
			this->AutoScaleMode = System::Windows::Forms::AutoScaleMode::Font;
			this->ClientSize = System::Drawing::Size(812, 506);
			this->Controls->Add(this->comboBox1);
			this->Controls->Add(this->button4);
			this->Controls->Add(this->listView1);
			this->Controls->Add(this->numericUpDown1);
			this->Controls->Add(this->button3);
			this->Controls->Add(this->button2);
			this->Controls->Add(this->textBox1);
			this->Controls->Add(this->button1);
			this->Name = L"MyForm";
			this->Text = L"MyForm";
			this->Load += gcnew System::EventHandler(this, &MyForm::MyForm_Load);
			(cli::safe_cast<System::ComponentModel::ISupportInitialize^>(this->numericUpDown1))->EndInit();
			this->ResumeLayout(false);
			this->PerformLayout();

		}
#pragma endregion
	private: System::Void button1_Click(System::Object^  sender, System::EventArgs^  e) {

		xlnt::workbook wbOut;
		std::string dest_filename = "output.xlsx";
		xlnt::worksheet wsOut = wbOut.active_sheet();


		folderBrowserDialog1->ShowDialog();
		String^ s = folderBrowserDialog1->SelectedPath;


		std::string os;
		for (int i = 0; i < Directory::GetFiles(s)->LongLength; ++i)
		{

			FileInfo^ FI = gcnew FileInfo(Directory::GetFiles(s)[i]);
			ListViewItem^ L1= gcnew ListViewItem();

			const char* chars1 = (const char*)(Marshal::StringToHGlobalAnsi(FI->Name)).ToPointer();
			os = chars1;
			Marshal::FreeHGlobal(IntPtr((void*)chars1));
			wsOut.cell(xlnt::cell_reference(1, i + 1)).value(os);
			L1->Text = FI->Name;


			wsOut.cell(xlnt::cell_reference(2, i + 1)).value(FI->Length);
			L1->SubItems->Add(Convert::ToString(FI->Length));

			const char* chars4 = (const char*)(Marshal::StringToHGlobalAnsi(Convert::ToString(FI->LastAccessTime))).ToPointer();
			os = chars4;
			Marshal::FreeHGlobal(IntPtr((void*)chars4));
			wsOut.cell(xlnt::cell_reference(3, i + 1)).value(os);
			L1->SubItems->Add(Convert::ToString(FI->LastAccessTime));

			const char* chars2 = (const char*)(Marshal::StringToHGlobalAnsi(FI->FullName)).ToPointer();
			os = chars2;
			Marshal::FreeHGlobal(IntPtr((void*)chars2));
			wsOut.cell(xlnt::cell_reference(4, i + 1)).value(os);
			L1->SubItems->Add(FI->FullName);

			const char* chars3 = (const char*)(Marshal::StringToHGlobalAnsi(FI->Extension)).ToPointer();
			os = chars3;
			Marshal::FreeHGlobal(IntPtr((void*)chars3));
			wsOut.cell(xlnt::cell_reference(5, i + 1)).value(os);
			L1->SubItems->Add(FI->Extension);

			listView1->Items->Add(L1);

			//textBox1->Text = textBox1->Text + Convert::ToString(i+1) + ":  " + FI->ToString() + "\r\n";
		}




		numericUpDown1->Minimum = 1;
		numericUpDown1->Maximum = Directory::GetFiles(s)->LongLength + 1;
		wbOut.save(dest_filename);
	};
	private: System::Void button2_Click(System::Object^  sender, System::EventArgs^  e) {
		std::string s = "output.xlsx";
		int len;
		int slength = (int)s.length() + 1;
		len = MultiByteToWideChar(CP_ACP, 0, s.c_str(), slength, 0, 0);
		wchar_t* buf = new wchar_t[len];
		MultiByteToWideChar(CP_ACP, 0, s.c_str(), slength, buf, len);
		std::wstring r(buf);
		delete[] buf;
		std::wstring stemp = r;
		ShellExecute(NULL, NULL, stemp.c_str(), NULL, NULL, SW_RESTORE);
	}
private: System::Void button3_Click(System::Object^  sender, System::EventArgs^  e) {
	xlnt::workbook wb;
	wb.load("output.xlsx");
	auto ws = wb.active_sheet();
	std::vector< std::vector<std::string> > theWholeSpreadSheet;
	for (auto row : ws.rows(false))
	{
		std::vector<std::string> aSingleRow;
		for (auto cell : row)
		{
			aSingleRow.push_back(cell.to_string());
		}
		theWholeSpreadSheet.push_back(aSingleRow);
	}


	int len;
	int slength = (int)theWholeSpreadSheet.at(Convert::ToInt16(numericUpDown1->Value)-1).at(0).length() + 1;
	len = MultiByteToWideChar(CP_ACP, 0, theWholeSpreadSheet.at(Convert::ToInt16(numericUpDown1->Value)-1).at(0).c_str(), slength, 0, 0);
	wchar_t* buf = new wchar_t[len];
	MultiByteToWideChar(CP_ACP, 0, theWholeSpreadSheet.at(Convert::ToInt16(numericUpDown1->Value)-1).at(0).c_str(), slength, buf, len);
	std::wstring r(buf);
	delete[] buf;
	std::wstring stemp = r;


	ShellExecute(NULL, NULL, stemp.c_str(), NULL, NULL, SW_RESTORE);

}
private: System::Void listView1_DoubleClick(System::Object^  sender, System::EventArgs^  e) {
	for (int i = 0; i < sizeof(listView1->Items) + 1; i++) {
		if (listView1->Items[i]->Selected == true)
		{
			textBox1->Text = listView1->Items[i]->Text;
				//ShellExecute(NULL, NULL, stemp.c_str(), NULL, NULL, SW_RESTORE);
		}
	}
}
private: System::Void comboBox1_SelectedValueChanged(System::Object^  sender, System::EventArgs^  e) {
	//for(int i =0;i<sizeof(listView1->Items)+1;i++)
		//textBox1->Text = textBox1->Text + listView1->Items[i]->SubItems[comboBox1->SelectedIndex]->Text + "\r\n";
	xlnt::workbook wb;
	wb.load("output.xlsx");
	xlnt::worksheet ws = wb.active_sheet();
	std::string buff[5];


	char bf;
	for(int k = sizeof(listView1->Items);k>=0;k--)
		for (int i = 0; i < k; i++)
		{

			std::string buff1 = ws.cell(xlnt::cell_reference(1, i)).to_string();
			std::string buff2 = ws.cell(xlnt::cell_reference(1, i + 1)).to_string();
			if (buff1 > buff2)
			{
				for (int k = 0; k < 5; k++)
					buff[k] = ws.cell(xlnt::cell_reference(1, k + 1)).to_string();
				for (int k = 0; k < 5; k++)
					ws.cell(xlnt::cell_reference(1, k)).to_string() = ws.cell(xlnt::cell_reference(1, k + 1)).to_string();
				for (int k = 0; k < 5; k++)
					ws.cell(xlnt::cell_reference(1, k + 1)).to_string() = buff[k];
			}
		}
	}
private: System::Void MyForm_Load(System::Object^  sender, System::EventArgs^  e) {
	//ListViewItem lv1 = new ListViewItem("lalalala",3);
	//lv1.SubItems->Add("123");
}
private: System::Void button4_Click(System::Object^  sender, System::EventArgs^  e) {
	
	
	
	xlnt::workbook wb;
	wb.load("output.xlsx");
	xlnt::worksheet ws = wb.active_sheet();
	//std::string i = ws.cell("A1").has_value();

	System::String^ fd = marshal_as<System::String^>(ws.cell(xlnt::cell_reference( 1, 4)).to_string());
	//String^ str3 = gcnew String(ws.cell("A1").to_string().c_str());
	
	textBox1->Text = textBox1->Text +fd;
	//textBox1->Text = comboBox1->Text;//Convert::ToString(comboBox1->SelectedValue);
	//comboBox1->Items[1];

	/*// Create three items and three sets of subitems for each item.
	ListViewItem^ item1 = gcnew ListViewItem("item1", 0);

	// Place a check mark next to the item.
	//item1->Checked = true;
	item1->SubItems->Add("1");
	item1->SubItems->Add("2");
	//item1->SubItems->Add("3");
	ListViewItem^ item2 = gcnew ListViewItem("item2", 1);
	item2->SubItems->Add("4");
	item2->SubItems->Add("5");
	//item2->SubItems->Add("6");
	ListViewItem^ item3 = gcnew ListViewItem("item3", 0);

	// Place a check mark next to the item.
	//item3->Checked = true;
	item3->SubItems->Add("7");
	item3->SubItems->Add("8");
	//item3->SubItems->Add("9");

	// Create columns for the items and subitems.
	// Width of -2 indicates auto-size.
	//listView1->Columns->Add("Item Column", -2, HorizontalAlignment::Left);
	//listView1->Columns->Add("Column 2", -2, HorizontalAlignment::Left);
	//listView1->Columns->Add("Column 3", -2, HorizontalAlignment::Left);
	//listView1->Columns->Add("Column 4", -2, HorizontalAlignment::Center);

	//Add the items to the ListView.
	array<ListViewItem^>^temp1 = { item1,item2,item3 };
	listView1->Items->AddRange(temp1);

	// Create two ImageList objects.
	/*ImageList^ imageListSmall = gcnew ImageList;
	ImageList^ imageListLarge = gcnew ImageList;

	// Initialize the ImageList objects with bitmaps.
	imageListSmall->Images->Add(Bitmap::FromFile("C:\\MySmallImage1.bmp"));
	imageListSmall->Images->Add(Bitmap::FromFile("C:\\MySmallImage2.bmp"));
	imageListLarge->Images->Add(Bitmap::FromFile("C:\\MyLargeImage1.bmp"));
	imageListLarge->Images->Add(Bitmap::FromFile("C:\\MyLargeImage2.bmp"));

	//Assign the ImageList objects to the ListView.
	listView1->LargeImageList = imageListLarge;
	listView1->SmallImageList = imageListSmall;*/

	// Add the ListView to the control collection.
	//this->Controls->Add(listView1);
	
	
	
	
	
	/*openFileDialog1->ShowDialog();
	
	FileInfo^ fi = gcnew FileInfo(openFileDialog1->FileName);

	textBox1->Text = textBox1->Text + Convert::ToString(fi->Length) + "\r\n";
	textBox1->Text = textBox1->Text + Convert::ToString(fi->CreationTime.Month) + "\r\n";
	textBox1->Text = textBox1->Text + Convert::ToString(fi->CreationTime.Day) + "\r\n";
	textBox1->Text = textBox1->Text + Convert::ToString(fi->CreationTime.Year) + "\r\n";
	textBox1->Text = textBox1->Text + Convert::ToString(fi->LastAccessTime.Month) + "\r\n";
	textBox1->Text = textBox1->Text + Convert::ToString(fi->LastAccessTime.Day) + "\r\n";
	textBox1->Text = textBox1->Text + Convert::ToString(fi->LastAccessTime.Year) + "\r\n";

	//Console::WriteLine("file size: {0}", fi->Length);

	//Console::Write("File creation date:  ");
	//Console::Write(fi->CreationTime.Month.ToString());
	//Console::Write(".{0}", fi->CreationTime.Day.ToString());
	//Console::WriteLine(".{0}", fi->CreationTime.Year.ToString());

	//Console::Write("Last access date:  ");
	//Console::Write(fi->LastAccessTime.Month.ToString());
	//Console::Write(".{0}", fi->LastAccessTime.Day.ToString());
	//Console::WriteLine(".{0}", fi->LastAccessTime.Year.ToString());
	//system("pause");*/

}

};
}



/*xlnt::workbook wb;
	wb.load("output.xlsx");
	auto ws = wb.active_sheet();
	String^ str;
	for (auto row : ws.rows(false))
	{
		for (auto cell : row)
		{
			str = gcnew String(cell.to_string().c_str());
			textBox1->Text = textBox1->Text + str + "\r\n";
		}
	}
}*/


/*for (int rowInt = 0; rowInt < theWholeSpreadSheet.size(); rowInt++)
	{
		for (int colInt = 0; colInt < theWholeSpreadSheet.at(rowInt).size(); colInt++)
		{
			str = gcnew String(theWholeSpreadSheet.at(rowInt).at(colInt).c_str()); // colInt = const
			textBox1->Text = textBox1->Text + str + "\r\n";
		}
	}*/