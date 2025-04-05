#pragma once
#include "AuthForm.h"
#include "GrafikForm.h"
#include "QueryForm.h"

namespace EmployeeAttendanceApp {

	using namespace System;
	using namespace System::ComponentModel;
	using namespace System::Collections;
	using namespace System::Windows::Forms;
	using namespace System::Data;
	using namespace System::Drawing;

	/// <summary>
	/// Summary for MainForm
	/// </summary>

	public ref class MainForm : public System::Windows::Forms::Form
	{
	public:
		MainForm(void)
		{
			InitializeComponent();
			//
			//TODO: Add the constructor code here
			//
			this->dataGridViewEmployees->CellValueChanged +=
				gcnew System::Windows::Forms::DataGridViewCellEventHandler(
					this, &MainForm::dataGridViewEmployees_CellValueChanged);

			this->dataGridViewEmployees->CellFormatting +=
				gcnew System::Windows::Forms::DataGridViewCellFormattingEventHandler(
					this, &MainForm::dataGridViewEmployees_CellFormatting);
		}

	protected:
		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		~MainForm()
		{
			if (components)
			{
				delete components;
			}
		}
	private: System::Windows::Forms::DataGridView^ dataGridViewEmployees;
	private: System::Windows::Forms::Button^ buttonLoad;
	private: System::Windows::Forms::Button^ buttonSave;
	private: System::Windows::Forms::Button^ buttonAddRow;
	private: System::Windows::Forms::Button^ buttonDeleteRow;
	private: System::Windows::Forms::MenuStrip^ menuStrip1;

	private: System::Windows::Forms::ToolStripMenuItem^ сохранитьToolStripMenuItem;
	private: System::Windows::Forms::Button^ authCard;
	String^ userInput;
	private: System::Windows::Forms::Button^ new_day;
	private: System::Windows::Forms::Label^ t_date;
	private: System::Windows::Forms::DataGridViewTextBoxColumn^ cardid;
	private: System::Windows::Forms::DataGridViewTextBoxColumn^ NameColumn;
	private: System::Windows::Forms::DataGridViewTextBoxColumn^ carddata;
	private: System::Windows::Forms::DataGridViewTextBoxColumn^ otdel;
	private: System::Windows::Forms::DataGridViewTextBoxColumn^ worktime;
	private: System::Windows::Forms::DataGridViewTextBoxColumn^ ArrivalColumn;
	private: System::Windows::Forms::DataGridViewTextBoxColumn^ LeaveColumn;
	private: System::Drawing::Printing::PrintDocument^ printDocument;

	private: System::Windows::Forms::PrintPreviewDialog^ printPreviewDialog;
	private: System::Windows::Forms::Button^ btnPrint;
	private: System::Windows::Forms::Button^ btnShowChart;
	private: System::Windows::Forms::Button^ btnOpenQueries;


	protected:
	private:
		/// <summary>
		/// Required designer variable.
		/// </summary>
		System::ComponentModel::Container ^components;

#pragma region Windows Form Designer generated code
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		void InitializeComponent(void)
		{
			System::ComponentModel::ComponentResourceManager^ resources = (gcnew System::ComponentModel::ComponentResourceManager(MainForm::typeid));
			this->dataGridViewEmployees = (gcnew System::Windows::Forms::DataGridView());
			this->cardid = (gcnew System::Windows::Forms::DataGridViewTextBoxColumn());
			this->NameColumn = (gcnew System::Windows::Forms::DataGridViewTextBoxColumn());
			this->carddata = (gcnew System::Windows::Forms::DataGridViewTextBoxColumn());
			this->otdel = (gcnew System::Windows::Forms::DataGridViewTextBoxColumn());
			this->worktime = (gcnew System::Windows::Forms::DataGridViewTextBoxColumn());
			this->ArrivalColumn = (gcnew System::Windows::Forms::DataGridViewTextBoxColumn());
			this->LeaveColumn = (gcnew System::Windows::Forms::DataGridViewTextBoxColumn());
			this->buttonLoad = (gcnew System::Windows::Forms::Button());
			this->buttonSave = (gcnew System::Windows::Forms::Button());
			this->buttonAddRow = (gcnew System::Windows::Forms::Button());
			this->buttonDeleteRow = (gcnew System::Windows::Forms::Button());
			this->menuStrip1 = (gcnew System::Windows::Forms::MenuStrip());
			this->сохранитьToolStripMenuItem = (gcnew System::Windows::Forms::ToolStripMenuItem());
			this->authCard = (gcnew System::Windows::Forms::Button());
			this->new_day = (gcnew System::Windows::Forms::Button());
			this->t_date = (gcnew System::Windows::Forms::Label());
			this->printDocument = (gcnew System::Drawing::Printing::PrintDocument());
			this->printPreviewDialog = (gcnew System::Windows::Forms::PrintPreviewDialog());
			this->btnPrint = (gcnew System::Windows::Forms::Button());
			this->btnShowChart = (gcnew System::Windows::Forms::Button());
			this->btnOpenQueries = (gcnew System::Windows::Forms::Button());
			(cli::safe_cast<System::ComponentModel::ISupportInitialize^>(this->dataGridViewEmployees))->BeginInit();
			this->menuStrip1->SuspendLayout();
			this->SuspendLayout();
			// 
			// dataGridViewEmployees
			// 
			this->dataGridViewEmployees->AutoSizeColumnsMode = System::Windows::Forms::DataGridViewAutoSizeColumnsMode::AllCells;
			this->dataGridViewEmployees->AutoSizeRowsMode = System::Windows::Forms::DataGridViewAutoSizeRowsMode::AllCells;
			this->dataGridViewEmployees->BackgroundColor = System::Drawing::Color::PaleGoldenrod;
			this->dataGridViewEmployees->ColumnHeadersHeightSizeMode = System::Windows::Forms::DataGridViewColumnHeadersHeightSizeMode::AutoSize;
			this->dataGridViewEmployees->Columns->AddRange(gcnew cli::array< System::Windows::Forms::DataGridViewColumn^  >(7) {
				this->cardid,
					this->NameColumn, this->carddata, this->otdel, this->worktime, this->ArrivalColumn, this->LeaveColumn
			});
			this->dataGridViewEmployees->Dock = System::Windows::Forms::DockStyle::Fill;
			this->dataGridViewEmployees->GridColor = System::Drawing::Color::DimGray;
			this->dataGridViewEmployees->Location = System::Drawing::Point(0, 24);
			this->dataGridViewEmployees->Name = L"dataGridViewEmployees";
			this->dataGridViewEmployees->RowTemplate->Resizable = System::Windows::Forms::DataGridViewTriState::True;
			this->dataGridViewEmployees->Size = System::Drawing::Size(1584, 737);
			this->dataGridViewEmployees->TabIndex = 0;
			this->dataGridViewEmployees->CellContentClick += gcnew System::Windows::Forms::DataGridViewCellEventHandler(this, &MainForm::dataGridView1_CellContentClick);
			this->dataGridViewEmployees->RowValidating += gcnew System::Windows::Forms::DataGridViewCellCancelEventHandler(this, &MainForm::dataGridViewEmployees_RowValidating);
			// 
			// cardid
			// 
			this->cardid->HeaderText = L"Номер карты";
			this->cardid->Name = L"cardid";
			this->cardid->Width = 92;
			// 
			// NameColumn
			// 
			this->NameColumn->HeaderText = L"ФИО";
			this->NameColumn->Name = L"NameColumn";
			this->NameColumn->Width = 59;
			// 
			// carddata
			// 
			this->carddata->HeaderText = L"Срок действия карты";
			this->carddata->Name = L"carddata";
			this->carddata->Width = 129;
			// 
			// otdel
			// 
			this->otdel->HeaderText = L"Отдел";
			this->otdel->Name = L"otdel";
			this->otdel->Width = 63;
			// 
			// worktime
			// 
			this->worktime->HeaderText = L"Время рабочего дня";
			this->worktime->Name = L"worktime";
			this->worktime->Resizable = System::Windows::Forms::DataGridViewTriState::True;
			this->worktime->Width = 107;
			// 
			// ArrivalColumn
			// 
			this->ArrivalColumn->HeaderText = L"Время прихода";
			this->ArrivalColumn->Name = L"ArrivalColumn";
			// 
			// LeaveColumn
			// 
			this->LeaveColumn->HeaderText = L"Время ухода";
			this->LeaveColumn->Name = L"LeaveColumn";
			this->LeaveColumn->Width = 88;
			// 
			// buttonLoad
			// 
			this->buttonLoad->BackColor = System::Drawing::Color::FromArgb(static_cast<System::Int32>(static_cast<System::Byte>(255)), static_cast<System::Int32>(static_cast<System::Byte>(255)),
				static_cast<System::Int32>(static_cast<System::Byte>(192)));
			this->buttonLoad->FlatStyle = System::Windows::Forms::FlatStyle::Flat;
			this->buttonLoad->ForeColor = System::Drawing::Color::Brown;
			this->buttonLoad->Location = System::Drawing::Point(486, 0);
			this->buttonLoad->Name = L"buttonLoad";
			this->buttonLoad->Size = System::Drawing::Size(143, 24);
			this->buttonLoad->TabIndex = 1;
			this->buttonLoad->Text = L"Загрузить из файла";
			this->buttonLoad->UseVisualStyleBackColor = false;
			this->buttonLoad->Click += gcnew System::EventHandler(this, &MainForm::buttonload_Click);
			// 
			// buttonSave
			// 
			this->buttonSave->BackColor = System::Drawing::Color::FromArgb(static_cast<System::Int32>(static_cast<System::Byte>(255)), static_cast<System::Int32>(static_cast<System::Byte>(255)),
				static_cast<System::Int32>(static_cast<System::Byte>(192)));
			this->buttonSave->FlatStyle = System::Windows::Forms::FlatStyle::Flat;
			this->buttonSave->ForeColor = System::Drawing::Color::Brown;
			this->buttonSave->Location = System::Drawing::Point(337, 0);
			this->buttonSave->Name = L"buttonSave";
			this->buttonSave->RightToLeft = System::Windows::Forms::RightToLeft::No;
			this->buttonSave->Size = System::Drawing::Size(143, 24);
			this->buttonSave->TabIndex = 2;
			this->buttonSave->Text = L"Сохранить в файл";
			this->buttonSave->UseVisualStyleBackColor = false;
			this->buttonSave->Click += gcnew System::EventHandler(this, &MainForm::buttonSave_Click);
			// 
			// buttonAddRow
			// 
			this->buttonAddRow->BackColor = System::Drawing::Color::FromArgb(static_cast<System::Int32>(static_cast<System::Byte>(255)), static_cast<System::Int32>(static_cast<System::Byte>(255)),
				static_cast<System::Int32>(static_cast<System::Byte>(192)));
			this->buttonAddRow->FlatStyle = System::Windows::Forms::FlatStyle::Flat;
			this->buttonAddRow->ForeColor = System::Drawing::Color::Brown;
			this->buttonAddRow->Location = System::Drawing::Point(39, 0);
			this->buttonAddRow->Name = L"buttonAddRow";
			this->buttonAddRow->Size = System::Drawing::Size(143, 24);
			this->buttonAddRow->TabIndex = 3;
			this->buttonAddRow->Text = L"Добавить строку";
			this->buttonAddRow->UseVisualStyleBackColor = false;
			this->buttonAddRow->Click += gcnew System::EventHandler(this, &MainForm::buttonAddRow_Click);
			// 
			// buttonDeleteRow
			// 
			this->buttonDeleteRow->BackColor = System::Drawing::Color::FromArgb(static_cast<System::Int32>(static_cast<System::Byte>(255)), static_cast<System::Int32>(static_cast<System::Byte>(255)),
				static_cast<System::Int32>(static_cast<System::Byte>(192)));
			this->buttonDeleteRow->FlatStyle = System::Windows::Forms::FlatStyle::Flat;
			this->buttonDeleteRow->ForeColor = System::Drawing::Color::Brown;
			this->buttonDeleteRow->Location = System::Drawing::Point(188, 0);
			this->buttonDeleteRow->Name = L"buttonDeleteRow";
			this->buttonDeleteRow->Size = System::Drawing::Size(143, 24);
			this->buttonDeleteRow->TabIndex = 4;
			this->buttonDeleteRow->Text = L"Удалить строку";
			this->buttonDeleteRow->UseVisualStyleBackColor = false;
			this->buttonDeleteRow->Click += gcnew System::EventHandler(this, &MainForm::buttonDeleteRow_Click);
			// 
			// menuStrip1
			// 
			this->menuStrip1->Items->AddRange(gcnew cli::array< System::Windows::Forms::ToolStripItem^  >(1) { this->сохранитьToolStripMenuItem });
			this->menuStrip1->Location = System::Drawing::Point(0, 0);
			this->menuStrip1->Name = L"menuStrip1";
			this->menuStrip1->Size = System::Drawing::Size(1584, 24);
			this->menuStrip1->TabIndex = 5;
			this->menuStrip1->Text = L"menuStrip1";
			// 
			// сохранитьToolStripMenuItem
			// 
			this->сохранитьToolStripMenuItem->Image = (cli::safe_cast<System::Drawing::Image^>(resources->GetObject(L"сохранитьToolStripMenuItem.Image")));
			this->сохранитьToolStripMenuItem->Name = L"сохранитьToolStripMenuItem";
			this->сохранитьToolStripMenuItem->Size = System::Drawing::Size(28, 20);
			this->сохранитьToolStripMenuItem->TextImageRelation = System::Windows::Forms::TextImageRelation::ImageAboveText;
			this->сохранитьToolStripMenuItem->Click += gcnew System::EventHandler(this, &MainForm::сохранитьToolStripMenuItem_Click);
			// 
			// authCard
			// 
			this->authCard->BackColor = System::Drawing::Color::FromArgb(static_cast<System::Int32>(static_cast<System::Byte>(255)), static_cast<System::Int32>(static_cast<System::Byte>(255)),
				static_cast<System::Int32>(static_cast<System::Byte>(192)));
			this->authCard->FlatStyle = System::Windows::Forms::FlatStyle::Flat;
			this->authCard->ForeColor = System::Drawing::Color::Brown;
			this->authCard->Location = System::Drawing::Point(635, 0);
			this->authCard->Name = L"authCard";
			this->authCard->Size = System::Drawing::Size(143, 24);
			this->authCard->TabIndex = 6;
			this->authCard->Text = L"Авторизация карты";
			this->authCard->UseVisualStyleBackColor = false;
			this->authCard->Click += gcnew System::EventHandler(this, &MainForm::authCard_Click);
			// 
			// new_day
			// 
			this->new_day->BackColor = System::Drawing::Color::FromArgb(static_cast<System::Int32>(static_cast<System::Byte>(255)), static_cast<System::Int32>(static_cast<System::Byte>(255)),
				static_cast<System::Int32>(static_cast<System::Byte>(192)));
			this->new_day->FlatStyle = System::Windows::Forms::FlatStyle::Flat;
			this->new_day->ForeColor = System::Drawing::Color::Brown;
			this->new_day->Location = System::Drawing::Point(784, 0);
			this->new_day->Name = L"new_day";
			this->new_day->Size = System::Drawing::Size(143, 24);
			this->new_day->TabIndex = 7;
			this->new_day->Text = L"Начать новый день";
			this->new_day->UseVisualStyleBackColor = false;
			this->new_day->Click += gcnew System::EventHandler(this, &MainForm::new_day_Click);
			// 
			// t_date
			// 
			this->t_date->AutoSize = true;
			this->t_date->BackColor = System::Drawing::Color::White;
			this->t_date->Location = System::Drawing::Point(1380, 5);
			this->t_date->Name = L"t_date";
			this->t_date->Size = System::Drawing::Size(33, 13);
			this->t_date->TabIndex = 8;
			this->t_date->Text = L"Дата";
			// 
			// printDocument
			// 
			this->printDocument->PrintPage += gcnew System::Drawing::Printing::PrintPageEventHandler(this, &MainForm::printDocument_PrintPage);
			// 
			// printPreviewDialog
			// 
			this->printPreviewDialog->AutoScrollMargin = System::Drawing::Size(0, 0);
			this->printPreviewDialog->AutoScrollMinSize = System::Drawing::Size(0, 0);
			this->printPreviewDialog->ClientSize = System::Drawing::Size(1000, 800);
			this->printPreviewDialog->Enabled = true;
			this->printPreviewDialog->Icon = (cli::safe_cast<System::Drawing::Icon^>(resources->GetObject(L"printPreviewDialog.Icon")));
			this->printPreviewDialog->Name = L"printPreviewDialog";
			this->printPreviewDialog->ShowIcon = false;
			this->printPreviewDialog->StartPosition = System::Windows::Forms::FormStartPosition::CenterParent;
			this->printPreviewDialog->Text = L"Предпросмотр табеля";
			this->printPreviewDialog->Visible = false;
			// 
			// btnPrint
			// 
			this->btnPrint->BackColor = System::Drawing::Color::FromArgb(static_cast<System::Int32>(static_cast<System::Byte>(255)), static_cast<System::Int32>(static_cast<System::Byte>(255)),
				static_cast<System::Int32>(static_cast<System::Byte>(192)));
			this->btnPrint->FlatStyle = System::Windows::Forms::FlatStyle::Flat;
			this->btnPrint->ForeColor = System::Drawing::Color::Brown;
			this->btnPrint->Location = System::Drawing::Point(933, 0);
			this->btnPrint->Name = L"btnPrint";
			this->btnPrint->Size = System::Drawing::Size(143, 24);
			this->btnPrint->TabIndex = 9;
			this->btnPrint->Text = L"Печать";
			this->btnPrint->UseVisualStyleBackColor = false;
			this->btnPrint->Click += gcnew System::EventHandler(this, &MainForm::btnPrint_Click);
			// 
			// btnShowChart
			// 
			this->btnShowChart->BackColor = System::Drawing::Color::FromArgb(static_cast<System::Int32>(static_cast<System::Byte>(255)), static_cast<System::Int32>(static_cast<System::Byte>(255)),
				static_cast<System::Int32>(static_cast<System::Byte>(192)));
			this->btnShowChart->FlatStyle = System::Windows::Forms::FlatStyle::Flat;
			this->btnShowChart->ForeColor = System::Drawing::Color::Brown;
			this->btnShowChart->Location = System::Drawing::Point(1082, 0);
			this->btnShowChart->Name = L"btnShowChart";
			this->btnShowChart->Size = System::Drawing::Size(143, 24);
			this->btnShowChart->TabIndex = 10;
			this->btnShowChart->Text = L"Показать график";
			this->btnShowChart->UseVisualStyleBackColor = false;
			this->btnShowChart->Click += gcnew System::EventHandler(this, &MainForm::btnShowChart_Click);
			// 
			// btnOpenQueries
			// 
			this->btnOpenQueries->BackColor = System::Drawing::Color::FromArgb(static_cast<System::Int32>(static_cast<System::Byte>(255)), static_cast<System::Int32>(static_cast<System::Byte>(255)),
				static_cast<System::Int32>(static_cast<System::Byte>(192)));
			this->btnOpenQueries->FlatStyle = System::Windows::Forms::FlatStyle::Flat;
			this->btnOpenQueries->ForeColor = System::Drawing::Color::Brown;
			this->btnOpenQueries->Location = System::Drawing::Point(1231, 0);
			this->btnOpenQueries->Name = L"btnOpenQueries";
			this->btnOpenQueries->Size = System::Drawing::Size(143, 23);
			this->btnOpenQueries->TabIndex = 11;
			this->btnOpenQueries->Text = L"Найти...";
			this->btnOpenQueries->UseVisualStyleBackColor = false;
			this->btnOpenQueries->Click += gcnew System::EventHandler(this, &MainForm::btnOpenQueries_Click);
			// 
			// MainForm
			// 
			this->AutoScaleDimensions = System::Drawing::SizeF(6, 13);
			this->AutoScaleMode = System::Windows::Forms::AutoScaleMode::Font;
			this->BackColor = System::Drawing::Color::FromArgb(static_cast<System::Int32>(static_cast<System::Byte>(255)), static_cast<System::Int32>(static_cast<System::Byte>(255)),
				static_cast<System::Int32>(static_cast<System::Byte>(128)));
			this->ClientSize = System::Drawing::Size(1584, 761);
			this->Controls->Add(this->btnOpenQueries);
			this->Controls->Add(this->btnShowChart);
			this->Controls->Add(this->btnPrint);
			this->Controls->Add(this->t_date);
			this->Controls->Add(this->new_day);
			this->Controls->Add(this->authCard);
			this->Controls->Add(this->buttonDeleteRow);
			this->Controls->Add(this->buttonAddRow);
			this->Controls->Add(this->buttonSave);
			this->Controls->Add(this->buttonLoad);
			this->Controls->Add(this->dataGridViewEmployees);
			this->Controls->Add(this->menuStrip1);
			this->Icon = (cli::safe_cast<System::Drawing::Icon^>(resources->GetObject(L"$this.Icon")));
			this->MainMenuStrip = this->menuStrip1;
			this->Name = L"MainForm";
			this->StartPosition = System::Windows::Forms::FormStartPosition::CenterScreen;
			this->Text = L"2с: Учёт сотрудников";
			this->WindowState = System::Windows::Forms::FormWindowState::Maximized;
			this->Load += gcnew System::EventHandler(this, &MainForm::MainForm_Load);
			(cli::safe_cast<System::ComponentModel::ISupportInitialize^>(this->dataGridViewEmployees))->EndInit();
			this->menuStrip1->ResumeLayout(false);
			this->menuStrip1->PerformLayout();
			this->ResumeLayout(false);
			this->PerformLayout();
		}
#pragma endregion
	private: System::Void dataGridView1_CellContentClick(System::Object^ sender, System::Windows::Forms::DataGridViewCellEventArgs^ e) {
	}
	private: System::Void buttonload_Click(System::Object^ sender, System::EventArgs^ e) {
		OpenFileDialog^ openFileDialog = gcnew OpenFileDialog();
		openFileDialog->Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*";
		if (openFileDialog->ShowDialog() == System::Windows::Forms::DialogResult::OK) {
			dataGridViewEmployees->Rows->Clear();
			array<String^>^ lines = System::IO::File::ReadAllLines(openFileDialog->FileName);
			for each (String ^ line in lines) {
				array<String^>^ parts = line->Split(';');
				dataGridViewEmployees->Rows->Add(parts);
			}
		}
	}
	private: System::Void buttonSave_Click(System::Object^ sender, System::EventArgs^ e) {
		SaveFileDialog^ saveFileDialog = gcnew SaveFileDialog();
		saveFileDialog->Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*";
		if (saveFileDialog->ShowDialog() == System::Windows::Forms::DialogResult::OK) {

			System::Collections::Generic::List<String^>^ lines = gcnew System::Collections::Generic::List<String^>();

			for each (DataGridViewRow ^ row in dataGridViewEmployees->Rows) {
				if (!row->IsNewRow) {
					System::Text::StringBuilder^ sb = gcnew System::Text::StringBuilder();
					for each (DataGridViewCell ^ cell in row->Cells) {
						sb->Append(cell->Value);
						sb->Append(";");
					}
					lines->Add(sb->ToString()->TrimEnd(';'));
				}
			}
			System::IO::File::WriteAllLines(saveFileDialog->FileName, lines->ToArray());
		}
	}
private: System::Void buttonAddRow_Click(System::Object^ sender, System::EventArgs^ e) {
	dataGridViewEmployees->Rows->Add();
}
private: System::Void buttonDeleteRow_Click(System::Object^ sender, System::EventArgs^ e) {
	if (dataGridViewEmployees->SelectedRows->Count > 0) {
		if (dataGridViewEmployees->SelectedRows->Count > 0) {
			auto row = dataGridViewEmployees->SelectedRows[0];
			if (!row->IsNewRow) {
				dataGridViewEmployees->Rows->Remove(row);
			}
			else {
				MessageBox::Show("Нельзя удалить незаполненную новую строку");
			}
		}
		else {
			MessageBox::Show("Пожалуйста, выберите строку для удаления");
		}
	}
}
private: System::Void dataGridViewEmployees_RowValidating(System::Object^ sender, System::Windows::Forms::DataGridViewCellCancelEventArgs^ e) {
  DataGridViewRow^ row = dataGridViewEmployees->Rows[e->RowIndex];
    if (row->IsNewRow) return; // Пропускаем пустую строку

    auto name = row->Cells["NameColumn"]->Value;
    auto arrivalDate = row->Cells["ArrivalColumn"]->Value;
    auto leaveDate = row->Cells["LeaveColumn"]->Value;
}
private: System::Void сохранитьToolStripMenuItem_Click(System::Object^ sender, System::EventArgs^ e) {
	CheckCardExpiration();
	String^ documentsPath = Environment::GetFolderPath(Environment::SpecialFolder::MyDocuments);
	String^ filePath = System::IO::Path::Combine(documentsPath, "data.txt");
	System::Collections::Generic::List<String^>^ lines = gcnew System::Collections::Generic::List<String^>();
	lines->Add(t_date->Text);
	for each (DataGridViewRow ^ row in dataGridViewEmployees->Rows) {
		if (!row->IsNewRow) {
			System::Text::StringBuilder^ sb = gcnew System::Text::StringBuilder();
			for each (DataGridViewCell ^ cell in row->Cells) {
				sb->Append(cell->Value ? cell->Value->ToString() : "");
				sb->Append(";");
			}
			lines->Add(sb->ToString()->TrimEnd(';'));
		}
	}
	System::IO::File::WriteAllLines(filePath, lines->ToArray());
}
private: System::Void MainForm_Load(System::Object^ sender, System::EventArgs^ e) {
	String^ documentsPath = Environment::GetFolderPath(Environment::SpecialFolder::MyDocuments);
	String^ filePath = System::IO::Path::Combine(documentsPath, "data.txt");
	CheckCardExpiration();
	if (System::IO::File::Exists(filePath)) {
		try {
			array<String^>^ lines = System::IO::File::ReadAllLines(filePath);
			if (lines->Length > 0) {
				t_date->Text = lines[0];
				for (int i = 1; i < lines->Length; i++) {
					array<String^>^ parts = lines[i]->Split(';');
					dataGridViewEmployees->Rows->Add(parts);
				}
			}
		}
		catch (Exception^ ex) {
			MessageBox::Show("Ошибка при чтении файла: " + ex->Message);
			t_date->Text = DateTime::Now.ToString("dd.MM.yyyy");
		}
	}
	else {
		try {
			array<String^>^ lines = { DateTime::Now.ToString("dd.MM.yyyy") };
			System::IO::File::WriteAllLines(filePath, lines);
			t_date->Text = lines[0];
			MessageBox::Show("Файл данных не был найден, был создан новый файл в " + filePath);
		}
		catch (Exception^ ex) {
			MessageBox::Show("Ошибка при создании файла: " + ex->Message);
			t_date->Text = DateTime::Now.ToString("dd.MM.yyyy");
		}
	}
}

private: System::Void authCard_Click(System::Object^ sender, System::EventArgs^ e) {
	AuthForm^ authForm = gcnew AuthForm();
	if (authForm->ShowDialog() == System::Windows::Forms::DialogResult::OK) {
		String^ cardNumber = authForm->getUserInput();
		bool found = false;

		for each (DataGridViewRow ^ row in dataGridViewEmployees->Rows) {
			if (!row->IsNewRow && row->Cells["cardid"]->Value != nullptr &&
				row->Cells["cardid"]->Value->ToString() == cardNumber) {
				found = true;
				DateTime currentTime = DateTime::Now;
				String^ currentDate = currentTime.ToString("yyyy-MM-dd");
				bool hasArrivalToday = false;
				DateTime lastArrivalTime;
				// Получаем последнюю запись прихода, если есть
				if (row->Cells["ArrivalColumn"]->Value != nullptr &&
					!String::IsNullOrEmpty(row->Cells["ArrivalColumn"]->Value->ToString())) {
					array<String^>^ arrivals = row->Cells["ArrivalColumn"]->Value->ToString()->Split(
						gcnew array<String^>{" | "}, StringSplitOptions::RemoveEmptyEntries);

					if (arrivals->Length > 0) {
						String^ lastArrival = arrivals[arrivals->Length - 1];
						if (DateTime::TryParse(lastArrival, lastArrivalTime)) {
							String^ lastArrivalDate = lastArrivalTime.ToString("yyyy-MM-dd");
							hasArrivalToday = (lastArrivalDate == currentDate);
						}
					}
				}
				if (!hasArrivalToday && row->Cells["ArrivalColumn"]->Value != nullptr) {
					// Обнуляем рабочее время при смене дня
					row->Cells["worktime"]->Value = "0";
				}

				// Определяем, нужно ли записывать приход или уход
				bool shouldRecordArrival = true;
				if (row->Cells["ArrivalColumn"]->Value != nullptr &&
					!String::IsNullOrEmpty(row->Cells["ArrivalColumn"]->Value->ToString())) {
					array<String^>^ arrivals = row->Cells["ArrivalColumn"]->Value->ToString()->Split(
						gcnew array<String^>{" | "}, StringSplitOptions::RemoveEmptyEntries);
					array<String^>^ leaves = row->Cells["LeaveColumn"]->Value != nullptr ?
						row->Cells["LeaveColumn"]->Value->ToString()->Split(
							gcnew array<String^>{" | "}, StringSplitOptions::RemoveEmptyEntries) :
						gcnew array<String^>(0);

					shouldRecordArrival = (arrivals->Length <= leaves->Length);
				}

				// Записываем время прихода или ухода
				if (shouldRecordArrival) {
					// Добавляем время прихода
					String^ newValue = row->Cells["ArrivalColumn"]->Value != nullptr ?
						row->Cells["ArrivalColumn"]->Value->ToString() + " | " + currentTime.ToString("yyyy-MM-dd HH:mm:ss") :
						currentTime.ToString("yyyy-MM-dd HH:mm:ss");
					row->Cells["ArrivalColumn"]->Value = newValue;
					MessageBox::Show("Приход зарегистрирован: " + currentTime.ToString("HH:mm:ss"));
				}
				else {
					// Добавляем время ухода
					String^ newValue = row->Cells["LeaveColumn"]->Value != nullptr ?
						row->Cells["LeaveColumn"]->Value->ToString() + " | " + currentTime.ToString("yyyy-MM-dd HH:mm:ss") :
						currentTime.ToString("yyyy-MM-dd HH:mm:ss");
					row->Cells["LeaveColumn"]->Value = newValue;
					MessageBox::Show("Уход зарегистрирован: " + currentTime.ToString("HH:mm:ss"));

					// Рассчитываем общее отработанное время за день
					if (row->Cells["ArrivalColumn"]->Value != nullptr &&
						row->Cells["LeaveColumn"]->Value != nullptr) {
						array<String^>^ arrivals = row->Cells["ArrivalColumn"]->Value->ToString()->Split(
							gcnew array<String^>{" | "}, StringSplitOptions::RemoveEmptyEntries);
						array<String^>^ leaves = row->Cells["LeaveColumn"]->Value->ToString()->Split(
							gcnew array<String^>{" | "}, StringSplitOptions::RemoveEmptyEntries);

						TimeSpan totalWorkedTime = TimeSpan::Zero;

						// Считаем время только для текущего дня
						for (int i = 0; i < Math::Min(arrivals->Length, leaves->Length); i++) {
							DateTime arrival, leave;
							if (DateTime::TryParse(arrivals[i], arrival) &&
								DateTime::TryParse(leaves[i], leave)) {
								if (arrival.ToString("yyyy-MM-dd") == currentDate) {
									totalWorkedTime = totalWorkedTime.Add(leave - arrival);
								}
							}
						}

						// Обновляем общее рабочее время
						double totalHours = totalWorkedTime.TotalHours;
						row->Cells["worktime"]->Value = totalHours.ToString("F2") + " часов";
					}
				}
				break;
			}
		}

		if (!found) {
			MessageBox::Show("Сотрудник с номером карты " + cardNumber + " не найден.");
		}
	}
}
private: System::Void new_day_Click(System::Object^ sender, System::EventArgs^ e) {
    System::Windows::Forms::DialogResult dialogResult = MessageBox::Show(
        "Закрыть смену и начать следующий день?", 
        "Подтверждение", 
        MessageBoxButtons::YesNo, 
        MessageBoxIcon::Question);
    
    if (dialogResult == System::Windows::Forms::DialogResult::Yes) {
        String^ currentDateStr = t_date->Text;
        DateTime currentDate;
        if (!DateTime::TryParse(currentDateStr, currentDate)) {
            MessageBox::Show("Ошибка формата даты в label!");
            return;
        }
        String^ documentsPath = Environment::GetFolderPath(Environment::SpecialFolder::MyDocuments);
        String^ fileName = currentDate.ToString("dd.MM.yyyy") + " employee data.txt";
        String^ filePath = System::IO::Path::Combine(documentsPath, fileName);
        
        try {
            System::IO::StreamWriter^ sw = gcnew System::IO::StreamWriter(filePath);
            sw->WriteLine("Номер карты;ФИО;Дата;Отработанное время");
            for each (DataGridViewRow ^ row in dataGridViewEmployees->Rows) {
                if (!row->IsNewRow && row->Cells["worktime"]->Value != nullptr) {
                    String^ cardId = row->Cells["cardid"]->Value ? row->Cells["cardid"]->Value->ToString() : "";
                    String^ name = row->Cells["NameColumn"]->Value ? row->Cells["NameColumn"]->Value->ToString() : "";
                    String^ workTime = row->Cells["worktime"]->Value->ToString();
                    
                    sw->WriteLine(String::Format("{0};{1};{2};{3}", 
                        cardId, name, currentDateStr, workTime));
                }
            }
            sw->Close();
        }
        catch (Exception^ ex) {
            MessageBox::Show("Ошибка при сохранении данных: " + ex->Message);
            return;
        }
        for each (DataGridViewRow ^ row in dataGridViewEmployees->Rows) {
            if (!row->IsNewRow) {
                row->Cells["worktime"]->Value = "";
                row->Cells["ArrivalColumn"]->Value = "";
                row->Cells["LeaveColumn"]->Value = "";
            }
        }
        DateTime nextDate = currentDate.AddDays(1);
        t_date->Text = nextDate.ToString("dd.MM.yyyy");
        сохранитьToolStripMenuItem_Click(nullptr, nullptr);
        MessageBox::Show("Данные за " + currentDateStr + " выгружены в " + filePath);
    }
}
private: System::Void btnPrint_Click(System::Object^ sender, System::EventArgs^ e) {
	if (dataGridViewEmployees->Rows->Count == 0 ||
		(dataGridViewEmployees->Rows->Count == 1 && dataGridViewEmployees->Rows[0]->IsNewRow)) {
		MessageBox::Show("Нет данных для печати!", "Ошибка",
			MessageBoxButtons::OK, MessageBoxIcon::Warning);
		return;
	}

	printPreviewDialog->Document = printDocument;
	printPreviewDialog->ShowDialog();
}
private: System::Void printDocument_PrintPage(System::Object^ sender,
	System::Drawing::Printing::PrintPageEventArgs^ e) {
	System::Drawing::Font^ headerFont = gcnew System::Drawing::Font("Arial", 16, System::Drawing::FontStyle::Bold);
	System::Drawing::Font^ textFont = gcnew System::Drawing::Font("Arial", 10);
	System::Drawing::SolidBrush^ brush = gcnew System::Drawing::SolidBrush(System::Drawing::Color::Black);
	System::Drawing::Pen^ pen = gcnew System::Drawing::Pen(System::Drawing::Color::Black, 1);
	System::Drawing::StringFormat^ format = gcnew System::Drawing::StringFormat();
	format->Alignment = System::Drawing::StringAlignment::Near;
	format->LineAlignment = System::Drawing::StringAlignment::Center;
	format->Trimming = System::Drawing::StringTrimming::Word;
	format->FormatFlags = System::Drawing::StringFormatFlags::LineLimit;
	float leftMargin = e->MarginBounds.Left;
	float topMargin = e->MarginBounds.Top;
	float yPos = topMargin;
	e->Graphics->DrawString("ТАБЕЛЬ УЧЕТА РАБОЧЕГО ВРЕМЕНИ", headerFont, brush, leftMargin, yPos);
	yPos += headerFont->GetHeight(e->Graphics) + 20;
	e->Graphics->DrawString("Период: " + t_date->Text, textFont, brush, leftMargin, yPos);
	yPos += textFont->GetHeight(e->Graphics) + 20;
	array<float>^ columnWidths = gcnew array<float>{
		40,   // №
			200,  // ФИО
			120,  // Отдел
			120,  // Приход
			120,  // Уход
			100   // Отработано
	};

	array<String^>^ headers = gcnew array<String^>{"№", "ФИО", "Отдел", "Приход", "Уход", "Отработано"};
	float baseRowHeight = textFont->GetHeight(e->Graphics) * 1.5f;
	float xPos = leftMargin;
	for (int i = 0; i < headers->Length; i++) {
		System::Drawing::RectangleF rect(xPos, yPos, columnWidths[i], baseRowHeight);
		e->Graphics->DrawRectangle(pen, rect.X, rect.Y, rect.Width, rect.Height);
		e->Graphics->DrawString(headers[i], textFont, brush, rect, format);
		xPos += columnWidths[i];
	}
	yPos += baseRowHeight;
	int rowNumber = 1;
	for each (DataGridViewRow ^ row in dataGridViewEmployees->Rows) {
		if (row->IsNewRow) continue;
		float maxCellHeight = baseRowHeight;
		array<float>^ cellHeights = gcnew array<float>(headers->Length);
		xPos = leftMargin;

		for (int col = 0; col < headers->Length; col++) {
			String^ value = "";
			switch (col) {
			case 0: value = rowNumber.ToString(); break;
			case 1: value = safe_cast<String^>(row->Cells["NameColumn"]->Value); break;
			case 2: value = safe_cast<String^>(row->Cells["otdel"]->Value); break;
			case 3: value = FormatDateTimeValue(safe_cast<String^>(row->Cells["ArrivalColumn"]->Value)); break;
			case 4: value = FormatDateTimeValue(safe_cast<String^>(row->Cells["LeaveColumn"]->Value)); break;
			case 5: value = safe_cast<String^>(row->Cells["worktime"]->Value); break;
			}

			System::Drawing::RectangleF cellRect(xPos, yPos, columnWidths[col], 1000); // Большая высота для измерения
			System::Drawing::SizeF textSize = e->Graphics->MeasureString(value, textFont, cellRect.Size, format);
			cellHeights[col] = textSize.Height + 4; // Небольшой отступ

			if (cellHeights[col] > maxCellHeight) {
				maxCellHeight = cellHeights[col];
			}

			xPos += columnWidths[col];
		}

		// Рисуем строку с вычисленной высотой
		xPos = leftMargin;
		for (int col = 0; col < headers->Length; col++) {
			String^ value = "";
			switch (col) {
			case 0: value = rowNumber.ToString(); break;
			case 1: value = safe_cast<String^>(row->Cells["NameColumn"]->Value); break;
			case 2: value = safe_cast<String^>(row->Cells["otdel"]->Value); break;
			case 3: value = FormatDateTimeValue(safe_cast<String^>(row->Cells["ArrivalColumn"]->Value)); break;
			case 4: value = FormatDateTimeValue(safe_cast<String^>(row->Cells["LeaveColumn"]->Value)); break;
			case 5: value = safe_cast<String^>(row->Cells["worktime"]->Value); break;
			}

			System::Drawing::RectangleF cellRect(xPos, yPos, columnWidths[col], maxCellHeight);
			e->Graphics->DrawRectangle(pen, cellRect.X, cellRect.Y, cellRect.Width, cellRect.Height);
			e->Graphics->DrawString(value, textFont, brush, cellRect, format);
			xPos += columnWidths[col];
		}

		yPos += maxCellHeight;
		rowNumber++;

		if (yPos > e->MarginBounds.Bottom - 100) {
			e->HasMorePages = true;
			return;
		}
	}
	yPos += 30;
	e->Graphics->DrawString("Ответственный: ___________________", textFont, brush, leftMargin, yPos);
	yPos += textFont->GetHeight(e->Graphics);
	e->Graphics->DrawString("Дата составления: " + DateTime::Now.ToString("dd.MM.yyyy"),
		textFont, brush, leftMargin, yPos);

	e->HasMorePages = false;
}
private: String^ FormatDateTimeValue(String^ input) {
	DateTime dateValue;
	if (DateTime::TryParse(input, dateValue)) {
		return dateValue.ToString("yyyy-MM-dd HH:mm");
	}
	return input;
}
private: System::Void btnShowChart_Click(System::Object^ sender, System::EventArgs^ e) {
	GrafikForm^ grafikForm = gcnew GrafikForm(this->dataGridViewEmployees);
	grafikForm->Show();
}
private:
	void CheckCardExpiration()
	{
		DateTime currentDate = DateTime::Now;

		for each (DataGridViewRow ^ row in dataGridViewEmployees->Rows)
		{
			if (!row->IsNewRow && row->Cells["carddata"]->Value != nullptr)
			{
				String^ expirationDateStr = row->Cells["carddata"]->Value->ToString();
				DateTime expirationDate;

				if (DateTime::TryParse(expirationDateStr, expirationDate))
				{
					TimeSpan timeLeft = expirationDate.Date - currentDate.Date;
					int daysLeft = timeLeft.Days;
					row->Cells["carddata"]->Style->BackColor = Color::White;
					if (daysLeft <= 0)
					{
						// Срок истек (красный цвет)
						row->Cells["carddata"]->Style->BackColor = Color::LightCoral;

						// Если сегодня день истечения срока
						if (daysLeft == 0)
						{
							String^ employeeName = row->Cells["NameColumn"]->Value != nullptr ?
								row->Cells["NameColumn"]->Value->ToString() : "неизвестный сотрудник";
							MessageBox::Show(
								String::Format("Срок действия карты {0} истёк!", employeeName),
								"Внимание",
								MessageBoxButtons::OK,
								MessageBoxIcon::Warning);
						}
					}
					else if (daysLeft <= 7)
					{
						// До истечения срока неделя или меньше (оранжевый цвет)
						row->Cells["carddata"]->Style->BackColor = Color::Orange;
					}
				}
			}
		}
	}
	private: System::Void dataGridViewEmployees_CellValueChanged(
		System::Object^ sender,
		System::Windows::Forms::DataGridViewCellEventArgs^ e)
	{
		if (e->ColumnIndex == dataGridViewEmployees->Columns["carddata"]->Index && e->RowIndex >= 0)
		{
			CheckCardExpiration();
		}
	}
	private: System::Void dataGridViewEmployees_CellFormatting(
		System::Object^ sender,
		System::Windows::Forms::DataGridViewCellFormattingEventArgs^ e)
	{
		if (e->ColumnIndex == dataGridViewEmployees->Columns["carddata"]->Index && e->RowIndex >= 0)
		{
			DataGridViewRow^ row = dataGridViewEmployees->Rows[e->RowIndex];
			if (!row->IsNewRow && row->Cells["carddata"]->Value != nullptr)
			{
				String^ expirationDateStr = row->Cells["carddata"]->Value->ToString();
				DateTime expirationDate;
				DateTime currentDate = DateTime::Now;

				if (DateTime::TryParse(expirationDateStr, expirationDate))
				{
					TimeSpan timeLeft = expirationDate.Date - currentDate.Date;
					int daysLeft = timeLeft.Days;

					if (daysLeft <= 0)
					{
						e->CellStyle->BackColor = Color::LightCoral;
					}
					else if (daysLeft <= 7)
					{
						e->CellStyle->BackColor = Color::Orange;
					}
				}
			}
		}
	}
private: System::Void btnOpenQueries_Click(System::Object^ sender, System::EventArgs^ e) {
	EmployeeAttendanceApp::QueryForm^ queryForm = gcnew EmployeeAttendanceApp::QueryForm(this->dataGridViewEmployees);
	queryForm->ShowDialog();
}
};
}