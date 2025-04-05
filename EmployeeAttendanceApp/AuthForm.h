#pragma once

namespace EmployeeAttendanceApp {

	using namespace System;
	using namespace System::ComponentModel;
	using namespace System::Collections;
	using namespace System::Windows::Forms;
	using namespace System::Data;
	using namespace System::Drawing;

	/// <summary>
	/// Summary for AuthForm
	/// </summary>
	public ref class AuthForm : public System::Windows::Forms::Form
	{
	public:
		AuthForm(void)
		{
			InitializeComponent();
			//
			//TODO: Add the constructor code here
			//
		}

	protected:
		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		~AuthForm()
		{
			if (components)
			{
				delete components;
			}
		}
	private: System::Windows::Forms::TextBox^ inputTextBox;
	private: System::Windows::Forms::Button^ confirmButton;
	protected:

	private:
		String^ userInput;  // Переменная для хранения введенных данных
	private: System::Windows::Forms::Label^ label1;
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
			System::ComponentModel::ComponentResourceManager^ resources = (gcnew System::ComponentModel::ComponentResourceManager(AuthForm::typeid));
			this->inputTextBox = (gcnew System::Windows::Forms::TextBox());
			this->confirmButton = (gcnew System::Windows::Forms::Button());
			this->label1 = (gcnew System::Windows::Forms::Label());
			this->SuspendLayout();
			// 
			// inputTextBox
			// 
			this->inputTextBox->Location = System::Drawing::Point(106, 140);
			this->inputTextBox->Name = L"inputTextBox";
			this->inputTextBox->Size = System::Drawing::Size(258, 20);
			this->inputTextBox->TabIndex = 0;
			// 
			// confirmButton
			// 
			this->confirmButton->Location = System::Drawing::Point(153, 192);
			this->confirmButton->Name = L"confirmButton";
			this->confirmButton->Size = System::Drawing::Size(146, 44);
			this->confirmButton->TabIndex = 1;
			this->confirmButton->Text = L"Ввести данные";
			this->confirmButton->UseVisualStyleBackColor = true;
			this->confirmButton->Click += gcnew System::EventHandler(this, &AuthForm::confirmButton_Click);
			// 
			// label1
			// 
			this->label1->AutoSize = true;
			this->label1->Font = (gcnew System::Drawing::Font(L"Microsoft Sans Serif", 15.75F, System::Drawing::FontStyle::Regular, System::Drawing::GraphicsUnit::Point,
				static_cast<System::Byte>(204)));
			this->label1->Location = System::Drawing::Point(119, 81);
			this->label1->Name = L"label1";
			this->label1->Size = System::Drawing::Size(230, 25);
			this->label1->TabIndex = 2;
			this->label1->Text = L"Введите номер карты";
			// 
			// AuthForm
			// 
			this->AutoScaleDimensions = System::Drawing::SizeF(6, 13);
			this->AutoScaleMode = System::Windows::Forms::AutoScaleMode::Font;
			this->BackColor = System::Drawing::Color::PaleGoldenrod;
			this->ClientSize = System::Drawing::Size(460, 345);
			this->Controls->Add(this->label1);
			this->Controls->Add(this->confirmButton);
			this->Controls->Add(this->inputTextBox);
			this->FormBorderStyle = System::Windows::Forms::FormBorderStyle::None;
			this->Icon = (cli::safe_cast<System::Drawing::Icon^>(resources->GetObject(L"$this.Icon")));
			this->Name = L"AuthForm";
			this->StartPosition = System::Windows::Forms::FormStartPosition::CenterScreen;
			this->Text = L"AuthForm";
			this->ResumeLayout(false);
			this->PerformLayout();

		}

		// Обработчик нажатия кнопки
		void confirmButton_Click(System::Object^ sender, System::EventArgs^ e)
		{
			// Сохраняем введенные данные
			this->userInput = this->inputTextBox->Text;
			// Закрываем форму с результатом OK
			this->DialogResult = System::Windows::Forms::DialogResult::OK;
			this->Close();
		}
	public:
		// Метод для получения введенных данных
		String^ getUserInput()
		{
			return this->userInput;
		}
	};
}