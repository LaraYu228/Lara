// QueryForm.h
#pragma once

#include <Windows.h>
#include <vcclr.h>
#include "AuthForm.h"
#include "GrafikForm.h"

namespace EmployeeAttendanceApp {

    using namespace System;
    using namespace System::ComponentModel;
    using namespace System::Collections;
    using namespace System::Windows::Forms;
    using namespace System::Data;
    using namespace System::Drawing;
    using namespace Microsoft::VisualBasic;

    public ref class QueryForm : public System::Windows::Forms::Form
    {
    public:
        QueryForm(System::Windows::Forms::DataGridView^ sourceDataGridView)
        {
            InitializeComponent();
            this->sourceDataGridView = sourceDataGridView;
        }

    private:
        System::Windows::Forms::DataGridView^ sourceDataGridView;
        System::Windows::Forms::DataGridView^ resultDataGridView;
        System::Windows::Forms::ComboBox^ cmbQueryType;
        System::Windows::Forms::Button^ btnExecuteQuery;
        System::Windows::Forms::Label^ label1;
        System::Windows::Forms::TextBox^ txtDepartment;
        System::ComponentModel::Container^ components;

        void InitializeComponent(void)
        {
            this->resultDataGridView = (gcnew System::Windows::Forms::DataGridView());
            this->cmbQueryType = (gcnew System::Windows::Forms::ComboBox());
            this->btnExecuteQuery = (gcnew System::Windows::Forms::Button());
            this->label1 = (gcnew System::Windows::Forms::Label());
            this->txtDepartment = (gcnew System::Windows::Forms::TextBox());
            (cli::safe_cast<System::ComponentModel::ISupportInitialize^>(this->resultDataGridView))->BeginInit();
            this->SuspendLayout();
            // 
            // resultDataGridView
            // 
            this->resultDataGridView->Anchor = static_cast<System::Windows::Forms::AnchorStyles>((((System::Windows::Forms::AnchorStyles::Top | System::Windows::Forms::AnchorStyles::Bottom)
                | System::Windows::Forms::AnchorStyles::Left)
                | System::Windows::Forms::AnchorStyles::Right));
            this->resultDataGridView->ColumnHeadersHeightSizeMode = System::Windows::Forms::DataGridViewColumnHeadersHeightSizeMode::AutoSize;
            this->resultDataGridView->Location = System::Drawing::Point(12, 80);
            this->resultDataGridView->Name = L"resultDataGridView";
            this->resultDataGridView->Size = System::Drawing::Size(760, 369);
            this->resultDataGridView->TabIndex = 0;
            // 
            // cmbQueryType
            // 
            this->cmbQueryType->DropDownStyle = System::Windows::Forms::ComboBoxStyle::DropDownList;
            this->cmbQueryType->FormattingEnabled = true;
            this->cmbQueryType->Items->AddRange(gcnew cli::array< System::Object^  >(4) {
                L"Выберите запрос", L"Сотрудники с истекающим сроком карты",
                    L"Сотрудники, отработавшие менее 4 часов", L"Сотрудники по отделу"
            });
            this->cmbQueryType->Location = System::Drawing::Point(12, 24);
            this->cmbQueryType->Name = L"cmbQueryType";
            this->cmbQueryType->Size = System::Drawing::Size(300, 21);
            this->cmbQueryType->TabIndex = 1;
            this->cmbQueryType->SelectedIndexChanged += gcnew System::EventHandler(this, &QueryForm::cmbQueryType_SelectedIndexChanged);
            // 
            // btnExecuteQuery
            // 
            this->btnExecuteQuery->BackColor = System::Drawing::Color::FromArgb(static_cast<System::Int32>(static_cast<System::Byte>(255)), static_cast<System::Int32>(static_cast<System::Byte>(255)),
                static_cast<System::Int32>(static_cast<System::Byte>(192)));
            this->btnExecuteQuery->FlatStyle = System::Windows::Forms::FlatStyle::Flat;
            this->btnExecuteQuery->ForeColor = System::Drawing::Color::Brown;
            this->btnExecuteQuery->Location = System::Drawing::Point(327, 24);
            this->btnExecuteQuery->Name = L"btnExecuteQuery";
            this->btnExecuteQuery->Size = System::Drawing::Size(120, 23);
            this->btnExecuteQuery->TabIndex = 2;
            this->btnExecuteQuery->Text = L"Выполнить";
            this->btnExecuteQuery->UseVisualStyleBackColor = false;
            this->btnExecuteQuery->Click += gcnew System::EventHandler(this, &QueryForm::btnExecuteQuery_Click);
            // 
            // label1
            // 
            this->label1->AutoSize = true;
            this->label1->Location = System::Drawing::Point(12, 8);
            this->label1->Name = L"label1";
            this->label1->Size = System::Drawing::Size(74, 13);
            this->label1->TabIndex = 3;
            this->label1->Text = L"Тип запроса:";
            // 
            // txtDepartment
            // 
            this->txtDepartment->Location = System::Drawing::Point(12, 51);
            this->txtDepartment->Name = L"txtDepartment";
            this->txtDepartment->Size = System::Drawing::Size(300, 20);
            this->txtDepartment->TabIndex = 4;
            this->txtDepartment->Visible = false;
            // 
            // QueryForm
            // 
            this->AutoScaleDimensions = System::Drawing::SizeF(6, 13);
            this->AutoScaleMode = System::Windows::Forms::AutoScaleMode::Font;
            this->BackColor = System::Drawing::Color::PaleGoldenrod;
            this->ClientSize = System::Drawing::Size(784, 461);
            this->Controls->Add(this->txtDepartment);
            this->Controls->Add(this->label1);
            this->Controls->Add(this->btnExecuteQuery);
            this->Controls->Add(this->cmbQueryType);
            this->Controls->Add(this->resultDataGridView);
            this->Name = L"QueryForm";
            this->ShowIcon = false;
            this->StartPosition = System::Windows::Forms::FormStartPosition::CenterParent;
            this->Text = L"Запрос";
            (cli::safe_cast<System::ComponentModel::ISupportInitialize^>(this->resultDataGridView))->EndInit();
            this->ResumeLayout(false);
            this->PerformLayout();

        }

    private:
        System::Void cmbQueryType_SelectedIndexChanged(System::Object^ sender, System::EventArgs^ e)
        {
            txtDepartment->Visible = (cmbQueryType->SelectedIndex == 3);
        }

        System::Void btnExecuteQuery_Click(System::Object^ sender, System::EventArgs^ e)
        {
            if (cmbQueryType->SelectedIndex <= 0)
            {
                MessageBox::Show("Выберите тип запроса", "Ошибка",
                    MessageBoxButtons::OK, MessageBoxIcon::Warning);
                return;
            }

            resultDataGridView->Rows->Clear();
            resultDataGridView->Columns->Clear();

            // Копируем структуру столбцов из исходной таблицы
            for each (DataGridViewColumn ^ column in sourceDataGridView->Columns)
            {
                if (column->Name != "ArrivalColumn" && column->Name != "LeaveColumn")
                {
                    resultDataGridView->Columns->Add(
                        (DataGridViewColumn^)column->Clone());
                }
            }

            DateTime currentDate = DateTime::Now;

            switch (cmbQueryType->SelectedIndex)
            {
            case 1: // Сотрудники с истекающим сроком карты
            {
                for each (DataGridViewRow ^ row in sourceDataGridView->Rows)
                {
                    if (!row->IsNewRow && row->Cells["carddata"]->Value != nullptr)
                    {
                        String^ expirationDateStr = row->Cells["carddata"]->Value->ToString();
                        DateTime expirationDate;

                        if (DateTime::TryParse(expirationDateStr, expirationDate))
                        {
                            TimeSpan timeLeft = expirationDate.Date - currentDate.Date;
                            int daysLeft = timeLeft.Days;

                            if (daysLeft <= 7)
                            {
                                AddRowToResult(row);
                            }
                        }
                    }
                }
                break;
            }

            case 2: // Сотрудники, отработавшие менее 4 часов
            {
                for each (DataGridViewRow ^ row in sourceDataGridView->Rows)
                {
                    if (!row->IsNewRow && row->Cells["worktime"]->Value != nullptr)
                    {
                        String^ workTimeStr = row->Cells["worktime"]->Value->ToString();
                        double workHours = 0;

                        if (workTimeStr->Contains("часов"))
                        {
                            String^ numericPart = workTimeStr->Substring(0, workTimeStr->IndexOf(" "));
                            Double::TryParse(numericPart, workHours);
                        }
                        else
                        {
                            Double::TryParse(workTimeStr, workHours);
                        }

                        if (workHours > 0 && workHours < 4)
                        {
                            AddRowToResult(row);
                        }
                    }
                }
                break;
            }

            case 3: // Сотрудники по отделу
            {
                String^ department = txtDepartment->Text;

                if (!String::IsNullOrEmpty(department))
                {
                    for each (DataGridViewRow ^ row in sourceDataGridView->Rows)
                    {
                        if (!row->IsNewRow && row->Cells["otdel"]->Value != nullptr)
                        {
                            String^ rowDepartment = row->Cells["otdel"]->Value->ToString();

                            if (rowDepartment->Equals(department, StringComparison::CurrentCultureIgnoreCase))
                            {
                                AddRowToResult(row);
                            }
                        }
                    }
                }
                break;
            }
            }
        }

        void AddRowToResult(DataGridViewRow^ sourceRow)
        {
            int index = resultDataGridView->Rows->Add();
            DataGridViewRow^ newRow = resultDataGridView->Rows[index];

            for each (DataGridViewColumn ^ column in resultDataGridView->Columns)
            {
                if (column->Index < sourceRow->Cells->Count)
                {
                    newRow->Cells[column->Index]->Value =
                        sourceRow->Cells[column->Name]->Value;
                }
            }
        }
    };
}