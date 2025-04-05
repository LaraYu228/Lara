#pragma once

#include <Windows.h>

namespace EmployeeAttendanceApp {

    using namespace System;
    using namespace System::ComponentModel;
    using namespace System::Collections;
    using namespace System::Windows::Forms;
    using namespace System::Data;
    using namespace System::Drawing;
    using namespace System::Windows::Forms::DataVisualization::Charting;

    public ref class GrafikForm : public System::Windows::Forms::Form
    {
    public:
        GrafikForm(System::Windows::Forms::DataGridView^ dataGridView)
        {
            InitializeComponent();
            this->dataGridView = dataGridView;
            InitializeChart();
            UpdateChart();
        }

    protected:
        ~GrafikForm()
        {
            if (components)
            {
                delete components;
            }
        }

    private:
        System::Windows::Forms::DataGridView^ dataGridView;
        System::ComponentModel::Container^ components;

        // Элементы управления
        System::Windows::Forms::DataVisualization::Charting::Chart^ chartAttendance;
        System::Windows::Forms::Button^ btnUpdateChart;
        System::Windows::Forms::ComboBox^ cmbChartType;
        System::Windows::Forms::Panel^ panelTop;

        void InitializeChart()
        {
            chartAttendance->Series->Clear();
            chartAttendance->Titles->Clear();

            if (chartAttendance->ChartAreas->Count == 0)
            {
                chartAttendance->ChartAreas->Add("ChartArea1");
            }

            chartAttendance->Titles->Add("График посещаемости сотрудников");
            chartAttendance->Titles[0]->Font = gcnew System::Drawing::Font("Arial", 12, System::Drawing::FontStyle::Bold);
            cmbChartType->SelectedIndex = 0;
        }

        void UpdateChart()
        {
            if (dataGridView == nullptr || dataGridView->Rows->Count == 0)
            {
                MessageBox::Show("Нет данных для отображения графика", "Информация",
                    MessageBoxButtons::OK, MessageBoxIcon::Information);
                return;
            }

            chartAttendance->Series->Clear();
            Series^ series = gcnew Series();
            series->Name = "Посещаемость";
            series->ChartArea = "ChartArea1";
            series->Legend = "Legend1";

            switch (cmbChartType->SelectedIndex)
            {
            case 0: series->ChartType = SeriesChartType::Column; break;
            case 1: series->ChartType = SeriesChartType::Pie; break;
            case 2: series->ChartType = SeriesChartType::Line; break;
            default: series->ChartType = SeriesChartType::Column; break;
            }

            // Считаем общее количество часов для расчета процентов
            double totalHours = 0;
            for each (DataGridViewRow ^ row in dataGridView->Rows)
            {
                if (!row->IsNewRow && row->Cells["NameColumn"]->Value != nullptr && row->Cells["worktime"]->Value != nullptr)
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
                    totalHours += workHours;
                }
            }

            // Добавляем точки данных
            for each (DataGridViewRow ^ row in dataGridView->Rows)
            {
                if (!row->IsNewRow && row->Cells["NameColumn"]->Value != nullptr && row->Cells["worktime"]->Value != nullptr)
                {
                    String^ name = row->Cells["NameColumn"]->Value->ToString();
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

                    if (workHours > 0)
                    {
                        // Добавляем точку с подписью
                        DataPoint^ point = gcnew DataPoint();
                        point->SetValueXY(name, workHours);

                        // Форматируем подпись
                        if (series->ChartType == SeriesChartType::Pie && totalHours > 0)
                        {
                            double percentage = (workHours / totalHours) * 100;
                            point->Label = String::Format("{0}\n{1:0.#}%", name, percentage);
                            point->LegendText = String::Format("{0} ({1:0.#}%)", name, percentage);
                        }
                        else
                        {
                            point->Label = name;
                        }

                        series->Points->Add(point);
                    }
                }
            }

            if (series->Points->Count > 0)
            {
                chartAttendance->Series->Add(series);

                // Настройки для круговой диаграммы
                if (series->ChartType == SeriesChartType::Pie)
                {
                    series->Label = "#VALX\n#PERCENT{P0}";
                    series->LegendText = "#VALX (#PERCENT{P0})";
                    series->IsValueShownAsLabel = true;
                    series->Font = gcnew System::Drawing::Font("Arial", 8);
                    series->LabelForeColor = Color::Black;

                    // Настройка отображения подписей
                    series["PieLabelStyle"] = "Outside";
                    series["PieLineColor"] = "Black";
                }
                else
                {
                    chartAttendance->ChartAreas["ChartArea1"]->AxisX->LabelStyle->Angle = -45;
                    chartAttendance->ChartAreas["ChartArea1"]->AxisX->Interval = 1;
                    chartAttendance->ChartAreas["ChartArea1"]->AxisY->Title = "Отработанные часы";
                }
            }
        }

#pragma region Windows Form Designer generated code
        void InitializeComponent(void)
        {
            this->chartAttendance = (gcnew System::Windows::Forms::DataVisualization::Charting::Chart());
            this->btnUpdateChart = (gcnew System::Windows::Forms::Button());
            this->cmbChartType = (gcnew System::Windows::Forms::ComboBox());
            this->panelTop = (gcnew System::Windows::Forms::Panel());
            (cli::safe_cast<System::ComponentModel::ISupportInitialize^>(this->chartAttendance))->BeginInit();
            this->panelTop->SuspendLayout();
            this->SuspendLayout();
            // 
            // chartAttendance
            // 
            this->chartAttendance->BackColor = System::Drawing::Color::PaleGoldenrod;
            this->chartAttendance->BorderlineColor = System::Drawing::Color::Goldenrod;
            this->chartAttendance->BorderlineDashStyle = System::Windows::Forms::DataVisualization::Charting::ChartDashStyle::Solid;
            this->chartAttendance->BorderlineWidth = 2;
            this->chartAttendance->Dock = System::Windows::Forms::DockStyle::Fill;
            this->chartAttendance->Location = System::Drawing::Point(0, 50);
            this->chartAttendance->Name = L"chartAttendance";
            this->chartAttendance->Size = System::Drawing::Size(800, 523);
            this->chartAttendance->TabIndex = 0;
            this->chartAttendance->Text = L"График посещаемости";
            // 
            // btnUpdateChart
            // 
            this->btnUpdateChart->BackColor = System::Drawing::Color::FromArgb(static_cast<System::Int32>(static_cast<System::Byte>(255)), static_cast<System::Int32>(static_cast<System::Byte>(255)),
                static_cast<System::Int32>(static_cast<System::Byte>(192)));
            this->btnUpdateChart->FlatStyle = System::Windows::Forms::FlatStyle::Flat;
            this->btnUpdateChart->ForeColor = System::Drawing::Color::Brown;
            this->btnUpdateChart->Location = System::Drawing::Point(12, 12);
            this->btnUpdateChart->Name = L"btnUpdateChart";
            this->btnUpdateChart->Size = System::Drawing::Size(120, 23);
            this->btnUpdateChart->TabIndex = 1;
            this->btnUpdateChart->Text = L"Обновить график";
            this->btnUpdateChart->UseVisualStyleBackColor = false;
            this->btnUpdateChart->Click += gcnew System::EventHandler(this, &GrafikForm::btnUpdateChart_Click);
            // 
            // cmbChartType
            // 
            this->cmbChartType->BackColor = System::Drawing::Color::PaleGoldenrod;
            this->cmbChartType->DropDownStyle = System::Windows::Forms::ComboBoxStyle::DropDownList;
            this->cmbChartType->FormattingEnabled = true;
            this->cmbChartType->Items->AddRange(gcnew cli::array< System::Object^  >(3) {
                L"Столбчатая диаграмма", L"Круговая диаграмма",
                    L"Линейный график"
            });
            this->cmbChartType->Location = System::Drawing::Point(138, 12);
            this->cmbChartType->Name = L"cmbChartType";
            this->cmbChartType->Size = System::Drawing::Size(180, 21);
            this->cmbChartType->TabIndex = 2;
            this->cmbChartType->SelectedIndexChanged += gcnew System::EventHandler(this, &GrafikForm::cmbChartType_SelectedIndexChanged);
            // 
            // panelTop
            // 
            this->panelTop->BackColor = System::Drawing::Color::PaleGoldenrod;
            this->panelTop->Controls->Add(this->btnUpdateChart);
            this->panelTop->Controls->Add(this->cmbChartType);
            this->panelTop->Dock = System::Windows::Forms::DockStyle::Top;
            this->panelTop->Location = System::Drawing::Point(0, 0);
            this->panelTop->Name = L"panelTop";
            this->panelTop->Size = System::Drawing::Size(800, 50);
            this->panelTop->TabIndex = 3;
            // 
            // GrafikForm
            // 
            this->AutoScaleDimensions = System::Drawing::SizeF(6, 13);
            this->AutoScaleMode = System::Windows::Forms::AutoScaleMode::Font;
            this->BackColor = System::Drawing::Color::FromArgb(static_cast<System::Int32>(static_cast<System::Byte>(255)), static_cast<System::Int32>(static_cast<System::Byte>(255)),
                static_cast<System::Int32>(static_cast<System::Byte>(128)));
            this->ClientSize = System::Drawing::Size(800, 573);
            this->Controls->Add(this->chartAttendance);
            this->Controls->Add(this->panelTop);
            this->Name = L"GrafikForm";
            this->StartPosition = System::Windows::Forms::FormStartPosition::CenterScreen;
            this->Text = L"График посещаемости";
            (cli::safe_cast<System::ComponentModel::ISupportInitialize^>(this->chartAttendance))->EndInit();
            this->panelTop->ResumeLayout(false);
            this->ResumeLayout(false);

        }
#pragma endregion

    private:
        System::Void btnUpdateChart_Click(System::Object^ sender, System::EventArgs^ e) {
            UpdateChart();
        }

        System::Void cmbChartType_SelectedIndexChanged(System::Object^ sender, System::EventArgs^ e) {
            UpdateChart();
        }
    };
}