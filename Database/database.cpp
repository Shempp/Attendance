#include "database.h"
#include "settingsdialog.h"
#include "ui_database.h"

database::database(QWidget *parent) :
    QMainWindow(parent),
    m_ui(new Ui::database),
    m_settings(new SettingsDialog),
    m_serial(new QSerialPort(this))
{
    m_ui->setupUi(this);

    QRegExp exp("[а-яёА-ЯЁ]+");
    m_ui->enterLineEdit->setValidator(new QRegExpValidator(exp, this));
    m_ui->enterLineEdit->setMaxLength(20);

    m_settings->setWindowModality(Qt::ApplicationModal);
    m_settings->setMinimumSize(315, 210);
    m_settings->setMaximumSize(315, 210);

    m_ui->disconnectButton->setEnabled(false);
    m_ui->idLineEdit->setEnabled(false);
    m_ui->addButton->setEnabled(false);
    m_ui->enterLineEdit->setEnabled(false);
    m_ui->idLineEdit->setReadOnly(true);

    initActionsConnections();

    connect(m_serial, &QSerialPort::errorOccurred, this, &database::handleError);
    connect(m_serial, &QSerialPort::readyRead, this, &database::readData);
    connect(m_ui->enterLineEdit, SIGNAL(textChanged(QString)), this, SLOT(lineChanged(QString)));
}

database::~database()
{
    delete m_ui;
    delete m_settings;
}

void database::openSerialPort()
{
    const SettingsDialog::Settings p = m_settings->settings();
    m_serial->setPortName(p.name);
    m_serial->setBaudRate(p.baudRate);
    m_serial->setDataBits(p.dataBits);
    m_serial->setParity(p.parity);
    m_serial->setStopBits(p.stopBits);
    m_serial->setFlowControl(p.flowControl);

    if (m_serial->open(QIODevice::ReadWrite))
    {
        m_ui->idLineEdit->setEnabled(true);
        m_ui->connectButton->setEnabled(false);
        m_ui->disconnectButton->setEnabled(true);
        m_ui->settingsButton->setEnabled(false);
        m_ui->enterLineEdit->setEnabled(true);
        m_ui->progressBar->setValue(100);

        showStatusMessage(tr("Подключено к %1 : %2, %3, %4, %5, %6")
                          .arg(p.name).arg(p.stringBaudRate).arg(p.stringDataBits)
                          .arg(p.stringParity).arg(p.stringStopBits).arg(p.stringFlowControl));
    }
    else
    {
        QMessageBox::critical(this, tr("Error"), m_serial->errorString());
        showStatusMessage("Ошибка открытия COM-порта. Убедитесь, что выбран правильный COM-порт,"
                           " а также заданы правильные параметры для подключения устройства");
    }
}

void database::closeSerialPort()
{
    if (m_serial->isOpen())
        m_serial->close();

    m_ui->connectButton->setEnabled(true);
    m_ui->disconnectButton->setEnabled(false);
    m_ui->settingsButton->setEnabled(true);
    m_ui->addButton->setEnabled(false);
    m_ui->idLineEdit->setEnabled(false);
    m_ui->enterLineEdit->setEnabled(false);
    m_ui->progressBar->setValue(0);
    m_ui->idLineEdit->clear();
    m_ui->enterLineEdit->clear();
    showStatusMessage(tr("Устройство отключено"));
}

void database::readData()
{
    QByteArray incomingData;

    m_ui->idLineEdit->clear();

    if(m_serial->canReadLine())
        incomingData = m_serial->readLine();

    m_ui->idLineEdit->setText(incomingData);

    QString text = m_ui->enterLineEdit->text();
    if (!text.isEmpty())
        m_ui->addButton->setEnabled(true);
}

void database::handleError(QSerialPort::SerialPortError error)
{
    if (error == QSerialPort::ResourceError) {
        QMessageBox::critical(this, tr("Error"), m_serial->errorString());
        closeSerialPort();
    }
}

void database::initActionsConnections()
{
    connect(m_ui->connectButton, &QPushButton::clicked, this, &database::openSerialPort);
    connect(m_ui->disconnectButton, &QPushButton::clicked, this, &database::closeSerialPort);
    connect(m_ui->settingsButton, &QPushButton::clicked, m_settings, &SettingsDialog::show);
    connect(m_ui->helpButton, &QPushButton::clicked, this, &database::helpClicked);
    connect(m_ui->addButton, &QPushButton::clicked, this, &database::addClicked);
    connect(m_ui->synchButton, &QPushButton::clicked, this, &database::synchClicked);
}

void database::showStatusMessage(const QString &string)
{
    QMessageBox::information(this, "Database", string);
}

void database::lineChanged(QString str)
{
    QString text = m_ui->idLineEdit->text();
    if ((m_ui->disconnectButton->isEnabled()) && (!text.isEmpty()))
        m_ui->addButton->setEnabled(!str.isEmpty());
}

void database::helpClicked()
{
    QMessageBox::information(this, tr("About"), "Database v1.3\noltsirkov@mail.ru");
}

void database::addClicked()
{
    QString idLine = m_ui->idLineEdit->text();
    QString nameLine = m_ui->enterLineEdit->text();
    string bufId;
    wstring bufName;

    QMessageBox::warning(this, tr("Error"),
                                 "Файл формата excel Database.xls не существует или не находится в папке с программой.");

    QFile file ("Database.xls");
    if (!file.exists())
    {
        QMessageBox::critical(this, tr("Error"),
                              "Файл формата excel Database.xls не существует или не находится в папке с программой.");
        return;
    }

    BasicExcel xls("Database.xls");
    BasicExcelWorksheet* sheet = xls.GetWorksheet(0);
    BasicExcelCell* cell;

    size_t maxRows = sheet->GetTotalRows();
    size_t maxCols = sheet->GetTotalCols();

    cell = sheet->Cell(0, 0);
    bufId = cell->Type();

    if ((maxRows == 1) && (maxCols == 1) && (bufId.front() == 0))
    {
        XLSFormatManager fmtMgr(xls);
        ExcelFont fontBold;
        fontBold._weight = FW_BOLD;
        CellFormat fmtBold(fmtMgr);
        fmtBold.set_font(fontBold);

        sheet->SetColWidth(0, 4000);
        sheet->SetColWidth(1, 4000);

        cell = sheet->Cell(0, 0);
        cell->SetFormat(fmtBold);
        cell->SetString("ID");

        cell = sheet->Cell(0, 1);
        cell->SetFormat(fmtBold);
        cell->SetWString(L"Фамилия");
    }

    for (size_t row = 1; row < maxRows; row++)
    {
        cell = sheet->Cell(row, 0);
        bufId = cell->GetString();

        if (bufId == idLine.toStdString())
        {
            QString infBox;
            QString QBufName = QString::fromStdWString(bufName);

            cell = sheet->Cell(row, 1);
            bufName= cell->GetWString();

            infBox = "Идентификатор \"" + idLine.remove(8,2) + "\" уже существует с фамилией \"" + QBufName +
                    "\". \nХотите перезаписать фамилию?";
            QMessageBox *msg = new QMessageBox(QMessageBox::Warning, "Предупреждение", infBox,
                                               QMessageBox::Yes | QMessageBox::No);
            int choise = msg->exec();
            delete msg;

            if(choise == QMessageBox::Yes)
            {
                cell->SetWString(nameLine.toStdWString().c_str());
                xls.SaveAs("Database.xls");

                m_ui->idLineEdit->clear();
                m_ui->enterLineEdit->clear();
                m_ui->addButton->setEnabled(false);

                QMessageBox::information(this, "Database", ("Фамилия \"" + nameLine + "\" с ID \"" + idLine.remove(8,2) +
                                                     "\" успешно занесена в файл Database.xls"));
                return;
            }
            else
                return;
        }
    }

    cell = sheet->Cell(maxRows, 0);
    cell->SetString(idLine.toStdString().c_str());

    cell = sheet->Cell(maxRows, 1);
    cell->SetWString(nameLine.toStdWString().c_str());

    xls.SaveAs("Database.xls");

    m_ui->idLineEdit->clear();
    m_ui->enterLineEdit->clear();
    m_ui->addButton->setEnabled(false);

    QMessageBox::information(this, "Database", ("Фамилия \"" + nameLine + "\" с ID \"" + idLine.remove(8,2) +
                                         "\" успешно занесена в файл Database.xls"));
}

bool database::checkLang(int n, QString str)
{
    for (int i = 0; i < n; i++)
    {
        if (!(str[i] >= 32 && str[i] <= 127))
        {
            return false;
        }
    }
    return true;
}

void database::synchClicked()
{
    QFile file ("Database.xls");
    if (!file.exists())
    {
        QMessageBox::critical(this, tr("Error"),
                              "Файл формата excel Database.xls не существует или не находится в папке с программой.");
        return;
    }

    QMessageBox::warning(this, "Предупреждение", "Выберите файл формата excel, который нужно синхронизировать с базой данных."
                                           "\nПримечание: файлы Database.xls и выбранный файл должны быть закрыты."
                                           " Также они не должны содержать русских символов в пути к файлам.");

    QString str = QFileDialog::getOpenFileName(NULL, "Проводник", "", "Excel files(*.xls)");
    if (str.contains("Database.xls"))
    {
        QMessageBox::critical(this, tr("Error"),
                              "Был выбран файл базы данных Database.xls \nВыберите файл,"
                              " в который необходимо скопировать фамилии и повторите попытку.");
        return;
    }
    else
    {
        if (checkLang(str.size(), str) == false)
        {
            QMessageBox::critical(this, tr("Error"),
                                  "Путь к выбранному файлу формата excel содержит русские символы.");
            return;
        }

        string bufId;
        string bufIdOut;
        wstring wBufName;

        BasicExcel xls("Database.xls");
        BasicExcelWorksheet* sheet = xls.GetWorksheet(0);
        BasicExcelCell* cell;

        BasicExcel xls_out(str.toStdString().c_str());
        BasicExcelWorksheet* sheetOut = xls_out.GetWorksheet(0);
        BasicExcelCell* cellOut;

        size_t maxRows = sheet->GetTotalRows();
        size_t maxCols = sheet->GetTotalCols();
        size_t maxRowsOut = sheetOut->GetTotalRows();

        if ((maxRows <= 1) || (maxCols <= 1))
        {
            QMessageBox::critical(this, tr("Error"),
                                  "Файл формата excel Database.xls не содержит записей о фамилиях.");
            return;
        }

        for (int row = 1; row < maxRows; row++)
        {
            cell = sheet->Cell(row, 0);
            bufId = cell->GetString();
            for (int rowOut = 7; rowOut < maxRowsOut; rowOut++)
            {
                cellOut = sheetOut->Cell(rowOut, 1);
                bufIdOut = cellOut->GetString();
                if (bufId == bufIdOut)
                {
                    cell = sheet->Cell(row, 1);
                    wBufName = cell->GetWString();
                    cellOut = sheetOut->Cell(rowOut, 2);
                    cellOut->SetWString(wBufName.c_str());
                }
            }
        }
        xls_out.Save();
        QMessageBox::information(this, "Database", "Синхронизация завершена.");
    }
}
