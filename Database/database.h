#ifndef DATABASE_H
#define DATABASE_H

#include <QMainWindow>
#include <QSerialPort>
#include <QMessageBox>
#include <QFileDialog>

#include <ExcelFormat.h>

QT_BEGIN_NAMESPACE

namespace Ui {
class database;
}

using namespace ExcelFormat;

QT_END_NAMESPACE

class SettingsDialog;

class database : public QMainWindow
{
    Q_OBJECT

public:
    explicit database(QWidget *parent = 0);
    ~database();

private slots:
    void lineChanged(QString str);
    void handleError(QSerialPort::SerialPortError error);
    void readData();
    void openSerialPort();
    void closeSerialPort();
    void helpClicked();
    void addClicked();
    void synchClicked();

private:
    void showStatusMessage(const QString &message);
    void initActionsConnections();
    bool checkLang(int n, QString str);

    Ui::database *m_ui = nullptr;
    QSerialPort *m_serial = nullptr;
    SettingsDialog *m_settings = nullptr;
};

#endif // DATABASE_H
