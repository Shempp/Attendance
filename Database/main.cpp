#include "database.h"
#include <QApplication>

int main(int argc, char *argv[])
{
    QApplication a(argc, argv);
    database w;
    w.setMinimumSize(485, 135);
    w.setMaximumSize(485, 135);
    w.show();

    return a.exec();
}
