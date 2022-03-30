#ifndef MAINWIDGET_H
#define MAINWIDGET_H

#include <QWidget>
#include <QLabel>
#include <QVBoxLayout>
#include <QPushButton>
#include <QMessageBox>
#include <QFileDialog>
#include <QFile>
#include <QTextStream>
#include "xlsxdocument.h"
#include "xlsxcellrange.h"
#include "xlsxchart.h"

class MainWidget : public QWidget
{
    Q_OBJECT

public:
    MainWidget();
    void selectFile();

private:
    QPushButton * StartBtn;
};
#endif // MAINWIDGET_H
