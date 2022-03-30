#include "MainWidget.h"

MainWidget::MainWidget()
{
    setWindowTitle("BendingData-1.0");
    this->resize(300, 400);

    StartBtn = new QPushButton("开始处理");

    QVBoxLayout *v = new QVBoxLayout(this);
    v->addWidget(StartBtn);
    this->setLayout(v);

    connect(StartBtn, &QPushButton::clicked, this, &MainWidget::selectFile);
}

void MainWidget::selectFile()
{
    QFileDialog * f = new QFileDialog(this);
    f->setWindowTitle("选择数据文件*.txt");
    f->setNameFilter("*.txt");
    f->setViewMode(QFileDialog::Detail);

    QString filePath;
    if(f->exec() == QDialog::Accepted)
        filePath = f->selectedFiles()[0];

    QString data, force, def1, def2, def3, disp, disp1, disp2, disp3, disp4, stress, strain;
    QFile file(filePath);
    file.open(QIODevice::ReadOnly);
    QTextStream readFile(&file);

    QFile temp1("temp.txt");
    temp1.open(QIODevice::WriteOnly);
    QTextStream writeTemp(&temp1);

    int rowCount = 0;

    while(!readFile.atEnd())
    {
        rowCount++;
        readFile >> data;

        if(rowCount > 105)
        {
            readFile >> force >> def1 >> def2 >> def3 >> disp >> disp1 >> disp2 >> disp3 >> disp4 >> stress >> strain;
            if(force == "=========================分析信息==========================" && rowCount > 105) break;
            writeTemp << force << " " << def1 << " " << def2 << " " << def3 << " " << disp << " " << disp1 << " " << disp2 << " " << disp3 << " " << disp4 << "\n";
        }
    }

    file.close();
    temp1.close();

    QFile temp2("temp.txt");
    temp2.open(QIODevice::ReadOnly);
    QTextStream readTemp(&temp2);

    QXlsx::Document xlsx1("model.xlsx");

    for(int i = 2; i < rowCount - 104; i++)
    {
        for(int j = 1; j < 10; j++)
        {
            readTemp >> data;
            double num = data.toDouble();
            xlsx1.write(i, j, num);
        }
    }

    temp2.remove();
    xlsx1.saveAs("output.xlsx");

    QXlsx::Document xlsx2("model.xlsx");
    QXlsx::Document xlsx3("output.xlsx");

    for(int i = 2; i < rowCount - 104; i++)
    {
        for(int j = 11; j < 14; j++)
        {
            QString num = xlsx2.read(i, j).toString();
            xlsx3.write(i, j, num);
        }
    }
    for(int i = 5; i < 20; i++)
    {
        QString num = xlsx2.read(i, 17).toString();
        xlsx3.write(i, 17, num);
    }
    xlsx3.write(10, 17, "=INDEX(K:K,Q9)");

    QXlsx::Chart *scatterChart = xlsx3.insertChart(20, 15, QSize(500, 300));
    scatterChart->setChartType(QXlsx::Chart::CT_Scatter);
    scatterChart->addSeries(QXlsx::CellRange("K2:L15000"));

    xlsx3.save();

    QMessageBox::information(this, "处理完毕", "数据处理完毕");
}
