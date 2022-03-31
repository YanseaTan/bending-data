#include "MainWidget.h"

MainWidget::MainWidget()
{
    setWindowIcon(QIcon(":/logo.ico"));
    setWindowTitle("BendingData-1.0");
    this->resize(280, 150);

    //默认试验参数
    span = 150;
    section = 40;

    SelectFileBtn = new QPushButton("选择文件");

    QGroupBox * Menu = menu();

    QVBoxLayout *v = new QVBoxLayout(this);
    v->addWidget(Menu);
    v->addWidget(SelectFileBtn);
    this->setLayout(v);

    connect(SelectFileBtn, &QPushButton::clicked, this, &MainWidget::selectFile);
}

QGroupBox * MainWidget::menu()
{
    QGroupBox * box = new QGroupBox("试验参数");

    Text1 = new QLabel;
    Text2 = new QLabel;
    MM1 = new QLabel;
    MM2 = new QLabel;
    Text1->setText("四点弯曲的跨度：");
    Text2->setText("截面高度：");
    MM1->setText("mm");
    MM2->setText("mm");

    Line1 = new QLineEdit;
    Line2 = new QLineEdit;
    //在输入框内输入好默认的试验参数
    Line1->setText("150");
    Line2->setText("40");

    QGridLayout * grid = new QGridLayout;
    grid->addWidget(Text1, 0, 0);
    grid->addWidget(Line1, 0, 1);
    grid->addWidget(MM1, 0, 2);
    grid->addWidget(Text2, 1, 0);
    grid->addWidget(Line2, 1, 1);
    grid->addWidget(MM2, 1, 2);
    box->setLayout(grid);

    //当用户按下回车键，或者鼠标点击输入框外的其它位置时对试验参数进行更新
    connect(Line1, &QLineEdit::editingFinished, this, &MainWidget::updateParam);
    connect(Line2, &QLineEdit::editingFinished, this, &MainWidget::updateParam);

    return box;
}

void MainWidget::updateParam()
{
    //将用户输入好的文本转换为 int 赋值给变量
    span = Line1->text().toInt();
    section = Line2->text().toInt();
}

void MainWidget::selectFile()
{
    //弹出文件对话框进行文件选择
    QFileDialog * f = new QFileDialog(this);
    f->setWindowTitle("选择数据文件*.txt");
    f->setNameFilter("*.txt");
    f->setViewMode(QFileDialog::Detail);

    //当用户选择文件后，将路径保存到 filePath中
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

    //去除原文件除数据外的冗余内容，将数据暂存在 temp.txt 中
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

    //将 temp.txt 文件中的数据导入到 xlsx 文件中
    for(int i = 2; i < rowCount - 104; i++)
    {
        for(int j = 1; j < 10; j++)
        {
            readTemp >> data;
            double num = data.toDouble();
            xlsx1.write(i, j, num);
        }
    }
    //写入试验参数
    xlsx1.write(1, 17, span);
    xlsx1.write(2, 17, section);

    temp2.remove();

    QString outputPath;
    //这一步会让 filePath 的长度也减少4，因此要放在后面进行操作，否则会影响前面 filePath 的读取。
    int index = filePath.length()-4;
    //创建输出文件路径与命名，与原文件只有文件后缀不同
    outputPath = filePath.replace(index, 4, ".xlsx");
    xlsx1.saveAs(outputPath);

    QXlsx::Document xlsx2("model.xlsx");
    QXlsx::Document xlsx3(outputPath);

    //将部分公式复制到需要输出的文件中
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

    //创建散点图
    QXlsx::Chart *scatterChart = xlsx3.insertChart(20, 15, QSize(500, 300));
    scatterChart->setChartType(QXlsx::Chart::CT_Scatter);
    scatterChart->addSeries(QXlsx::CellRange("K2:L15000"));

    xlsx3.save();

    QMessageBox::information(this, "处理完毕", "数据处理完毕");
}
