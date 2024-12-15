#include "widget.h"
#include "OpenXLSX.hpp"
#include "ui_widget.h"
#include "settings.h"
#include <QMessageBox>
#include <QFileDialog>
#include <QInputDialog>
#include <cstdlib>

Widget::Widget(QWidget *parent):QWidget(parent),ui(new Ui::Widget)
{
    ui->setupUi(this);
    this->setWindowTitle(QString("BOM2BSH Tool - ") + QString(VERSION));
    connect(ui->infileb,&QPushButton::clicked,this,qOverload<>(&Widget::onInfilebuttonClicked));
    connect(ui->outfileb,&QPushButton::clicked,this,qOverload<>(&Widget::onOutfilebuttonClicked));
    connect(ui->load,&QPushButton::clicked,this,qOverload<>(&Widget::onLoadbuttonClicked));
    connect(ui->save,&QPushButton::clicked,this,qOverload<>(&Widget::onSavebuttonClicked));
    connect(ui->autogen,&QPushButton::clicked,this,qOverload<>(&Widget::onAutobuttonClicked));
}
void Widget::onInfilebuttonClicked(void)
{
    QString filename=QFileDialog::getOpenFileName(this,"选择输入文件",QDir::currentPath(),"XLSX表格文件(*.xlsx)");
    ui->infile->setText(filename);
}
void Widget::onOutfilebuttonClicked(void)
{
    QString filename=QFileDialog::getSaveFileName(this,"选择输出文件",QDir::currentPath(),"XLSX表格文件(*.xlsx)");
    ui->outfile->setText(filename);
}
int Widget::onLoadbuttonClicked(void)
{
    std::string str,str2,type;
    std::map<std::string,int> headercol,whrow;
    char cstr[128];
    int i,j,cnt;
    QString filename=ui->infile->text();
    QFileInfo fileinfo(filename);
    OpenXLSX::XLDocument doc;
    QStringList workbooks;
    QString workbookname,tempqs;
    ui->progress->setValue(0);
    if(fileinfo.exists()&&fileinfo.isFile())
    {
        doc.open(filename.toStdString());
        auto sheetnames=doc.workbook().worksheetNames();
        if(sheetnames.size()>1)
        {
            for(auto s:doc.workbook().worksheetNames())
            {
                workbooks<<QString::fromStdString(s);
            }
            workbookname=QInputDialog::getItem(this,"工作表不唯一","请选择工作表",workbooks,0,false);
        }
        else
        {
            workbookname=QString::fromStdString(sheetnames[0]);
        }
        auto workbook=doc.workbook().worksheet(workbookname.toStdString());
        while(ui->table->rowCount())
        {
            ui->table->removeRow(0);
        }
        ui->progress->setMaximum(workbook.rowCount());
        for(i=1;i<=workbook.columnCount();++i)
        {
            str=workbook.cell(2,i).getString();
            if(str=="料号"||str=="PN")
            {
                headercol["lh"]=i;
            }
            else if(str=="描述")
            {
                headercol["ms"]=i;
            }
            else if(str=="位号")
            {
                headercol["wh"]=i;
            }
            else if(str=="数量"||str=="单耗")
            {
                headercol["sl"]=i;
            }
            else if(str=="工序")
            {
                headercol["gx"]=i;
            }
            else if(str=="元件")
            {
                headercol["yj"]=i;
            }
        }
        for(i=3;i<=workbook.rowCount();++i)
        {
            if(workbook.cell(i,1).getString()=="贴片物料")
            {
                type="SMT";
                continue;
            }
            else if(workbook.cell(i,1).getString()=="插件物料")
            {
                type="HOL";
                continue;
            }
            str=workbook.cell(i,headercol["wh"]).getString();
            cnt=0;
            j=0;
            while(j<str.size())
            {
                ui->table->insertRow(ui->table->rowCount());
                str2.clear();
                while(j<str.size()&&(str[j]>='A'&&str[j]<='Z'||str[j]>='0'&&str[j]<='9'))
                {
                    str2+=str[j];
                    j++;
                }
                ui->table->setItem(ui->table->rowCount()-1,0,new QTableWidgetItem(QString::fromStdString(str2)));
                if(whrow.count(str2)>=1)
                {
                    sprintf(cstr,"发现位号出现问题：位号%s第一次出现于第%d行，第二次出现于第%d行",str2.c_str(),whrow[str2],i);
                    QMessageBox::critical(this,"错误",cstr);
                    return 0;
                }
                whrow[str2]=i;
                if(j<str.size()-2&&str[j]==','&&str[j+1]==' '&&(str[j+2]>='A'&&str[j+2]<='Z'||str[j+2]>='0'&&str[j+2]<='9'))
                {
                    j+=2;
                }
                else if(j==str.size()||j==str.size()-1&&str[j]==' ')
                {
                    j=str.size();
                }
                else
                {
                    sprintf(cstr,"发现分隔符出现问题：位于第%d行。你打算怎么做：",i);
                    tempqs=QInputDialog::getItem(this,"警告",cstr,{"忽略","终止"},0,false);
                    if(tempqs=="忽略")
                    {
                        while(j<str.size()&&!(str[j]>='A'&&str[j]<='Z'||str[j]>='0'&&str[j]<='9'))
                        {
                            j++;
                        }
                    }
                    else
                    {
                        QMessageBox::critical(this,"错误","任务被用户取消","确定");
                        return 0;
                    }
                }
                ui->table->setItem(ui->table->rowCount()-1,1,new QTableWidgetItem(QString::fromStdString(workbook.cell(i,headercol["lh"]).getString())));
                ui->table->setItem(ui->table->rowCount()-1,2,new QTableWidgetItem(QString::fromStdString(workbook.cell(i,headercol["ms"]).getString())));
                ui->table->setItem(ui->table->rowCount()-1,3,new QTableWidgetItem("1"));
                ui->table->setItem(ui->table->rowCount()-1,4,new QTableWidgetItem("A"));
                if(headercol.count("gx")==1)
                {
                    if(workbook.cell(i,headercol["gx"]).getString()=="贴片")
                    {
                        str2="SMD";
                    }
                    else if(workbook.cell(i,headercol["gx"]).getString()=="插件")
                    {
                        str2="HOL";
                    }
                    else if(workbook.cell(i,headercol["gx"]).getString()=="PCB")
                    {
                        str2="PCB";
                    }
                    else
                    {
                        sprintf(cstr,"发现未知工序问题：位于第%d行。请手动修改表格。",i);
                        QMessageBox::warning(this,"警告",cstr,"确定");
                        str2="";
                    }
                }
                else
                {
                    if(workbook.cell(i,headercol["yj"]).getString()=="PCB")
                    {
                        str2="PCB";
                    }
                    else if(type.length()==0)
                    {
                        sprintf(cstr,"发现未知工序问题：位于第%d行。请手动修改表格。",i);
                        QMessageBox::warning(this,"警告",cstr,"确定");
                        str2="";
                    }
                    else
                    {
                        str2=type;
                    }
                }
                ui->table->setItem(ui->table->rowCount()-1,5,new QTableWidgetItem(QString::fromStdString(str2)));
                cnt++;
            }
            sprintf(cstr,"%d",cnt);
            if(workbook.cell(i,headercol["sl"]).getString()=="")
            {
                break;
            }
            else if(workbook.cell(i,headercol["sl"]).getString()!=cstr)
            {
                sprintf(cstr,"发现数量位号不对等问题：位于第%d行。你打算怎么做：",i);
                tempqs=QInputDialog::getItem(this,"警告",cstr,{"忽略","终止"},0,false);
                if(tempqs!="忽略")
                {
                    QMessageBox::critical(this,"错误","任务被用户取消","确定");
                    return 0;
                }
            }
            ui->progress->setValue(i);
        }
        ui->progress->setValue(ui->progress->maximum());
        return 1;
    }
    else
    {
        QMessageBox::critical(this,"错误","输入文件不存在","确定");
        return 0;
    }
}
int Widget::onSavebuttonClicked(void)
{
    QString filename=ui->outfile->text();
    OpenXLSX::XLDocument doc;
    if(filename.isEmpty())
    {
        QMessageBox::critical(this,"错误","输出文件不能为空","确定");
        return 0;
    }
    doc.create(filename.toStdString());
    int i,j;
    doc.workbook().addWorksheet("Sheet");
    for(auto name:doc.workbook().sheetNames())
    {
        if(name!="Sheet")
        {
            doc.workbook().deleteSheet(name);
        }
    }
    auto workbook=doc.workbook().worksheet("Sheet");
    ui->progress->setValue(0);
    ui->progress->setMaximum(ui->table->rowCount()+1);
    workbook.cell(1,1).value()="CAD-REF";
    workbook.cell(1,2).value()="MATERIALNUMBER";
    workbook.cell(1,3).value()="DESIGNATION";
    workbook.cell(1,4).value()="AMOUNT";
    workbook.cell(1,5).value()="REVISION";
    workbook.cell(1,6).value()="MOUNTING";
    ui->progress->setValue(1);
    for(i=2;i<=ui->table->rowCount()+1;++i)
    {
        for(j=1;j<=6;++j)
        {
            workbook.cell(i,j).value()=ui->table->item(i-2,j-1)->text().toStdString();
        }
        ui->progress->setValue(i);
    }
    doc.save();
    return 1;
}
void Widget::onAutobuttonClicked(void)
{
    if(this->onLoadbuttonClicked())
    {
        if(this->onSavebuttonClicked())
        {
            QMessageBox::information(this,"成功","指令执行成功");
        }
    }
}
Widget::~Widget()
{
    delete ui;
}
