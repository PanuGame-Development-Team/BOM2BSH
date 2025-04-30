#include "widget.h"
#include "xlnt/xlnt.hpp"
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
    std::string workbookname,str,str2,type;
    std::map<std::string,int> headercol,whrow;
    char cstr[128];
    int i,j,cnt;
    QString filename=ui->infile->text();
    QFileInfo fileinfo(filename);
    xlnt::workbook doc;
    xlnt::worksheet ws;
    QStringList workbooks;
    QString tempqs;
    ui->progress->setValue(0);
    if(fileinfo.exists()&&fileinfo.isFile())
    {
        doc.load(filename.toStdString());
        if(doc.sheet_titles().size()>1)
        {
            for(std::string title:doc.sheet_titles())
            {
                workbooks<<QString::fromStdString(title);
            }
            workbookname=QInputDialog::getItem(this,"工作表不唯一","请选择工作表",workbooks,0,false).toStdString();
        }
        else
        {
            workbookname=doc.sheet_titles()[0];
        }
        ws=doc.sheet_by_title(workbookname);
        while(ui->table->rowCount())
        {
            ui->table->removeRow(0);
        }
        ui->progress->setMaximum(ws.highest_row());
        for(i=1;i<=ws.highest_column();++i)
        {
            str=ws.cell(i,2).to_string();
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
        for(i=3;i<=ws.highest_row();++i)
        {
            if(ws.cell(1,i).to_string()=="贴片物料")
            {
                type="SMT";
                continue;
            }
            else if(ws.cell(1,i).to_string()=="插件物料")
            {
                type="HOL";
                continue;
            }
            str=ws.cell(headercol["wh"],i).to_string();
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
                ui->table->setItem(ui->table->rowCount()-1,1,new QTableWidgetItem(QString::fromStdString(ws.cell(headercol["lh"],i).to_string())));
                ui->table->setItem(ui->table->rowCount()-1,2,new QTableWidgetItem(QString::fromStdString(ws.cell(headercol["ms"],i).to_string())));
                ui->table->setItem(ui->table->rowCount()-1,3,new QTableWidgetItem("1"));
                ui->table->setItem(ui->table->rowCount()-1,4,new QTableWidgetItem("A"));
                if(headercol.count("gx")==1)
                {
                    if(ws.cell(headercol["gx"],i).to_string()=="贴片")
                    {
                        str2="SMD";
                    }
                    else if(ws.cell(headercol["gx"],i).to_string()=="插件")
                    {
                        str2="HOL";
                    }
                    else if(ws.cell(headercol["gx"],i).to_string()=="PCB")
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
                    if(ws.cell(headercol["yj"],i).to_string()=="PCB")
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
            if(ws.cell(headercol["sl"],i).to_string()=="")
            {
                break;
            }
            else if(ws.cell(headercol["sl"],i).value<int>()!=cnt)
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
    xlnt::workbook doc;
    xlnt::worksheet ws;
    if(filename.isEmpty())
    {
        QMessageBox::critical(this,"错误","输出文件不能为空","确定");
        return 0;
    }
    int i,j;
    doc.create_sheet();
    ws=doc.active_sheet();
    ui->progress->setValue(0);
    ui->progress->setMaximum(ui->table->rowCount()+1);
    ws.cell(1,1).value("CAD-REF");
    ws.cell(2,1).value("MATERIALNUMBER");
    ws.cell(3,1).value("DESIGNATION");
    ws.cell(4,1).value("AMOUNT");
    ws.cell(5,1).value("REVISION");
    ws.cell(6,1).value("MOUNTING");
    ui->progress->setValue(1);
    for(i=2;i<=ui->table->rowCount()+1;++i)
    {
        for(j=1;j<=6;++j)
        {
            ws.cell(j,i).value(ui->table->item(i-2,j-1)->text().toStdString());
        }
        ui->progress->setValue(i);
    }
    doc.save(filename.toStdString());
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
