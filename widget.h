#ifndef WIDGET_H
#define WIDGET_H

#include <QWidget>

QT_BEGIN_NAMESPACE
namespace Ui {
class Widget;
}
QT_END_NAMESPACE

class Widget : public QWidget
{
    Q_OBJECT

public:
    Widget(QWidget *parent = nullptr);
    void onInfilebuttonClicked(void);
    void onOutfilebuttonClicked(void);
    int onLoadbuttonClicked(void);
    int onSavebuttonClicked(void);
    void onAutobuttonClicked(void);
    ~Widget();

private:
    Ui::Widget *ui;
};
#endif // WIDGET_H
