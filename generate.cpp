#include <iostream>
#include <windows.h>
#include "libxl.h"
#include <sys/types.h>
#include <sys/stat.h>
#include <map>
#include <vector>


using namespace libxl;
using namespace std;
using Base = std::map <std::wstring, int>;

const wchar_t* filename = L"2.xlsx"; //файл склада
const wchar_t* filename_out = L"99.xlsx"; //файл инвойса
const wchar_t* filename_otchet = L"10.xlsx"; //файл продаж

//находит количество юнтов  если есть
int IsExist(std::wstring unit, Base& base) {
    if (base.count(unit)) {
        int result = base[unit];
        base.erase(unit); // убираем уже найденую позиции, что бы повторно по ней не искать.
        return result;
 }
    return 0;
}

struct PositionName {
    int y; // позиция наименование
    int x; // позиция наименование
    int q; // позиция количества
};

const wchar_t* NameFile(const wchar_t* filename, PositionName pos, const wchar_t* text) {
    Book* otchet = xlCreateXMLBook();
    if (otchet->load(filename)) {
        std::wcout << filename << L" - Книга открыта" << endl;
        Sheet* sheet = otchet->getSheet(0);

        auto name = sheet->readStr(pos.y, pos.x);
        std::wcout << text << L" : " << name << endl << endl;
        otchet->release();
        return name;

    }
}

Base FileToMap(int start_line, int start_line_name, const wchar_t* filename, PositionName pos) {
    Book* otchet = xlCreateXMLBook();

    if (otchet->load(filename) ){                 

       Sheet* sheet = otchet->getSheet(0);
       auto name = sheet->readStr(pos.y, pos.x);

       Base result;
      static int i = 0;
       int line = start_line;    
       std::wstring unit;
       int q=0;
       for(int ii=start_line; ii<50+start_line; ++ii){            
           if (sheet->readStr(line, start_line_name)) {                                              
                unit = sheet->readStr(line, start_line_name);
                q = sheet->readNum(line, pos.q);                       
                ++line; ++i;
           }          
            result[unit] = q;
        }
        otchet->release();
        return result;
    }
return {};
}

struct Sbor{
 //  const wchar_t* name;
    std::wstring name;
    int need;
    int exist;
};

std::vector<Sbor> DoCheckList(Base& stock, const Base& list) {
    int i1 = 0;
    std::vector<Sbor> result;
    for (auto& unit : list) {
        int quantity = IsExist(unit.first, stock);
        if (quantity > 0) {
            wcout << L"Найдено : " << ++i1 << ")" << unit.first << L" Надо : " << unit.second << L" В наличии : " << quantity << endl;
            Sbor a;
            a.name = unit.first;
            a.need = unit.second;
            a.exist = quantity;
            result.push_back(a);            
        }
    }
    return result;
}


void Invoice(std::wstring& name_stock, const std::vector<Sbor>& list1, int q) {
    Book* book = xlCreateXMLBook();
    Font* boldFont = book->addFont();
    boldFont->setBold();

    Font* titleFont = book->addFont();
    titleFont->setName(L"Arial Black");
    titleFont->setSize(14);

    Format* titleFormat = book->addFormat();
    titleFormat->setFont(titleFont);

    Format* headerFormat = book->addFormat();
    headerFormat->setAlignH(ALIGNH_CENTER);
    headerFormat->setBorder(BORDERSTYLE_THIN);
    headerFormat->setFont(boldFont);
    headerFormat->setFillPattern(FILLPATTERN_SOLID);
    headerFormat->setPatternForegroundColor(COLOR_TAN);

    Format* descriptionFormat = book->addFormat();
    descriptionFormat->setBorderLeft(BORDERSTYLE_THIN);

    Format* amountFormat = book->addFormat();
    amountFormat->setNumFormat(NUMFORMAT_CURRENCY_NEGBRA);
    amountFormat->setBorderLeft(BORDERSTYLE_THIN);
    amountFormat->setBorderRight(BORDERSTYLE_THIN);

    Format* totalLabelFormat = book->addFormat();
    totalLabelFormat->setBorderTop(BORDERSTYLE_THIN);
    totalLabelFormat->setAlignH(ALIGNH_RIGHT);
    totalLabelFormat->setFont(boldFont);

    Format* totalFormat = book->addFormat();
    totalFormat->setNumFormat(NUMFORMAT_CURRENCY_NEGBRA);
    totalFormat->setBorder(BORDERSTYLE_THIN);
    totalFormat->setFont(boldFont);
    totalFormat->setFillPattern(FILLPATTERN_SOLID);
    totalFormat->setPatternForegroundColor(COLOR_YELLOW);

    Format* signatureFormat = book->addFormat();
    signatureFormat->setAlignH(ALIGNH_CENTER);
    signatureFormat->setBorderTop(BORDERSTYLE_THIN);

    Sheet* sheet = book->addSheet(L"Invoice");
    if (sheet)
    {
        sheet->setCol(0, 0, 3);
        sheet->setCol(1, 1, 60);
        sheet->setCol(2, 2, 3);
        sheet->setCol(3, 3, 6);
        sheet->setCol(4, 4, 6);

        wchar_t* stock_name = name_stock.data();
        sheet->writeStr(2, 1, stock_name, titleFormat);
        sheet->writeStr(3, 1, L"Наименование ");
        sheet->writeStr(3, 3, L"Пордажи ");
        sheet->writeStr(3, 4, L"СКЛАД ");
        int i = 4;
        for (const auto& el : list1) {

            wstring item_name = el.name;
            wchar_t* iname = item_name.data();
            sheet->writeNum(i, 0, i - 3);
            sheet->writeNum(i, 3, el.need);
            sheet->writeNum(i, 4, el.exist);
            sheet->writeStr(i++, 1, iname);

        }
        sheet->writeStr(i, 1, L"Итого ");
        sheet->writeNum(i, 2, q);
        sheet->setPrintArea(0, i, 0, 4);
        sheet->setPrintGridlines();
    }

    if (book->save(filename_out))
    {
        ::ShellExecute(NULL, L"open", filename_out, NULL, NULL, SW_SHOW);
    }

    book->release();
}



int main()
{
    setlocale(LC_ALL, "ru_RU.UTF-8");
    setlocale(LC_ALL, "Russian");
    
    // заполняем базу продаж
    // сложное заполнение т.к. в триал версии функция readSTR выполняется лимитое количество раз
    // будь они прокляты :)
    PositionName pos{ 6, 2, 4 };
    int start_line = 9;
    int start_line_name = 0;
    wstring name_stock = NameFile(filename_otchet, pos, L"Заказ по ");
    Base otchet, temp;
    otchet = FileToMap(start_line, start_line_name, filename_otchet, pos);
  
    for (int a = 1; a < 50; ++a) {
        start_line += 50;
        temp = FileToMap(start_line, start_line_name, filename_otchet, pos);
        otchet.insert(temp.begin(), temp.end());
   
    }   
    wcout << endl << L"Размер базы заказа : " << otchet.size() << endl;

    //заполняем базу склада
    // сложное заполнение т.к. в триал версии функция readSTR выполняется лимитое количество раз
    // будь они прокляты :)
    PositionName pos1{ 2, 14, 14 };
    start_line = 8;
    start_line_name = 5;
    Base b1,b2;
    NameFile(filename, pos1,L"Формируется по складу");
    b1 = FileToMap(start_line, start_line_name, filename, pos1);
   
    for(int a=1; a<50; ++a){
        start_line += 50;
        b2 = FileToMap(start_line, start_line_name, filename, pos1);
        b1.insert(b2.begin(), b2.end());
    }    
    wcout << endl << L"Размер базы склада : " << b1.size() <<endl;
    
    // сопостовляем базу продаж с базой склада
    std::vector<Sbor> list1 =  DoCheckList(b1, otchet);
    wcout << L" Позиций  : " << list1.size() <<endl;
    int q=0;
    for (const auto& el : list1) {
        q += el.need;
    }
    wcout << L" Общее количество   : " << q <<endl;

    
    // подготовка инвойса
    Invoice(name_stock,list1,q);
    
    std::cin.get();
    return 0;
    
}
