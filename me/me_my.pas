// ************************************************
// Сокращение рутиной работы. Routines for formulas.
// Author:    	Андрей А. Кандауров (Andrey A. Kandaurov)
// Companiya: 	Santig
// e-mail:  	  san@santig.ru
// URL:       	http://santig.ru
// License:   	zlib
// ************************************************

(* Copyright (c) 2015-2021 Santig

This software is provided 'as-is', without any express or implied
warranty. In no event will the authors be held liable for any damages
arising from the use of this software.

Permission is granted to anyone to use this software for any purpose,
including commercial applications, and to alter it and redistribute it
freely, subject to the following restrictions:

1. The origin of this software must not be misrepresented; you must not
   claim that you wrote the original software. If you use this software
   in a product, an acknowledgement in the product documentation would be
   appreciated but is not required.
2. Altered source versions must be plainly marked as such, and must not be
   misrepresented as being the original software.
3. This notice may not be removed or altered from any source distribution. *)

unit me_my;

{$mode objfpc}{$H+}{$X+}
{$WARN 4105 off : Implicit string type conversion with potential data loss from "$1" to "$2"}
{$ModeSwitch advancedrecords}

(*
FileExists('/home/user/a.zip'); - Для проверки существования файла
CleanAndExpandFilename(name) - добавляет в name путь до exe, удаляет пробелы
CleanAndExpandDirectory      - добавляет путь до exe, делает правильные слеши, добавляет в конец \
IncludeTrailingPathDelimiter(path)
AppendPathDelim(file)        - добавляет слэш в конце пути
TrimAndExpandDirectory       - добавляет слэш, даже к файлу
SwitchPathDelims             - делает правильные слэши
LastDelimiter                - возвращает индекс последнего (самого правого) вхождения
                               любого из символов Delimters в строке S, иначе 0.
IsDelimiter                  - является ли индексный символ в строке S символом- разделителем,
                               переданным в разделителях. Если индекс выходит за пределы диапазона, возвращается значение False .
IsValidIdent(str; AllowDots, StrictDots: Boolean = False): Boolean; - корректный идентификатор Pascal,
                               AllowDots=True - разрешает точки в любом месте,
                               StrictDots=True - запрещает точки по краям.
FindAllFiles                 - поиск всех файлов в папке (и подпапках) LazUtils
FindDefaultExecutablePath    - поиск exe в PATH (GetEnvironmentVariableUTF8('PATH');) FileUtil
GetAllFilesMask;             - win: *.*   Lin *  {fileutil.inc}

ForceDirectoriesUTF8         - Создать папки с подпапками

// Показ составляющих частей этого полного имени  {LazFileUtils}
  fullFileName := ExtractFilePath(
  fullFileName := 'C:\MyProg\Projects\Unit1.pas';
  ShowMessage('Диск       = '+ExtractFileDrive (fullFileName));   C:
  ShowMessage('Путь       = '+ExtractFilePath  (fullFileName));   C:\MyProg\Projects\
  ShowMessage('Каталог    = '+ExtractFileDir   (fullFileName));   C:\MyProg\Projects
  ShowMessage('-каталог   = '+ChompPathDelim   (fullFileName));   C:\MyProg\
  ShowMessage('Имя+расш   = '+ExtractFileName  (fullFileName));   Unit1.pas
  ShowMessage('Имя        = '+ExtractFileNameOnly(fullFileName)); Unit1
  ShowMessage('Расширение = '+ExtractFileExt   (fullFileName));   .pas
  ShowMessage('без Расширения='+ExtractFileNameWithoutExt(fullFileName)); C:\MyProg\Projects\Unit1
s := ExtractFilePath(Application.ExeName)

  SwapCase(Str): String        - конвертирует нижний в верхний регистр (LCLProc)
  StringCase
  TextToSingleLine (str):str   - мультилинию в строку, оставляет по 1 пробелу

  :=Format(statTxt, [RowCount.ToString]));
  {$ModeSwitch advancedrecords} - что бы записать в record метод
   AnimateWindow(MonthCalendar1.Handle, 200, AW_SLIDE or AW_VER_POSITIVE);

В Двумерном массиве задаются СТРОКИ потом СТОЛБЦЫ.

procedure TForm1.StringGrid1DrawCell(Sender: TObject; aCol, aRow: Integer; aRect: TRect; aState: TGridDrawState);
begin
  if (gdFocused in aState) and (StringGrid1.EditorMode) then begin
    StringGrid1.Editor.Color := clBlack;
    StringGrid1.Editor.Font.Color := clWhite;
  end;
end;
*)

{$WARN 5044 off : Symbol "$1" is not portable}
interface

uses
  {$IFDEF UNIX} {$IFDEF UseCThreads} cthreads, {$ENDIF} {$ENDIF}
  Classes, Windows, Forms, Dialogs, Buttons, Grids, LCLProc, Graphics, SysUtils,
  LazFileUtils, StdCtrls, ComCtrls, LazUTF8, FileInfo, Winsock, registry,
  RegExpr, strutils, eventlog;

type
  TarrStr  = array of string;
  TarrChar = array of char;
  TsetChar = set of char;

{запись объекта}
type
  // записать заначение в дерево TTreeView:
  //TreeView1.Items[1].Data := TTreeDates.Create('значение');
  // Получаем сохраненное ранее значение:
  //ShowMessage( TTreeDates(TreeView1.Items[1].Data).Value );
  TTreeDates = class
  private
    FValue: variant;
  public
    constructor Create(Value: variant);
    property Value: variant read FValue;
  end;

  {Версия исполняемого файла, например:
  Caption:= vers.FileVersion; }

  { TFileVersInit }

  TFileVersInit = record
    Company, FileDescription: string;
    InternalName, LegalCopyright, OriginalFilename: string;
    ProductName, ProductVersion: string;
    FileVersion: string;
    v_star, v_mlad, v_reviz, v_bild: word;

    //true - если versionNew > versionOLD, сравнивать до релиза: verN (max<=4)
    function compare(const versionNew, versionOld: string; verN: byte = 3): boolean;
  end;
//Версия файла
function FileVersInit: TFileVersInit;


{ID и ключ windows}
type
  TWinInf = record
    //Windows Product Name
    Name:      string;
    //Windows Product ID
    ProductID: string;
    BuildGUID: string;
    //Windows 8 Key
    keyWin:    string;
    //номер сборки
    BuildNumber: integer;
    // версия:  MajorVersion.MinorVersion
    CurrentVersion: string;
    //старшая версия
    MajorVersion: integer;
    //младшая версия
    MinorVersion: integer;
    //платформа
    Platforms: integer;
    //сервис паки
    CSDVersion: ShortString;
    //поля генератора случайных чисел
    //Хэшь от строки= gen2(7)+BuildGUID+"-"+      577+1571: function TfSD.identif: boolean; + procedure TfSD.Vhod;
    //               "gen3(5)-"+                  function winInfo(ke: byte = 0): TWinInf;
    //               ключ_win+"-gen1(5)"+"-"+ID+  1885: procedure TfSD.Vhod;
    //В работе winInf.g1 или (winInf.g2 + winInf.g3);
    g1, g2, g3: string;
  end;
//информация о виндовс
function winInfo: TWinInf;

var
  winInf: TWinInf;
//----TWinInf-----------------------


type
  {Текст в Статус баре окна, +логирование, +автоматическое отслеживание
     за разделом лога в доч.формах.  }

  {Логирование myFs.Log Инициализируется в основной программе (1 раз), см myFs}

  { TstatusBarMy }

  {Для статуса:
  1) создать статус на форме (StatusBar1) и добавить панели (по умолч. 1)
  2) stat:  TstatusBarMy; //определить переменную статусбара в var
  3) stat.Free; //разрушить в форме по OnDestroy
  4) stat  := TstatusBarMy.CreateStat(StatusBar1); //определить в FormCreate

  5) при необходимости определить раздел. Для дочерних форм можно своё имя
     stat.razdel := 'ПСА';
     //stat.poleN := 0;   //в какое поле статуса заносить данные по умолчанию

  Использовать:
      stat.Txt(s); stat.Log(s,stat.poleN, s1) - текст ТОЛЬКО в Статус бар
      stat.Log(s);      - в статусс бар и лог файл
      stat.Warning(s);  - в статусс бар и лог файл
      stat.Error(s);    - в статусс бар и лог файл
  }

  TstatusBarMy = class
    razdelOld: string;
  private
    //Статусбар формы. Определить при открытии формы
    statusBarForms: TStatusBar;
    // номер основного поля для текста
    textHints:      byte;

    // Читать раздел
    function statusIdentRead: string;
    // Задать раздел
    procedure statusIdentWrite(razdel: string);
  public
    constructor CreateStat(stat: TStatusBar);
    destructor Destroy; override;
    //текст только в Статус бар
    procedure Txt(Text: string; ApoleN: byte = 255);
    //запись в статусс бар и лог файл
    procedure Log(Text: string; ApoleN: byte = 255; strVlog: string = '');
    //запись в статусс бар и лог файл
    procedure Warning(Text: string; ApoleN: byte = 255; strVlog: string = '');
    //запись в статусс бар и лог файл
    procedure Error(Text: string; ApoleN: byte = 255; strVlog: string = '');

    // Задать раздел
    property razdel: string read statusIdentRead write statusIdentWrite;
    // номер основного поля для текста
    property poleN: byte read textHints write textHints default 0;
  end;


{Шифрование строк}

type
  compl = record
    sh: string;
    n:  integer;
  end;

  // шифровать строки в record, использовать так:
  //sh1 := CSh.Create;
  //    c := sh1.shifr(s, k);
  //    _ms([c.sh, sh1.AnShifr(c, k)]);
  //sh1.Free;
  CSh = class
    //private
  public
    //шифрование строки, кол-во лидирующих нулей сохраняется
    function shifr(Astr, Akey: string): compl;
    //Расшифровать
    function AnShifr(strShifr: compl; Akey: string): string;
  end;

{Работа с шифрованием строк на базе одной случайной строки}

// сгенерировать шаблон случайных символов
procedure shifrShablGen(var shabl: string);
//Кодировать строку случ паролем (строка это позиции в шаблоне)
function shifrKod(val: string; var shabl: string; pswLen: word = 10): string;
//Кодировать строку случ паролем и шаблоном (строка это позиции в шаблоне)
//спрятать пароль и шаблон в строку
function shifrShablKod(Source: string; pswLen: word = 10): string;
//ДеКодировать строку со спрятанным паролем
function shifrDeKod(Source, shabl: string): string;
//ДеКодировать строку со спрятанным паролем и шаблоном
function shifrShablDeKod(Source: string): string;

const
  shifrLenMax = 5;

var
  shifrShabl: string;

{шифрование-------------}

type
  // all - учесть всё; notPrint - убрать не печатные; tchk - убрать точки на конце
  TValidIdentifier = set of (viFile, viPath, viNotPrint);
// Конвертирует строку в правильный Pascal-идентификатор
function ValidStr(aValue: string; Flags: TValidIdentifier = [viFile]): string;


// Вывод значений на экран. Обрамить всё в []
procedure _ms(const Args: array of const; err: string = '??');
//Преобразовать переменные в строку
function toStr(const Args: array of const; raz: string = ' ';
  err: string = '??'): string;
 //Преобразовать одномерный массив в строку
 //function toStr(Arr: char; raz: string = ' '; err: string = ''): string;
function toStr(const Arr: array of double; raz: string = ' '; err: string = ''): string;
function toStr(const Arr: array of word; raz: string = ' '; err: string = ''): string;
function toStr(const Arr: array of integer; raz: string = ' ';
  err: string = ''): string;
 //function toStr(Arr: array of real; raz: string = ' '; err: string = ''): string;
 //function toStr(Arr: array of byte; raz: string = ' '; err: string = ''): string;

//Перевод в строку. Строки обрамляются obraml = ''...''
function toStrBD(const Arr: array of const; err: string = 'NULL';
  razgr: string = ', '; obraml: string = ''''): string;

//Преобразовать переменные в число
function toInt(Args: array of const; err: integer = -1): integer;
//Преобразовать переменную в число и вернуть строкой
function toInt(Args: array of const; err: integer = -1): string; overload;
//Преобразовать переменные в variant
function toPer(Args: array of const; err: variant): variant;

// Получить текст с клавиатуры
function _inp(Caption, Podpis: string; var Value: string): boolean;
//проверить ввод c клавы, если пусто или ошибка, то default
procedure _inpValid(var Value: string; default: string;
  flag: TValidIdentifier = [viNotPrint]);
//выдаёт индекс первого расхождения в строках
function CompareLen(s1, s2: string): integer;



 // постепенно увеличивает форму с шагом: shag
 // добавить в событие формы: FormActivate (onActivate)
procedure formMaxresize(FormVar: TForm; shag: byte = 10);

// стандартный StringReplace. Была замена - True
function StringReplaceBool(var str: string; oldStr, newStr: string;
  flag: TReplaceFlags = [rfReplaceAll, rfIgnoreCase]): boolean;
//найти и выделить из строки номер
function strNamber(strNambers: string): string;
// Находим и выделяем дату из строки (new=true - при ошибке подставляет системную дату)
// MessageDlg('Дата документа','Установленна текущая дата',mtWarning,mbOK,0);
function strDate(strDates: string; newDate: boolean = True): string;
// удалить дату из строки вместе с: от и г.
procedure strDateDel(var strValue: string; Data: string);

// Инициализировать папку в AppData\Local\
function appDataInit(folder: string): string;

const
  myCDig      = '0123456789';
  myCCharH    = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
  myCCharL    = 'abcdefghijklmnopqrstuvwxyz';
  myCHex      = '0123456789ABCDEF';
  myCHexL     = '0123456789abcdef';
  myCBrackets = '()[]{}';
  myCZnaki    = myCBrackets + '.,<>!?/\;"''@#$%^&*-+|:=_`~№';
  myCnotFile  = '/\:|?*"<>';
//Генерировать случайное значение
function strGen(leng: byte = 7; chars: string = myCDig + myCCharL): string;

{ работа с ПК}

// Получить имя компьютера
function GetComputerNetName: string;
// Получить имя пользователя машины
function GetSystemUserName: string;
// узнать МАС адрес ПК
function GetMACAddress: string;
//Получить хост ПК
function GetMyHostName: string;
//-----------------------------------


// Сохранить компонент (класс) в файл
function ClassSaveFile(RootObject: TComponent; const FileName: TFileName): boolean;
// Прочитать компонент (класс) из файла
function ClassLoadFile(RootObject: TComponent; const FileName: TFileName): boolean;

{StringGrid}

//Добавить новую колонку с указанными параметрами
procedure GridAddCol(var grid: TStringGrid; const ParamZnach: array of const);
//удалить строки с повторяющимися значениями в столбце
function gridDelPovtor(var sg: TStringGrid; const col: integer): boolean;


const
  //набор табуляций
  tb1 = #9;
  tb2 = #9#9;
  tb3 = #9#9#9;
  tb4 = #9#9#9#9;

  nL = #$D#$A;     // Новая строка в сообщениях
  nR = LineEnding; //#10#13; system  //Новая строка



var
  { Версия файла}
  //Company:      string;
  //FileDescription: string;
  versFile: TFileVersInit;
  //InternalName: string;
  //LegalCopyright: string;
  //OriginalFilename: string;
  //ProductName:  string;
  //ProductVersion: string;
  {------------------}

  //настройка для работы с датой
  myFormatDate: TFormatSettings = (CurrencyFormat: 3;
  NegCurrFormat: 8;
  ThousandSeparator: ' ';
  DecimalSeparator: '.';
  CurrencyDecimals: 2;
  DateSeparator: '-';
  TimeSeparator: ':';
  ListSeparator: ',';
  CurrencyString: 'р.';
  ShortDateFormat: 'dd.mm.yyyy';
  LongDateFormat: 'dd" "mmmm" "yyyy';
  TimeAMString: 'AM';
  TimePMString: 'PM';
  ShortTimeFormat: 'hh:nn';
  LongTimeFormat: 'hh:nn:ss';
  ShortMonthNames: ('Янв', 'Фев', 'Март', 'Апр',
    'Май', 'Июнь', 'Июль', 'Авг', 'Сент', 'Окт',
    'Нояб', 'Дек');
  LongMonthNames: ('Январь', 'Февраль', 'Март',
    'Апрель', 'Май', 'Июнь', 'Июль', 'Август',
    'Сентябрь', 'Октябрь', 'Ноябрь', 'Декабрь');
  ShortDayNames: ('Вс', 'Пн', 'Вт', 'Ср', 'Чт', 'Пт', 'Сб');
  LongDayNames: ('Воскресенье', 'Понедельник',
    'Вторник', 'Среда', 'Четверг', 'Пятница',
    'Суббота');
  TwoDigitYearCenturyWindow: 50;
  );

//-------------- var end ---------------------



implementation

uses m_myFs;

var
  // идентификация в логе
  logRazd: string = '-> "me_my"' + #9;
//function tryFunc(fu: Function):


procedure _ms(const Args: array of const; err: string = '??');
var
  ss: string = '';
  sf: string;
  i:  byte = 0;
begin
  sf := '%.' + Length(IntToStr(Length(Args) - 1)).ToString + 'd)   ';
  for i := Low(Args) to High(Args) do
    begin
    ss += Format(sf, [i + 1]);
    //ss += (Args[i].Name+': ';
      try
      ss += toStr(Args[i], ' ', err);
      except
      ss += err;
      end;
    ss += nL;
    end;
  ShowMessage(ss);
end;

//Преобразовать переменные в строку
function toStr(const Args: array of const; raz: string = ' ';
  err: string = '??'): string;
var
  s:  string = '';
  ss: string = '';
  i:  byte;
begin
  for i := Low(Args) to High(Args) do
    begin
      try
      //if Args[i] is type

      case Args[i].VType of
        vtInteger: s := dbgs(Args[i].vinteger);
        vtInt64: s := dbgs(Args[i].VInt64^);
        vtQWord: s := dbgs(Args[i].VQWord^);
        vtBoolean: s := dbgs(Args[i].vboolean);
        vtExtended: s := dbgs(Args[i].VExtended^);
        {$ifdef FPC_CURRENCY_IS_INT64}
        // MWE:
        // fpc 2.x has troubles in choosing the right dbgs()
        // so we convert here
        vtCurrency: s := dbgs(int64(Args[i].vCurrency^) / 10000, 4);
  {$else}
        vtCurrency: ss := dbgs(Args[i].vCurrency^);
  {$endif}
        vtString: s := Args[i].VString^;
        vtAnsiString: s := ansistring(Args[i].VAnsiString);
        vtChar: s := Args[i].VChar;
        vtPChar: s := Args[i].VPChar;
        vtUnicodeString: s :=
            ansistring(UnicodeString(Args[i].VUnicodeString));
        //DbgS(Args[i].VUnicodeString);
        vtPWideChar: s := {%H-}s {%H-} + Args[i].VPWideChar;
        vtWideChar: s := ansistring(Args[i].VWideChar);
        vtWidestring: s := ansistring(WideString(Args[i].VWideString));
        vtObject: s := DbgSName(Args[i].VObject);
        vtClass: s := DbgSName(Args[i].VClass);
        vtPointer: s := Dbgs(Args[i].VPointer);
        vtVariant: s := DbgS(Args[i].VVariant)
        else
          s := err; //'??';
        end;
      except
      s := err;
      end;
    if i = High(Args) then
      raz := '';
    ss    += Trim(s) + raz;
    end;
  Result := ss;
end;

 ////Преобразовать одномерный массив в строку
 //function toStr(Arr: char; raz: string = ' '; err: string = ''): string;
 //var
 //  s: string = '';
 //  i: byte;
 //  ch: char;
 //begin
 //  for ch in Arr do; // i := Low(Arr) to High(Arr) do
 //    begin
 //      try
 //      s += toStr([arr[i]], raz, err);
 //      except
 //      s += err;
 //      end;
 //    if i <> High(Arr) then
 //      s += raz;
 //    end;
 //  Result := Trim(s);
 //end;

//Преобразовать одномерный массив в строку
function toStr(const Arr: array of double; raz: string = ' '; err: string = ''): string;
var
  s: string = '';
  i: byte;
begin
  for i := Low(Arr) to High(Arr) do
    begin
      try
      s += toStr([arr[i]], raz, err);
      except
      s += err;
      end;
    if i <> High(Arr) then
      s += raz;
    end;
  Result := Trim(s);
end;

//Преобразовать одномерный массив в строку
function toStr(const Arr: array of word; raz: string = ' '; err: string = ''): string;
var
  s: string = '';
  i: byte;
begin
  for i := Low(Arr) to High(Arr) do
    begin
      try
      s += toStr([arr[i]], raz, err);
      except
      s += err;
      end;
    if i <> High(Arr) then
      s += raz;
    end;
  Result := Trim(s);
end;

//Преобразовать одномерный массив в строку
function toStr(const Arr: array of integer; raz: string = ' ';
  err: string = ''): string;
var
  s: string = '';
  i: byte;
begin
  for i := Low(Arr) to High(Arr) do
    begin
      try
      s += toStr([arr[i]], raz, err);
      except
      s += err;
      end;
    if i <> High(Arr) then
      s += raz;
    end;
  Result := Trim(s);
end;

//Преобразовать одномерный массив в строку
function toStr(Arr: array of byte; raz: string = ' '; err: string = ''): string;
var
  s: string = '';
  i: byte;
begin
  for i := Low(Arr) to High(Arr) do
    begin
      try
      s += toStr([arr[i]], raz, err);
      except
      s += err;
      end;
    if i <> High(Arr) then
      s += raz;
    end;
  Result := Trim(s);
end;

//Перевод в строку. Строки обрамляются obraml = ''...''
function toStrBD(const Arr: array of const; err: string = 'NULL';
  razgr: string = ', '; obraml: string = ''''): string;
var
  s: string = '';
  i: byte;
begin
  for i := Low(Arr) to High(Arr) do
    begin
      try
      case Arr[i].VType of
        vtString, vtAnsiString, vtChar, vtPChar, vtPWideChar,
        vtWideChar, vtWidestring, vtObject, vtClass:
          s += obraml + toStr(Arr[i], '', err) + obraml;
          //vtBoolean: s += toStr([Arr[i].VInteger], razgr, err)
        else
          s += toStr(Arr[i], razgr, err);
        end;
      except
      s += err;
      end;
    if i <> High(Arr) then
      s += razgr;
    end;
  s := Trim(s);
  if s.IsEmpty then
    s    := err;
  Result := StringReplace(s, '''NULL''', 'NULL', [rfReplaceAll, rfIgnoreCase]);
end;

// ----------------------------- toStr -------------------------------

{  toInt }

//Преобразовать переменную в число
function toInt(Args: array of const; err: integer = -1): integer;
var
  i:   integer;
  s:   string;
  rez: string = '';
  b:   boolean = False;
begin
  Result := err;
  s      := toStr(Args[0], '', '-1');
  if s = 'True' then
    Exit(1);
  if s = 'False' then
    exit(0);
    try
    for i := 1 to Length(s) do
      if s[i] in strutils.DigitChars then
        begin
        b   := True;
        rez := rez + s[i];
        end
      else
        if b then
          Break;
    if not TryStrToInt(rez, Result) then
      Result := err;
    except
    Result := err;
    end;
end;

//Преобразовать переменную в число и вернуть строкой
function toInt(Args: array of const; err: integer = -1): string;
begin
  Result := IntToStr(toInt(Args[0], err));
end;

// ---------------------------- toInt ---------------------------------

//Преобразовать переменные в variant
function toPer(Args: array of const; err: variant): variant;
begin
  //for i := Low(Args) to High(Args) do
    try
    case Args[0].VType of
      vtInteger: Result := variant(Args[0].vinteger);
      //      vtInt64: Result := dbgs(Args[i].VInt64^);
      //      vtQWord: Result := dbgs(Args[i].VQWord^);
      //      vtBoolean: Result := dbgs(Args[i].vboolean);
      //      vtExtended: Result := dbgs(Args[i].VExtended^);
      //      {$ifdef FPC_CURRENCY_IS_INT64}
      //      // MWE:
      //      // fpc 2.x has troubles in choosing the right dbgs()
      //      // so we convert here
      //      vtCurrency: s := dbgs(int64(Args[i].vCurrency^) / 10000, 4);
      //{$else}
      //      vtCurrency: ss := dbgs(Args[i].vCurrency^);
      //{$endif}
      //      vtString: Result := Args[i].VString^;
      //      vtAnsiString: Result := ansistring(Args[i].VAnsiString);
      //      vtChar: Result := Args[i].VChar;
      //      vtPChar: Result := Args[i].VPChar;
      //      vtUnicodeString: Result := ansistring(UnicodeString(Args[i].VUnicodeString));
      //      //DbgS(Args[i].VUnicodeString);
      //      vtPWideChar: Result := {%H-}Result {%H-} + Args[i].VPWideChar;
      //      vtWideChar: Result := ansistring(Args[i].VWideChar);
      //      vtWidestring: Result := ansistring(WideString(Args[i].VWideString));
      //      vtObject: Result := DbgSName(Args[i].VObject);
      //      vtClass: Result := DbgSName(Args[i].VClass);
      //      vtPointer: Result := Dbgs(Args[i].VPointer);
      vtVariant: Result := variant(Args[0].VVariant^)
      else
        Result := err; //'??';
      end;
    except
    Result := err;
    end;
end;
// ------------------------ toPer --------------------------------


//информация о виндовс
function winInfo: TWinInf;
var
  Reg:  TRegistry;
  regKeyNT, valb, regKey: string;
  date: TBytes = nil;

  function myConvertToKey(dByte: TBytes): string;   //TMemoryStream
  const
    KeyOffset = 52;
  var
    j, Cur, y, Last, isWin8: integer;
    Chars, keypart1, ins: string;
    winKeyOutput:  string = '';
    a, b, c, d, e: string;
  begin
    isWin8 := trunc(dByte[66] / 6) and 1;
    dByte[66] := (dByte[66] and $F7) or ((isWin8 and 2) * 4);
    j     := 24;
    Chars := 'BCDFGHJKMPQRTVWXY2346789';
    repeat
      Cur := 0;
      y   := 14;
      repeat
        Cur := Cur * 256;
        Cur := dByte[y + KeyOffset] + Cur;
        dByte[y + KeyOffset] := trunc(Cur / 24);
        Cur := Cur mod 24;
        y   := y - 1;
      until y < 0;
      j    := j - 1;
      winKeyOutput := Copy(Chars, Cur + 1, 1) + winKeyOutput;
      Last := Cur;
    until j < 0;
    if (isWin8 = 1) then
      begin
      keypart1 := Copy(winKeyOutput, 2, Last);
      ins      := 'N';
      winKeyOutput := StringReplace(winKeyOutput, keypart1, keypart1 + ins, []);
      //2, 1, 0);
      if Last = 0 then
        winKeyOutput := ins + winKeyOutput;
      end;
    a      := Copy(winKeyOutput, 1, 5);
    b      := Copy(winKeyOutput, 6, 5);
    c      := Copy(winKeyOutput, 11, 5);
    d      := Copy(winKeyOutput, 16, 5);
    e      := Copy(winKeyOutput, 21, 5);
    Result := a + '-' + b + '-' + c + '-' + d + '-' + e;
  end;

begin
  regKeyNT := '\SOFTWARE\Microsoft\Windows NT\CurrentVersion';
  regKey   := '\SOFTWARE\Microsoft\Windows\CurrentVersion';
  Reg      := TRegistry.Create;
  Reg.RootKey := HKEY_LOCAL_MACHINE;
  //HKEY_CURRENT_USER; REG_KEY_DONT_SILENT_FAIL
    try
    if IsWindow(Application.Handle) = True then
      begin
      Reg.Access := KEY_WOW64_64KEY;
      Reg.OpenKeyReadOnly(regKeyNT);
      end
    else
      begin
      Reg.Access := KEY_WOW64_32KEY;
      Reg.OpenKeyReadOnly(regKey);
      end;

    with Result do
      begin
      //Windows Product ProductID
        try
        ProductID := WinCPToUTF8(reg.ReadString('ProductId'));
        except
        ProductID := '';
        end;
      //Windows Product Name
        try
        Name := Reg.ReadString('ProductName');
        except
        Name := '';
        end;
        try
        BuildGUID := reg.ReadString('BuildGUID');
        except
        BuildGUID := '';
        end;

      //Windows 8 Key
      valb   := 'DigitalProductId';
      keyWin := '';
      if Reg.ValueExists(valb) then
          try
          SetLength(date, Reg.GetDataSize(Valb));
          reg.ReadBinaryData(valb, Pointer(date)^, Reg.GetDataSize(Valb));
          g3     := strGen(5, myCHex);
          keyWin := myConvertToKey(date);
          except
          keyWin := 'BBBBB-BBBBB-BBBBB-BBBBB-BBBBB';
          end;
      //номер сборки
      BuildNumber  := Win32BuildNumber;
      //старшая версия
      MajorVersion := Win32MajorVersion;
      //младшая версия
      MinorVersion := Win32MinorVersion;
        try
        CurrentVersion := reg.ReadString('CurrentVersion');
        except
        CurrentVersion := '';
        end;
      //платформа
      Platforms  := Win32Platform;
      CSDVersion := Win32CSDVersion;
      end;
    finally
    Reg.CloseKey;
    Reg.Free;
    end;
  //with Result do
  //  _ms([key8, Id, Name, 'BuildGUID ' + BuildGUID, 'BuildNumber ',
  //    BuildNumber, 'MajorVersion', MajorVersion, 'MinorVersion', MinorVersion,
  //    'CurrentVersion', CurrentVersion, 'Platforms', Platforms, 'UBR', UBR,
  //    'CSDVersion', CSDVersion, 'GetMyHostName', GetMyHostName]);
end;

// Инициализировать папку в AppData\Local\
function appDataInit(folder: string): string;
var
  s: string = '';
begin
  if not folder.IsEmpty then
    begin
    {$IFDEF windows}
    s := GetEnvironmentVariableUTF8('USERPROFILE') + '\AppData\Local\';
  {$ELSE}
    s := AppendPathDelim(GetEnvironmentVariableUTF8('HOME'));
  {$endif}
    s := SwitchPathDelims(s + folder, pdsSystem);
    if not DirectoryExistsUTF8(s) then
      CreateDir(s);
    end;
  Result := AppendPathDelim(s);
end;

//Генерировать случайное значение
function strGen(leng: byte = 7; chars: string = myCDig + myCCharL): string;
var
  i:    integer;
  chs:  string = '';
  a, n: integer;
begin
  Result := '';
  for i := Low(chars) to High(chars) do
    chs += chars[i];
  n     := High(chs);

  Randomize;
  for i := 1 to leng do
    begin
    a      := Random(n) + 1;
    Result += chs[a];
    end;
end;

 { TFileVersInit }
 //Версия файла
function FileVersInit: TFileVersInit;
var
  FileVers_: TFileVersionInfo;
  tmp: TFileVersInit;
  sl:  TStringList;
begin
  FillChar(tmp, SizeOf(tmp), 0);
  FileVers_ := TFileVersionInfo.Create(nil);
  with FileVers_.VersionStrings do
      try
      FileVers_.ReadFileInfo;
      tmp.Company := Values['CompanyName'];
      tmp.FileDescription := Values['FileDescription'];
      tmp.InternalName := Values['InternalName'];
      tmp.LegalCopyright := Values['LegalCopyright'];
      tmp.OriginalFilename := Values['OriginalFilename'];
      tmp.ProductName := Values['ProductName'];
      tmp.ProductVersion := Values['ProductVersion'];
      tmp.FileVersion := Values['FileVersion'];
      sl := TStringList.Create;
      sl.Delimiter := '.';
      sl.StrictDelimiter := True;
      sl.DelimitedText := tmp.FileVersion;
      if sl.Count > 0 then
        begin
        tmp.v_star  := StrToIntDef(sl[0], 0);
        tmp.v_mlad  := StrToIntDef(sl[1], 0);
        tmp.v_reviz := StrToIntDef(sl[2], 0);
        tmp.v_bild  := StrToIntDef(sl[3], 0);
        end;
      sl.Free;
      finally
      FileVers_.Free;
      end;
  Result := tmp;
end;

//true - если versionNew > versionOLD, сравнивать до релиза: verN (max<=4)
function TFileVersInit.compare(const versionNew, versionOld: string;
  //например: Разница в не совместимости версий!!! Удалить файл.
  //if not ver.compare(IniPropStorage1.ReadString('vers', ''), '0.8.2', 3) then
  verN: byte = 3): boolean;
var
  sN, sO: TStringList;
  i:      integer;
begin
  Result := False;
  sN     := TStringList.Create;
  sN.Delimiter := '.';
  sN.StrictDelimiter := True;
  sN.DelimitedText := versionNew;
  for i := sN.Count to 4 - 1 do
    sN.Add('0');

  sO := TStringList.Create;
  sO.Delimiter := '.';
  sO.StrictDelimiter := True;
  sO.DelimitedText := versionOld;
  for i := sO.Count to 4 - 1 do
    sO.Add('0');

  //_ms(['New: '+sN[0]+'.'+sN[1]+'.'+sN[2]+'.'+sN[3],
  //'Old: '+sO[0]+'.'+sO[1]+'.'+sO[2]+'.'+sO[3],
  //'sN[1]>sO[1]: '+ tostr([StrToIntDef(sN[1], 0) > StrToIntDef(sO[1], 0)]),
  //'sN[2]>sO[2]: '+ tostr([StrToIntDef(sN[2], 0) > StrToIntDef(sO[2], 0)])]);
    try
    if StrToIntDef(sN[0], 0) > StrToIntDef(sO[0], 0) then
      exit(True);
    if (verN >= 2) and (StrToIntDef(sN[1], 0) > StrToIntDef(sO[1], 0)) then
      exit(True);
    if (verN >= 3) and (StrToIntDef(sN[2], 0) > StrToIntDef(sO[2], 0)) then
      exit(True);
    if (verN >= 4) and (StrToIntDef(sN[3], 0) > StrToIntDef(sO[3], 0)) then
      exit(True);
    finally
    sN.Free;
    sO.Free;
    end;
end;

{ TstatusBarMy }

function TstatusBarMy.statusIdentRead: string;
begin
  Result := myFs.Log.Identification;
end;

procedure TstatusBarMy.statusIdentWrite(razdel: string);
begin
  //запомним старый раздел, чтоб потом его восстаовить
  if Assigned(myFs) then
    begin
    razdelOld := myFs.Log.Identification;
    myFs.Log.Identification := razdel;
    end;
end;

constructor TstatusBarMy.CreateStat(stat: TStatusBar);
begin
  inherited Create;
  statusBarForms := stat;
end;

destructor TstatusBarMy.Destroy;
begin
    try
    if Assigned(myFs) then
      myFs.Log.Identification := razdelOld;
    except
    on E: Exception do
      _ms(['Exception: ', E.ClassName, E.Message]);
    end;
  inherited;// Destroy;
end;

//текст только в Статус бар
procedure TstatusBarMy.Txt(Text: string; ApoleN: byte = 255);
var
  n: byte;
begin
  if ApoleN = 255 then
    n := poleN
  else
    n := ApoleN;
  statusBarForms.Panels[n].Text := Text;
  statusBarForms.Update;
end;

//запись в статусс бар и лог файл
procedure TstatusBarMy.Log(Text: string; ApoleN: byte = 255; strVlog: string = '');
begin
  txt(Text, ApoleN);
  if strVlog.IsEmpty then
    myFs.Log.Info(Text)
  else
    myFs.Log.Info(Text + ' :: ' + strVlog);
end;

//запись в статус бар и лог файл
procedure TstatusBarMy.Warning(Text: string; ApoleN: byte = 255; strVlog: string = '');
begin
  txt(Text, ApoleN);
  if strVlog.IsEmpty then
    myFs.Log.Warning(Text)
  else
    myFs.Log.Warning(Text + ' :: ' + strVlog);
end;

//запись в статусс бар и лог файл
procedure TstatusBarMy.Error(Text: string; ApoleN: byte = 255; strVlog: string = '');





begin
  txt(Text, ApoleN);
  if strVlog.IsEmpty then
    myFs.Log.Error(Text)
  else
    myFs.Log.Error(Text + ' :: ' + strVlog);
end;

 //--------- Статус end -------
 { ------------- TstatusBarMy ----------------- }


{ CSh }

//шифрование строки, кол-во лидирующих нулей сохраняется
function CSh.shifr(Astr, Akey: string): compl;
var
  i: integer = 1;
begin
  Result.n  := 0;
  Result.sh := XorEncode(Akey, Astr);
  while Result.sh[i] = '0' do
    begin
    Inc(Result.n);
    Delete(Result.sh, i, i);
    end;
end;

//Расшифровать
function CSh.AnShifr(strShifr: compl; Akey: string): string;
var
  i: integer;
begin
  Result := '';
  for i := 1 to strShifr.n do
    Result += '0';
  Result   := XorDecode(Akey, Result + strShifr.sh);
end;

//----------------------------------------

//{ TTreeDates }
constructor TTreeDates.Create(Value: variant);
begin
  FValue := Value;
end;

{Работа с шифрованием строк на базе одной случайной строки}

// сгенерировать шаблон случайных символов
procedure shifrShablGen(var shabl: string);
var
  g: string;
begin
  Randomize;
  shabl := '';
  while Length(shabl) < 15 do
    begin
    g := strGen(1, copy(myCHexL, 2, 15));
    if Pos(g, shabl) = 0 then
      shabl += g;
    end;
end;

//Кодировать строку (16 777 215 byte) случ паролем (строка это позиции в шаблоне)
function shifrKod(val: string; var shabl: string; pswLen: word = 10): string;
var
  i:      integer;
  psw:    string = '';
  pswHex: string = '';
  tekPoz, brDelt: integer;
  kod:    string = '';
  ch:     char;
begin
  Randomize;
  if val.IsEmpty then
    Exit('');
  //если шаблон ещё пуст, то создать его
  if shabl.IsEmpty then
    shifrShablGen(shabl);

  // генерируем пароль
  for i := 1 to pswLen do
    begin
    tekPoz := Random(255);
    psw    += chr(tekPoz);
    pswHex += AnsiLowerCase(intToHex(tekPoz, 2));
    end;

  //кодируем в hex - для переноса в программу
  kod := XorEncode(psw, val);

  //спрятать код в строку
  i      := 1;
  tekPoz := 0;
  while i <= pswLen * 2 - 1 do
    begin
    //сгенерировать относительную позицию для вставки нового символа
    repeat
      if i = 1 then
        brDelt := Random(Length(kod) - i) + i
      else
        brDelt := Random(254) - 127;
    until ((brDelt + tekPoz) >= 1) and ((brDelt + tekPoz) <= Length(kod));
    tekPoz += brDelt;

    //вставляем 2 символа пароля
    Insert(pswHex[i] + pswHex[i + 1], kod, tekPoz);
    //вставляем переход в пределах +-127
    if i = 1 then
      Insert('ff', kod, tekPoz)
    else
      Insert(AnsiLowerCase(IntToHex(brDelt + 127, 2)), kod, tekPoz);

    Inc(i, 2);
    end;
  kod += AnsiLowerCase(IntToHex(tekPoz, shifrLenMax));

  Result := '';
  //перестроить текст на основе шаблона
  for ch in kod do
    Result += AnsiLowerCase(IntToHex(Pos(ch, shabl), 1));
end;

//Кодировать строку случ паролем и шаблоном (строка это позиции в шаблоне)
//спрятать пароль и шаблон в строку
function shifrShablKod(Source: string; pswLen: word): string;
var
  s: string = '';
begin
  Result := shifrKod(Source, s, pswLen);
  Result := Copy(s, 1, 5) + Result + Copy(s, 6, 16);
end;

//ДеКодировать строку со спрятанным паролем
function shifrDeKod(Source, shabl: string): string;
var
  kod: string = '';
  ch:  char;
  adr, adr1: integer;
begin
  Result := '';
  if Source.IsEmpty then
    Exit('');
  Result := '';

  //перестроить текст на основе шаблона
  for ch in Source do
    if ch = '0' then
      kod += chr(48)
    else
      kod += shabl[Hex2Dec(ch)];

  // взять первый адрес
  adr := Hex2Dec(copy(kod, Length(kod) - shifrLenMax + 1, shifrLenMax));
  Delete(kod, Length(kod) - shifrLenMax + 1, shifrLenMax);

  //найти пароль в шифрованном сообщении
  repeat
    adr1 := Hex2Dec(copy(kod, adr, 2));
    Delete(kod, adr, 2);
    Insert(chr(StrTointDef('$' + copy(kod, adr, 2), Ord(' '))), Result, 1);
    Delete(kod, adr, 2);
    adr += 127 - adr1;
  until adr1 = 255;

  Result := XorDecode(Result, kod);
end;

//ДеКодировать строку со спрятанным паролем и шаблоном
function shifrShablDeKod(Source: string): string;
var
  s: string = '';
begin
  s := Copy(Source, 1, 5) + Copy(Source, Length(Source) - 9, 15);
  Delete(Source, 1, 5);
  Delete(Source, Length(Source) - 9, 15);
  Result := shifrDeKod(Source, s);
end;


// Конвертирует строку в правильный Pascal-идентификатор
function ValidStr(aValue: string; Flags: TValidIdentifier = [viFile]): string;
var
  nameFile: string = '';
begin
  if aValue.IsEmpty and (viFile in Flags) then
    aValue := strGen(7, myCCharL + myCDig);
  if aValue.IsEmpty and (viPath in Flags) then
    begin
    Result := '';
    exit;
    end;

  if (viFile in Flags) or (viPath in Flags) then
    begin
    if viPath in Flags then
      begin
      if aValue[1] = '.' then
        aValue := Copy(aValue, 2, Length(aValue));
      //имя в пути сканировать особым образом
      nameFile := ExtractFileName(aValue);
      aValue   := ExtractFilePath(aValue);
      // В имени нельзя использовать некоторые символы из пути
      aValue   := StringReplace(aValue, '/', '_', [rfReplaceAll, rfIgnoreCase]);
      end
    else
      begin
      aValue := StringReplace(aValue, '\', '_', [rfReplaceAll, rfIgnoreCase]);
      aValue := StringReplace(aValue, '/', '_', [rfReplaceAll, rfIgnoreCase]);
      aValue := StringReplace(aValue, ':', '_', [rfReplaceAll, rfIgnoreCase]);
      end;
    aValue := StringReplace(aValue, '*', ' ', [rfReplaceAll, rfIgnoreCase]);
    aValue := StringReplace(aValue, '<', ' ', [rfReplaceAll, rfIgnoreCase]);
    aValue := StringReplace(aValue, '>', ' ', [rfReplaceAll, rfIgnoreCase]);
    aValue := StringReplace(aValue, '|', ' ', [rfReplaceAll, rfIgnoreCase]);
    aValue := StringReplace(aValue, '"', '', [rfReplaceAll, rfIgnoreCase]);
    aValue := StringReplace(aValue, '?', '', [rfReplaceAll, rfIgnoreCase]);
    if not aValue.IsEmpty then
      if aValue[Length(aValue)] = '.' then
        aValue := Copy(aValue, 1, Length(aValue) - 1);
    end;
  if (viFile in Flags) or (viPath in Flags) or (viNotPrint in Flags) then
    begin
    aValue := StringReplace(aValue, nR, ' ', [rfReplaceAll, rfIgnoreCase]);
    aValue := StringReplace(aValue, tb1, '  ', [rfReplaceAll, rfIgnoreCase]);
    aValue := StringReplace(aValue, tb2, '    ', [rfReplaceAll, rfIgnoreCase]);
    aValue := StringReplace(aValue, tb3, '      ', [rfReplaceAll, rfIgnoreCase]);
    aValue := StringReplace(aValue, nL, ' ', [rfReplaceAll, rfIgnoreCase]);
    end;
  //при сканировании пути, имя просканировать отдельно, если оно есть
  if (viPath in Flags) and not nameFile.IsEmpty then
    aValue := Trim(aValue) + Trim(ValidStr(nameFile, [viFile]));
  Result   := Trim(aValue);

end; {ValidStr}

// Получить текст с клавиатуры
function _inp(Caption, Podpis: string; var Value: string): boolean;
begin
  Result := False;
    try
    Result := InputQuery(Caption, Podpis, Value);
    except
    Value := '';
    end;
end;

//проверить ввод c клавы, если пусто или ошибка, то default
procedure _inpValid(var Value: string; default: string; flag: TValidIdentifier);

begin
    try
    Value := ValidStr(Value, flag);
    if (Value.IsEmpty) then
      Value := default;
    except
    Value := default;
    end;
end;

//выдаёт индекс первого расхождения в строках
function CompareLen(s1, s2: string): integer;
var
  i, a, b: integer;
begin
  a      := Length(s1);
  b      := Length(s2);
  Result := Max(a, b) - Abs(a - b) + 1;
  for i := 1 to Min(a, b) do
    if s1[i] <> s2[i] then
      begin
      Result := i;
      Break;
      end;
end;


 // постепенно увеличивает форму с шагом: shag
 // добавить в событие формы: FormActivate
procedure formMaxresize(FormVar: TForm; shag: byte = 10);
var
  w, h: integer;
  iw:   integer = 0;
  ih:   integer = 0;
begin
  if FormVar = nil then
    exit;
  w := FormVar.Width;
  h := FormVar.Height;
  FormVar.Width := 0;
  FormVar.Height := 0;
  //FormVar.Resiz;
  Application.ProcessMessages;
  while (iw <= w) or (ih <= h) do
    begin
    if iw <= w then
      begin
      Inc(iw, shag);
      FormVar.Width := iw;
      end;
    if ih <= h then
      begin
      Inc(ih, shag);
      FormVar.Height := ih;
      end;
    //FormVar.Resiz;
    Application.ProcessMessages;
    end;
  FormVar.Width  := w;
  FormVar.Height := h;
end;

//найти и выделить из строки номер
function strNamber(strNambers: string): string;
var
  t_nom: string = '';
begin
  if strNambers.IsEmpty then
    Exit('');
  //if Pos('№', strNambers) > 0 then
  //  t_nom := Trim(Copy(strNambers, Pos('№', strNambers) + Length('№'),
  //    Length(strNambers)))
  //else
  //  if Pos('#', strNambers) > 0 then
  //    t_nom := Trim(Copy(strNambers, Pos('#', strNambers) + Length('#'),
  //      Length(strNambers)));

  //if t_nom.IsEmpty then
  with TRegExpr.Create do
      try
      Expression := '[#№]+([\d|\s])(.+?)\s';
      if Exec(strNambers) then
        begin
        t_nom := StringReplace(Match[0], '№', '', [rfReplaceAll]);
        t_nom := StringReplace(t_nom, '#', '', [rfReplaceAll]);
        end
      else
        begin
        Expression := '[\w]*[\d]+[\w]*'; //'\s*\d(.+?)\s*';
        if Exec(strNambers) then
          t_nom := Match[0];
        end;
      finally
      Free;
      end;
  t_nom := Trim(t_nom);
  if Pos(' ', t_nom) > 0 then
    t_nom := Trim(Copy(t_nom, 1, Pos(' ', t_nom) - 1));
  Result  := t_nom;
end;

// Выделяет дату из строки (newDate=true - подставляет системную дату при '')
// MessageDlg('Дата документа','Установленна текущая дата',mtWarning,mbOK,0);
function strDate(strDates: string; newDate: boolean = True): string;
var
  t_dat: string = '';
  reg:   TRegExpr;
begin
  reg := TRegExpr.Create;
  reg.Expression := '(\d{1,4}[-./\*\\]){1,2}\d{2,4}';
  with Reg do
      try
      if Exec(strDates) then
        begin
        t_dat := Match[0];
        //проверить, что это действительно дата
        t_dat := DateToStr(StrToDate(t_dat));
        end;
      except
        try
        t_dat := Match[1];
        t_dat := DateToStr(StrToDate(t_dat));
        except
          try
          t_dat := Match[2];
          t_dat := DateToStr(StrToDate(t_dat));
          except
          t_dat := '';
          end;
        end;
      end;
  reg.Free;
  t_dat := Trim(t_dat);
  if newDate and t_dat.IsEmpty then
    t_dat := DateToStr(Date);
  Result  := t_dat;
end;

// стандартный StringReplace. Была замена - True
function StringReplaceBool(var str: string; oldStr, newStr: string;
  flag: TReplaceFlags = [rfReplaceAll, rfIgnoreCase]): boolean;
var
  b: boolean = False;
  s: string;
begin
  s   := str;
  str := StringReplace(s, oldStr, newStr, flag);
  if s <> str then
    b    := True;
  Result := b;
end;

// удалить дату (data) из строки strValue вместе с: от и г.
procedure strDateDel(var strValue: string; Data: string);
begin
  //'[ \t]+(от )*' + Data + '([ \t]*г*\.*[ \t]+)';
  if not strValue.IsEmpty then
    if not (StringReplaceBool(strValue, ' от ' + Data + ' г. ',
      ' ', [rfIgnoreCase]) or StringReplaceBool(strValue, '\tот ' +
      Data + ' г. ', ' ', [rfIgnoreCase]) or StringReplaceBool(strValue,
      ' от ' + Data + ' г.\t', ' ', [rfIgnoreCase]) or
      StringReplaceBool(strValue, '\tот ' + Data + ' г.\t', ' ',
      [rfIgnoreCase])) then

      if not (StringReplaceBool(strValue, ' от ' + Data + ' г ',
        ' ', [rfIgnoreCase]) or StringReplaceBool(strValue, '\tот ' +
        Data + ' г ', ' ', [rfIgnoreCase]) or
        StringReplaceBool(strValue, ' от ' + Data + ' г\t', ' ',
        [rfIgnoreCase]) or StringReplaceBool(strValue, '\tот ' +
        Data + ' г\t', ' ', [rfIgnoreCase])) then

        if not (StringReplaceBool(strValue, ' от ' + Data + ' ',
          ' ', [rfIgnoreCase]) or StringReplaceBool(strValue, '\tот ' +
          Data + ' ', ' ', [rfIgnoreCase]) or StringReplaceBool(strValue,
          ' от ' + Data + '\t', ' ', [rfIgnoreCase]) or
          StringReplaceBool(strValue, '\tот ' + Data + '\t', ' ',
          [rfIgnoreCase])) then

          if not (StringReplaceBool(strValue, ' от ' + Data +
            '.', '.', [rfIgnoreCase]) or StringReplaceBool(
            strValue, '\tот ' + Data + '.', '.', [rfIgnoreCase])) then

            if not (StringReplaceBool(strValue, ' ' + Data +
              ' г. ', ' ', [rfIgnoreCase]) or
              StringReplaceBool(strValue, ' ' + Data + ' г. ',
              ' ', [rfIgnoreCase]) or StringReplaceBool(strValue,
              '\t' + Data + ' г.\t', ' ', [rfIgnoreCase]) or
              StringReplaceBool(strValue, '\t' + Data + ' г.\t',
              ' ', [rfIgnoreCase])) then

              if not (StringReplaceBool(strValue, ' ' + Data +
                ' г ', ' ', [rfIgnoreCase]) or
                StringReplaceBool(strValue, ' ' + Data + ' г ',
                ' ', [rfIgnoreCase]) or StringReplaceBool(
                strValue, '\t' + Data + ' г\t', ' ', [rfIgnoreCase]) or
                StringReplaceBool(strValue, '\t' + Data + ' г\t',
                ' ', [rfIgnoreCase])) then

                if not (StringReplaceBool(strValue, Data +
                  ' г. ', ' ', [rfIgnoreCase]) or StringReplaceBool(
                  strValue, Data + ' г.\t', ' ', [rfIgnoreCase]) or
                  StringReplaceBool(strValue, Data + ' г ',
                  ' ', [rfIgnoreCase]) or StringReplaceBool(
                  strValue, Data + ' г\t', ' ', [rfIgnoreCase])) then

                  if not (StringReplaceBool(strValue, ' ' +
                    Data + ' ', ' ', [rfIgnoreCase]) or
                    StringReplaceBool(strValue, '\t' + Data +
                    '\t', ' ', [rfIgnoreCase])) then

                    if not (StringReplaceBool(
                      strValue, Data + ' ', ' ', [rfIgnoreCase]) or
                      StringReplaceBool(strValue, Data + '\t',
                      ' ', [rfIgnoreCase])) then
                      StringReplaceBool(strValue, Data, ' ', [rfIgnoreCase]);

  strValue := Trim(strValue);
end;

{--------- Работа с ПК ----------}

// Получить имя компьютера
function GetComputerNetName: string;
var
  Buf:  array[0..MAXBYTE] of char;
  Size: DWORD;
begin
  Size := SizeOf(Buf);
  if GetComputerName(Buf, Size) then
    Result := Buf
  else
    Result := '';
end;

// Получить имя пользователя машины
function GetSystemUserName: string;
var
  UserName: array[0..MAXBYTE] of char;
  Size:     DWORD;
begin
  Size := SizeOf(UserName);
  if Windows.GetUserName(@UserName, Size) then
    Result := string(UserName)
  else
    Result := '';
end;

// узнать МАС адрес ПК
function GetMACAddress: string;
const
  WSVer = $101;
var
  wsaData: TWSAData;
  P:   PHostEnt;
  Buf: array [0..127] of char;
begin
  Result := '';
  wsaData.wVersion := 0;
  if WSAStartup(WSVer, wsaData) = 0 then
    begin
    if GetHostName(@Buf, 128) = 0 then
      begin
      P := GetHostByName(@Buf);
      if P <> nil then
        Result := iNet_ntoa(PInAddr(p^.h_addr_list^)^);
      end;
    P := nil;
    WSACleanup;
    end;
end;

//Получить хост ПК
function GetMyHostName: string;
const
  WSVer = $101;
var
  wsaData: TWSAData;
  P:   PHostEnt;
  Buf: array [0..127] of char;
begin
  Result := '';
  wsaData.wVersion := 0;
  if WSAStartup(WSVer, wsaData) = 0 then
    begin
    if GetHostName(@Buf, 128) = 0 then
      begin
      P := GetHostByName(@Buf);
      if P <> nil then
        Result := string(p^.h_name);
      end;
    WSACleanup;
    P := nil;
    end;
end;

{---------------}


 // Сохранить компонент (класс) в файл
 // Унаследовать корневой класс от TComponent
function ClassSaveFile(RootObject: TComponent; const FileName: TFileName): boolean;
var
  FileStream: TFileStream;
  MemStream:  TMemoryStream;
begin
  Result     := False;
  FileStream := TFileStream.Create(FileName, fmCreate);
  MemStream  := TMemoryStream.Create;
    try
    MemStream.WriteComponent(RootObject);
    MemStream.Position := 0;
    ObjectBinaryToText(MemStream, FileStream);
    Result := True;
    finally
    MemStream.Free;
    FileStream.Free;
    end;
end;

// Прочитать компонент (класс) из файла
function ClassLoadFile(RootObject: TComponent; const FileName: TFileName): boolean;
var
  FileStream: TFileStream;
  MemStream:  TMemoryStream;
begin
  Result     := False;
  FileStream := TFileStream.Create(FileName, 0);
  MemStream  := TMemoryStream.Create;
    try
    ObjectTextToBinary(FileStream, MemStream);
    MemStream.Position := 0;
    MemStream.ReadComponent(RootObject);
    Result := True;
    finally
    MemStream.Free;
    FileStream.Free;
    end;
end;

//Добавить новую колонку с указанными параметрами
procedure GridAddCol(var grid: TStringGrid; const ParamZnach: array of const);
var
  ColEnd, i: integer;
  nk: boolean = True;
  s:  string;
begin
  //запоминаем кол-во колонок, а потом добавляем колонку
  ColEnd := grid.ColCount;
  grid.Columns.Add;
  for i := Low(ParamZnach) to High(ParamZnach) - 1 do
    begin
    nk := not nk;
    if nk then
      Continue;
    s := toStr(ParamZnach[i]);
    case toStr(ParamZnach[i]) of
      'text', 't', 'T': grid.Columns.Items[ColEnd].Title.Caption :=
          toStr(ParamZnach[i + 1]);
      'w', 'W': grid.Columns.Items[ColEnd].Width := toInt(ParamZnach[i + 1]);
      'Color', 'c', 'C': grid.Columns.Items[ColEnd].Color :=
          StringToColor(toStr(ParamZnach[i + 1]));
      end;
    end;
end;

//удалить строки с повторяющимися значениями в столбце
function gridDelPovtor(var sg: TStringGrid; const col: integer): boolean;
var
  j: integer;
begin
  Result := False;
  for j := sg.RowCount - 1 downto 2 do
    if sg.Cells[1 + 1, j] = sg.Cells[1 + 1, j - 1] then
      begin
      sg.DeleteRow(j);
      Result := True;
      end;
end;


end.
