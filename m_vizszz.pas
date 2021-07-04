 // ************************************************************************
 //                 m_VizSzz.pas
 //                 ------------
// Редактирование *.xml файла ONEPLAN Sazon, если недостаточно лицензий.
// Преобразование выгруженного эксель файла из Визир в xml ONEPLAN Sazon
 // Author:        Андрей А. Кандауров (Andrey A. Kandaurov)
 // Companiya:     Santig
 // e-mail:      san@santig.ru
 // URL:           http://santig.ru
 // License:       zlib
 // Create:      04.01.2021
 // Last update: 2021.07.04

// * Для работы с экселем (*.xlsx;*.xls) используется avemey.com
//   https://avemey.com/zexmlss/index.php?lang=ru
// * Панели ExpandPanels, компонент TMyRollOut от Alexander Roth, Massimo Magnano
 //   http://wiki.lazarus.freepascal.org/TMyRollOut_and_ExpandPanel
 // ************************************************************************

(* Copyright (c) 2021 Santig

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

unit m_VizSzz;

{$mode objfpc}{$H+}{$X+}
//{$ModeSwitch advancedrecords}
{$WARN 5024 off : Parameter "$1" not used}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, StdCtrls, ComCtrls,
  Grids, IniPropStorage, ExtCtrls, Menus, comobj, FileUtil, LazUTF8, Types,
  Laz_AVL_Tree, LazFileUtils, LConvEncoding, ExpandPanels,
  zexlsx, zexmlss, zeSave,
  DOM, xmlutils, xmliconv_windows,
  me_my, m_myFs, me_xml, me_exl, m_type;

type

  { TFVizSzz }

  TFVizSzz = class(TForm)
    btZOZavto: TButton;
    btZOZdel: TButton;
    btZOZzdan: TButton;
    ColorDialog1: TColorDialog;
    f_bAddRow: TButton;
    f_diap: TToggleBox;
    f_frec: TStringGrid;
    f_load: TButton;
    f_Tparam: TStringGrid;
    f_Tnad: TStringGrid;
    f_Topis: TStringGrid;
    f_Tkt: TStringGrid;
    f_Tzdan: TStringGrid;
    f_Tzoz: TStringGrid;
    f_xmlOpen: TButton;
    f_xmlSave: TButton;
    GroupBoxEdit: TGroupBox;
    IniPropStorage1: TIniPropStorage;
    license: TMemo;
    license1: TMemo;
    MPOne_Save: TMenuItem;
    MPOne_2: TMenuItem;
    MPOne_AutoCol: TMenuItem;
    MPOne_1: TMenuItem;
    MPDiap_del: TMenuItem;
    MPOne_Del: TMenuItem;
    RollOutLic: TMyRollOut;
    RollOutDoc: TMyRollOut;
    OpenDialog1: TOpenDialog;
    OpenDlgXml: TOpenDialog;
    Panel1: TPanel;
    f_PCTabl: TPageControl;
    PopupM_diap: TPopupMenu;
    PopupM_Tdata: TPopupMenu;
    ProgressBar1: TProgressBar;
    SaveDialog1: TSaveDialog;
    StatusBar: TStatusBar;
    opis:  TTabSheet;
    param: TTabSheet;
    zdan:  TTabSheet;
    toch:  TTabSheet;
    zoz:   TTabSheet;
    nadp:  TTabSheet;
    xl:    TZEXMLSS;
    procedure btZOZavtoClick(Sender: TObject);
    procedure btZOZdelClick(Sender: TObject);
    procedure btZOZzdanClick(Sender: TObject);
    procedure FormActivate(Sender: TObject);
    procedure FormClose(Sender: TObject; var CloseAction: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure f_bAddRowClick(Sender: TObject);
    procedure f_diapChange(Sender: TObject);
    procedure f_loadClick(Sender: TObject);
    procedure f_TparamButtonClick(Sender: TObject; aCol, aRow: integer);
    procedure f_TparamDrawCell(Sender: TObject; aCol, aRow: integer;
      aRect: TRect; aState: TGridDrawState);
    procedure f_TzdanButtonClick(Sender: TObject; aCol, aRow: integer);
    procedure f_TzdanDrawCell(Sender: TObject; aCol, aRow: integer;
      aRect: TRect; aState: TGridDrawState);
    procedure f_TzozButtonClick(Sender: TObject; aCol, aRow: integer);
    procedure f_TzozDrawCell(Sender: TObject; aCol, aRow: integer;
      aRect: TRect; aState: TGridDrawState);
    procedure f_xmlOpenClick(Sender: TObject);
    procedure f_xmlSaveClick(Sender: TObject);
    procedure MPDiap_delClick(Sender: TObject);
    procedure MPOne_AutoColClick(Sender: TObject);
    procedure MPOne_DelClick(Sender: TObject);
    procedure MPOne_SaveClick(Sender: TObject);
    procedure RollOutLicCollapse(Sender: TObject);
    procedure RollOutLicPreExpand(Sender: TObject);
  private
    //Очистить таблицы
    procedure ClearTable;
    //создать уникальных владелцев РЭС
    procedure exlRazbor(const sh: TZSheet);
    // Добавить всю строку из экселя (новая антенна)
    procedure rowAdd(const sh: TZSheet; nRow: integer);
    //определить принадлежность к технологии
    function technologiya(f, vidRES: string): string;
    //Получить данные из таблицы частот
    function GetDateFrec(const f: string; col: byte): string;
    //ПДУ тип
    function PDU_Type(f: string): string;
    //ПДУ значение
    function PDU_znach(f: string): string;
    function colorOper(vlavelec: string): string;

    //Открыть и обработать эксель файл
    procedure loadExe(nameExel: string);
    //распознать столбцы в экселе
    function colExlSet(const sh: TZSheet): boolean;
    //Сохранить данные в xml Файл
    procedure saveXML(nameXML: string);
    procedure saveVisible;
  protected
    procedure progres;
    procedure progressEnd;
    //Установить начальные значения и отобразить
    procedure progressSet(maxProgress: integer);
  end;


var
  FVizSzz: TFVizSzz;

  stat:     TstatusBarMy;
  //массив с отработанными строками из эксель
  mOk:      array of integer;
  //файл для сохранения диапазонов
  diapFile: string = 'VizSzz_diap.txt';
  // при запуске был передан путь к файлу
  paramFile: string = '';
  paramFileDel: boolean = False;
  //Cохранили данные
  SaveXMLBool: boolean = False;
  _colorSelected: TColor = $00009898;
  _colorFix: TColor = clRed;  //$00B0E4EF;
  vers:     string;
  ver:      TFileVersInit;

implementation

{$R *.lfm}

{ TFVizSzz }

procedure TFVizSzz.FormCreate(Sender: TObject);
begin
  stat := TstatusBarMy.CreateStat(StatusBar);
  stat.razdel := 'VizSzz';
  stat.poleN := 0;
end;

procedure TFVizSzz.FormClose(Sender: TObject; var CloseAction: TCloseAction);
begin
  IniPropStorage1.WriteString('vers', vers);
    try
    f_frec.SaveToFile(diapFile);
    stat.Log('Диапазоны сохранены в файл: ' + diapFile);
    except
    stat.Warning(
      'Диапазоны НЕ сохранены в файл (ошибка доступа): '
      + diapFile);
    end;
end;

procedure TFVizSzz.FormActivate(Sender: TObject);
begin
  ver     := FileVersInit;
  vers    := ver.FileVersion;
  Caption := 'VizSzz ' + vers;
  iPaIni(f_Tparam);
  iZdIni;
  // Разница в не совместимости версий!!!
  if not ver.compare(IniPropStorage1.ReadString('vers', ''), '0.9.5', 3) then
    begin
    myFs.Log.Log('Удалён файл диапазонов: ' + diapFile);
    DeleteFileUTF8(diapFile);
    myFs.Log.Log('Удалён файл ini: ' + IniPropStorage1.IniFileName);
    IniPropStorage1.Active := False;
    DeleteFileUTF8(IniPropStorage1.IniFileName);
    IniPropStorage1.Active := True;
    end;

  if FileExists(diapFile) then
      try
      f_frec.LoadFromFile(diapFile);
      stat.Log('Диапазоны загружены из файла: ' +
        diapFile);
      except
      stat.Warning('Диапазоны НЕ загружены из файла: ' +
        diapFile);
      end
  else
    stat.Log('Файла с диапазонами не было: ' + diapFile);
  stat.Log(UTF8encode(#169) + ' Santig. Редактор *.xml ONEPLAN Sazon. 2021 г.');

  f_Topis.SelectedColor  := _colorSelected;
  f_Tparam.SelectedColor := _colorSelected;
  f_Tzdan.SelectedColor  := _colorSelected;
  f_Tkt.SelectedColor    := _colorSelected;
  f_Tzoz.SelectedColor   := _colorSelected;
  f_Tnad.SelectedColor   := _colorSelected;
  f_frec.SelectedColor   := _colorSelected;

  //_ms([Application.ParamCount]);
  if Application.ParamCount > 0 then
    if (Application.Params[1] = '-f') or (Application.Params[1] = '-fd') then
      begin
      paramFile := Application.Params[2];
      if Application.Params[1] = '-fd' then
        paramFileDel := True;
      end
    else
      paramFile := Application.Params[1]
  else
    f_load.SetFocus;

  f_PCTabl.ActivePageIndex := 1;

  // был задан параметр: файл эксель
  paramFile := Trim(paramFile);
  if not paramFile.IsEmpty then
    begin
    loadExe(paramFile);
    f_xmlSaveClick(f_xmlSave);
    if paramFileDel then
      DeleteFileUTF8(paramFile);
    if SaveXMLBool then
      Close;
    //иначе не закрывать, чтобы можно было отредактировать при необходимости
    end;
end;

procedure TFVizSzz.btZOZavtoClick(Sender: TObject);
var
  i: integer;
begin
  f_Tzoz.RowCount := 1;
  if f_Tparam.RowCount < 1 then
    exit;

  for i := 1 to f_Tparam.RowCount - 1 do
    begin
    f_Tzoz.RowCount := f_Tzoz.RowCount + 1;
    f_Tzoz.Cells[0 + 1, i] := '-1';
    f_Tzoz.Cells[1 + 1, i] := f_Tparam.Cells[iPa.h_An.n, i];
    f_Tzoz.Cells[2 + 1, i] :=
      ColorToString(RGBToColor(StrToInt(strGen(2, myCDig)),
      StrToInt(strGen(2, myCDig)), StrToInt(strGen(2, myCDig))));
    f_Tzoz.Cells[3 + 1, i] := 'ЗОЗ';
    end;

  f_Tzoz.RowCount := f_Tzoz.RowCount + 1;
  f_Tzoz.Cells[0 + 1, 1] := '-1';
  f_Tzoz.Cells[1 + 1, 1] := '2';
  f_Tzoz.Cells[2 + 1, 1] := ColorToString(clRed);
  f_Tzoz.Cells[3 + 1, 1] := 'СЗЗ';

  f_Tzoz.SortColRow(True, 1 + 1);
  gridDelPovtor(f_Tzoz, 1 + 1);
end;

procedure TFVizSzz.btZOZdelClick(Sender: TObject);
begin
  //if f_Tzoz.RowCount<1 then
  if f_Tzoz.Row < 1 then
    exit;

  f_Tzoz.DeleteRow(f_Tzoz.Row);
end;

//добавить ЗОЗ над зданием
procedure TFVizSzz.btZOZzdanClick(Sender: TObject);
var
  i: integer;
begin
  for i := 1 to f_Tzdan.RowCount - 1 do
    begin
    f_Tzoz.RowCount := f_Tzoz.RowCount + 1;
    f_Tzoz.Cells[0 + 1, i] := '-1';
    f_Tzoz.Cells[1 + 1, i] := f_Tzdan.Cells[3 + 1, i];
    f_Tzoz.Cells[2 + 1, i] :=
      ColorToString(RGBToColor(StrToInt(strGen(2, myCDig)),
      StrToInt(strGen(2, myCDig)), StrToInt(strGen(2, myCDig))));
    f_Tzoz.Cells[3 + 1, i] := 'ЗОЗ над зданием ' +
      f_Tzdan.Cells[3 + 1, i] + ' м.';
    end;
  gridDelPovtor(f_Tzoz, 1 + 1);
end;

procedure TFVizSzz.ClearTable;
begin
  f_Tparam.RowCount := 1;
  f_Tzdan.RowCount  := 1;
  f_Tkt.RowCount    := 1;
  f_Tzoz.RowCount   := 1;
  f_Tnad.RowCount   := 1;

end;

procedure TFVizSzz.FormDestroy(Sender: TObject);
begin
  stat.Free;
  myFs.Free;
  if Assigned(docXML) then
    docXML.Free;
end;

procedure TFVizSzz.f_bAddRowClick(Sender: TObject);
begin
  case f_PCTabl.ActivePageIndex of
    //0: f_Topis.RowCount  := f_Topis.RowCount + 1;  не добавлять!!!
    1: f_Tparam.RowCount := f_Tparam.RowCount + 1;
    2: f_Tzdan.RowCount  := f_Tzdan.RowCount + 1;
    3: f_Tkt.RowCount    := f_Tkt.RowCount + 1;
    4: f_Tzoz.RowCount   := f_Tzoz.RowCount + 1;
    5: f_Tnad.RowCount   := f_Tnad.RowCount + 1;
    end;
  Application.ProcessMessages;
end;

procedure TFVizSzz.f_diapChange(Sender: TObject);
begin
  f_frec.Visible := not f_frec.Visible;
end;

procedure TFVizSzz.f_loadClick(Sender: TObject);
begin
  f_xmlSave.Visible := False;
  if opendialog1.Execute then
    begin
    Caption := 'VizSzz ' + vers + '   ..\' + ExtractFileName(
      ExtractFileDir(OpenDialog1.FileName)) + '\' + ExtractFileNameOnly(
      OpenDialog1.FileName);
    ClearTable;
    loadExe(OpenDialog1.FileName);
    saveVisible;
    end;
end;

procedure TFVizSzz.f_TparamButtonClick(Sender: TObject; aCol, aRow: integer);
begin
  if (aCol = iPa.color_Se.n) and ColorDialog1.Execute then //Button
    begin
    f_Tparam.Cells[aCol, aRow] := ColorToString(Colordialog1.Color);
    //store cell colour in array
    f_Tparam.Invalidate; //Could also use 'Repaint' te force DrawCell event
    end;
end;

procedure TFVizSzz.f_TparamDrawCell(Sender: TObject; aCol, aRow: integer;
  aRect: TRect; aState: TGridDrawState);
begin
  if (aRow > 0) then         //Use DrawCell to paint rectangle
    //перерисовать цвет в ячейке
    if (ACol = iPa.color_Se.n) then
      begin                    //Get colour from array
      f_Tparam.canvas.Brush.Color :=
        StringToColorDef(f_Tparam.Cells[aCol, aRow], clMaroon);// '8388608');
      f_Tparam.canvas.FillRect(aRect);   //Paint Cell
      //f_Tparam.Cells[aCol, aRow] := ColorToString(clMaroon); // не ставить - не закрывается
      end;
end;

procedure TFVizSzz.f_TzdanButtonClick(Sender: TObject; aCol, aRow: integer);
begin
  if (aCol = iZd.color.n) and ColorDialog1.Execute then //Button
    begin
    f_Tzoz.Cells[aCol, aRow] := ColorToString(ColorDialog1.Color);
    //store cell colour in array
    f_Tzoz.Invalidate; //Could also use 'Repaint' te force DrawCell event
    end;
end;

procedure TFVizSzz.f_TzdanDrawCell(Sender: TObject; aCol, aRow: integer;
  aRect: TRect; aState: TGridDrawState);
begin
  if (aRow > 0) then         //Use DrawCell to paint rectangle
    //перерисовать цвет в ячейке
    if (ACol = iZd.color.n) then
      begin                    //Get colour from array
      f_Tzdan.canvas.Brush.Color :=
        StringToColorDef(f_Tzdan.Cells[aCol, aRow], clMaroon);// '8388608');
      f_Tzdan.canvas.FillRect(aRect);   //Paint Cell
      //f_Tzdan.Cells[aCol, aRow] := ColorToString(clMaroon); // не ставить - не закрывается приложение
      end;
end;

procedure TFVizSzz.f_TzozButtonClick(Sender: TObject; aCol, aRow: integer);
begin
  if (aCol = 2 + 1) and ColorDialog1.Execute then //Button
    begin
    f_Tzoz.Cells[aCol, aRow] := ColorToString(ColorDialog1.Color);
    //store cell colour in array
    f_Tzoz.Invalidate; //Could also use 'Repaint' te force DrawCell event
    end;
end;

procedure TFVizSzz.f_TzozDrawCell(Sender: TObject; aCol, aRow: integer;
  aRect: TRect; aState: TGridDrawState);
begin
  if (aRow > 0) then         //Use DrawCell to paint rectangle
    //перерисовать цвет в ячейке
    if (ACol = 2 + 1) then
      begin                    //Get colour from array
      f_Tzoz.canvas.Brush.Color :=
        StringToColorDef(f_Tzoz.Cells[aCol, aRow], clBlue);// '8388608');
      f_Tzoz.canvas.FillRect(aRect);   //Paint Cell
      end;
end;

procedure TFVizSzz.f_xmlOpenClick(Sender: TObject);
var
  fileXml, tmpS: string;
  nod, nodTmp, nodTmp2, nodTmp3: TDOMNode;
  r, i:    integer;
  nodRoot: TDOMElement;
  tmpI:    PtrInt;
begin
  fileXml := xmlOpensFile(Sender, 'ONEPLAN Sazon|*.xml|все файлы|*.*');
  if fileXml.IsEmpty then
    begin
    stat.Error('Файл не открывается: ' + fileXml);
    exit;
    end;

  OpenDialog1.FileName := fileXml;
  Caption := 'VizSzz ' + vers + '   ..\' + ExtractFileNameOnly(
    ExtractFileDir(OpenDlgXml.FileName)) + '\' + ExtractFileNameOnly(fileXml);
  ClearTable;

  if not Assigned(docXML) then
    begin
    stat.Error('Файл пуст: ' + fileXml, stat.poleN, 'docXML = nil');
    exit;
    end;

  progressSet(f_PCTabl.PageCount);
  stat.Log('0. Создание переменной xml');

  if not Assigned(docXML) then
    docXML := TXMLDocument.Create;
  nodRoot  := docXML.DocumentElement; // = <Document>
  nod      := nodRoot.FindNode('SITES').FindNode('site');

  stat.Log('1. Чтение описания ПРТО');
  //перебираем все атрибуты и записываем в таблицу
  for i := 1 to f_Topis.RowCount - 1 do
    f_Topis.Cells[3, i] := atrRead(nod, f_Topis.Cells[1, i]);
  progres;

  stat.Log('2. Чтение параметров', stat.poleN, 'f_Tparam');
  nodTmp := me_xml.findZnachAtr(nod, 'id', f_Tparam.RowCount.ToString);
  while assigned(nodTmp) do
    begin
    r := f_Tparam.RowCount;
    f_Tparam.RowCount := r + 1;
    //_ms(['while',nodTmp.NodeName, atrRead(nodTmp,iPa.vkl_An.x,'--')]);

    f_Tparam.Cells[iPa.tech_Se.n, r]     := atrRead(nodTmp, iPa.tech_Se.x, '');
    f_Tparam.Cells[iPa.PDUparam_Se.n, r] :=
      f_Tparam.Columns[iPa.PDUparam_Se.n].PickList.Strings[StrToInt(
      atrRead(nodTmp, iPa.PDUparam_Se.x, '1'))];
    f_Tparam.Cells[iPa.PDUznach_Se.n, r] := atrRead(nodTmp, iPa.PDUznach_Se.x, '');
    f_Tparam.Cells[iPa.color_Se.n, r]    := atrRead(nodTmp, iPa.color_Se.x, '');

    nodTmp2 := nodTmp.FindNode('TRX');
    f_Tparam.Cells[iPa.P_Tr.n, r] := atrRead(nodTmp2, iPa.P_Tr.x, '0', '');
    f_Tparam.Cells[iPa.f_Tr.n, r] := atrRead(nodTmp2, iPa.f_Tr.x, '0', '');

    nodTmp2 := nodTmp.FindNode('combiner');
    f_Tparam.Cells[iPa.bPassiv_Co.n, r] := atrRead(nodTmp2, iPa.bPassiv_Co.x, '', '0');
    f_Tparam.Cells[iPa.CombinerType_Co.n, r] :=
      atrRead(nodTmp2, iPa.CombinerType_Co.x, '');
    // кол-во передатчиков           (9)
    f_Tparam.Cells[iPa.kolPrd_An_Co.n, r] := atrRead(nodTmp2, iPa.kolPrd_An_Co.x, '');

    nodTmp2 := nodTmp.FindNode('ANTENNA');
    f_Tparam.Cells[iPa.prd_An.n, r] := atrRead(nodTmp2, iPa.prd_An.x, '1', '0');
    f_Tparam.Cells[iPa.azim_An.n, r] := atrRead(nodTmp2, iPa.azim_An.x, '0', '');
    f_Tparam.Cells[iPa.h_An.n, r] := atrRead(nodTmp2, iPa.h_An.x, '0', '');
    f_Tparam.Cells[iPa.TiltM_An.n, r] := atrRead(nodTmp2, iPa.TiltM_An.x, '0');
    f_Tparam.Cells[iPa.K_An.n, r] := atrRead(nodTmp2, iPa.K_An.x, '0', '');
    f_Tparam.Cells[iPa.bLen_An.n, r] := atrRead(nodTmp2, iPa.bLen_An.x, '', '0');
    f_Tparam.Cells[iPa.bPogon_An.n, r] := atrRead(nodTmp2, iPa.bPogon_An.x, '', '0');
    f_Tparam.Cells[iPa.b_An.n, r] := atrRead(nodTmp2, iPa.b_An.x, '0', '');
    f_Tparam.Cells[iPa.vkl_An.n, r] := atrRead(nodTmp2, iPa.vkl_An.x, '');
    f_Tparam.Cells[iPa.koordN_An.n, r] :=
      atrRead(nodTmp2, iPa.koordN_An.x, atrRead(nod, iPa.koordN_An.x, ''), '');
    f_Tparam.Cells[iPa.koordE_An.n, r] :=
      atrRead(nodTmp2, iPa.koordE_An.x, atrRead(nod, iPa.koordE_An.x, ''), '');
    f_Tparam.Cells[iPa.kolPrd_An_Co.n, r] :=
      atrRead(nodTmp2, iPa.kolPrd_An_Co.x, '1', '0');
    f_Tparam.Cells[iPa.CombinerType_An.n, r] :=
      atrRead(nodTmp2, iPa.CombinerType_An.x, '', '0');
    f_Tparam.Cells[iPa.calcType_An.n, r] := atrRead(nodTmp2, iPa.calcType_An.x, '', '0');
    f_Tparam.Cells[iPa.secZaprAzim_An.n, r] :=
      atrRead(nodTmp2, iPa.secZaprAzim_An.x, '', '0');
    f_Tparam.Cells[iPa.secZaprShir_An.n, r] :=
      atrRead(nodTmp2, iPa.secZaprShir_An.x, '', '0');
    f_Tparam.Cells[iPa.typeApert_An.n, r] :=
      f_Tparam.Columns[iPa.typeApert_An.n].PickList.Strings[StrToInt(
      atrRead(nodTmp, iPa.typeApert_An.x, '0'))];
    f_Tparam.Cells[iPa.KND_An.n, r] := atrRead(nodTmp2, iPa.KND_An.x, '', '0');
    f_Tparam.Cells[iPa.KND_a_An.n, r] := atrRead(nodTmp2, iPa.KND_a_An.x, '', '0');
    f_Tparam.Cells[iPa.KND_D_An.n, r] := atrRead(nodTmp2, iPa.KND_D_An.x, '', '0');
    f_Tparam.Cells[iPa.ugRasGor_An.n, r] := atrRead(nodTmp2, iPa.ugRasGor_An.x, '', '0');
    f_Tparam.Cells[iPa.ugRasVer_An.n, r] := atrRead(nodTmp2, iPa.ugRasVer_An.x, '', '0');
    f_Tparam.Cells[iPa.storApertGor_An.n, r] :=
      atrRead(nodTmp2, iPa.storApertGor_An.x, '', '0');
    f_Tparam.Cells[iPa.storApertVert_An.n, r] :=
      atrRead(nodTmp2, iPa.storApertVert_An.x, '', '0');
    f_Tparam.Cells[iPa.impF_An.n, r] := atrRead(nodTmp2, iPa.impF_An.x, '', '0');
    f_Tparam.Cells[iPa.impT_An.n, r] := atrRead(nodTmp2, iPa.impT_An.x, '', '0');
    f_Tparam.Cells[iPa.impP_An.n, r] := atrRead(nodTmp2, iPa.impP_An.x, '', '0');
    f_Tparam.Cells[iPa.Model_An.n, r] := atrRead(nodTmp2, iPa.Model_An.x, '');
    f_Tparam.Cells[iPa.Model_id_An.n, r] := atrRead(nodTmp2, iPa.Model_id_An.x, '', '0');
    f_Tparam.Cells[iPa.pol_An.n, r] := atrRead(nodTmp2, iPa.pol_An.x, '', '0');
    f_Tparam.Cells[iPa.modu_An.n, r] := atrRead(nodTmp2, iPa.modu_An.x, '');
    tmpS    := atrRead(nodTmp2, iPa.vlad_An.x, '');
    tmpI    := Pos(';', tmpS);
    //не найден koordN_Anразделитель
    if tmpI = 0 then
      f_Tparam.Cells[iPa.vlad_An.n, r] := tmpS
    else
      begin
      f_Tparam.Cells[iPa.BS_No.n, r]   := Copy(tmpS, 1, tmpI - 1);
      f_Tparam.Cells[iPa.vlad_An.n, r] := Copy(tmpS, tmpI + 1, Length(tmpS));
      end;

    nodTmp3 := nodTmp.FindNode('DNA_vert');
    // ДНА верт
    f_Tparam.Cells[iPa.DNAvert_DNA.n, r] := atrRead(nodTmp3, iPa.DNAvert_DNA.x, '');
    //ДНА гор
    nodTmp3 := nodTmp.FindNode('DNA_horz');
    f_Tparam.Cells[iPa.DNAhorz_DNA.n, r] := atrRead(nodTmp3, iPa.DNAhorz_DNA.x, '');

    // пЕРЕХОД К СЛЕДУЮЩЕМУ УЗЛУ
    nodTmp := findZnachAtr(nod, 'id', f_Tparam.RowCount.ToString);
    end;
  progres;

  //--------------------------
  nod := nodRoot.FindNode('SAZON_PARAMS');

  stat.Log('3. Чтение зданий');
  nodTmp := nod.FindNode('MAP_POLYGONS');
  if Assigned(nodTmp) then
    if nodTmp.HasChildNodes then
      begin
      nodTmp2 := nodTmp.FirstChild;
      while assigned(nodTmp2) do
        begin
        r := f_Tzdan.RowCount;
        f_Tzdan.RowCount := r + 1;
        f_Tzdan.Cells[iZd.vkl.n, r] := atrRead(nodTmp2, iZd.vkl.x, '-1');
        f_Tzdan.Cells[iZd.cap.n, r] := atrRead(nodTmp2, iZd.cap.x, '');
        f_Tzdan.Cells[iZd.typ.n, r] :=
          f_Tzdan.Columns[iZd.typ.n - 1].PickList.Strings[StrToInt(
          atrRead(nodTmp2, iZd.typ.x, '1')) - 1];
        f_Tzdan.Cells[iZd.h.n, r] := atrRead(nodTmp2, iZd.h.x, '');
        f_Tzdan.Cells[iZd.color.n, r] := atrRead(nodTmp2, iZd.color.x, '');
        f_Tzdan.Cells[iZd.otrajBool.n, r] := atrRead(nodTmp2, iZd.otrajBool.x, '-1');
        f_Tzdan.Cells[iZd.otrajKoef.n, r] := atrRead(nodTmp2, iZd.otrajKoef.x, '0.5');
        f_Tzdan.Cells[iZd.LossBool.n, r] := atrRead(nodTmp2, iZd.LossBool.x, '0');
        f_Tzdan.Cells[iZd.LossZnach.n, r] := atrRead(nodTmp2, iZd.LossZnach.x, '', '0');
        f_Tzdan.Cells[iZd.Loss_m.n, r] := atrRead(nodTmp2, iZd.Loss_m.x, '', '0');
        f_Tzdan.Cells[iZd.LossP2346.n, r] := atrRead(nodTmp2, iZd.LossP2346.x, '0');
        f_Tzdan.Cells[iZd.koment.n, r] := atrRead(nodTmp2, iZd.koment.x, '');
        if nodTmp2.HasChildNodes then
          begin
          tmpS    := '';
          nodTmp3 := nodTmp2.FirstChild;
          while assigned(nodTmp3) do
            begin
            tmpS    += atrRead(nodTmp3, 'lat', '') + ';';
            tmpS    += atrRead(nodTmp3, 'lon', '') + ';';
            // пЕРЕХОД К СЛЕДУЮЩЕМУ УЗЛУ
            nodTmp3 := nodTmp3.NextSibling;
            end;
          f_Tzdan.Cells[iZd.tochki.n, r] := tmpS;
          end;
        // пЕРЕХОД К СЛЕДУЮЩЕМУ УЗЛУ
        nodTmp2 := nodTmp2.NextSibling;
        end;
      end;
  progres;

  //--------------------------------------
  stat.Log('4. Чтение контрольных точек');
  nodTmp := nod.FindNode('CONTROL_POINTS');
  if Assigned(nodTmp) then
    if nodTmp.HasChildNodes then
      begin
      nodTmp2 := nodTmp.FirstChild;
      while assigned(nodTmp2) do
        begin
        r := f_Tkt.RowCount;
        f_Tkt.RowCount := r + 1;
        f_Tkt.Cells[0 + 1, r] := atrRead(nodTmp2, 'checked', '');
        f_Tkt.Cells[1 + 1, r] := atrRead(nodTmp2, 'name', '');
        f_Tkt.Cells[2 + 1, r] := atrRead(nodTmp2, 'height', '');
        f_Tkt.Cells[7 + 1, r] := atrRead(nodTmp2, 'measure', '');
        f_Tkt.Cells[5 + 1, r] :=
          f_Tkt.Columns[5].PickList.Strings[StrToInt(
          atrRead(nodTmp2, 'm_type', '1')) - 1];
        f_Tkt.Cells[3 + 1, r] := atrRead(nodTmp2, 'lat', '');
        f_Tkt.Cells[4 + 1, r] := atrRead(nodTmp2, 'lon', '');
        f_Tkt.Cells[6 + 1, r] := atrRead(nodTmp2, 'comment', '');

        // пЕРЕХОД К СЛЕДУЮЩЕМУ УЗЛУ
        nodTmp2 := nodTmp2.NextSibling;
        end;
      end;
  progres;

  //--------------------------------------
  stat.Log('5. Чтение высот ЗОЗ');
  nodTmp := nod.FindNode('LEVELS');
  if Assigned(nodTmp) then
    if nodTmp.HasChildNodes then
      begin
      nodTmp2 := nodTmp.FirstChild;
      while assigned(nodTmp2) do
        begin
        r := f_Tzoz.RowCount;
        f_Tzoz.RowCount := r + 1;
        f_Tzoz.Cells[0 + 1, r] := atrRead(nodTmp2, 'checked', '');
        f_Tzoz.Cells[1 + 1, r] := atrRead(nodTmp2, 'height', '');
        f_Tzoz.Cells[2 + 1, r] := atrRead(nodTmp2, 'color', '');
        f_Tzoz.Cells[3 + 1, r] := atrRead(nodTmp2, 'comment', '');

        // пЕРЕХОД К СЛЕДУЮЩЕМУ УЗЛУ
        nodTmp2 := nodTmp2.NextSibling;
        end;
      end;
  progres;

  //-----------------------------------
  stat.Log('6. Чтение надписей');
  nodTmp := nod.FindNode('MAP_CAPTIONS');
  if Assigned(nodTmp) then
    if nodTmp.HasChildNodes then
      begin
      nodTmp2 := nodTmp.FirstChild;
      while assigned(nodTmp2) do
        begin
        r := f_Tnad.RowCount;
        f_Tnad.RowCount := r + 1;
        f_Tnad.Cells[0 + 1, r] := atrRead(nodTmp2, 'checked', '');
        f_Tnad.Cells[1 + 1, r] := atrRead(nodTmp2, 'caption', '');
        f_Tnad.Cells[2 + 1, r] := atrRead(nodTmp2, 'lat', '');
        f_Tnad.Cells[3 + 1, r] := atrRead(nodTmp2, 'lon', '');
        f_Tnad.Cells[4 + 1, r] := atrRead(nodTmp2, 'comment', '');
        // пЕРЕХОД К СЛЕДУЮЩЕМУ УЗЛУ
        nodTmp2 := nodTmp2.NextSibling;
        end;
      end;
  progres;

  stat.Log('Файл прочитан.');
  progressEnd;
  saveVisible;
end;

//распознать столбцы в экселе
function TFVizSzz.colExlSet(const sh: TZSheet): boolean;
var
  i: integer;

  //поиск в строке данных val
  function findExl(val: string): integer;
  var
    k:   integer;
    dat: string;
  begin
    Result := -1;
    for k := 0 to sh.ColCount - 1 do
      begin
      dat := Trim(exlCellDataTry(sh.Cell[k, 0]));
      if val = dat then
        begin
        Result := k;
        Break;
        end;
      end;
  end;

begin
  Result := True;
  i      := findExl('№ П/П');
  if i >= 0 then
    iPa.nnExcel_No.exl := i
  else
    Exit(False);
  i := findExl('Номер канала или БС');
  if i >= 0 then
    iPa.BS_No.exl := i
  else
    Exit(False);
  i := findExl('Владелец РЭС');
  if i >= 0 then
    iPa.vlad_An.exl := i
  else
    Exit(False);
  i := findExl('Вид РЭС');
  if i >= 0 then
    iPa.vidRES_No.exl := i
  else
    Exit(False);
  i := findExl('Частота ПРД, МГц');
  if i >= 0 then
    iPa.f_Tr.exl := i
  else
    Exit(False);
  i := findExl('Мощность, Вт');
  if i >= 0 then
    iPa.P_Tr.exl := i
  else
    Exit(False);
  i := findExl('Азимут');
  if i >= 0 then
    iPa.azim_An.exl := i
  else
    Exit(False);
  i := findExl('Высота подвеса антенны, м');
  if i >= 0 then
    iPa.h_An.exl := i
  else
    Exit(False);
  i := findExl('Номер св-ва о регистрации РЭС');
  if i >= 0 then
    iPa.svid_No.exl := i
  else
    Exit(False);
  i := findExl(
    'Дата окончания действия св-ва регистрации РЭС');
  if i >= 0 then
    iPa.svidDO_No.exl := i
  else
    Exit(False);
  i := findExl('Модель');
  if i >= 0 then
    iPa.prd_An.exl := i
  else
    Exit(False);
  i := findExl('Место установки');
  if i >= 0 then
    iPa.adr_Si.exl := i
  else
    Exit(False);
end;

//Открыть и обработать эксель файл
procedure TFVizSzz.loadExe(nameExel: string);
var
  i, n: integer;
begin
  f_Tparam.Clean([gzNormal, gzFixedRows]);
  f_Tparam.RowCount := 1;
  if not FileExistsUTF8(nameExel) then
    begin
    stat.Error('Файл не найден: ' + nameExel);
    exit;
    end;

  stat.Log('открытие excel для обработки: ' +
    ExtractFileNameOnly(nameExel),
    stat.poleN, 'Выбран файл: ' + nameExel);

    try
    xl := TZEXMLSS.Create(nil);
    if xl = nil then
      begin
      stat.Error('Файл не прочитался: ' + nameExel);
      exit;
      end;
    i := ReadXLSX(xl, nameExel);
    if i <> 0 then
      begin
      stat.Error(
        'При чтении документа возникла ошибка! Код: '
        +
        i.ToString);
      if i = 2 then
        begin
        stat.Error(
          'Для продолжения необходимо закрыть документ!');
        MessageDlg('Ошибка!',
          'Для продолжения необходимо закрыть документ: ',
          mtError, [mbOK], 0);
        end
      else
        MessageDlg('Ошибка!',
          'При чтении документа возникла ошибка: ' +
          i.ToString, mtError, [mbOK], 0);
      exit;
      end;

    if xl.Sheets.Count < 1 then
      begin
      stat.Warning('Листов в документе не обнаружено!');
      MessageDlg('Предупреждение!',
        'Листов в документе не обнаружено!',
        mtWarning, [mbOK], 0);
      exit;
      end;

    //Распознание столбцов в экселе (именно здесь, до stat.Log)
    if not colExlSet(xl.Sheets[0]) then
      begin
      stat.Error('Не все столбцы найдены.');
      MessageDlg('Ошибка!',
        'Наверное это другой файл или включена группировка'
        + nL +
        'Oтключите группировку и сформируйте Excel заново',
        mtError, [mbOK], 0);
      exit;
      end;

    stat.Log('Открыт лист: ' + xl.Sheets[0].Title);
    progressSet(xl.Sheets[0].RowCount - 2);
    stat.log('Найдено позиций: ' + ProgressBar1.Max.ToString);
    Application.ProcessMessages;

    //--------------------------
    if Assigned(xl) then
      exlRazbor(xl.Sheets[0]);
    //--------------------------

    finally
    FreeAndNil(xl);
    end;

  stat.log('Обработано РЭС: ' + ProgressBar1.Max.ToString +
    '->' + (f_Tparam.RowCount - 1).ToString);

  //СОРТИРОВКА СТРОК по владельцу
  f_Tparam.SortColRow(True, iPa.vlad_An.n);
  ////сортировать по технологии
  //n := 1;
  //for i := 1 to f_Tparam.RowCount - 2 do
  //  if (f_Tparam.Cells[iPa.azim_An.n, i] <> f_Tparam.Cells[iPa.azim_An.n, i + 1]) then
  //    begin
  //    f_Tparam.SortColRow(True, iPa.tech_Se.n, n, i);
  //    n := i + 1;
  //    end;
  //f_Tparam.SortColRow(True, iPa.tech_Se.n, n, i + 1);
  //сортировать по виду РЭС
  n := 1;
  for i := 1 to f_Tparam.RowCount - 2 do
    if (f_Tparam.Cells[iPa.vlad_An.n, i] <> f_Tparam.Cells[iPa.vlad_An.n, i + 1]) then
      begin
      f_Tparam.SortColRow(True, iPa.vidRES_No.n, n, i);
      n := i + 1;
      end;
  f_Tparam.SortColRow(True, iPa.vidRES_No.n, n, i + 1);

  //сортировать по азимуту
  n := 1;
  for i := 1 to f_Tparam.RowCount - 2 do
    if (f_Tparam.Cells[iPa.vidRES_No.n, i] <>
      f_Tparam.Cells[iPa.vidRES_No.n, i + 1]) then
      begin
      f_Tparam.SortColRow(True, iPa.azim_An.n, n, i);
      n := i + 1;
      end;
  f_Tparam.SortColRow(True, iPa.azim_An.n, n, i + 1);
  //-------------------------

  saveVisible;
end;

procedure TFVizSzz.exlRazbor(const sh: TZSheet);
var
  n, i_vlad, nObr, i: integer;
  kolEx: integer = 0;

  //если в массиве mOk найдено число id - True
  function idFind(id: integer): boolean;
  var
    j: integer;
  begin
    Result := False;
    for j := 1 to Length(mOk) - 1 do
      if id = mOk[j] then
        begin
        Result := True;
        break;
        end;
  end;

  //добавленна новая строка
  function strNext: boolean;
  var
    i: integer;
  begin
    Result := False;
    for i := 1 to kolEx - 1 do
      begin
      Result := not idFind(i);
      if Result then
        begin
        rowAdd(sh, i);
        break;
        end;
      end;
  end;

begin
  //убрать заголовок из экселя
  SetLength(mOk, 1);
  mOk[0] := 0;
  progres;

  //сколько строк всего в файле
  //for n := 0 to sh.RowCount - 1 do
  //  if not exlCellDataTry(sh.Cell[1, n]).IsEmpty then
  //    Inc(kolEx);
  kolEx := sh.RowCount - 2;

  while Length(mOk) < kolEx do
    begin
    progres;
    if not strNext then
      Continue;

    nObr := mOk[High(mOk)];
    // vlad: поиск по владельцам
    for i_vlad := 1 to kolEx - 1 do
      if idFind(i_vlad) then
        Continue
      else
      //1. vlad:
        if Trim(exlCellDataTry(sh.Cell[iPa.vlad_An.exl, i_vlad])) =
          Trim(exlCellDataTry(sh.Cell[iPa.vlad_An.exl, nObr])) then
          //2. vid:
          if Trim(exlCellDataTry(sh.Cell[iPa.vidRES_No.exl, i_vlad])) =
            Trim(exlCellDataTry(sh.Cell[iPa.vidRES_No.exl, nObr])) then
            //3. tech (frec):
            if technologiya(exlCellDataTry(sh.Cell[iPa.f_Tr.exl, i_vlad]),
              exlCellDataTry(sh.Cell[iPa.vidRES_No.exl, i_vlad])) =
              f_Tparam.Cells[iPa.tech_Se.n, f_Tparam.RowCount - 1] then
              // P:
              if Trim(exlCellDataTry(sh.Cell[iPa.P_Tr.exl, i_vlad])) =
                Trim(exlCellDataTry(sh.Cell[iPa.P_Tr.exl, nObr])) then
                // azim:
                if Trim(exlCellDataTry(sh.Cell[iPa.azim_An.exl, i_vlad])) =
                  Trim(exlCellDataTry(sh.Cell[iPa.azim_An.exl, nObr])) then
                  // h:
                  if Trim(exlCellDataTry(sh.Cell[iPa.h_An.exl, i_vlad])) =
                    Trim(exlCellDataTry(sh.Cell[iPa.h_An.exl, nObr])) then
                    // sv:
                    if Trim(exlCellDataTry(sh.Cell[iPa.svid_No.exl, i_vlad])) =
                      Trim(exlCellDataTry(sh.Cell[iPa.svid_No.exl, nObr])) then
                      begin
                      //Увеличить кол-во передатчиков
                      progres;
                      n := StrToInt(f_Tparam.Cells[iPa.kolPrd_An_Co.n,
                        f_Tparam.RowCount - 1]) + 1;
                      f_Tparam.Cells[iPa.kolPrd_An_Co.n, f_Tparam.RowCount - 1] :=
                        n.ToString;
                      SetLength(mOk, Length(mOk) + 1);
                      mOk[High(mOk)] := i_vlad;

                      f_Tparam.Cells[iPa.nnExcel_No.n, f_Tparam.RowCount - 1] :=
                        f_Tparam.Cells[iPa.nnExcel_No.n, f_Tparam.RowCount - 1] +
                        ', ' + exlCellDataTry(sh.Cell[iPa.nnExcel_No.exl, i_vlad]);
                      end;
    end;

  for i := 1 to f_Tparam.RowCount - 1 do
    f_Tparam.Cells[0, i] := i.ToString;
end;

// Добавить всю строку из экселя (новая антенна)
procedure TFVizSzz.rowAdd(const sh: TZSheet; nRow: integer);
var
  n: integer;
  Slope, Gain, DNA_vert, DNA_horz, modul: string;
  vlad, R_KND, R_Diam, R_rask: string;
  antModel: string = '';

  procedure dopRasch(f, vid: string);
  begin
    //'беспроводного доступа', Result := 'Wi-Fi'
    case technologiya(f, vid) of
      'GSM', 'UMTS', 'LTE':
        begin
        Slope    := '-1';
        antModel := 'Huawei ASI4518R4';
        //поиск частоты по убыванию
        if Pos('2600', GetDateFrec(f_Tparam.Cells[iPa.f_Tr.n, n], 1)) > 0 then
          begin
          Gain     := '17.7';
          DNA_vert :=
            '0=0.08;1=0.08;2=0.87;3=2.6;4=5.34;5=9.48;6=16.05;7=16.96;8=13.71;9=13.03;10=13.93;11=17.01;12=25.31;13=26.21;14=19.93;15=18.66;16=20.76;17=28.55;18=28.4;19=20.98;20=17.93;21=17.38;22=19.32;23=23.75;24=31.57;25=29.91;26=23.68;27=22.5;28=24.52;29=28.16;30=27.38;31=23.1;32=20.8;33=19.91;34=19.77;35=21.14;36=24.61;37=28.44;38=32.14;39=38.99;40=37.77;41=42.3;42=31.18;43=26.29;44=24.69;45=23.09;46=21.15;47=20.46;48=21.27;49=22.68;50=24.28;51=27.65;52=33.68;53=32.56;54=30.32;55=29.24;56=26.72;57=23.6;58=21.51;59=20.69;60=20.57;61=20.7;62=21.05;63=21.4;64=21.26;65=20.76;66=20.84;67=22.1;68=24.16;69=25.62;70=25.64;71=25.54;72=26.88;73=29.84;74=32.18;75=31.89;76=30.62;77=31;78=33.66;79=33.65;80=32.43;81=34.26;82=38.28;83=41.21;84=41.89;85=44.15;86=43.6;87=36.03;88=32.9;89=32.53;90=34.24;91=38.79;92=46.58;93=40.84;94=36.55;95=35.02;96=36.07;97=39.72;98=44.88;99=47.06;100=46.19;101=47.13;102=49.12;103=48.81;104=43.77;105=40.54;106=41.1;107=43.73;108=43.41;109=39.83;110=36.56;111=36.34;112=39.43;113=42.91;114=44.68;115=40.68;116=36.53;117=36.09;118=37.66;119=37.78;120=34.91;121=32.4;122=32.47;123=35.91;124=41.6;125=40.14;126=38.25;127=38.63;128=41.63;129=42.8;130=41.6;131=45.77;132=48.37;133=41.82;134=40.53;135=40.82;136=40.57;137=39.48;138=38.75;139=40.3;140=44.94;141=43.93;142=42.42;143=47.86;144=48.98;145=41.86;146=41.5;147=45.53;148=61.25;149=46.14;150=39.69;151=37.48;152=38.46;153=42.69;154=43.77;155=41.88;156=43.94;157=43.08;158=40.23;159=41.4;160=46.44;161=44.71;162=40.25;163=38.91;164=40.2;165=41.74;166=41.49;167=43.82;168=51.74;169=49.24;170=48.41;171=45.29;172=40.52;173=39.48;174=40.06;175=38.54;176=37.47;177=36.62;178=33.76;179=31.69;180=31.22;181=31.68;182=32.85;183=35.08;184=37.74;185=39.13;186=39.75;187=41.69;188=45.87;189=50.01;190=61.37;191=50.43;192=44.67;193=43.87;194=43.86;195=44.99;196=49.58;197=53.49;198=49.31;199=44.38;200=41.4;201=39.13;202=37.55;203=37.64;204=38.68;205=38.53;206=36.52;207=35;208=35.97;209=37.61;210=35.76;211=36.12;212=40.57;213=40.37;214=39.43;215=42.76;216=42.15;217=40.51;218=40.63;219=42.24;220=45.4;221=41.09;222=40.05;223=43.02;224=39.87;225=38.11;226=41.57;227=43.38;228=38.99;229=38.85;230=42.09;231=46.48;232=44.18;233=40.02;234=37.34;235=35.59;236=35;237=35.87;238=36.99;239=37.49;240=38.95;241=39.79;242=38.06;243=36.26;244=34.22;245=32.45;246=31.28;247=30.17;248=29.24;249=29.22;250=30.19;251=31.07;252=31.23;253=31.88;254=33.08;255=32.89;256=31.53;257=30.39;258=29.37;259=28.25;260=27.37;261=26.74;262=26.29;263=26.21;264=26.82;265=28.18;266=29.07;267=27.66;268=25.9;269=25.27;270=25.69;271=26.79;272=28.3;273=29.97;274=31.22;275=31.51;276=30.81;277=29.37;278=28.22;279=27.46;280=27.05;281=27.53;282=27.46;283=25.96;284=25.27;285=25.45;286=25.93;287=26.02;288=24.1;289=22.36;290=22.31;291=23.21;292=23.35;293=21.94;294=20.68;295=21.06;296=23.19;297=25.47;298=25.39;299=24.22;300=24.59;301=27.4;302=30.74;303=30.72;304=27.69;305=25.28;306=25.77;307=29.01;308=31.7;309=30.17;310=26.83;311=24.2;312=23.04;313=23.34;314=23.89;315=23.46;316=24.69;317=30.85;318=33.51;319=30.41;320=36.07;321=50.15;322=52.35;323=34.34;324=27.36;325=25.77;326=27.89;327=37.43;328=33.26;329=28.67;330=32.99;331=32.5;332=24.65;333=22.02;334=20.92;335=19.67;336=18.24;337=17.49;338=17.1;339=16.22;340=15.9;341=17.37;342=21.04;343=28.2;344=42.51;345=29.33;346=28.01;347=25.78;348=23.17;349=22.66;350=22.31;351=21.38;352=20.65;353=18.49;354=14.22;355=10.08;356=6.75;357=4.05;358=2.01;359=0.72';
          DNA_horz :=
            '0=0.71;1=0.63;2=0.54;3=0.45;4=0.35;5=0.26;6=0.18;7=0.11;8=0.05;9=0.01;10=0;11=0;12=0.03;13=0.07;14=0.14;15=0.22;16=0.32;17=0.43;18=0.56;19=0.68;20=0.81;21=0.94;22=1.07;23=1.19;24=1.31;25=1.43;26=1.53;27=1.64;28=1.75;29=1.85;30=1.97;31=2.09;32=2.22;33=2.37;34=2.53;35=2.71;36=2.91;37=3.13;38=3.38;39=3.65;40=3.94;41=4.26;42=4.59;43=4.95;44=5.33;45=5.73;46=6.15;47=6.59;48=7.05;49=7.51;50=8;51=8.49;52=9;53=9.51;54=10.04;55=10.57;56=11.1;57=11.64;58=12.17;59=12.7;60=13.24;61=13.78;62=14.31;63=14.84;64=15.38;65=15.92;66=16.46;67=17.01;68=17.55;69=18.11;70=18.68;71=19.26;72=19.88;73=20.52;74=21.21;75=21.96;76=22.77;77=23.66;78=24.64;79=25.71;80=26.87;81=28.08;82=29.28;83=30.34;84=31.09;85=31.39;86=31.22;87=30.71;88=30.01;89=29.26;90=28.53;91=27.85;92=27.23;93=26.69;94=26.22;95=25.84;96=25.53;97=25.3;98=25.14;99=25.05;100=25.01;101=24.99;102=24.99;103=24.96;104=24.9;105=24.79;106=24.64;107=24.46;108=24.26;109=24.09;110=23.95;111=23.88;112=23.88;113=23.97;114=24.13;115=24.36;116=24.65;117=24.96;118=25.25;119=25.5;120=25.65;121=25.71;122=25.67;123=25.56;124=25.41;125=25.26;126=25.14;127=25.06;128=25.04;129=25.09;130=25.19;131=25.34;132=25.55;133=25.8;134=26.08;135=26.4;136=26.75;137=27.13;138=27.53;139=27.94;140=28.35;141=28.75;142=29.11;143=29.41;144=29.63;145=29.78;146=29.86;147=29.88;148=29.86;149=29.81;150=29.74;151=29.64;152=29.5;153=29.31;154=29.09;155=28.84;156=28.58;157=28.36;158=28.19;159=28.11;160=28.15;161=28.32;162=28.64;163=29.12;164=29.74;165=30.47;166=31.23;167=31.88;168=32.22;169=32.14;170=31.67;171=31;172=30.3;173=29.68;174=29.23;175=28.96;176=28.89;177=29.03;178=29.39;179=29.97;180=30.77;181=31.78;182=32.95;183=34.17;184=35.16;185=35.56;186=35.21;187=34.35;188=33.34;189=32.39;190=31.61;191=31.02;192=30.62;193=30.4;194=30.36;195=30.49;196=30.8;197=31.3;198=31.98;199=32.89;200=34.04;201=35.49;202=37.3;203=39.54;204=42.24;205=44.86;206=45.51;207=43.75;208=41.58;209=39.81;210=38.5;211=37.58;212=36.95;213=36.59;214=36.43;215=36.46;216=36.63;217=36.93;218=37.29;219=37.68;220=38;221=38.17;222=38.15;223=37.91;224=37.51;225=37.04;226=36.56;227=36.15;228=35.84;229=35.64;230=35.56;231=35.55;232=35.56;233=35.5;234=35.29;235=34.85;236=34.22;237=33.46;238=32.66;239=31.91;240=31.28;241=30.81;242=30.51;243=30.42;244=30.56;245=30.96;246=31.64;247=32.65;248=34.08;249=36;250=38.42;251=40.66;252=40.61;253=38.27;254=35.76;255=33.74;256=32.21;257=31.08;258=30.27;259=29.74;260=29.43;261=29.32;262=29.37;263=29.54;264=29.79;265=30.05;266=30.27;267=30.38;268=30.34;269=30.16;270=29.86;271=29.5;272=29.13;273=28.78;274=28.48;275=28.24;276=28.06;277=27.95;278=27.87;279=27.84;280=27.82;281=27.8;282=27.76;283=27.69;284=27.55;285=27.32;286=26.98;287=26.52;288=25.93;289=25.22;290=24.42;291=23.57;292=22.68;293=21.8;294=20.95;295=20.12;296=19.35;297=18.62;298=17.94;299=17.3;300=16.69;301=16.12;302=15.57;303=15.04;304=14.52;305=14;306=13.49;307=12.98;308=12.47;309=11.98;310=11.5;311=11.03;312=10.59;313=10.16;314=9.75;315=9.37;316=9.01;317=8.66;318=8.33;319=8.02;320=7.72;321=7.43;322=7.14;323=6.86;324=6.58;325=6.31;326=6.04;327=5.76;328=5.49;329=5.21;330=4.92;331=4.64;332=4.35;333=4.06;334=3.77;335=3.48;336=3.2;337=2.92;338=2.66;339=2.4;340=2.17;341=1.95;342=1.76;343=1.58;344=1.43;345=1.3;346=1.2;347=1.11;348=1.05;349=1.01;350=0.98;351=0.96;352=0.95;353=0.94;354=0.94;355=0.93;356=0.91;357=0.88;358=0.84;359=0.78';
          modul    := '64QPSK';
          end
        else if Pos('2100', GetDateFrec(f_Tparam.Cells[iPa.f_Tr.n, n], 1)) > 0 then
            begin
            Gain     := '17.5';
            DNA_vert :=
              '0=0.06;1=0.07;2=0.6;3=1.69;4=3.55;5=6.62;6=11.68;7=19.98;8=18.38;9=13.73;10=11.84;11=11.8;12=13.6;13=17.86;14=25.87;15=21.34;16=16.22;17=14.03;18=13.59;19=14.46;20=16.41;21=19.51;22=23.49;23=24.31;24=21.92;25=20.52;26=20.08;27=20.01;28=20.22;29=20.95;30=22.38;31=24.34;32=25.8;33=25.47;34=24.09;35=22.42;36=20.5;37=18.62;38=17.15;39=16.23;40=15.88;41=16.04;42=16.67;43=17.7;44=19.22;45=21.43;46=24.62;47=29.08;48=35.31;49=46.9;50=42.86;51=37.5;52=37.17;53=40.92;54=48.34;55=41.92;56=38.16;57=36.69;58=36.01;59=35.27;60=34.13;61=32.85;62=31.88;63=31.53;64=31.98;65=33.29;66=35.3;67=37.63;68=39.89;69=42.18;70=44.67;71=45.25;72=42.79;73=40.64;74=39.56;75=39.07;76=38.69;77=38.16;78=37.69;79=37.66;80=38.08;81=38.52;82=39.07;83=40.69;84=44.51;85=52.27;86=63.31;87=61.85;88=66.35;89=55.4;90=47.88;91=42.78;92=39.82;93=38.63;94=38.81;95=39.97;96=42.02;97=45.35;98=50.45;99=54.3;100=51.19;101=48.07;102=46.64;103=46.21;104=45.31;105=44.09;106=43.69;107=44.05;108=44.02;109=42.67;110=40.64;111=39.05;112=38.44;113=38.86;114=40.15;115=41.91;116=43.34;117=43.93;118=44.37;119=44.71;120=44.03;121=42.97;122=42.39;123=41.66;124=40.05;125=38.31;126=37.35;127=37.49;128=38.65;129=40.48;130=42.3;131=42.41;132=40.61;133=39.25;134=39.29;135=40.62;136=42.73;137=44.11;138=43.05;139=41.69;140=42.04;141=44.03;142=44.46;143=42.55;144=41.73;145=42.35;146=43.59;147=44.95;148=46.34;149=47.24;150=47.84;151=46.5;152=42.04;153=38.68;154=37.46;155=38.44;156=41.62;157=44.18;158=41.18;159=38.45;160=37.41;161=37.91;162=39.74;163=41.72;164=41.17;165=38.8;166=36.93;167=36.37;168=37.06;169=37.8;170=37.17;171=36.31;172=36.42;173=37.75;174=40.59;175=46.15;176=59.88;177=56.34;178=55.67;179=54.83;180=49.77;181=48.38;182=49.86;183=52.12;184=52.25;185=48.87;186=44.17;187=40.93;188=39.57;189=39.79;190=40.71;191=41.29;192=42.39;193=45.59;194=47.71;195=44.64;196=43.86;197=46.55;198=52.88;199=51.69;200=47.62;201=46.88;202=50.07;203=53.86;204=46.19;205=43.25;206=43.88;207=47.47;208=52.37;209=51.06;210=44.53;211=39.72;212=37.26;213=36.78;214=37.8;215=38.89;216=38.43;217=37.92;218=39.14;219=43.17;220=52.59;221=54.26;222=49.45;223=46.27;224=43.35;225=42.49;226=43.9;227=45.77;228=44.53;229=42.65;230=41.61;231=41.19;232=41.62;233=43.27;234=45.44;235=45.43;236=44.03;237=43.74;238=45.28;239=47.65;240=44.27;241=39.22;242=35.76;243=33.56;244=32.31;245=31.9;246=32.37;247=33.72;248=35.69;249=37.48;250=38.82;251=40.98;252=46.32;253=66.69;254=45.42;255=41.71;256=41.96;257=43.53;258=40;259=36;260=33.95;261=33.51;262=34.34;263=36.02;264=37.6;265=38.07;266=37.46;267=36.2;268=34.99;269=34.48;270=34.84;271=35.85;272=37.05;273=37.78;274=37.63;275=36.67;276=35.48;277=34.66;278=34.46;279=34.66;280=34.61;281=33.75;282=32.19;283=30.55;284=29.25;285=28.41;286=27.87;287=27.36;288=26.65;289=25.78;290=25;291=24.48;292=24.18;293=23.94;294=23.67;295=23.57;296=23.96;297=25.08;298=26.89;299=28.81;300=29.92;301=30.57;302=31.19;303=30.18;304=27.55;305=25.38;306=24.22;307=23.82;308=23.67;309=23.25;310=22.32;311=21.06;312=19.83;313=18.89;314=18.38;315=18.37;316=18.92;317=20.06;318=21.71;319=23.64;320=25.72;321=27.79;322=27.96;323=25.52;324=23.39;325=22.54;326=22.81;327=23.66;328=24.06;329=23.09;330=21.27;331=19.6;332=18.57;333=18.31;334=18.91;335=20.68;336=24.45;337=32.16;338=32.12;339=28.15;340=29.75;341=40.85;342=30.12;343=23.98;344=21.08;345=19.44;346=18.72;347=19.07;348=20.52;349=22.63;350=24.41;351=25.08;352=23.84;353=19.18;354=13.78;355=9.42;356=6.12;357=3.63;358=1.81;359=0.62';
            DNA_horz :=
              '0=0;1=0;2=0.01;3=0.02;4=0.03;5=0.06;6=0.08;7=0.12;8=0.16;9=0.22;10=0.28;11=0.35;12=0.43;13=0.53;14=0.63;15=0.75;16=0.88;17=1.01;18=1.16;19=1.32;20=1.49;21=1.67;22=1.86;23=2.05;24=2.25;25=2.46;26=2.67;27=2.88;28=3.1;29=3.32;30=3.55;31=3.78;32=4.01;33=4.25;34=4.49;35=4.73;36=4.98;37=5.22;38=5.48;39=5.74;40=5.99;41=6.26;42=6.52;43=6.79;44=7.05;45=7.32;46=7.59;47=7.85;48=8.11;49=8.36;50=8.61;51=8.85;52=9.08;53=9.3;54=9.51;55=9.71;56=9.9;57=10.08;58=10.25;59=10.42;60=10.58;61=10.75;62=10.91;63=11.09;64=11.27;65=11.47;66=11.69;67=11.93;68=12.19;69=12.48;70=12.8;71=13.14;72=13.52;73=13.93;74=14.36;75=14.82;76=15.3;77=15.8;78=16.3;79=16.81;80=17.32;81=17.81;82=18.3;83=18.77;84=19.24;85=19.69;86=20.14;87=20.59;88=21.05;89=21.52;90=22.01;91=22.53;92=23.07;93=23.63;94=24.22;95=24.83;96=25.46;97=26.11;98=26.79;99=27.49;100=28.23;101=29.01;102=29.84;103=30.75;104=31.75;105=32.83;106=34.02;107=35.3;108=36.61;109=37.88;110=38.97;111=39.73;112=40.08;113=40.13;114=40;115=39.85;116=39.75;117=39.74;118=39.78;119=39.82;120=39.79;121=39.64;122=39.36;123=38.96;124=38.52;125=38.1;126=37.77;127=37.57;128=37.51;129=37.61;130=37.85;131=38.17;132=38.47;133=38.56;134=38.28;135=37.58;136=36.55;137=35.4;138=34.25;139=33.21;140=32.31;141=31.57;142=31;143=30.59;144=30.34;145=30.26;146=30.34;147=30.6;148=31.03;149=31.64;150=32.47;151=33.52;152=34.82;153=36.41;154=38.26;155=40.23;156=41.8;157=42.11;158=41.13;159=39.66;160=38.29;161=37.17;162=36.32;163=35.7;164=35.28;165=35.04;166=34.95;167=35;168=35.19;169=35.5;170=35.95;171=36.55;172=37.34;173=38.35;174=39.67;175=41.41;176=43.8;177=47.3;178=52.8;179=55.04;180=48.95;181=44.76;182=42.01;183=40.09;184=38.72;185=37.78;186=37.19;187=36.91;188=36.94;189=37.27;190=37.92;191=38.9;192=40.16;193=41.46;194=42.07;195=41.25;196=39.38;197=37.32;198=35.47;199=33.9;200=32.61;201=31.57;202=30.74;203=30.1;204=29.63;205=29.31;206=29.14;207=29.09;208=29.15;209=29.32;210=29.58;211=29.91;212=30.28;213=30.69;214=31.09;215=31.47;216=31.81;217=32.08;218=32.29;219=32.44;220=32.52;221=32.53;222=32.46;223=32.31;224=32.06;225=31.7;226=31.23;227=30.69;228=30.09;229=29.46;230=28.85;231=28.26;232=27.71;233=27.23;234=26.8;235=26.44;236=26.13;237=25.88;238=25.68;239=25.52;240=25.37;241=25.25;242=25.12;243=24.99;244=24.84;245=24.67;246=24.48;247=24.26;248=24.03;249=23.79;250=23.55;251=23.32;252=23.09;253=22.87;254=22.67;255=22.47;256=22.28;257=22.1;258=21.91;259=21.73;260=21.53;261=21.32;262=21.1;263=20.86;264=20.62;265=20.35;266=20.07;267=19.78;268=19.48;269=19.17;270=18.85;271=18.51;272=18.16;273=17.81;274=17.44;275=17.06;276=16.67;277=16.28;278=15.89;279=15.5;280=15.12;281=14.74;282=14.38;283=14.03;284=13.69;285=13.36;286=13.05;287=12.74;288=12.45;289=12.16;290=11.87;291=11.59;292=11.3;293=11.01;294=10.71;295=10.41;296=10.11;297=9.81;298=9.51;299=9.22;300=8.93;301=8.66;302=8.41;303=8.16;304=7.94;305=7.73;306=7.53;307=7.35;308=7.19;309=7.03;310=6.88;311=6.73;312=6.57;313=6.42;314=6.26;315=6.09;316=5.91;317=5.72;318=5.51;319=5.3;320=5.08;321=4.84;322=4.61;323=4.36;324=4.12;325=3.87;326=3.62;327=3.38;328=3.13;329=2.89;330=2.66;331=2.43;332=2.21;333=2;334=1.79;335=1.6;336=1.42;337=1.25;338=1.09;339=0.94;340=0.81;341=0.69;342=0.58;343=0.48;344=0.4;345=0.33;346=0.27;347=0.22;348=0.17;349=0.13;350=0.1;351=0.08;352=0.06;353=0.04;354=0.03;355=0.02;356=0.01;357=0;358=0;359=0';
            modul    := 'QPSK';
            end
          else if Pos('1800', GetDateFrec(f_Tparam.Cells[iPa.f_Tr.n, n], 1)) > 0 then
              begin
              Gain     := '17.1';
              modul    := '8-PSK(GMSK)';
              DNA_vert :=
                '0=0.03;1=0.06;2=0.49;3=1.43;4=2.93;5=5.08;6=8.05;7=12.34;8=18.6;9=19.78;10=15.55;11=13.33;12=12.61;13=13.08;14=14.79;15=17.96;16=22.56;17=24.42;18=21.11;19=18.49;20=17.38;21=17.65;22=19.23;23=22.14;24=26.02;25=26.6;26=23.01;27=20.28;28=18.87;29=18.43;30=18.65;31=19.26;32=19.85;33=19.82;34=18.95;35=17.67;36=16.45;37=15.56;38=15.12;39=15.16;40=15.62;41=16.51;42=17.94;43=20.04;44=22.79;45=25.47;46=26.86;47=26.85;48=25.91;49=24.56;50=23.53;51=23.14;52=23.3;53=23.8;54=24.54;55=25.62;56=27.1;57=28.75;58=30.18;59=31.4;60=32.84;61=34.68;62=36.54;63=38.04;64=39.44;65=41.12;66=42.94;67=44.12;68=43.8;69=42.18;70=40.58;71=39.93;72=40.41;73=41.59;74=42.47;75=42.6;76=42.6;77=43.04;78=44;79=44.89;80=44.97;81=44.89;82=44.83;83=43.57;84=41.39;85=39.56;86=38.5;87=38.3;88=38.75;89=39.37;90=39.91;91=40.93;92=42.81;93=44.17;94=42.83;95=40.18;96=37.75;97=36.16;98=35.62;99=36.02;100=36.93;101=37.67;102=37.68;103=37.07;104=36.42;105=36.16;106=36.25;107=36.32;108=36.05;109=35.5;110=35.04;111=34.97;112=35.31;113=35.76;114=35.88;115=35.47;116=34.87;117=34.5;118=34.59;119=35.08;120=35.73;121=36.22;122=36.4;123=36.42;124=36.64;125=37.34;126=38.68;127=40.69;128=43.21;129=45.94;130=48.52;131=50.74;132=52.19;133=51.77;134=49.21;135=46.22;136=44.14;137=43.35;138=43.79;139=45.02;140=45.98;141=45.71;142=44.63;143=43.67;144=43.08;145=42.41;146=41.34;147=40.46;148=40.51;149=41.77;150=44.15;151=46.51;152=46.2;153=44.08;154=42.42;155=41.71;156=41.89;157=42.65;158=43.46;159=43.55;160=42.44;161=40.56;162=38.64;163=37.19;164=36.49;165=36.64;166=37.65;167=39.28;168=40.91;169=41.73;170=41.68;171=41.71;172=42.95;173=46.67;174=54.06;175=47.77;176=41.51;177=37.35;178=34.48;179=32.79;180=32.22;181=32.78;182=34.62;183=38.35;184=45.47;185=45.88;186=40.3;187=37.94;188=36.94;189=36.05;190=34.61;191=33.07;192=32.06;193=31.82;194=32.35;195=33.7;196=36.13;197=40.24;198=46.83;199=50.48;200=47.97;201=47.27;202=47.15;203=46.14;204=44.29;205=42.37;206=40.61;207=39.04;208=38.13;209=38.56;210=40.92;211=45.16;212=45.31;213=43.05;214=44.19;215=51.17;216=46.01;217=39.73;218=37.29;219=36.64;220=36.58;221=36.54;222=36.86;223=37.98;224=39.75;225=41.21;226=41.48;227=41.26;228=41.96;229=44.69;230=50.75;231=55.32;232=52.1;233=52.42;234=53.69;235=54.01;236=55.87;237=58.86;238=59.4;239=55.3;240=50.63;241=49.56;242=54.58;243=55.06;244=45.8;245=42.81;246=41.9;247=40.22;248=37.84;249=36.51;250=36.58;251=37.69;252=38.95;253=39.31;254=38.5;255=37.25;256=36.05;257=34.88;258=33.8;259=33.07;260=32.77;261=32.64;262=32.3;263=31.68;264=30.95;265=30.39;266=30.18;267=30.36;268=30.73;269=31.04;270=31.18;271=31.29;272=31.57;273=32.06;274=32.56;275=32.78;276=32.59;277=32.15;278=31.82;279=31.9;280=32.43;281=33.14;282=33.5;283=33.14;284=32.5;285=32.32;286=33.07;287=34.74;288=36.26;289=35.61;290=33.67;291=32.29;292=31.67;293=30.87;294=28.9;295=26.51;296=24.72;297=23.77;298=23.45;299=23.21;300=22.5;301=21.43;302=20.5;303=19.99;304=19.89;305=19.9;306=19.8;307=19.77;308=20.24;309=21.49;310=23.44;311=25.48;312=27.2;313=29.68;314=35.52;315=36.48;316=29.12;317=26.12;318=25.42;319=25.76;320=25.52;321=24.04;322=22.23;323=20.78;324=19.99;325=19.88;326=19.77;327=18.73;328=17.3;329=16.63;330=17.17;331=18.82;332=20.87;333=22.64;334=24.94;335=27.89;336=26.8;337=24.62;338=24.83;339=27.88;340=35.1;341=48.01;342=35.73;343=32.73;344=34.24;345=49.89;346=33.36;347=26.93;348=24.28;349=22.87;350=20.95;351=17.72;352=14.03;353=10.7;354=7.91;355=5.57;356=3.63;357=2.11;358=1.03;359=0.35';
              DNA_horz :=
                '0=0.02;1=0.06;2=0.1;3=0.15;4=0.22;5=0.29;6=0.37;7=0.46;8=0.55;9=0.64;10=0.73;11=0.83;12=0.93;13=1.02;14=1.12;15=1.21;16=1.3;17=1.39;18=1.48;19=1.57;20=1.65;21=1.74;22=1.83;23=1.91;24=2;25=2.1;26=2.2;27=2.3;28=2.41;29=2.53;30=2.65;31=2.78;32=2.92;33=3.07;34=3.23;35=3.41;36=3.59;37=3.79;38=3.99;39=4.21;40=4.45;41=4.69;42=4.95;43=5.22;44=5.5;45=5.79;46=6.09;47=6.4;48=6.73;49=7.06;50=7.39;51=7.73;52=8.08;53=8.43;54=8.78;55=9.13;56=9.49;57=9.83;58=10.18;59=10.52;60=10.86;61=11.2;62=11.53;63=11.86;64=12.18;65=12.5;66=12.83;67=13.15;68=13.47;69=13.79;70=14.12;71=14.46;72=14.79;73=15.14;74=15.49;75=15.84;76=16.2;77=16.57;78=16.94;79=17.32;80=17.7;81=18.08;82=18.46;83=18.84;84=19.22;85=19.61;86=19.99;87=20.39;88=20.79;89=21.2;90=21.62;91=22.06;92=22.51;93=22.98;94=23.46;95=23.96;96=24.46;97=24.96;98=25.45;99=25.92;100=26.35;101=26.73;102=27.06;103=27.34;104=27.57;105=27.75;106=27.91;107=28.05;108=28.2;109=28.37;110=28.57;111=28.82;112=29.11;113=29.45;114=29.84;115=30.28;116=30.75;117=31.25;118=31.75;119=32.23;120=32.67;121=33.04;122=33.33;123=33.53;124=33.67;125=33.77;126=33.85;127=33.95;128=34.09;129=34.3;130=34.59;131=34.98;132=35.49;133=36.13;134=36.9;135=37.82;136=38.87;137=40.04;138=41.24;139=42.3;140=42.97;141=43.09;142=42.68;143=42;144=41.25;145=40.58;146=40.06;147=39.71;148=39.54;149=39.54;150=39.72;151=40.08;152=40.6;153=41.27;154=42.02;155=42.75;156=43.27;157=43.35;158=42.9;159=42.04;160=40.97;161=39.87;162=38.83;163=37.88;164=37.06;165=36.34;166=35.71;167=35.17;168=34.7;169=34.29;170=33.94;171=33.62;172=33.34;173=33.09;174=32.87;175=32.69;176=32.53;177=32.42;178=32.36;179=32.36;180=32.43;181=32.57;182=32.81;183=33.14;184=33.6;185=34.19;186=34.94;187=35.87;188=37.02;189=38.44;190=40.14;191=42.09;192=43.98;193=44.87;194=44.06;195=42.31;196=40.53;197=39.01;198=37.8;199=36.87;200=36.19;201=35.73;202=35.47;203=35.42;204=35.55;205=35.89;206=36.44;207=37.23;208=38.28;209=39.66;210=41.43;211=43.6;212=45.91;213=47.02;214=45.72;215=43.37;216=41.19;217=39.42;218=38.02;219=36.93;220=36.08;221=35.46;222=35.01;223=34.74;224=34.61;225=34.63;226=34.77;227=35.04;228=35.42;229=35.88;230=36.41;231=36.94;232=37.39;233=37.66;234=37.66;235=37.35;236=36.77;237=36.01;238=35.18;239=34.34;240=33.54;241=32.79;242=32.13;243=31.53;244=31.01;245=30.56;246=30.17;247=29.83;248=29.54;249=29.29;250=29.06;251=28.85;252=28.64;253=28.42;254=28.18;255=27.9;256=27.57;257=27.2;258=26.78;259=26.31;260=25.8;261=25.27;262=24.73;263=24.18;264=23.64;265=23.11;266=22.6;267=22.11;268=21.65;269=21.22;270=20.82;271=20.44;272=20.09;273=19.75;274=19.44;275=19.13;276=18.82;277=18.52;278=18.21;279=17.88;280=17.54;281=17.18;282=16.8;283=16.4;284=15.98;285=15.55;286=15.11;287=14.66;288=14.21;289=13.77;290=13.33;291=12.9;292=12.48;293=12.08;294=11.68;295=11.31;296=10.94;297=10.59;298=10.25;299=9.92;300=9.6;301=9.29;302=8.98;303=8.69;304=8.4;305=8.11;306=7.84;307=7.57;308=7.3;309=7.05;310=6.8;311=6.56;312=6.33;313=6.11;314=5.9;315=5.7;316=5.51;317=5.32;318=5.15;319=4.99;320=4.83;321=4.69;322=4.55;323=4.41;324=4.29;325=4.17;326=4.05;327=3.93;328=3.82;329=3.7;330=3.59;331=3.47;332=3.34;333=3.21;334=3.08;335=2.94;336=2.79;337=2.63;338=2.47;339=2.3;340=2.13;341=1.95;342=1.77;343=1.59;344=1.41;345=1.24;346=1.07;347=0.91;348=0.76;349=0.62;350=0.5;351=0.38;352=0.28;353=0.2;354=0.13;355=0.07;356=0.03;357=0.01;358=0;359=0';
              end
            else if Pos('800', GetDateFrec(f_Tparam.Cells[iPa.f_Tr.n, n], 1)) > 0 then
                begin
                Gain     := '17';
                modul    := '64QPSK';
                DNA_vert :=
                  '0=0.02;1=0.08;2=0.58;3=1.56;4=3.1;5=5.38;6=8.75;7=14.23;8=26.61;9=21.33;10=14.77;11=12.34;12=11.65;13=12.2;14=13.95;15=17.32;16=24.21;17=45.37;18=22.96;19=18.4;20=16.43;21=15.78;22=16.11;23=17.36;24=19.6;25=23.15;26=28.68;27=35.66;28=34.28;29=32.11;30=33.23;31=38.25;32=38.55;33=30.61;34=26.11;35=23.28;36=21.33;37=19.94;38=18.96;39=18.45;40=18.51;41=19.28;42=20.9;43=23.61;44=27.8;45=33.67;46=37.04;47=36.15;48=37.65;49=38.98;50=33.96;51=30.01;52=28.13;53=27.74;54=28.58;55=30.55;56=33.61;57=37.22;58=39.61;59=40.46;60=40.67;61=37.68;62=33.53;63=30.83;64=29.73;65=30.01;66=31.2;67=31.98;68=30.91;69=28.82;70=27.06;71=26.25;72=26.53;73=27.61;74=28.69;75=28.93;76=28.13;77=26.89;78=26.01;79=25.96;80=26.38;81=26.59;82=26.64;83=26.26;84=25.6;85=25.39;86=25.88;87=26.84;88=27.62;89=27.68;90=27.25;91=27.18;92=28.2;93=30.8;94=35.6;95=42.77;96=41.55;97=35.74;98=32.51;99=31.36;100=31.84;101=33.14;102=34.15;103=35.12;104=37.47;105=42.2;106=45.98;107=40.98;108=36.83;109=34.66;110=34.14;111=35;112=36.84;113=39.09;114=40.97;115=41.77;116=42.04;117=43.1;118=44.85;119=45.13;120=43.76;121=42.59;122=41.59;123=40;124=38.03;125=36.6;126=36.29;127=37.24;128=39.18;129=41.35;130=42.74;131=43.4;132=44.89;133=48.98;134=47.74;135=41.19;136=37.72;137=36.05;138=35.34;139=35.19;140=35.58;141=36.5;142=37.71;143=38.62;144=38.55;145=37.98;146=38.32;147=40.46;148=43.99;149=43.6;150=40.51;151=38.88;152=38.61;153=39.07;154=38.7;155=36.8;156=35.19;157=34.96;158=36.23;159=38.83;160=42.39;161=46.66;162=50.72;163=52.61;164=53.08;165=46.25;166=39.95;167=36.12;168=33.93;169=32.86;170=32.61;171=32.92;172=33.65;173=34.7;174=35.98;175=37.41;176=39.13;177=41.47;178=44.38;179=47.05;180=48.71;181=47.54;182=43.92;183=40.43;184=37.42;185=35.13;186=33.87;187=33.71;188=34.44;189=35.66;190=37.09;191=38.96;192=41.73;193=45.58;194=49;195=49.73;196=53.08;197=65.26;198=46.14;199=41.13;200=39.54;201=39.98;202=41.47;203=42.54;204=42.18;205=41.12;206=40.22;207=38.58;208=35.77;209=33.46;210=32.6;211=33.18;212=34.65;213=35.73;214=35.39;215=34.24;216=33.3;217=33.05;218=33.13;219=32.86;220=32.53;221=33.1;222=35.05;223=38.41;224=43.14;225=49.56;226=50.72;227=48.9;228=50.19;229=43.17;230=37.38;231=34.61;232=33.95;233=35.25;234=39.35;235=47.56;236=43.18;237=39.14;238=39.1;239=41.03;240=40.65;241=36.63;242=33.14;243=31.15;244=30.77;245=32.01;246=34.62;247=37.2;248=37.15;249=34.84;250=32.63;251=31.63;252=32.11;253=34.19;254=38.23;255=42.46;256=39.12;257=36.01;258=34.96;259=34.47;260=33.98;261=34.53;262=36.75;263=39.02;264=40.03;265=40.23;266=34.68;267=30.04;268=27.99;269=27.85;270=28.92;271=30.84;272=34.58;273=39.94;274=34.69;275=30.57;276=29.26;277=29.29;278=30.01;279=32.19;280=38.15;281=44.82;282=36.41;283=33.88;284=33.38;285=33.72;286=36.01;287=39.14;288=34;289=29.88;290=28.42;291=28.76;292=30.52;293=33.47;294=33.08;295=28.39;296=25.12;297=23.46;298=22.9;299=22.95;300=23.18;301=23.13;302=22.56;303=21.86;304=21.51;305=21.69;306=22.38;307=23.49;308=24.98;309=26.68;310=28.17;311=28.93;312=28.88;313=28.55;314=28.44;315=28.51;316=28.4;317=27.79;318=26.96;319=26.5;320=26.72;321=27.46;322=27.88;323=27.11;324=25.66;325=24.13;326=22.7;327=21.35;328=20.01;329=18.72;330=17.61;331=16.79;332=16.32;333=16.18;334=16.34;335=16.86;336=17.91;337=19.72;338=22.82;339=28.56;340=41.33;341=32.73;342=28.02;343=27.06;344=28.04;345=29.57;346=28.84;347=26.49;348=25.16;349=26.12;350=34.28;351=30.63;352=18.86;353=12.94;354=8.94;355=6.02;356=3.84;357=2.23;358=1.08;359=0.35';
                DNA_horz :=
                  '0=0.03;1=0.01;2=0;3=0;4=0;5=0.01;6=0.03;7=0.05;8=0.07;9=0.11;10=0.15;11=0.2;12=0.25;13=0.31;14=0.38;15=0.45;16=0.53;17=0.61;18=0.7;19=0.8;20=0.91;21=1.02;22=1.14;23=1.26;24=1.39;25=1.52;26=1.67;27=1.81;28=1.97;29=2.13;30=2.29;31=2.46;32=2.64;33=2.82;34=3.01;35=3.2;36=3.39;37=3.6;38=3.8;39=4.01;40=4.23;41=4.44;42=4.67;43=4.89;44=5.12;45=5.35;46=5.59;47=5.82;48=6.06;49=6.3;50=6.55;51=6.8;52=7.05;53=7.3;54=7.55;55=7.81;56=8.07;57=8.32;58=8.59;59=8.85;60=9.12;61=9.39;62=9.66;63=9.93;64=10.21;65=10.49;66=10.78;67=11.06;68=11.35;69=11.65;70=11.94;71=12.24;72=12.55;73=12.85;74=13.16;75=13.47;76=13.78;77=14.1;78=14.41;79=14.72;80=15.03;81=15.34;82=15.65;83=15.95;84=16.24;85=16.53;86=16.81;87=17.08;88=17.33;89=17.58;90=17.82;91=18.04;92=18.25;93=18.46;94=18.65;95=18.83;96=19;97=19.17;98=19.33;99=19.48;100=19.64;101=19.79;102=19.94;103=20.09;104=20.24;105=20.4;106=20.56;107=20.73;108=20.91;109=21.09;110=21.28;111=21.47;112=21.68;113=21.89;114=22.11;115=22.34;116=22.58;117=22.82;118=23.07;119=23.33;120=23.6;121=23.87;122=24.15;123=24.44;124=24.73;125=25.02;126=25.32;127=25.63;128=25.94;129=26.25;130=26.57;131=26.89;132=27.22;133=27.55;134=27.89;135=28.23;136=28.57;137=28.9;138=29.24;139=29.57;140=29.89;141=30.21;142=30.5;143=30.78;144=31.02;145=31.24;146=31.41;147=31.56;148=31.66;149=31.72;150=31.75;151=31.74;152=31.7;153=31.65;154=31.58;155=31.5;156=31.43;157=31.36;158=31.31;159=31.27;160=31.26;161=31.27;162=31.31;163=31.39;164=31.51;165=31.67;166=31.88;167=32.14;168=32.45;169=32.83;170=33.28;171=33.81;172=34.43;173=35.15;174=36;175=37.01;176=38.22;177=39.68;178=41.52;179=43.93;180=47.25;181=52.17;182=55.62;183=50.36;184=45.93;185=42.87;186=40.59;187=38.81;188=37.35;189=36.13;190=35.08;191=34.19;192=33.41;193=32.72;194=32.13;195=31.6;196=31.13;197=30.72;198=30.35;199=30.03;200=29.75;201=29.49;202=29.27;203=29.07;204=28.89;205=28.72;206=28.57;207=28.43;208=28.29;209=28.16;210=28.03;211=27.9;212=27.76;213=27.62;214=27.48;215=27.33;216=27.17;217=27.02;218=26.86;219=26.7;220=26.54;221=26.39;222=26.25;223=26.11;224=25.98;225=25.86;226=25.75;227=25.66;228=25.58;229=25.52;230=25.47;231=25.44;232=25.42;233=25.42;234=25.42;235=25.44;236=25.46;237=25.49;238=25.51;239=25.53;240=25.54;241=25.52;242=25.49;243=25.42;244=25.32;245=25.18;246=25.01;247=24.79;248=24.54;249=24.24;250=23.92;251=23.58;252=23.22;253=22.84;254=22.46;255=22.08;256=21.69;257=21.31;258=20.94;259=20.57;260=20.22;261=19.87;262=19.54;263=19.22;264=18.91;265=18.61;266=18.32;267=18.04;268=17.77;269=17.5;270=17.25;271=16.99;272=16.75;273=16.5;274=16.26;275=16.01;276=15.77;277=15.52;278=15.27;279=15.01;280=14.75;281=14.48;282=14.22;283=13.94;284=13.66;285=13.37;286=13.08;287=12.79;288=12.49;289=12.2;290=11.9;291=11.6;292=11.3;293=11;294=10.71;295=10.41;296=10.12;297=9.83;298=9.55;299=9.27;300=8.99;301=8.72;302=8.46;303=8.2;304=7.95;305=7.7;306=7.45;307=7.22;308=6.98;309=6.76;310=6.53;311=6.31;312=6.1;313=5.89;314=5.69;315=5.49;316=5.29;317=5.1;318=4.91;319=4.72;320=4.54;321=4.36;322=4.18;323=4.01;324=3.84;325=3.67;326=3.51;327=3.35;328=3.19;329=3.03;330=2.87;331=2.72;332=2.57;333=2.43;334=2.28;335=2.14;336=2.01;337=1.87;338=1.74;339=1.62;340=1.49;341=1.38;342=1.26;343=1.15;344=1.04;345=0.94;346=0.84;347=0.75;348=0.66;349=0.58;350=0.5;351=0.43;352=0.36;353=0.3;354=0.25;355=0.2;356=0.15;357=0.11;358=0.08;359=0.05';
                end
              else //if Pos('900', s[iPa.tech_Se.n]) > 0 then
                begin
                Gain     := '17.3';
                modul    := '8-PSK(GMSK)';
                DNA_vert :=
                  '0=0.06;1=0.09;2=0.76;3=2.16;4=4.53;5=8.42;6=15.98;7=28.44;8=14.31;9=10.33;10=9.04;11=9.4;12=11.32;13=15.29;14=21.85;15=20.63;16=15.43;17=12.92;18=12.03;19=12.35;20=13.85;21=16.78;22=21.74;23=27.05;24=24.42;25=21.58;26=20.97;27=21.87;28=23.47;29=24.27;30=23.45;31=22.12;32=21.19;33=20.88;34=21.27;35=22.5;36=25.05;37=30.3;38=34.95;39=26.98;40=22.08;41=19.61;42=18.73;43=19.07;44=20.27;45=21.81;46=23.13;47=23.69;48=23.21;49=22.03;50=20.93;51=20.53;52=21.22;53=23.16;54=26.47;55=30.55;56=31.8;57=30.65;58=30.42;59=29.72;60=27.2;61=24.48;62=22.74;63=22.25;64=22.96;65=24.27;66=24.89;67=23.99;68=22.71;69=22.29;70=23.07;71=24.87;72=26.12;73=24.18;74=21.55;75=20.29;76=20.37;77=21.75;78=25.31;79=30.28;80=29.52;81=27.71;82=26.37;83=26.36;84=28.57;85=30.34;86=30.68;87=29.84;88=28.17;89=26.4;90=26.09;91=26.88;92=27.82;93=28.42;94=28.66;95=28.16;96=26.5;97=25.47;98=25.7;99=26.96;100=29.18;101=32.99;102=34.92;103=31.1;104=28.6;105=27.9;106=28.66;107=31.37;108=35.77;109=35.85;110=32.68;111=31.2;112=31.93;113=35.21;114=38.11;115=35.89;116=33.9;117=33.89;118=37.03;119=43.91;120=40.23;121=37.22;122=37.46;123=39.46;124=45.63;125=45.89;126=40.82;127=43.25;128=55.15;129=46.31;130=41.55;131=39.42;132=40.09;133=43.39;134=41.11;135=38.2;136=38.18;137=39.72;138=41.49;139=40.18;140=38.32;141=39.79;142=46.29;143=56.73;144=60.45;145=49.17;146=44.3;147=45.9;148=56.7;149=59.56;150=55.45;151=59.14;152=65.19;153=50.18;154=44.64;155=43.14;156=43.35;157=42.52;158=40.2;159=38.91;160=40.02;161=43.94;162=48.06;163=46.93;164=44.22;165=42.07;166=39.53;167=36.7;168=35.47;169=36.62;170=40.02;171=44.5;172=47.42;173=48.8;174=51.66;175=47.86;176=42.99;177=43.03;178=49.08;179=59.2;180=48.22;181=44.8;182=42.46;183=40.05;184=37.77;185=36.73;186=37.32;187=39.09;188=40.75;189=41.18;190=41.31;191=42.94;192=45.31;193=44.25;194=41.94;195=40.45;196=39.43;197=38.47;198=37.8;199=38.39;200=40.92;201=44.86;202=43.98;203=39.5;204=36.94;205=36.14;206=36.04;207=36.04;208=36.47;209=37.48;210=38.69;211=39.45;212=39.84;213=41.01;214=42.95;215=43.46;216=41.56;217=39.63;218=40.01;219=43.84;220=45.83;221=43.11;222=43.79;223=46.25;224=44.9;225=41.11;226=38.96;227=39.75;228=43.97;229=49.12;230=50.2;231=48.38;232=45.38;233=43.81;234=41.97;235=40.68;236=40.17;237=39.19;238=38.22;239=37.51;240=37.24;241=38.52;242=41.39;243=42.29;244=38.66;245=34.87;246=33.3;247=33.95;248=36.07;249=38.86;250=39.31;251=35.45;252=32.7;253=31.5;254=31.41;255=32.88;256=36.38;257=43.03;258=44.38;259=34.74;260=31.31;261=31.3;262=33.11;263=34.68;264=32.9;265=28.62;266=26.11;267=25.67;268=26.51;269=28.46;270=32.87;271=33.92;272=28.06;273=24.89;274=23.7;275=24.31;276=27.44;277=34.07;278=32.37;279=26.74;280=24.21;281=23.77;282=25.14;283=27.35;284=27.19;285=24.7;286=22.82;287=22.33;288=23.18;289=24.99;290=26.67;291=26.88;292=26.11;293=25.79;294=26.39;295=27.8;296=29.83;297=32.74;298=36.7;299=39.8;300=41.22;301=42.06;302=37.19;303=32.73;304=30.61;305=30.35;306=30.97;307=29.31;308=25.88;309=23.53;310=22.55;311=22.49;312=22.65;313=22.4;314=22.1;315=22.59;316=24.5;317=28.31;318=33.2;319=31.91;320=27.32;321=23.69;322=21.17;323=19.93;324=20.07;325=21.6;326=24.12;327=25.78;328=25.53;329=25.64;330=27.22;331=29.55;332=29.47;333=25.18;334=20.46;335=16.92;336=14.59;337=13.32;338=12.88;339=13.03;340=13.6;341=14.64;342=16.44;343=19.52;344=24.65;345=31.67;346=33.37;347=33.49;348=31.42;349=26.05;350=22.95;351=22.4;352=23.57;353=20.62;354=14.25;355=9.36;356=5.88;357=3.41;358=1.69;359=0.59';
                DNA_horz :=
                  '0=0.17;1=0.13;2=0.09;3=0.06;4=0.03;5=0.01;6=0;7=0;8=0;9=0.01;10=0.03;11=0.05;12=0.09;13=0.13;14=0.17;15=0.23;16=0.29;17=0.36;18=0.44;19=0.52;20=0.62;21=0.72;22=0.82;23=0.94;24=1.06;25=1.19;26=1.32;27=1.46;28=1.61;29=1.76;30=1.92;31=2.09;32=2.26;33=2.44;34=2.62;35=2.81;36=3;37=3.2;38=3.4;39=3.61;40=3.82;41=4.04;42=4.26;43=4.49;44=4.72;45=4.95;46=5.19;47=5.43;48=5.68;49=5.93;50=6.18;51=6.44;52=6.7;53=6.97;54=7.23;55=7.5;56=7.78;57=8.06;58=8.34;59=8.62;60=8.91;61=9.2;62=9.48;63=9.78;64=10.07;65=10.36;66=10.66;67=10.95;68=11.25;69=11.54;70=11.83;71=12.13;72=12.42;73=12.71;74=13;75=13.29;76=13.58;77=13.87;78=14.15;79=14.44;80=14.73;81=15.03;82=15.32;83=15.62;84=15.92;85=16.22;86=16.53;87=16.84;88=17.15;89=17.47;90=17.79;91=18.11;92=18.43;93=18.76;94=19.09;95=19.41;96=19.73;97=20.05;98=20.36;99=20.66;100=20.96;101=21.25;102=21.53;103=21.8;104=22.06;105=22.31;106=22.55;107=22.78;108=23.01;109=23.24;110=23.46;111=23.69;112=23.92;113=24.16;114=24.4;115=24.66;116=24.93;117=25.22;118=25.53;119=25.86;120=26.2;121=26.57;122=26.97;123=27.39;124=27.83;125=28.3;126=28.8;127=29.31;128=29.85;129=30.4;130=30.97;131=31.54;132=32.1;133=32.64;134=33.15;135=33.61;136=34;137=34.32;138=34.55;139=34.69;140=34.75;141=34.73;142=34.66;143=34.55;144=34.41;145=34.26;146=34.1;147=33.95;148=33.8;149=33.67;150=33.56;151=33.46;152=33.38;153=33.32;154=33.28;155=33.25;156=33.24;157=33.25;158=33.27;159=33.31;160=33.37;161=33.44;162=33.52;163=33.62;164=33.73;165=33.86;166=34;167=34.16;168=34.33;169=34.53;170=34.75;171=35;172=35.27;173=35.58;174=35.92;175=36.3;176=36.73;177=37.22;178=37.79;179=38.43;180=39.19;181=40.06;182=41.12;183=42.39;184=43.99;185=46.07;186=49;187=53.73;188=65.96;189=58.91;190=51.01;191=46.82;192=43.95;193=41.75;194=39.96;195=38.47;196=37.18;197=36.07;198=35.08;199=34.21;200=33.43;201=32.73;202=32.11;203=31.56;204=31.06;205=30.63;206=30.25;207=29.92;208=29.64;209=29.41;210=29.22;211=29.08;212=28.99;213=28.94;214=28.93;215=28.97;216=29.05;217=29.17;218=29.34;219=29.55;220=29.79;221=30.08;222=30.4;223=30.74;224=31.1;225=31.46;226=31.81;227=32.12;228=32.36;229=32.51;230=32.55;231=32.47;232=32.27;233=31.97;234=31.58;235=31.14;236=30.67;237=30.19;238=29.71;239=29.24;240=28.8;241=28.37;242=27.98;243=27.61;244=27.26;245=26.93;246=26.62;247=26.31;248=26.02;249=25.74;250=25.45;251=25.15;252=24.85;253=24.54;254=24.22;255=23.87;256=23.52;257=23.15;258=22.77;259=22.38;260=21.98;261=21.57;262=21.17;263=20.76;264=20.36;265=19.97;266=19.58;267=19.2;268=18.84;269=18.48;270=18.14;271=17.81;272=17.49;273=17.19;274=16.9;275=16.61;276=16.34;277=16.08;278=15.83;279=15.58;280=15.34;281=15.1;282=14.87;283=14.64;284=14.4;285=14.17;286=13.94;287=13.7;288=13.46;289=13.21;290=12.96;291=12.71;292=12.44;293=12.18;294=11.91;295=11.64;296=11.36;297=11.08;298=10.81;299=10.53;300=10.25;301=9.97;302=9.69;303=9.42;304=9.15;305=8.89;306=8.62;307=8.36;308=8.11;309=7.87;310=7.62;311=7.39;312=7.16;313=6.93;314=6.71;315=6.49;316=6.28;317=6.07;318=5.87;319=5.67;320=5.48;321=5.29;322=5.1;323=4.92;324=4.74;325=4.57;326=4.39;327=4.22;328=4.05;329=3.89;330=3.72;331=3.56;332=3.4;333=3.24;334=3.08;335=2.93;336=2.78;337=2.63;338=2.48;339=2.34;340=2.2;341=2.06;342=1.92;343=1.79;344=1.66;345=1.53;346=1.41;347=1.29;348=1.17;349=1.06;350=0.96;351=0.85;352=0.76;353=0.66;354=0.58;355=0.49;356=0.42;357=0.35;358=0.28;359=0.23';
                end;
        end;
      'Технологическая сеть',
      'Фиксированной службы',
      'Транкинговая сеть', 'КВ':
        if (Pos('SLR5500', f_Tparam.Cells[iPa.prd_An.n, n]) > 0) or
          (Pos('SLR 5500', f_Tparam.Cells[iPa.prd_An.n, n]) > 0) then
          begin
          Gain     := '6';
          antModel := 'Петлевой вибратор (Radial DP2 VHF)';
          modul    := 'ЧМ';
          DNA_vert :=
            '0=0;1=0.027;2=0.1;3=0.173;4=0.263;5=0.37;6=0.493;7=0.631;8=0.784;9=0.952;10=1.135;11=1.333;12=1.545;13=1.773;14=2.018;15=2.278;16=2.556;17=2.85;18=3.162;19=3.74;20=4.363;21=5.028;22=5.732;23=6.472;24=7.244;25=8.045;26=8.873;27=9.725;28=10.597;29=30.999;30=30.121;31=29.242;32=28.365;33=27.49;34=26.62;35=25.756;36=24.902;37=24.059;38=23.231;39=22.42;40=21.629;41=20.863;42=20.126;43=19.42;44=18.753;45=18.129;46=17.554;47=17.033;48=16.57;49=16.169;50=15.833;51=15.567;52=15.373;53=15.255;54=15.216;55=15.26;56=15.388;57=15.598;58=15.886;59=16.249;60=16.682;61=17.182;62=17.743;63=18.36;64=19.029;65=19.742;66=20.494;67=21.278;68=22.091;69=22.928;70=23.785;71=24.659;72=25.546;73=26.443;74=27.349;75=28.26;76=29.174;77=30.089;78=31.004;79=30.993;80=30.982;81=30.971;82=30.96;83=30.949;84=30.937;85=30.926;86=30.915;87=30.904;88=30.893;89=30.882;90=30.871;91=30.86;92=30.849;93=30.838;94=30.827;95=30.816;96=30.805;97=30.794;98=30.783;99=30.772;100=30.761;101=30.75;102=30.739;103=30.728;104=30.717;105=30.706;106=30.695;107=30.684;108=30.673;109=30.36;110=30.049;111=29.738;112=29.429;113=29.123;114=28.82;115=28.52;116=28.226;117=27.937;118=27.654;119=27.379;120=27.111;121=26.852;122=26.602;123=26.363;124=26.135;125=25.919;126=25.716;127=25.526;128=25.351;129=25.192;130=25.049;131=24.923;132=24.816;133=24.727;134=24.659;135=24.611;136=24.582;137=24.573;138=24.581;139=24.606;140=24.647;141=24.703;142=24.772;143=24.855;144=24.949;145=25.053;146=25.168;147=25.291;148=25.422;149=25.559;150=25.702;151=25.85;152=26.001;153=26.156;154=26.311;155=26.468;156=19.678;157=16.727;158=16.01;159=15.292;160=14.575;161=13.861;162=13.15;163=12.919;164=12.691;165=12.466;166=12.248;167=12.036;168=11.834;169=11.642;170=11.462;171=11.296;172=11.145;173=11.011;174=10.895;175=10.799;176=10.725;177=10.673;178=10.646;179=10.644;180=10.67;181=10.722;182=10.799;183=10.9;184=11.023;185=11.167;186=11.331;187=11.514;188=11.714;189=11.931;190=12.163;191=12.408;192=12.665;193=12.933;194=13.211;195=13.496;196=13.788;197=14.086;198=14.386;199=14.689;200=15.567;201=16.845;202=18.71;203=27.173;204=26.963;205=26.753;206=26.545;207=26.341;208=26.14;209=25.946;210=25.758;211=25.578;212=25.408;213=25.247;214=25.099;215=24.964;216=24.843;217=24.738;218=24.65;219=24.58;220=24.529;221=24.499;222=24.491;223=24.504;224=24.538;225=24.591;226=24.663;227=24.754;228=24.861;229=24.986;230=25.126;231=25.281;232=25.45;233=25.633;234=25.829;235=26.036;236=26.254;237=26.482;238=26.72;239=26.966;240=27.221;241=27.482;242=27.75;243=28.024;244=28.303;245=28.586;246=28.873;247=29.163;248=29.455;249=29.748;250=30.043;251=30.337;252=30.359;253=30.381;254=30.403;255=30.425;256=30.447;257=30.469;258=30.491;259=30.513;260=30.535;261=30.557;262=30.579;263=30.601;264=30.623;265=30.645;266=30.667;267=30.689;268=30.711;269=30.733;270=30.755;271=30.777;272=30.799;273=30.821;274=30.843;275=30.865;276=30.887;277=30.909;278=30.931;279=30.953;280=30.078;281=29.203;282=28.331;283=27.462;284=26.6;285=25.747;286=24.906;287=24.081;288=23.274;289=22.489;290=21.731;291=21.003;292=20.311;293=19.658;294=19.052;295=18.497;296=17.999;297=17.564;298=17.197;299=16.901;300=16.681;301=16.541;302=16.484;303=16.508;304=16.609;305=16.784;306=17.03;307=17.342;308=17.717;309=18.151;310=18.64;311=19.179;312=19.763;313=20.386;314=21.045;315=21.734;316=22.45;317=23.189;318=23.948;319=24.723;320=25.512;321=26.311;322=27.118;323=27.93;324=28.746;325=29.562;326=30.378;327=25.967;328=21.328;329=16.353;330=11.324;331=10.242;332=9.188;333=8.167;334=7.185;335=6.248;336=5.361;337=4.53;338=4.109;339=3.708;340=3.33;341=2.975;342=2.646;343=2.341;344=2.058;345=1.797;346=1.556;347=1.333;348=1.128;349=0.94;350=0.769;351=0.614;352=0.475;353=0.353;354=0.247;355=0.159;356=0.09;357=0.039;358=0.009;359=0';
          DNA_horz :=
            '0=0;1=0;2=0;3=0;4=0;5=0;6=0;7=0;8=0;9=0;10=0;11=0;12=0;13=0;14=0;15=0;16=0;17=0;18=0;19=0;20=0;21=0;22=0;23=0;24=0;25=0;26=0;27=0;28=0;29=0;30=0;31=0;32=0;33=0;34=0;35=0;36=0;37=0;38=0;39=0;40=0;41=0;42=0;43=0;44=0;45=0;46=0;47=0;48=0;49=0;50=0;51=0;52=0;53=0;54=0;55=0;56=0;57=0;58=0;59=0;60=0;61=0;62=0;63=0;64=0;65=0;66=0;67=0;68=0;69=0;70=0;71=0;72=0;73=0;74=0;75=0;76=0;77=0;78=0;79=0;80=0;81=0;82=0;83=0;84=0;85=0;86=0;87=0;88=0;89=0;90=0;91=0;92=0;93=0;94=0;95=0;96=0;97=0;98=0;99=0;100=0;101=0;102=0;103=0;104=0;105=0;106=0;107=0;108=0;109=0;110=0;111=0;112=0;113=0;114=0;115=0;116=0;117=0;118=0;119=0;120=0;121=0;122=0;123=0;124=0;125=0;126=0;127=0;128=0;129=0;130=0;131=0;132=0;133=0;134=0;135=0;136=0;137=0;138=0;139=0;140=0;141=0;142=0;143=0;144=0;145=0;146=0;147=0;148=0;149=0;150=0;151=0;152=0;153=0;154=0;155=0;156=0;157=0;158=0;159=0;160=0;161=0;162=0;163=0;164=0;165=0;166=0;167=0;168=0;169=0;170=0;171=0;172=0;173=0;174=0;175=0;176=0;177=0;178=0;179=0;180=0;181=0;182=0;183=0;184=0;185=0;186=0;187=0;188=0;189=0;190=0;191=0;192=0;193=0;194=0;195=0;196=0;197=0;198=0;199=0;200=0;201=0;202=0;203=0;204=0;205=0;206=0;207=0;208=0;209=0;210=0;211=0;212=0;213=0;214=0;215=0;216=0;217=0;218=0;219=0;220=0;221=0;222=0;223=0;224=0;225=0;226=0;227=0;228=0;229=0;230=0;231=0;232=0;233=0;234=0;235=0;236=0;237=0;238=0;239=0;240=0;241=0;242=0;243=0;244=0;245=0;246=0;247=0;248=0;249=0;250=0;251=0;252=0;253=0;254=0;255=0;256=0;257=0;258=0;259=0;260=0;261=0;262=0;263=0;264=0;265=0;266=0;267=0;268=0;269=0;270=0;271=0;272=0;273=0;274=0;275=0;276=0;277=0;278=0;279=0;280=0;281=0;282=0;283=0;284=0;285=0;286=0;287=0;288=0;289=0;290=0;291=0;292=0;293=0;294=0;295=0;296=0;297=0;298=0;299=0;300=0;301=0;302=0;303=0;304=0;305=0;306=0;307=0;308=0;309=0;310=0;311=0;312=0;313=0;314=0;315=0;316=0;317=0;318=0;319=0;320=0;321=0;322=0;323=0;324=0;325=0;326=0;327=0;328=0;329=0;330=0;331=0;332=0;333=0;334=0;335=0;336=0;337=0;338=0;339=0;340=0;341=0;342=0;343=0;344=0;345=0;346=0;347=0;348=0;349=0;350=0;351=0;352=0;353=0;354=0;355=0;356=0;357=0;358=0;359=0';
          end
        else
          begin
          Gain     := '3.2';
          antModel := 'Штырь';
          modul    := 'ЧМ';
          DNA_vert :=
            '0=0;1=0.004;2=0.005;3=0.011;4=0.02;5=0.031;6=0.043;7=0.059;8=0.078;9=0.098;10=0.121;11=0.147;12=0.174;13=0.204;14=0.238;15=0.273;16=0.31;17=0.35;18=0.393;19=0.438;20=0.485;21=0.535;22=0.587;23=0.642;24=0.699;25=0.758;26=0.82;27=0.884;28=0.951;29=1.02;30=1.092;31=1.167;32=1.243;33=1.322;34=1.404;35=1.488;36=1.574;37=1.664;38=1.755;39=1.849;40=1.945;41=2.044;42=2.146;43=2.25;44=2.357;45=2.466;46=2.578;47=2.692;48=2.809;49=2.929;50=3.051;51=3.176;52=3.303;53=3.433;54=3.566;55=3.702;56=3.841;57=3.983;58=4.126;59=4.274;60=4.426;61=4.58;62=4.737;63=4.897;64=5.061;65=5.23;66=5.402;67=5.578;68=5.759;69=5.944;70=6.136;71=6.331;72=6.535;73=6.746;74=6.963;75=7.189;76=7.428;77=7.678;78=7.94;79=8.218;80=8.518;81=8.845;82=9.198;83=9.592;84=10.039;85=10.558;86=11.181;87=11.975;88=13.072;89=21.38;90=23.95;91=21.38;92=13.072;93=11.975;94=11.181;95=10.558;96=10.039;97=9.592;98=9.198;99=8.845;100=8.518;101=8.218;102=7.94;103=7.678;104=7.428;105=7.189;106=6.963;107=6.746;108=6.535;109=6.331;110=6.136;111=5.944;112=5.759;113=5.578;114=5.402;115=5.23;116=5.061;117=4.897;118=4.737;119=4.58;120=4.426;121=4.274;122=4.126;123=3.983;124=3.841;125=3.702;126=3.566;127=3.433;128=3.303;129=3.176;130=3.051;131=2.929;132=2.809;133=2.692;134=2.578;135=2.466;136=2.357;137=2.25;138=2.146;139=2.044;140=1.945;141=1.849;142=1.755;143=1.664;144=1.574;145=1.488;146=1.404;147=1.322;148=1.243;149=1.167;150=1.092;151=1.02;152=0.951;153=0.884;154=0.82;155=0.758;156=0.699;157=0.642;158=0.587;159=0.535;160=0.485;161=0.438;162=0.393;163=0.35;164=0.31;165=0.273;166=0.238;167=0.204;168=0.174;169=0.147;170=0.121;171=0.098;172=0.078;173=0.059;174=0.043;175=0.031;176=0.02;177=0.011;178=0.005;179=0.004;180=0;181=0.004;182=0.005;183=0.011;184=0.02;185=0.031;186=0.043;187=0.059;188=0.078;189=0.098;190=0.121;191=0.147;192=0.174;193=0.204;194=0.238;195=0.273;196=0.31;197=0.35;198=0.393;199=0.438;200=0.485;201=0.535;202=0.587;203=0.642;204=0.699;205=0.758;206=0.82;207=0.884;208=0.951;209=1.02;210=1.092;211=1.167;212=1.243;213=1.322;214=1.404;215=1.488;216=1.574;217=1.664;218=1.755;219=1.849;220=1.945;221=2.044;222=2.146;223=2.25;224=2.357;225=2.466;226=2.578;227=2.692;228=2.809;229=2.929;230=3.051;231=3.176;232=3.303;233=3.433;234=3.566;235=3.702;236=3.841;237=3.983;238=4.126;239=4.274;240=4.426;241=4.58;242=4.737;243=4.897;244=5.061;245=5.23;246=5.402;247=5.578;248=5.759;249=5.944;250=6.136;251=6.331;252=6.535;253=6.746;254=6.963;255=7.189;256=7.428;257=7.678;258=7.94;259=8.218;260=8.518;261=8.845;262=9.198;263=9.592;264=10.039;265=10.558;266=11.181;267=11.975;268=13.072;269=14.936;270=23.95;271=14.936;272=13.072;273=11.975;274=11.181;275=10.558;276=10.039;277=9.592;278=9.198;279=8.845;280=8.518;281=8.218;282=7.94;283=7.678;284=7.428;285=7.189;286=6.963;287=6.746;288=6.535;289=6.331;290=6.136;291=5.944;292=5.759;293=5.578;294=5.402;295=5.23;296=5.061;297=4.897;298=4.737;299=4.58;300=4.426;301=4.274;302=4.126;303=3.983;304=3.841;305=3.702;306=3.566;307=3.433;308=3.303;309=3.176;310=3.051;311=2.929;312=2.809;313=2.692;314=2.578;315=2.466;316=2.357;317=2.25;318=2.146;319=2.044;320=1.945;321=1.849;322=1.755;323=1.664;324=1.574;325=1.488;326=1.404;327=1.322;328=1.243;329=1.167;330=1.092;331=1.02;332=0.951;333=0.884;334=0.82;335=0.758;336=0.699;337=0.642;338=0.587;339=0.535;340=0.485;341=0.438;342=0.393;343=0.35;344=0.31;345=0.273;346=0.238;347=0.204;348=0.174;349=0.147;350=0.121;351=0.098;352=0.078;353=0.059;354=0.043;355=0.031;356=0.02;357=0.011;358=0.005;359=0.004';
          DNA_horz :=
            '0=0;1=0;2=0;3=0;4=0;5=0;6=0;7=0;8=0;9=0;10=0;11=0;12=0;13=0;14=0;15=0;16=0;17=0;18=0;19=0;20=0;21=0;22=0;23=0;24=0;25=0;26=0;27=0;28=0;29=0;30=0;31=0;32=0;33=0;34=0;35=0;36=0;37=0;38=0;39=0;40=0;41=0;42=0;43=0;44=0;45=0;46=0;47=0;48=0;49=0;50=0;51=0;52=0;53=0;54=0;55=0;56=0;57=0;58=0;59=0;60=0;61=0;62=0;63=0;64=0;65=0;66=0;67=0;68=0;69=0;70=0;71=0;72=0;73=0;74=0;75=0;76=0;77=0;78=0;79=0;80=0;81=0;82=0;83=0;84=0;85=0;86=0;87=0;88=0;89=0;90=0;91=0;92=0;93=0;94=0;95=0;96=0;97=0;98=0;99=0;100=0;101=0;102=0;103=0;104=0;105=0;106=0;107=0;108=0;109=0;110=0;111=0;112=0;113=0;114=0;115=0;116=0;117=0;118=0;119=0;120=0;121=0;122=0;123=0;124=0;125=0;126=0;127=0;128=0;129=0;130=0;131=0;132=0;133=0;134=0;135=0;136=0;137=0;138=0;139=0;140=0;141=0;142=0;143=0;144=0;145=0;146=0;147=0;148=0;149=0;150=0;151=0;152=0;153=0;154=0;155=0;156=0;157=0;158=0;159=0;160=0;161=0;162=0;163=0;164=0;165=0;166=0;167=0;168=0;169=0;170=0;171=0;172=0;173=0;174=0;175=0;176=0;177=0;178=0;179=0;180=0;181=0;182=0;183=0;184=0;185=0;186=0;187=0;188=0;189=0;190=0;191=0;192=0;193=0;194=0;195=0;196=0;197=0;198=0;199=0;200=0;201=0;202=0;203=0;204=0;205=0;206=0;207=0;208=0;209=0;210=0;211=0;212=0;213=0;214=0;215=0;216=0;217=0;218=0;219=0;220=0;221=0;222=0;223=0;224=0;225=0;226=0;227=0;228=0;229=0;230=0;231=0;232=0;233=0;234=0;235=0;236=0;237=0;238=0;239=0;240=0;241=0;242=0;243=0;244=0;245=0;246=0;247=0;248=0;249=0;250=0;251=0;252=0;253=0;254=0;255=0;256=0;257=0;258=0;259=0;260=0;261=0;262=0;263=0;264=0;265=0;266=0;267=0;268=0;269=0;270=0;271=0;272=0;273=0;274=0;275=0;276=0;277=0;278=0;279=0;280=0;281=0;282=0;283=0;284=0;285=0;286=0;287=0;288=0;289=0;290=0;291=0;292=0;293=0;294=0;295=0;296=0;297=0;298=0;299=0;300=0;301=0;302=0;303=0;304=0;305=0;306=0;307=0;308=0;309=0;310=0;311=0;312=0;313=0;314=0;315=0;316=0;317=0;318=0;319=0;320=0;321=0;322=0;323=0;324=0;325=0;326=0;327=0;328=0;329=0;330=0;331=0;332=0;333=0;334=0;335=0;336=0;337=0;338=0;339=0;340=0;341=0;342=0;343=0;344=0;345=0;346=0;347=0;348=0;349=0;350=0;351=0;352=0;353=0;354=0;355=0;356=0;357=0;358=0;359=0';
          end;
      'РРЛ', 'ЗССС':
        begin
        Gain     := '39.2';
        R_KND    := '38.9';
        R_Diam   := '0.6';
        R_rask   := '1.8';
        modul    := '64QAM';
        antModel := 'Mini-Link NT D=0,6';
        DNA_vert :=
          '0=0;0.1=0.03;0.2=0.13;0.4=0.52;0.6=1.17;0.8=2.07;1.1=3.92;1.4=6.35;1.7=9.36;2=12.96;2.4=23.72;2.7=24.93;3.1=26.32;3.5=27.52;3.9=28.57;4.3=29.49;4.8=30.5;5.2=31.21;5.7=32;6.2=32.7;6.7=33.33;7.2=33.89;7.7=34.4;8.2=34.86;8.8=35.36;9.3=35.73;9.9=36.14;10.5=36.52;11=36.81;11.6=37.13;12.2=37.42;12.9=37.74;13.5=38;14.1=38.24;14.8=38.5;15.4=38.71;16.1=38.94;16.8=39.17;17.5=39.38;18.2=39.58;18.9=39.77;19.6=39.96;20.3=40.14;21=40.32;21.8=40.52;22.5=40.68;23.3=40.87;24=41.03;24.8=41.22;25.6=41.4;26.4=41.57;27.1=41.73;27.9=41.9;28.8=42.1;29.6=42.28;30.4=42.45;31.2=42.62;32.1=42.82;32.9=43;33.8=43.19;34.6=43.37;35.5=43.57;36.4=43.78;37.3=43.98;38.2=44.19;39.1=44.4;40=44.61;40.9=44.83;41.8=45.05;42.7=45.27;43.7=45.52;44.6=45.75;45.5=45.98;46.5=46.25;47.4=46.49;48.4=46.64;49.4=46.7;50.4=46.77;51.3=46.84;52.3=46.93;53.3=47.04;54.3=47.15;55.3=47.27;56.4=47.42;57.4=47.57;58.4=47.74;59.4=47.91;60.5=48.12;61.5=48.32;62.6=48.56;63.6=48.79;64.7=49.06;65.8=49.35;66.8=49.63;67.9=49.96;69=50.31;70.1=50.68;71.2=51.07;72.3=51.48;73.4=51.91;74.5=52.36;75.7=52.87;76.8=53.35;77.9=53.85;79.1=54.38;80.2=54.86;81.3=55.32;82.5=55.76;83.7=56.13;84.8=56.36;86=56.49;87.2=56.48;88.3=56.34;89.5=56.06;90.7=55.68;91.9=55.22;93.1=54.71;94.3=54.18;95.5=53.64;96.8=53.06;98=52.55;99.2=52.05;100.4=51.57;101.7=51.08;102.9=50.66;104.2=50.23;105.4=49.85;106.7=49.48;107.9=49.15;109.2=48.82;110.5=48.52;111.8=48.25;113=48.01;114.3=47.78;115.6=47.56;116.9=47.37;118.2=47.2;119.5=47.05;120.8=46.92;122.1=46.81;123.5=46.71;124.8=46.63;126.1=46.57;127.5=46.53;128.8=46.51;130.1=46.51;131.5=46.52;132.8=46.56;134.2=46.61;135.6=46.69;136.9=46.77;138.3=46.89;139.7=47.03;141.1=47.19;142.4=47.36;143.8=47.56;145.2=47.79;146.6=48.05;148=48.33;149.4=48.63;150.8=48.97;152.3=49.36;153.7=49.76;155.1=50.19;156.5=50.65;158=51.19;159.4=51.72;160.9=52.33;162.3=52.93;163.8=53.59;165.2=54.22;166.7=54.88;168.1=55.45;169.6=55.97;171.1=56.33;172.6=56.5;174=56.45;175.5=56.18;177=55.74;178.5=55.16;180=54.52;181.5=55.16;183=55.74;184.5=56.18;186=56.45;187.4=56.5;188.9=56.33;190.4=55.97;191.9=55.45;193.3=54.88;194.8=54.22;196.2=53.59;197.7=52.93;199.1=52.33;200.6=51.72;202=51.19;203.5=50.65;204.9=50.19;206.3=49.76;207.7=49.36;209.2=48.97;210.6=48.63;212=48.33;213.4=48.05;214.8=47.79;216.2=47.56;217.6=47.36;218.9=47.19;220.3=47.03;221.7=46.89;223.1=46.77;224.4=46.69;225.8=46.61;227.2=46.56;228.5=46.52;229.9=46.51;231.2=46.51;232.5=46.53;233.9=46.57;235.2=46.63;236.5=46.71;237.9=46.81;239.2=46.92;240.5=47.05;241.8=47.2;243.1=47.37;244.4=47.56;245.7=47.78;247=48.01;248.2=48.25;249.5=48.52;250.8=48.82;252.1=49.15;253.3=49.48;254.6=49.85;255.8=50.23;257.1=50.66;258.3=51.08;259.6=51.57;260.8=52.05;262=52.55;263.2=53.06;264.5=53.64;265.7=54.18;266.9=54.71;268.1=55.22;269.3=55.68;270.5=56.06;271.7=56.34;272.8=56.48;274=56.49;275.2=56.36;276.3=56.13;277.5=55.76;278.7=55.32;279.8=54.86;280.9=54.38;282.1=53.85;283.2=53.35;284.3=52.87;285.5=52.36;286.6=51.91;287.7=51.48;288.8=51.07;289.9=50.68;291=50.31;292.1=49.96;293.2=49.63;294.2=49.35;295.3=49.06;296.4=48.79;297.4=48.56;298.5=48.32;299.5=48.12;300.6=47.91;301.6=47.74;302.6=47.57;303.6=47.42;304.7=47.27;305.7=47.15;306.7=47.04;307.7=46.93;308.7=46.84;309.6=46.77;310.6=46.7;311.6=46.64;312.6=46.49;313.5=46.25;314.5=45.98;315.4=45.75;316.3=45.52;317.3=45.27;318.2=45.05;319.1=44.83;320=44.61;320.9=44.4;321.8=44.19;322.7=43.98;323.6=43.78;324.5=43.57;325.4=43.37;326.2=43.19;327.1=43;327.9=42.82;328.8=42.62;329.6=42.45;330.4=42.28;331.2=42.1;332.1=41.9;332.9=41.73;333.6=41.57;334.4=41.4;335.2=41.22;336=41.03;336.7=40.87;337.5=40.68;338.2=40.52;339=40.32;339.7=40.14;340.4=39.96;341.1=39.77;341.8=39.58;342.5=39.38;343.2=39.17;343.9=38.94;344.6=38.71;345.2=38.5;345.9=38.24;346.5=38;347.1=37.74;347.8=37.42;348.4=37.13;349=36.81;349.5=36.52;350.1=36.14;350.7=35.73;351.2=35.36;351.8=34.86;352.3=34.4;352.8=33.89;353.3=33.33;353.8=32.7;354.3=32;354.8=31.21;355.2=30.5;355.7=29.49;356.1=28.57;356.5=27.52;356.9=26.32;357.3=24.93;357.6=23.72;358=12.96;358.3=9.36;358.6=6.35;358.9=3.92;359.2=2.07;359.4=1.17;359.6=0.52;359.8=0.13;359.9=0.03';
        DNA_horz :=
          '0=0;0.1=0.03;0.2=0.13;0.4=0.52;0.6=1.17;0.8=2.07;1.1=3.92;1.4=6.35;1.7=9.36;2=12.96;2.4=23.72;2.7=24.93;3.1=26.32;3.5=27.52;3.9=28.57;4.3=29.49;4.8=30.5;5.2=31.21;5.7=32;6.2=32.7;6.7=33.33;7.2=33.89;7.7=34.4;8.2=34.86;8.8=35.36;9.3=35.73;9.9=36.14;10.5=36.52;11=36.81;11.6=37.13;12.2=37.42;12.9=37.74;13.5=38;14.1=38.24;14.8=38.5;15.4=38.71;16.1=38.94;16.8=39.17;17.5=39.38;18.2=39.58;18.9=39.77;19.6=39.96;20.3=40.14;21=40.32;21.8=40.52;22.5=40.68;23.3=40.87;24=41.03;24.8=41.22;25.6=41.4;26.4=41.57;27.1=41.73;27.9=41.9;28.8=42.1;29.6=42.28;30.4=42.45;31.2=42.62;32.1=42.82;32.9=43;33.8=43.19;34.6=43.37;35.5=43.57;36.4=43.78;37.3=43.98;38.2=44.19;39.1=44.4;40=44.61;40.9=44.83;41.8=45.05;42.7=45.27;43.7=45.52;44.6=45.75;45.5=45.98;46.5=46.25;47.4=46.49;48.4=46.64;49.4=46.7;50.4=46.77;51.3=46.84;52.3=46.93;53.3=47.04;54.3=47.15;55.3=47.27;56.4=47.42;57.4=47.57;58.4=47.74;59.4=47.91;60.5=48.12;61.5=48.32;62.6=48.56;63.6=48.79;64.7=49.06;65.8=49.35;66.8=49.63;67.9=49.96;69=50.31;70.1=50.68;71.2=51.07;72.3=51.48;73.4=51.91;74.5=52.36;75.7=52.87;76.8=53.35;77.9=53.85;79.1=54.38;80.2=54.86;81.3=55.32;82.5=55.76;83.7=56.13;84.8=56.36;86=56.49;87.2=56.48;88.3=56.34;89.5=56.06;90.7=55.68;91.9=55.22;93.1=54.71;94.3=54.18;95.5=53.64;96.8=53.06;98=52.55;99.2=52.05;100.4=51.57;101.7=51.08;102.9=50.66;104.2=50.23;105.4=49.85;106.7=49.48;107.9=49.15;109.2=48.82;110.5=48.52;111.8=48.25;113=48.01;114.3=47.78;115.6=47.56;116.9=47.37;118.2=47.2;119.5=47.05;120.8=46.92;122.1=46.81;123.5=46.71;124.8=46.63;126.1=46.57;127.5=46.53;128.8=46.51;130.1=46.51;131.5=46.52;132.8=46.56;134.2=46.61;135.6=46.69;136.9=46.77;138.3=46.89;139.7=47.03;141.1=47.19;142.4=47.36;143.8=47.56;145.2=47.79;146.6=48.05;148=48.33;149.4=48.63;150.8=48.97;152.3=49.36;153.7=49.76;155.1=50.19;156.5=50.65;158=51.19;159.4=51.72;160.9=52.33;162.3=52.93;163.8=53.59;165.2=54.22;166.7=54.88;168.1=55.45;169.6=55.97;171.1=56.33;172.6=56.5;174=56.45;175.5=56.18;177=55.74;178.5=55.16;180=54.52;181.5=55.16;183=55.74;184.5=56.18;186=56.45;187.4=56.5;188.9=56.33;190.4=55.97;191.9=55.45;193.3=54.88;194.8=54.22;196.2=53.59;197.7=52.93;199.1=52.33;200.6=51.72;202=51.19;203.5=50.65;204.9=50.19;206.3=49.76;207.7=49.36;209.2=48.97;210.6=48.63;212=48.33;213.4=48.05;214.8=47.79;216.2=47.56;217.6=47.36;218.9=47.19;220.3=47.03;221.7=46.89;223.1=46.77;224.4=46.69;225.8=46.61;227.2=46.56;228.5=46.52;229.9=46.51;231.2=46.51;232.5=46.53;233.9=46.57;235.2=46.63;236.5=46.71;237.9=46.81;239.2=46.92;240.5=47.05;241.8=47.2;243.1=47.37;244.4=47.56;245.7=47.78;247=48.01;248.2=48.25;249.5=48.52;250.8=48.82;252.1=49.15;253.3=49.48;254.6=49.85;255.8=50.23;257.1=50.66;258.3=51.08;259.6=51.57;260.8=52.05;262=52.55;263.2=53.06;264.5=53.64;265.7=54.18;266.9=54.71;268.1=55.22;269.3=55.68;270.5=56.06;271.7=56.34;272.8=56.48;274=56.49;275.2=56.36;276.3=56.13;277.5=55.76;278.7=55.32;279.8=54.86;280.9=54.38;282.1=53.85;283.2=53.35;284.3=52.87;285.5=52.36;286.6=51.91;287.7=51.48;288.8=51.07;289.9=50.68;291=50.31;292.1=49.96;293.2=49.63;294.2=49.35;295.3=49.06;296.4=48.79;297.4=48.56;298.5=48.32;299.5=48.12;300.6=47.91;301.6=47.74;302.6=47.57;303.6=47.42;304.7=47.27;305.7=47.15;306.7=47.04;307.7=46.93;308.7=46.84;309.6=46.77;310.6=46.7;311.6=46.64;312.6=46.49;313.5=46.25;314.5=45.98;315.4=45.75;316.3=45.52;317.3=45.27;318.2=45.05;319.1=44.83;320=44.61;320.9=44.4;321.8=44.19;322.7=43.98;323.6=43.78;324.5=43.57;325.4=43.37;326.2=43.19;327.1=43;327.9=42.82;328.8=42.62;329.6=42.45;330.4=42.28;331.2=42.1;332.1=41.9;332.9=41.73;333.6=41.57;334.4=41.4;335.2=41.22;336=41.03;336.7=40.87;337.5=40.68;338.2=40.52;339=40.32;339.7=40.14;340.4=39.96;341.1=39.77;341.8=39.58;342.5=39.38;343.2=39.17;343.9=38.94;344.6=38.71;345.2=38.5;345.9=38.24;346.5=38;347.1=37.74;347.8=37.42;348.4=37.13;349=36.81;349.5=36.52;350.1=36.14;350.7=35.73;351.2=35.36;351.8=34.86;352.3=34.4;352.8=33.89;353.3=33.33;353.8=32.7;354.3=32;354.8=31.21;355.2=30.5;355.7=29.49;356.1=28.57;356.5=27.52;356.9=26.32;357.3=24.93;357.6=23.72;358=12.96;358.3=9.36;358.6=6.35;358.9=3.92;359.2=2.07;359.4=1.17;359.6=0.52;359.8=0.13;359.9=0.03';
        end;
      'ТВ', 'РВ':
        modul := 'ЧМ';
      'Любительское РЭС':
        modul := 'ЧМ'
      else
        begin
        Gain  := '10';
        modul := '';
        end;
      end;
  end;

begin
  Slope    := '0';
  R_KND    := '0';
  R_Diam   := '0';
  R_rask   := '0';
  Gain     := '10';
  DNA_vert := '';
  DNA_horz := '';
  modul    := '';
  antModel := '';

  SetLength(mOk, Length(mOk) + 1);
  mOk[High(mOk)] := nRow;

  vlad := exlCellDataTry(sh.Cell[iPa.vlad_An.exl, nRow]);
  n    := f_Tparam.RowCount;
  f_Tparam.RowCount := n + 1;
  f_Tparam.Cells[iPa.vkl_An.n, n] := '-1';
  f_Tparam.Cells[iPa.nnExcel_No.n, n] :=
    exlCellDataTry(sh.Cell[iPa.nnExcel_No.exl, nRow]);
  f_Tparam.Cells[iPa.BS_No.n, n] := exlCellDataTry(sh.Cell[iPa.BS_No.exl, nRow]);
  f_Tparam.Cells[iPa.vlad_An.n, n] := vlad;
  f_Tparam.Cells[iPa.vidRES_No.n, n] := exlCellDataTry(sh.Cell[iPa.vidRES_No.exl, nRow]);
  f_Tparam.Cells[iPa.tech_Se.n, n] :=
    technologiya(exlCellDataTry(sh.Cell[iPa.f_Tr.exl, nRow]),
    exlCellDataTry(sh.Cell[iPa.vidRES_No.exl, nRow]));
  f_Tparam.Cells[iPa.f_Tr.n, n] :=
    GetDateFrec(exlCellDataTry(sh.Cell[iPa.f_Tr.exl, nRow]), 1);
  f_Tparam.Cells[iPa.P_Tr.n, n] := exlCellDataTry(sh.Cell[iPa.P_Tr.exl, nRow]);
  f_Tparam.Cells[iPa.azim_An.n, n] := exlCellDataTry(sh.Cell[iPa.azim_An.exl, nRow]);
  f_Tparam.Cells[iPa.h_An.n, n] := exlCellDataTry(sh.Cell[iPa.h_An.exl, nRow]);
  f_Tparam.Cells[iPa.kolPrd_An_Co.n, n] := '1';
  f_Tparam.Cells[iPa.svid_No.n, n] := exlCellDataTry(sh.Cell[iPa.svid_No.exl, nRow]);
  f_Tparam.Cells[iPa.svidDO_No.n, n] :=
    exlCellDataTry(sh.Cell[iPa.svidDO_No.exl, nRow]);
  f_Tparam.Cells[iPa.prd_An.n, n] := exlCellDataTry(sh.Cell[iPa.prd_An.exl, nRow]);
  f_Tparam.Cells[iPa.adr_Si.n, n] := exlCellDataTry(sh.Cell[iPa.adr_Si.exl, nRow]);
  f_Tparam.Cells[iPa.PDUparam_Se.n, n] :=
    f_Tparam.Columns.Items[iPa.PDUparam_Se.n].PickList.Strings[StrToIntDef(
    PDU_Type(f_Tparam.Cells[iPa.f_Tr.n, n]), 1)];
  f_Tparam.Cells[iPa.PDUznach_Se.n, n] := PDU_znach(f_Tparam.Cells[iPa.f_Tr.n, n]);
  f_Tparam.Cells[iPa.color_Se.n, n] := colorOper(vlad);

  dopRasch(f_Tparam.Cells[iPa.f_Tr.n, n], f_Tparam.Cells[iPa.vidRES_No.n, n]);
  f_Tparam.Cells[iPa.K_An.n, n]     := Gain;
  f_Tparam.Cells[iPa.TiltM_An.n, n] := Slope;
  f_Tparam.Cells[iPa.modu_An.n, n]  := modul;
  f_Tparam.Cells[iPa.KND_An.n, n]   := R_KND;
  f_Tparam.Cells[iPa.KND_D_An.n, n] := R_Diam;
  f_Tparam.Cells[iPa.KND_a_An.n, n] := R_rask;
  f_Tparam.Cells[iPa.Model_An.n, n] := antModel;
  f_Tparam.Cells[iPa.DNAhorz_DNA.n, n] := DNA_horz;
  f_Tparam.Cells[iPa.DNAvert_DNA.n, n] := DNA_vert;
end;


//определить принадлежность к технологии
function TFVizSzz.technologiya(f, vidRES: string): string;
begin
  if Pos('GSM', vidRES) > 0 then
    Result := 'GSM'
  else if Pos('UMTS', vidRES) > 0 then
      Result := 'UMTS'
    else if Pos('LTE', vidRES) > 0 then
        Result := 'LTE'
      else if Pos('транкингов', vidRES) > 0 then
          Result := 'Транкинговая сеть'
        else if Pos('технологического', vidRES) > 0 then
            Result := 'Технологическая сеть'
          else if Pos('радиорелейная', vidRES) > 0 then
              Result := 'РРЛ'
            else if (Pos('звукового радиовещания', vidRES) > 0) then
                Result := 'РВ'
              else if (Pos('телевизионного', vidRES) > 0) then
                  Result := 'ТВ'
                else if (Pos('беспроводного доступа',
                    vidRES) > 0) then
                    Result := 'Wi-Fi'
                  else if (Pos('земная станция', vidRES) > 0) then //VSAT
                      Result := 'ЗССС'
                    else if (Pos('9 кГц-30МГц', vidRES) > 0) then
                        Result := 'КВ'
                      else if (Pos('любительское', vidRES) > 0) then
                          Result := 'Любительское РЭС'
                        else if Pos('фиксированной', vidRES) > 0 then
                            Result := 'Фиксированной службы'
                          else
                            Result := GetDateFrec(f, 0);
end;

 //Получить данные из таблицы частот diapazon
 // по F, рез. col - номер колонки из инспектора
function TFVizSzz.GetDateFrec(const f: string; col: byte): string;
var
  i: integer;
  n: integer;
begin
  Result := '';
  n      := toInt([f]);
  if (not f.IsEmpty) then
    for i := f_frec.RowCount - 1 downto 1 do
      if f_frec.Cells[4 + 1, i] = '1' then
        if (f_frec.Cells[2 + 1, i].ToDouble <= n) and
          (n <= f_frec.Cells[3 + 1, i].ToDouble) then
          begin
          Result := Trim(f_frec.Cells[col + 1, i]);
          Break;
          end;
  if Result.IsEmpty then
    Result := IntToStr(toInt([f], 0));
end;

procedure TFVizSzz.f_xmlSaveClick(Sender: TObject);
begin
  SaveXMLBool := False;
  //в файле были данные...
  if f_xmlSave.Visible then
    if paramFile.IsEmpty then
      //запуск ПО без параметров
      saveXML(OpenDialog1.FileName)
    else
      //ПО было запущено с параметрами
      begin
      SaveDialog1.FileName   := ExtractFileNameOnly(paramFile);
      SaveDialog1.InitialDir :=
        GetEnvironmentVariable('USERPROFILE') + '\Documents';
      if SaveDialog1.Execute then
        saveXML(SaveDialog1.FileName);
      end
  else
    MessageDlg('Внимание!', 'Нет данных для сохранения.',
      mtError, [mbOK], 0);
end;

procedure TFVizSzz.MPDiap_delClick(Sender: TObject);
begin
  f_frec.DeleteRow(f_frec.Row);
end;

procedure TFVizSzz.MPOne_AutoColClick(Sender: TObject);
begin
  if MPOne_AutoCol.Checked then
    f_Tparam.AutoFillColumns := True
  else
    f_Tparam.AutoFillColumns := False;
end;

procedure TFVizSzz.MPOne_DelClick(Sender: TObject);
begin
  f_Tparam.DeleteRow(f_Tparam.Row);
end;

procedure TFVizSzz.MPOne_SaveClick(Sender: TObject);
var
  i, j: integer;
  f:    TextFile;
  s:    string;
begin
  progressSet(f_Tparam.RowCount - 1);

  if OpenDialog1.FileName.IsEmpty then
    OpenDialog1.FileName := GetEnvironmentVariable('USERPROFILE') +
      '\Documents\VizSzz.csv';


  s := ExtractFileNameWithoutExt(OpenDialog1.FileName) + '.csv';
  stat.log('Попытка сохранить в файл: ' + s);
  AssignFile(f, s);
    try
    f_Tparam.BeginUpdate;
    Rewrite(f);
    CloseFile(f);
    Append(f);

    for j := 0 to f_Tparam.RowCount - 1 do
      begin
      for i := 0 to f_Tparam.ColCount - 1 do
        if (j > 0) and (i = 0) then
          Write(f, j.ToString + ';')
        else
          Write(f, UTF8ToCP1251(f_Tparam.Cells[i, j], False) + ';');
      Write(f, nR);
      progres;
      end;
    CloseFile(f);
    f_Tparam.EndUpdate;
    stat.log('Файл сохранён: ' + s);
    except
    progressEnd;
    stat.Error('Не удалось сохранить:' + s);
    end;
  progressEnd;
end;

procedure TFVizSzz.RollOutLicCollapse(Sender: TObject);
begin
  //RollOutLic.Width := 172;
end;

procedure TFVizSzz.RollOutLicPreExpand(Sender: TObject);
begin
  // RollOutLic.Width:=Width;
  //RollOutLic.Height := Height;
end;

//Определить цвет для владельца
function TFVizSzz.colorOper(vlavelec: string): string;
begin
  if Pos('МегаФон', vlavelec) > 0 then
    Result := '32768'
  else
    if Pos('ВымпелКом', vlavelec) > 0 then
      Result := '59624'
    else
      if Pos('МТС', vlavelec) > 0 then
        Result := '255'
      else
        if Pos('Мобайл', vlavelec) > 0 then
          Result := '0001'
        else
          Result := '8388608';
end;

//ПДУ тип
function TFVizSzz.PDU_Type(f: string): string;
var
  fr: integer;
begin
  fr := toInt([f], 0) * 100;
  case fr of
    30001..30000000:
      Result := '1'
    else
      Result := '0';
    end;
end;

//ПДУ значение
function TFVizSzz.PDU_znach(f: string): string;
var
  fr: integer;
begin
  fr := toInt([f], 0) * 100;
  case fr of
    3..30:
      Result := '25';
    31..300:
      Result := '15';
    301..3000:
      Result := '10';
    3001..30000:
      Result := '3'
    else
      Result := '10';
    end;
end;

//Сохранить данные в xml Файл
procedure TFVizSzz.saveXML(nameXML: string);
var
  s, ss, koor, fNew: string;
  i:  integer;
  root, nod1, nod2, nod3, nod4, nod5: TDOMElement;
  sl: TStringList;
  b:  boolean;
  f, fTmp: TextFile;// file of string[36];
begin
  //nameXML:= ValidStr(DateToStr(Date, myFormatDate));
  if nameXML.IsEmpty then
    nameXML := 'ONEPLAN Sazon.xml';

  if not Assigned(docXML) then
    docXML := TXMLDocument.Create;

  docXML.SetHeaderData(xmlVersion10, 'windows-1251');
  //docXML.StylesheetType := 'text/xsl'; //инача две строки записывает
  docXML.StylesheetHRef := '';
  //для laz2_
  ////docXML.XMLVersion := '1.0';
  ////docXML.Encoding := 'windows-1251';
  ////DOMt:=docXML.CreateTextNode('<?xml:stylesheet type="text/xsl" href=""?>');

  stat.Log('xml открыт');
  stat.Log('Подготовка к сохранению данных в "' +
    nameXML + '"');
  root := docXML.CreateElement('Document');
  //docXML.AppendChild(root);
  root := docXML.DocumentElement; // назначить текущим

  stat.Log('1. Сохранение параметров проекта...');
  nod1 := docXML.CreateElement('SITES');
  root.AppendChild(nod1);

  stat.Log('Сохранение писания проекта...');
  nod2 := docXML.CreateElement('site');
  for i := 1 to f_Topis.RowCount - 1 do
    AtrWin(nod2, f_Topis.Cells[1, i], f_Topis.Cells[3, i]);
  nod1.AppendChild(nod2);

  stat.Log('Запись секторов...');
  //if f_Tparam.RowCount >1 then
  for i := 1 to f_Tparam.RowCount - 1 do
    begin
    nod3 := docXML.CreateElement('SECTOR');
    AtrWin(nod3, 'id', i.ToString);
    AtrWin(nod3, 'name', 'cell_' + i.ToString);
    AtrWin(nod3, iPa.tech_Se.x, f_Tparam.Cells[iPa.tech_Se.n, i]);
    AtrWin(nod3, iPa.PDUparam_Se.x,
      IntToStr(f_Tparam.Columns[iPa.PDUparam_Se.n].PickList.IndexOf(
      f_Tparam.Cells[iPa.PDUparam_Se.n, i])));
    //Cells[iPa.PDUparam_Se.n, i]);
    AtrWin(nod3, iPa.PDUznach_Se.x, f_Tparam.Cells[iPa.PDUznach_Se.n, i]);
    AtrWin(nod3, iPa.color_Se.x, f_Tparam.Cells[iPa.color_Se.n, i]);

    nod4 := docXML.CreateElement('TRX');
    AtrWin(nod4, iPa.P_Tr.x, f_Tparam.Cells[iPa.P_Tr.n, i], '0');
    AtrWin(nod4, iPa.f_Tr.x, f_Tparam.Cells[iPa.f_Tr.n, i], '0');
    nod3.AppendChild(nod4);

    nod4 := docXML.CreateElement('combiner');
    AtrWin(nod4, iPa.bPassiv_Co.x, f_Tparam.Cells[iPa.bPassiv_Co.n, i], '0');
    AtrWin(nod4, iPa.CombinerType_Co.x, f_Tparam.Cells[iPa.CombinerType_Co.n, i], '0');
    //AtrWin(nod4, iPa.typeAnt_An_Co.x, f_Tparam.Cells[iPa.typeAnt_An_Co.n, i]);
    AtrWin(nod4, iPa.kolPrd_An_Co.x, f_Tparam.Cells[iPa.kolPrd_An_Co.n, i]);
    nod3.AppendChild(nod4);

    nod4 := docXML.CreateElement('ANTENNA');
    AtrWin(nod4, 'id', i.ToString);
    AtrWin(nod4, iPa.prd_An.x, f_Tparam.Cells[iPa.prd_An.n, i]);
    AtrWin(nod4, iPa.azim_An.x, f_Tparam.Cells[iPa.azim_An.n, i]);
    AtrWin(nod4, iPa.h_An.x, f_Tparam.Cells[iPa.h_An.n, i]);
    AtrWin(nod4, iPa.TiltM_An.x, f_Tparam.Cells[iPa.TiltM_An.n, i], '0');
    AtrWin(nod4, iPa.K_An.x, f_Tparam.Cells[iPa.K_An.n, i]);
    AtrWin(nod4, iPa.bLen_An.x, f_Tparam.Cells[iPa.bLen_An.n, i], '0');
    AtrWin(nod4, iPa.bPogon_An.x, f_Tparam.Cells[iPa.bPogon_An.n, i], '0');
    AtrWin(nod4, iPa.b_An.x, f_Tparam.Cells[iPa.b_An.n, i], '0');
    AtrWin(nod4, iPa.vkl_An.x, f_Tparam.Cells[iPa.vkl_An.n, i]);
    if not Trim(f_Tparam.Cells[iPa.koordN_An.n, i]).IsEmpty then
      AtrWin(nod4, iPa.koordN_An.x, f_Tparam.Cells[iPa.koordN_An.n, i])
    else
      AtrWin(nod4, iPa.koordN_An.x, f_Topis.Cells[3, 12]);
    if not Trim(f_Tparam.Cells[iPa.koordE_An.n, i]).IsEmpty then
      AtrWin(nod4, iPa.koordE_An.x, f_Tparam.Cells[iPa.koordE_An.n, i])
    else
      AtrWin(nod4, iPa.koordE_An.x, f_Topis.Cells[3, 13]);
    //AtrWin(nod4, iPa.typeAnt_An_Co.x, f_Tparam.Cells[iPa.typeAnt_An_Co.n, i]);
    AtrWin(nod4, iPa.kolPrd_An_Co.x, f_Tparam.Cells[iPa.kolPrd_An_Co.n, i]);
    AtrWin(nod4, iPa.CombinerType_An.x, f_Tparam.Cells[iPa.CombinerType_An.n, i], '0');
    AtrWin(nod4, iPa.calcType_An.x, f_Tparam.Cells[iPa.calcType_An.n, i], '0');
    AtrWin(nod4, iPa.secZaprAzim_An.x, f_Tparam.Cells[iPa.secZaprAzim_An.n, i], '0');
    AtrWin(nod4, iPa.secZaprShir_An.x, f_Tparam.Cells[iPa.secZaprShir_An.n, i], '0');
    AtrWin(nod4, iPa.typeApert_An.x, f_Tparam.Cells[iPa.typeApert_An.n, i], '0');
    AtrWin(nod4, iPa.KND_An.x, f_Tparam.Cells[iPa.KND_An.n, i], '0');
    AtrWin(nod4, iPa.KND_a_An.x, f_Tparam.Cells[iPa.KND_a_An.n, i], '0');
    AtrWin(nod4, iPa.KND_D_An.x, f_Tparam.Cells[iPa.KND_D_An.n, i], '0');
    AtrWin(nod4, iPa.ugRasGor_An.x, f_Tparam.Cells[iPa.ugRasGor_An.n, i], '0');
    AtrWin(nod4, iPa.ugRasVer_An.x, f_Tparam.Cells[iPa.ugRasVer_An.n, i], '0');
    AtrWin(nod4, iPa.storApertGor_An.x, f_Tparam.Cells[iPa.storApertGor_An.n, i], '0');
    AtrWin(nod4, iPa.storApertVert_An.x, f_Tparam.Cells[iPa.storApertVert_An.n, i], '0');
    AtrWin(nod4, iPa.impF_An.x, f_Tparam.Cells[iPa.impF_An.n, i], '0');
    AtrWin(nod4, iPa.impT_An.x, f_Tparam.Cells[iPa.impT_An.n, i], '0');
    AtrWin(nod4, iPa.impP_An.x, f_Tparam.Cells[iPa.impP_An.n, i], '0');
    AtrWin(nod4, iPa.Model_An.x, f_Tparam.Cells[iPa.Model_An.n, i]);
    AtrWin(nod4, iPa.Model_id_An.x, f_Tparam.Cells[iPa.Model_id_An.n, i], '0');
    AtrWin(nod4, iPa.pol_An.x, f_Tparam.Cells[iPa.pol_An.n, i], '0');
    AtrWin(nod4, iPa.modu_An.x, f_Tparam.Cells[iPa.modu_An.n, i]);
    AtrWin(nod4, iPa.vlad_An.x, f_Tparam.Cells[iPa.BS_No.n, i] +
      '; ' + f_Tparam.Cells[iPa.vlad_An.n, i]);
    nod3.AppendChild(nod4);

    nod5 := docXML.CreateElement('DNA_vert');
    AtrWin(nod5, iPa.DNAvert_DNA.x, f_Tparam.Cells[iPa.DNAvert_DNA.n, i]);
    nod4.AppendChild(nod5);
    nod5 := docXML.CreateElement('DNA_horz');
    AtrWin(nod5, iPa.DNAhorz_DNA.x, f_Tparam.Cells[iPa.DNAhorz_DNA.n, i]);
    nod4.AppendChild(nod5);

    nod2.AppendChild(nod3);
    stat.Log('Сектор № ' + i.ToString + ' записан');
    end;

  stat.Log('****************************************************');
  stat.Log('2. Сохранение СЗЗ, КТ, зданий проекта...');
  nod1 := docXML.CreateElement('SAZON_PARAMS');
  AtrWin(nod1, 'step', '1');
  root.AppendChild(nod1);

  stat.Log('2.1. Сохранение СЗЗ и ЗОЗ проекта...    **************');
  nod2 := docXML.CreateElement('LEVELS');
  //AtrWin(nod2, f_Tzoz.Cells[1, i], f_Topis.Cells[3, i]);
  if f_Tzoz.RowCount > 1 then
    for i := 1 to f_Tzoz.RowCount - 1 do
      begin
      nod3 := docXML.CreateElement('item');
      AtrWin(nod3, 'checked', f_Tzoz.Cells[0 + 1, i]);
      AtrWin(nod3, 'height', f_Tzoz.Cells[1 + 1, i]);
      AtrWin(nod3, 'color', f_Tzoz.Cells[2 + 1, i]);
      AtrWin(nod3, 'comment', f_Tzoz.Cells[3 + 1, i]);
      nod2.AppendChild(nod3);
      end;
  nod1.AppendChild(nod2);

  stat.Log('2.2. Сохранение КТ проекта...          **************');
  nod2 := docXML.CreateElement('CONTROL_POINTS');
  if f_Tkt.RowCount > 1 then
    for i := 1 to f_Tkt.RowCount - 1 do
      begin
      nod3 := docXML.CreateElement('item');
      AtrWin(nod3, 'checked', f_Tkt.Cells[0 + 1, i]);
      AtrWin(nod3, 'name', f_Tkt.Cells[1 + 1, i]);
      AtrWin(nod3, 'height', f_Tkt.Cells[2 + 1, i]);
      AtrWin(nod3, 'measure', f_Tkt.Cells[7 + 1, i]);
      AtrWin(nod3, 'm_type', IntToStr(f_Tkt.Columns[5].PickList.IndexOf(
        f_Tkt.Cells[5 + 1, i]))); //f_Tkt.Cells[7 + 1, i]);
      AtrWin(nod3, 'lat', f_Tkt.Cells[3 + 1, i]);
      AtrWin(nod3, 'lon', f_Tkt.Cells[4 + 1, i]);
      AtrWin(nod3, 'comment', f_Tkt.Cells[6 + 1, i]);
      nod2.AppendChild(nod3);
      end;
  nod1.AppendChild(nod2);

  stat.Log('2.3. Сохранение Надписей проекта...      **************');
  nod2 := docXML.CreateElement('MAP_CAPTIONS');
  if f_Tnad.RowCount > 1 then
    for i := 1 to f_Tnad.RowCount - 1 do
      begin
      nod3 := docXML.CreateElement('item');
      AtrWin(nod3, 'checked', f_Tnad.Cells[0 + 1, i]);
      AtrWin(nod3, 'caption', f_Tnad.Cells[1 + 1, i]);
      AtrWin(nod3, 'lat', f_Tnad.Cells[2 + 1, i]);
      AtrWin(nod3, 'lon', f_Tnad.Cells[3 + 1, i]);
      AtrWin(nod3, 'comment', f_Tnad.Cells[4 + 1, i]);
      nod2.AppendChild(nod3);
      end;
  nod1.AppendChild(nod2);

  stat.Log('2.4. Сохранение Зданий, дорог... проекта...      **************');
  nod2 := docXML.CreateElement('MAP_POLYGONS');
  if f_Tzdan.RowCount > 1 then
    for i := 1 to f_Tzdan.RowCount - 1 do
      begin
      nod3 := docXML.CreateElement('item');
      AtrWin(nod3, iZd.vkl.x, f_Tzdan.Cells[iZd.vkl.n, i]);
      AtrWin(nod3, iZd.cap.x, f_Tzdan.Cells[iZd.cap.n, i]);
      AtrWin(nod3, iZd.typ.x,
        IntToStr(f_Tzdan.Columns[iZd.typ.n - 1].PickList.IndexOf(
        f_Tzdan.Cells[iZd.typ.n, i]))); //f_Tzdan.Cells[iZd.typ.n, i]);
      AtrWin(nod3, iZd.h.x, f_Tzdan.Cells[iZd.h.n, i]);
      AtrWin(nod3, iZd.otrajBool.x, f_Tzdan.Cells[iZd.otrajBool.n, i]);
      AtrWin(nod3, iZd.otrajKoef.x, f_Tzdan.Cells[iZd.otrajKoef.n, i]);
      AtrWin(nod3, iZd.color.x, f_Tzdan.Cells[iZd.color.n, i]);
      AtrWin(nod3, iZd.LossBool.x, f_Tzdan.Cells[iZd.LossBool.n, i]);
      AtrWin(nod3, iZd.LossZnach.x, f_Tzdan.Cells[iZd.LossZnach.n, i]);
      AtrWin(nod3, iZd.Loss_m.x, f_Tzdan.Cells[iZd.Loss_m.n, i]);
      AtrWin(nod3, iZd.LossP2346.x, f_Tzdan.Cells[iZd.LossP2346.n, i]);
      AtrWin(nod3, iZd.koment.x, f_Tzdan.Cells[iZd.koment.n, i]);
      nod2.AppendChild(nod3);

      {проверить: м/б запись в условие занести}
      //разбиваем строку по ';' и записываем атрибуты
      sl := TStringList.Create;
      sl.AddDelimitedText(f_Tzdan.Cells[iZd.tochki.n, i], ';', True);
      b := False;
      for koor in sl do
        begin
        b := not b;
        if b then
          nod4 := docXML.CreateElement('point')
        else
          nod3.AppendChild(nod4);
        AtrWin(nod4, iZd.tochki.x, koor);
        end;
      sl.Free;
      end;
  nod1.AppendChild(nod2);

  stat.Log('Сохранение xml файла...');
  ss := ExtractFileNameWithoutExt(nameXML) + '~0~-.xml';
  xmlSaves(ss);//Зараза, сохраняется только в UTF8

  fNew := ExtractFileNameWithoutExt(nameXML) + '+.xml';
  stat.Log('Пересохранение в коировке windows-1251');
  AssignFile(f, fNew);
  Rewrite(f);
  AssignFile(fTmp, ss);
  Reset(fTmp);
    try
    WriteLn(f, '<?xml version="1.0" encoding="windows-1251"?>');
    while not EOF(fTmp) do
      begin
      ReadLn(fTmp, s);
      WriteLn(f, UTF8ToCP1251(s));
      end;
    SaveXMLBool := True;
    finally
    CloseFile(f);
    CloseFile(fTmp);
    end;

  //в случае чего, временный файл в UTF8 останентся на диске
  if SaveXMLBool then
    DeleteFile(ss);

  stat.Log('Файл ' + fNew + ' сохранён ! ! !');
end;

procedure TFVizSzz.saveVisible;
begin
  if f_Tparam.RowCount > 1 then
    begin
    f_xmlSave.Visible := True;
    f_xmlSave.SetFocus;
    end
  else
    f_xmlSave.Visible := False;
  progressEnd;
end;

//Установить начальные значения и отобразить
procedure TFVizSzz.progressSet(maxProgress: integer);
begin
  ProgressBar1.Max      := maxProgress;
  ProgressBar1.Position := 0;
  ProgressBar1.Visible  := True;
  Application.ProcessMessages;
end;

procedure TFVizSzz.progres;
begin
  ProgressBar1.StepIt;
  ProgressBar1.Update;
  Application.ProcessMessages;
end;

procedure TFVizSzz.progressEnd;
begin
  ProgressBar1.Visible := False;
  Application.ProcessMessages;
end;

end.
