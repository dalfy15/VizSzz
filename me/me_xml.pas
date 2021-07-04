 // ************************************************
 // Сокращение рутиной работы. Routines for formulas.
 // Author:        Андрей А. Кандауров (Andrey A. Kandaurov)
 // Companiya:     Santig
 // e-mail:        san@santig.ru
 // URL:           http://santig.ru
 // License:       zlib
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

(*
* (перекоментировать encoding в файле XMLWrite)
*
* Cвойства FirstChild и NextSibling - чтобы шагать вперед по дереву
*           LastChild и PreviousSibling - назад с конца дерева
* FindNode - ищется первый узел верхнего уровня с подходящим именем
*    GetElementsByTagName - создается объект TDOMNodeList,
*            который после использования должен быть освобождён
*  GetNamedItem - для поиска аттрибута по имени в текущем узле.
*)
unit me_xml;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, Dialogs, LazUTF8, LConvEncoding, LazFileUtils,
  Forms,
  DOM, XMLWrite, xmlutils, XMLRead, xmliconv_windows,
  me_my;

// открыть xml файл через диалог
function xmlOpensFile(Sender: TObject;
  filter: string = 'xml файлы|*.xml|все файлы|*.*'): string;
// Загружаем XML в переменную для обработки. True - загрузился
function xmlOpens(file_name: string): boolean;
// Сохраняем XML в файл. True - сохранился
function xmlSaves(file_name: string): boolean;
//Чтение атрибута, если его нет, то возвращает noAttr. Если result=aValue, то noAttr
function atrRead(const aNode: TDOMNode; Attr: string; noAttr: string = '';
  aValue: string = '~,:/'): string;
//установить атрибут в системной кодировке
procedure AtrWin(var NOD: TDOMElement; const XMLteg, Value: string; empt: string = '');
//Поиск в дочерних узлах узла aNod узла с атрибутом Atribut = Value
function findZnachAtr(const uzel: TDOMNode; Atribut, aValue: string): TDOMNode;


var
  docXML: TXMLDocument = nil;

implementation

var
  OpenDialogXML: TOpenDialog;

// открыть xml файл через диалог
function xmlOpensFile(Sender: TObject;
  filter: string = 'xml файлы|*.xml|все файлы|*.*'): string;
begin
  Result := '';
  OpenDialogXML := TOpenDialog.Create(nil);
  OpenDialogXML.Title := 'Открыть xml-файл';
  OpenDialogXML.DefaultExt := '.xml';
  OpenDialogXML.Filter := filter;//'kml файлы|*.kml';
  if OpenDialogXML.Execute then
    if xmlOpens(OpenDialogXML.FileName) then
      Result := OpenDialogXML.FileName;
  OpenDialogXML.Free;
end;

// Загружаем XML в переменную для обработки. True - загрузился
function xmlOpens(file_name: string): boolean;
begin
  Result := False;
  //docXML := nil; //TXMLDocument.Create; уже содержится в ReadXMLFile
  if FileExistsUTF8(file_name) then
    if Assigned(docXML) then
      FreeAndNil(docXML);
    try
    ReadXMLFile(docXML, UTF8ToSys(file_name));
    Result := True;
    except
    //log.Error('Ошибка открытия файла. (function xmlOpens)');
    docXML.Free;
    end;
end;

// Сохраняем XML в файл. True - сохранился
function xmlSaves(file_name: string): boolean;
begin
  if docXML = nil then
    Exit(False);

    try
    WriteXMLFile(docXML, file_name);
    Result := True;
    except
    Result := False;
    end;
end;

//Чтение атрибута, если его нет, то возвращает noAttr. Если result=aValue, то noAttr
function atrRead(const aNode: TDOMNode; Attr: string; noAttr: string = '';
  aValue: string = '~,:/'): string;
var
  AttrNode: TDOMNode = nil;
begin
  Result := noAttr;
    try
    //_ms([aNode.NodeName]);
    //в текущем узле aNode пытаемся найти аттрибут Attr
    if Assigned(aNode) then
      if aNode.HasAttributes then
        AttrNode := aNode.Attributes.GetNamedItem(DOMString(SysToUTF8(Attr)));
    //если такой узел создан (и не равен nil)
    if Assigned(AttrNode) then
      if AttrNode.NodeValue <> '' then
        begin
        Result := UTF8ToSys(AttrNode.NodeValue); //UTF8ToCP1251
        if Result = aValue then
          Result := noAttr;
        end
      else
        Result := '';
    except
    Result := noAttr;
    end;
end;

//установить атрибут в системной кодировке
procedure AtrWin(var NOD: TDOMElement; const XMLteg, Value: string; empt: string = '');
var
  s: string;
begin
  s := Value;
  if s.IsEmpty then
    s := empt;
  TDOMElement(nod).SetAttribute(DOMString(XMLteg), DOMString(UTF8ToSys(s)));
end;

//Поиск в дочерних узлах узла aNod узла с атрибутом Atribut = Value
function findZnachAtr(const uzel: TDOMNode; Atribut, aValue: string): TDOMNode;
var
  nod: TDOMNode;
  aValueNew: string;
begin
  Result := nil;
  //если не сможет прочитать атрибут, то при сравнении с пустым атрибутом вернёт nil
  aValueNew := aValue + '1';
  //Если есть дочерние узлы
  nod := uzel.FirstChild;
  while assigned(nod) do
    begin
    if atrRead(nod, Atribut, aValueNew) = aValue then
      begin
      Result := nod;
      //_ms(['"findZnachAtr"', uzel.NodeName + ' (uzel)', nod.NodeName +' (aValue - ' + aValue +')']);
      Break;
      end;
    nod := nod.NextSibling;
    end;
end;

end.
