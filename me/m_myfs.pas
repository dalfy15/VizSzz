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
{************************************************
 Модуль с дополнительными компонентами:
    * Логирование ПО;
    * Работа со статусом окна.

 //Logs Инициализация в основной программе (один раз), тамже разрушается:
 Добавиь в проект форму myFs из me\m_myFs и в самом проекте lpr поставить её
 до загрузки основной формы (Настройка формы: FormStyle=fsSplash;):
    myFs:=TmyFs.Create(nil); myFs.Hide;
 После, уничтожить форму в OnDestroy в главной форме:   myFs.Free;

// можно не использовать параметры для Лога:
//myFs.Log.FileName:='Logs.log';
//myFs.Log.Identification:='-';  - устанавливается в stat.razdel := 'ПСА';
//myFs.Log.LogType:=ltFile;     //запись в файл
//myFs.Log.AppendContent:=True; //дописывание (умолч: False)
//myFs.Log.Active := True;      - активация происходит при создании формы
}
unit m_myFs;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, eventlog, Forms, Controls, Graphics, Dialogs,
  IniPropStorage, StdCtrls, ComCtrls, ExtCtrls, FileUtil, LazFileUtils,
  LConvEncoding,
  me_my;

type
  LogTypeFile = eventlog.TLogType;

  { TmyFs }

  TmyFs = class(TForm)
    IniF: TIniPropStorage;
    Log:  TEventLog;
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
  private
    //прочитать имя файла лога
    function logFileNameRead: string;
    //присвоить логу имя файла
    procedure logFileNameWrite(logFileNames: string);
    //Записать данные в ini
    procedure iniWriteStri(section, Ident, Value: string);
    //Прочитать строку из ini
    function iniReadStr(section, Ident, DefValue: string): string;
  public
    //скинуть лог к ехе в папку log
    procedure logCopyLog(log_file: string; log_dir: string = 'log');
    //Оставить в папке log не более n log-файлов
    procedure logDelN(n: integer = 10; log_dir: string = 'log');
    // Удалить лог. Если vkl=bool, то по окончанию включить лог
    procedure logDelete(vkl: boolean = False);
    // Удалить ini
    procedure iniDelete;
    //имя фала логa
    property logFileName: string read logFileNameRead write logFileNameWrite;
    //Прочитать/записать строку в ini
    //property iniStrng: string read iniReadStr write iniWriteStri;
  end;

 var
  myFs: TmyFs;


implementation

{$R *.lfm}

procedure TmyFs.FormCreate(Sender: TObject);
begin
  //nameLog := 'Logs' + Time.CurrentSec100OfDay.ToString + '.log';
  Log.Identification := 'Лог';
  Log.FileName := me_my.ValidStr(AnsiToUtf8('logs ' + DateTimeToStr(Now) + '.log'), [viFile]);
  log.Active := True;
end;


procedure TmyFs.FormDestroy(Sender: TObject);
begin
  log.Active := False;
  DeleteFile(Log.FileName);
end;

function TmyFs.logFileNameRead: string;
begin
  Result := Log.FileName;
end;

procedure TmyFs.logFileNameWrite(logFileNames: string);
var
  b: boolean;
begin
  b := log.Active;
  log.Active := False;
  Log.FileName := logFileNames;
  Log.LogType := LogTypeFile.ltFile;
  log.Active := b;
end;

//скинуть лог (log_file) к ехе в папку log_dir
procedure TmyFs.logCopyLog(log_file: string; log_dir: string = 'log');
var
  s: string;
begin
  if not FileExists(log_file) then
    exit;

  log.Active := False;
  //остановить лог если надо получить доступ к файлу
  s := ExtractFilePath(Application.ExeName) + log_dir + '\';
    try
    //CreateDir(s);
    s += me_my.ValidStr('logs ' + DateTimeToStr(Now) + '.log', [viFile]);
    CopyFile(log_file, s, [cffOverwriteFile, cffCreateDestDirectory,
      cffPreserveTime]);

    //после копирования - удаляем
    DeleteFile(PChar(log_file));
    except;
    end;
end;

procedure TmyFs.logDelN(n: integer = 10; log_dir: string = 'log');
var
  sl: TStringList;
  i:  integer;
begin
  if log_dir = 'log' then
    log_dir := ExtractFilePath(Application.ExeName) + log_dir;

  sl := TStringList.Create;
  FindAllFiles(sl, log_dir);
  if Assigned(sl) then
    if sl.Count >= n then
      for i := 0 to n - 1 do
        DeleteFile(PChar(sl[i]));
  sl.Free;
end;

// Удалить лог. Если vkl=bool, то по окончанию включить лог
procedure TmyFs.logDelete(vkl: boolean = False);
begin
  log.Active := False;
  DeleteFileUTF8(Log.FileName);
  if vkl then
    log.Active := True;
end;

// Удалить ini
procedure TmyFs.iniDelete;
begin
  log.Log('Удаляется файл настроек ini: ' + IniF.IniFileName);
  DeleteFileUTF8(IniF.IniFileName);
end;

//Записать данные в ini
procedure TmyFs.iniWriteStri(section, Ident, Value: string);
var
  sec: string;
begin
  sec := IniF.IniSection;
  IniF.IniSection := section;
  IniF.WriteString(Ident, Value);
  IniF.IniSection := sec;
end;

//Прочитать строку из ini
function TmyFs.iniReadStr(section, Ident, DefValue: string): string;
var
  sec: string;
begin
  sec    := IniF.IniSection;
  IniF.IniSection := section;
  Result := IniF.ReadString(Ident, DefValue);
  IniF.IniSection := sec;
end;

end.
