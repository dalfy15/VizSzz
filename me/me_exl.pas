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
 unit me_exl;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils,
  me_my, zexlsx, zexmlss, zeSave;

//Если ячеёка не существует, то выдать err
function exlCellDataTry(const cel: TZCell; err: string = ''): string;


implementation

//Если ячеёка не существует, то выдать err
function exlCellDataTry(const cel: TZCell; err: string = ''): string;
begin
  if Assigned(cel) then
    Result := cel.Data
  else
    Result := err;
end;


end.

