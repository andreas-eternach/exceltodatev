//This file is part of exceltodatev.

//  exceltodatev is free software: you can redistribute it and/or modify
//  it under the terms of the GNU General Public License as published by
//  the Free Software Foundation, either version 3 of the License, or
//  (at your option) any later version.
//
//  Foobar is distributed in the hope that it will be useful,
//  but WITHOUT ANY WARRANTY; without even the implied warranty of
//  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
//  GNU General Public License for more details.
//
//  You should have received a copy of the GNU General Public License
//  along with Foobar.  If not, see <http://www.gnu.org/licenses/>.
//
//(c) 1996-2009 Andreas Eternach (andreas.eternach@google.com)
program excproj;

uses
  Forms,
  excel in 'excel.pas' {Form1},
  Nesy24b in 'Nesy24b.pas',
  infos in 'infos.pas' {info},
  errorlo in 'errorlo.pas' {ErrorLog},
  RegExpr in '3rdparty\RegExpr.pas',
  options in 'options.pas' {OptionDialog};

{$R *.RES}

begin
  Application.Initialize;
  Application.Title := 'Excelkonverter';
  Application.CreateForm(TForm1, Form1);
  Application.CreateForm(TErrorLog, ErrorLog);
  Application.CreateForm(TOptionDialog, OptionDialog);
  Application.Run;
end.
