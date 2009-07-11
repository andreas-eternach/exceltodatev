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
unit infos;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls;

type
  Tinfo = class(TForm)
    Memo1: TMemo;
    Label1: TLabel;
    Label2: TLabel;
    Button1: TButton;
    procedure Button1Click(Sender: TObject);
    procedure creat(Sender: TObject);
  private
    { Private-Deklarationen }
  public
    { Public-Deklarationen }
  end;

var
  info: Tinfo;

implementation

{$R *.DFM}

procedure Tinfo.Button1Click(Sender: TObject);
begin
 close;
end;

procedure Tinfo.creat(Sender: TObject);
begin
	Memo1.Lines.Add('Datev-Konverter Version 3.0 Beta');
	Memo1.Lines.Add('');
	Memo1.Lines.Add('(c) ''96-''01 by:');
	Memo1.Lines.Add('Andreas Eternach');
	Memo1.Lines.Add('Ortrunweg 5');
	Memo1.Lines.Add('04279 Leipzig');
	Memo1.Lines.Add('');
	Memo1.Lines.Add('Bekannte Bugs:');
	Memo1.Lines.Add('noch keine');
	Memo1.Lines.Add('');
end;

end.
