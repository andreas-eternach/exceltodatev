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
unit errorlo;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls;

type
  TErrorLog = class(TForm)
    ErrorMemo: TMemo;
    warnungen: TCheckBox;
    fehler: TCheckBox;
    Schliessen: TButton;
    infosmsgs: TCheckBox;
    procedure infosmsgsClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure SchliessenClick(Sender: TObject);
    procedure fehlerClick(Sender: TObject);
    procedure warnungenClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure ErrorMemoClick(Sender: TObject);
  private
    { Private-Deklarationen }
  public
    tempLines : TStringList;
    liste     : TStringList;
    { Public-Deklarationen }
  end;

var
  ErrorLog: TErrorLog;

implementation

{$R *.DFM}


procedure TErrorLog.FormShow(Sender: TObject);
var i : integer;
begin
  tempLines.Clear;
  for i:=0 to liste.Count-1 do begin;
    if ((warnungen.Checked)and(Pos('W',liste.Strings[i][1])=1)) then
     begin;
      tempLines.Add(liste.Strings[i]);
      tempLines.Add('');
     end;
    if ((fehler.Checked)and
      ((Pos('F',liste.Strings[i])=1)or(Pos('S',liste.Strings[i][1])=1))) then
        begin;
          tempLines.Add(liste.Strings[i]);
          tempLines.Add('');
        end;
    // Display Info-Messages?
    if ((infosmsgs.Checked)and
      (Pos('I',liste.Strings[i][1])=1)) then
        begin;
          tempLines.Add(liste.Strings[i]);
          tempLines.Add('');
        end;

  end;
  ErrorMemo.Lines.Text:=tempLines.Text;
end;

procedure TErrorLog.SchliessenClick(Sender: TObject);
begin
  Self.ModalResult:=12;
end;

procedure TErrorLog.fehlerClick(Sender: TObject);
begin
 FormShow(Sender);
end;

procedure TErrorLog.warnungenClick(Sender: TObject);
begin
 FormShow(Sender);
end;

procedure TErrorLog.FormCreate(Sender: TObject);
begin
 tempLines:=TStringList.Create;
end;

procedure TErrorLog.FormDestroy(Sender: TObject);
begin
  tempLines.free;
end;

procedure TErrorLog.ErrorMemoClick(Sender: TObject);
begin
  ErrorMemo.SelStart
end;

procedure TErrorLog.infosmsgsClick(Sender: TObject);
begin
 FormShow(Sender);
end;

end.
