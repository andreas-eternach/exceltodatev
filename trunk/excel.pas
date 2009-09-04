//This file is part of exceltodatev.

//  exceltodatev is free software: you can redistribute it and/or modify
//  it under the terms of the GNU General Public License as published by
//  the Free Software Foundation, either version 3 of the License, or
//  (at your option) any later version.
//
//  exceltodatev is distributed in the hope that it will be useful,
//  but WITHOUT ANY WARRANTY; without even the implied warranty of
//  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
//  GNU General Public License for more details.
//
//  You should have received a copy of the GNU General Public License
//  along with exceltodatev .  If not, see <http://www.gnu.org/licenses/>.
//
//(c) 1996-2009 Andreas Eternach (andreas.eternach@google.com)
unit excel;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, DdeMan,nesy24b, ComCtrls,infos, Menus, options, DateUtils;

//
//Typen für Zahlenangaben im Hauptformular
//
const C_DATE       : integer = 1;
const C_YEAR       : integer = 2;
const C_VORLNUMBER : integer = 3;

type
  TForm1 = class(TForm)
    client: TDdeClientConv;
    item: TDdeClientItem;
    btnStart: TButton;
    inhalt: TComboBox;
    MNR: TEdit;
    Bearb: TEdit;
    Vorlauf: TEdit;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    Label6: TLabel;
    refreshFromExcel: TButton;
    fortschritt: TProgressBar;
    MainMenu1: TMainMenu;
    Datei1: TMenuItem;
    Beenden1: TMenuItem;
    Info1: TMenuItem;
    Optionen1: TMenuItem;
    N1: TMenuItem;
    cbYears: TComboBox;
    start: TComboBox;
    ende: TComboBox;
    yearErrorText: TLabel;
    endeErrorText: TLabel;
    startErrorText: TLabel;
    procedure startChange(Sender: TObject);
    procedure startSelect(Sender: TObject);
    procedure endeSelect(Sender: TObject);
    procedure endeChange(Sender: TObject);
    procedure cbYearsSelect(Sender: TObject);
    procedure cbYearsChange(Sender: TObject);
    procedure Optionen1Click(Sender: TObject);
    procedure Info1Click(Sender: TObject);
    procedure Beenden1Click(Sender: TObject);
    procedure konvertieren(Sender: TObject);
    procedure zeige(Sender: TObject);
    procedure refreshFromExcelClick(Sender: TObject);

    procedure DateTest (Date, Year, Meldung : string);
    function  stringkonvert(s : string; typ : integer; meldung : string) : string;
    procedure wndproc(var Message : TMessage);override;

    procedure Button3Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);

  private
    { Private declarations. }
    k        : import;
    into     : Tinfo;
    isYearValid : boolean;
    const clStateError : TColor = $9090FF;
    const clStateOk : TColor = $90FF90;
    const clStateUnknown : TColor = $90FFFF;

    var currSelectedYear : integer;
    procedure ncactivate(var Msg : TMessage);message WM_NCACTIVATE;
    procedure ncc(var Msg : TMessage);message WM_NOTIFY;
    procedure setUiEnabledState(newState:boolean);
    procedure fillControls();
    function getComboText(cb : TComboBox) : String;
    procedure checkDateCombo(Sender : TComboBox; errorLabel : TLabel);
  public
    { Public-Deklarationen }
    client1 : TDDeClientConv;
  end;

var
  Form1: TForm1;

implementation

uses errorlo;

procedure tform1.ncc(var Msg : TMessage);
begin;
 inherited;
end;
procedure tform1.ncactivate(var msg : TMessage);
begin;
 inherited;
 // disabled, since not required any more
 //if Active=false then Height:=319
 //   else Height:=25;
end;


{$R *.DFM}
//testet Datum auf logische Gültigkeit
procedure TForm1.DateTest (Date, Year, Meldung : string);
var Jahr, Monat, Tag : integer;
    Datum            : TDateTime;
begin;
  try
    //nachsehen, ob jahr zwei oder 4stellig
    if Length (Year) = 2 then Jahr := 1900
      else Jahr := 0;
    //Jahr ergänzen
    Jahr := Jahr + StrToInt (Year);
    //Tage berechnen
    Tag := StrToInt (Copy (Date, 1, 2));
    //Monate berechnen
    Monat := StrToInt (Copy (Date, 3, 2));
    //in Datum konvertieren
    if (not TryEncodeDate (Jahr, Monat, Tag, Datum)) then
      raise EConvertError.Create(Format('Kein gültiges Datum: Jahr:%d, Monat:%d, Tag:%d', [Jahr, Monat, Tag]));
    //Exception abfangen und erneut auslösen
  except
    on EConvertError do
      begin;
        Raise EConvertError.Create (Meldung);
      end;
  end;
end;

//konvertiert Vorlaufnummer, ... in richtiges Format
function TForm1.stringkonvert(s : string; typ : integer; meldung : string) : string;
var wert : integer;
    temp : string;
begin;
  try
   StrToInt (s);
  except
   on EConvertError do
    begin;
      //Behandlung der verschiedenen Typen, die Kommas, Punkte enthalten dürfen
      case typ of
        1 : begin;
              //Tage bearbeiten
              temp := Copy (s, 1, Pos ('.', s) - 1);
              if ((Length (temp) < 1) or (Length (temp) > 2)) then
                 Raise EConvertError.Create (Meldung);
              s := Copy (s, Length (temp) + 2, Length (s) - Length (temp) + 1);
              if (Length (temp) = 1) then temp := '0' + temp;
              //Monat bearbeiten
              temp := temp + Copy (s, 1, Pos ('.',s) - 1);
              if ((Length (temp) < 3) or (Length (temp) > 4)) then
                 Raise EConvertError.Create (Meldung);
              s := Copy (s, Length (temp), Length (s) - (Length (temp) - 1));
              if (Length (temp) = 3) then Insert ('0', temp, 3);
              //nach letztem Komma noch Einträge?
              if (Length(s) <> 0) then
                Raise EConvertError.Create (Meldung);
              s := temp;
              //jetzt in Integer kopierbar?
              try
               StrToInt (s);
              except
               on EConvertError do
                 begin;
                  Raise EConvertError.Create (Meldung);
                 end;
               end;
            end;
        2 : Raise EConvertError.Create (Meldung);
        3 : Raise EConvertError.Create (Meldung);
      end;
    end;
  end;
  case typ of
   1 : begin;
         temp := Copy (s, 1, 2);
         wert := StrToInt (temp);
         if (wert > 31) then
          Raise EConvertError.Create (Meldung);
         temp := Copy (s, 3, 2);
         wert :=StrToInt (temp);
         if (wert > 12) then
          Raise EConvertError.Create (Meldung);
       end;
   2 : begin;
         //Jahr berarbeiten
         wert:=StrToInt (s);
         if wert > 99 then
           Raise EConvertError.Create (Meldung);
         s := Format ('%.2d', [wert]);
       end;
   3 : begin;
         //Vorlaufnummer bearbeiten
         wert:=StrToInt (s);
         if ((wert > 9999) or (wert = 0)) then
           Raise EConvertError.Create (Meldung);
         s := Format ('%.4d', [wert]);
       end;
  end;
  Result := s;
end;

//liest Daten aus Steuerelementen
//prüft diese Daten
//erzeugt Thread für Konvertierung
//startet Konvertierung
procedure TForm1.konvertieren(Sender: TObject);
begin
  client1:=TDDeClientConv.create(self);

  k:=import.create(Handle,'',client,client1, OptionDialog.getTemplateDir, OptionDialog.getSaveDir);
  try
    //Listboxinhalt prüfen
    if inhalt.ItemIndex = -1 then
      Raise EConvertError.Create ('Sie haben keine Arbeitsmappe ausgewählt.');
    //Editfelder auf Fehler prüfen
    k.startzeitraum := stringkonvert(start.Text, C_DATE, 'Ungültiges Startdatum.');
    k.endezeitraum := stringkonvert(ende.Text, C_DATE, 'Ungültiges Endedatum.');
    k.bearbeiter := bearb.text;
    k.jahr := stringkonvert(cbYears.Text, C_YEAR, 'Ungültiges Jahr.');
    k.vorlauf := stringkonvert(vorlauf.text, C_VORLNUMBER, 'Ungültige Vorlaufnummer.');
    k.beraternummer := '115024';
    //Datum auf logische Gültigkeit testen
    DateTest (k.startzeitraum, k.jahr, 'Ungültiges Startdatum (dieser Tag existiert nicht).');
    DateTest (k.endezeitraum, k.jahr, 'Ungültiges Endedatum (dieser Tag existiert nicht).');

    //Open the DDE connection to MS-Excel
    client1.ConnectMode := ddeAutomatic;
    client1.Setlink ('Excel',inhalt.Items[inhalt.Itemindex]);

    // disable the UI
    setUiEnabledState(false);
    //Thread starten
    k.Execute;
    //auf Thread warten
    while k.fertig=false do Application.ProcessMessages;
    //wenn bei Konvertierung aufgetreten sind
    if k.ShowErrorLog then
     begin;
      Errorlog.liste:=k.fehlermeldungen;
      Visible:=false;
      Errorlog.ShowModal;
      Visible:=true;
     end;
    k.Free;
  except
    on e : EConvertError do
     begin;
      k.Free;
      application.Messagebox(pchar(e.Message),'Fehler',MB_OK);
     end;
  end;
end;

procedure TForm1.zeige(Sender: TObject);
begin
 if inhalt.Items.Count=0 then
   refreshFromExcelClick(Sender);
end;

//liest verfügbare Arbeitsmappen aus Excel neu ein
procedure TForm1.refreshFromExcelClick(Sender: TObject);
var s,s1,p1 : string;
    i       : integeR;
begin
 //inhalt der Combobox loeschen
 inhalt.clear;
 client.closelink;
 client.ConnectMode:=ddeManual;
 client.SetLink('Excel','System');
 client.openlink;
 if ((client.ddeservice<>'Excel')and(client.ddetopic<>'System')) then
  Application.MessageBox('Kann keine Verbindung zu Excel herstellen/neditieren Sie evtl. gerade eine Zelle?','Fehler',IDOK);
 s:=client.requestdata('Topics');
 i:=10;
 p1:=chr(9);
 while ((i<>0)and(s<>'')) do
  begin;
   i:=pos(p1,s);
   s1:=copy(s,0,i-1);
   s:=copy(s,i+1,strlen(pchar(s)));
   if (s1<>'') then inhalt.items.Add(s1);
  end;
end;

procedure TForm1.wndproc(var Message : TMessage);
begin;
 inherited wndproc(Message);
 if (Message.msg=55555) then
  begin;
   Fortschritt.Max:=MEssage.LParam;
   fortschritt.position:=Message.WParam;
  end;
 if (Message.msg=55556) then
  begin;
   setUiEnabledState(true);
  end;
end;
procedure TForm1.setUiEnabledState(newState: boolean);
begin
   btnStart.Enabled:=newState;
   Datei1.Enabled:=newState;
end;
procedure TForm1.Button3Click(Sender: TObject);
begin
 ErrorLog.Free;
 Close;
end;

var
mHandle :THandle; // Mutexhandle

procedure TForm1.FormCreate(Sender: TObject);
begin
  // initialize basic constants / members
  isYearValid := false;

  // initialize complex / dependent state of object / form
  ErrorLog:=TErrorLog.Create(Self);
  ErrorLog.Hide;
  fillControls;
end;

procedure TForm1.fillControls();
  var currDate : TDateTime;
      yearList : TStringList;
      i, year  : integer;
      currYearStr : String;

begin;
  // determine the current year
  currDate := DateUtils.Today();
  year := DateUtils.YearOf(currDate);
  yearList := TStringList.Create;

  // fill last ten years into the list
  for i := 1 to 10 do
  begin
    currYearStr := IntToStr(year);
    yearList.Add(Copy(currYearStr, 3, 2));
    year := year - 1;
  end;
  cbYears.Items := yearList;

  // select latest year in list
  cbYears.SelText := cbYears.Items[0];
end;

procedure TForm1.FormDestroy(Sender: TObject);
begin
  //
end;

procedure TForm1.Beenden1Click(Sender: TObject);
begin
  Button3Click(Sender);
end;

procedure TForm1.Info1Click(Sender: TObject);
begin
 if (into<>nil) then
  begin;
   into.destroy;
   into:=nil;
  end;
 if (into=nil) then
  begin;
   into:=Tinfo.create(Self);
   into.Showmodal;
  end;
end;

procedure TForm1.Optionen1Click(Sender: TObject);
begin
  // show the option-dialog.
  OptionDialog.ShowModal;
end;

procedure TForm1.cbYearsChange(Sender: TObject);
var currMonthStart, nextMonthEnd  : TDateTime;
    month : integer;
begin
  try
    currSelectedYear := StrToInt(getComboText(cbYears));
    if (currSelectedYear < 0) or (currSelectedYear > 99) then
      raise EConvertError.Create('Jahr muss größer als 0 und kleiner 100 sein.');
    cbYears.Color := clStateOk;
    // calculate begin and end for the single monthes of the selected year
    ende.Text := '';
    start.Text := '';
    // 2000-related corrections
    if (currSelectedYear < 50) then
      currSelectedYear := currSelectedYear + 100;
    currSelectedYear := currSelectedYear + 1900;
    isYearValid := true;
    currMonthStart := EncodeDateTime(currSelectedYear, 1, 1, 0, 0, 0, 0);
    for month := 1 to 12 do
    begin
      currMonthStart := IncMonth(currMonthStart, 1);
      start.Items.Add('01' + Format('%.*d', [2, month]));
      nextMonthEnd := IncDay(currMonthStart, -1);
      ende.Items.Add(Format('%.*d%.*d', [2, DayOf(nextMonthEnd), 2, month]));
    end;
    ende.SelText := ende.Items[0];
    start.SelText := start.Items[0];
    yearErrorText.Caption := '';
    yearErrorText.Visible := false;
  except
    // ignore integer parsing-errors, may happen
    // if the year contains an invalid valud
    on exc : EConvertError do
    begin
      cbYears.Color := clStateError;
      yearErrorText.Caption := exc.Message;
      yearErrorText.Visible := true;
      isYearValid := false;
    end;
  end;
end;

procedure TForm1.cbYearsSelect(Sender: TObject);
begin
  cbYearsChange(Sender);
end;

function TForm1.getComboText(cb : TComboBox) : String;
begin
    Result := cb.Text
end;

procedure TForm1.checkDateCombo(Sender : TComboBox; errorLabel : TLabel);
  var endeText : string;
      endeDate : TDateTime;
      tag      : integer;
begin
  // check the end date for compatibility
  errorLabel.Visible := false;
  errorLabel.Caption := '';
  Sender.Color := clStateOk;
  if (isYearValid) then
  begin
    endeText := getComboText(Sender);
    if (Length(endeText) < 4) then
    begin
      errorLabel.Visible := true;
      errorLabel.Color := clStateUnknown;
      Sender.Color := clStateUnknown;
      errorLabel.Caption := 'Weniger als 4 Stellen eingegeben.';
      exit;
    end;
    try
      tag := StrToInt(endeText);
      endeDate := EncodeDateTime(currSelectedYear, tag mod 100, tag div 100, 0, 0, 0, 0);
    except
      on e : EConvertError do
      begin
      errorLabel.Visible := true;
      errorLabel.Color := clStateError;
      Sender.Color := clStateError;
      errorLabel.Caption := e.Message;
      end;
    end;
  end;
end;

procedure TForm1.endeChange(Sender: TObject);
begin
  checkDateCombo(ende, endeErrorText);
end;

procedure TForm1.endeSelect(Sender: TObject);
begin
  endeChange(Sender);
end;

procedure TForm1.startSelect(Sender: TObject);
begin
  startChange(Sender);
end;

procedure TForm1.startChange(Sender: TObject);
begin
  checkDateCombo(start, startErrorText);
end;

initialization

 mHandle:= CreateMutex(nil, True, 'Excel-DDE');
  if GetLastError = ERROR_ALREADY_EXISTS then
   begin
    MessageDlg('Excel-Konverter läuft bereits!', mtInformation, [mbOK], 0);
    halt;
   end;

finalization

 if mHandle <> 0 then CloseHandle(mHandle);


end.
