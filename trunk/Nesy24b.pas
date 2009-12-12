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

unit nesy24b;
interface

uses Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, DdeMan, RegExpr;
     type headerRecognizeColCmpCmd = (exactLowerCaseMatch, startLowerCaseMatch);
     type a=^w;
          b=^z;
          w=array[0..64000] of byte;
          z=array[0..1023] of byte;
     const colAddionsPrefix = 'zusatz:';
     const dataColNames : array[1..9] of string=('Datum','Gegenkonto','Konto','Soll',
                                               'Haben','Text','Beleg1','Beleg2', colAddionsPrefix);
     const colHeadCmpCmd : array[1..9] of headerRecognizeColCmpCmd= (
                                              exactLowerCaseMatch, exactLowerCaseMatch,
                                              exactLowerCaseMatch, exactLowerCaseMatch,
                                              exactLowerCaseMatch, exactLowerCaseMatch,
                                              exactLowerCaseMatch, exactLowerCaseMatch,
                                              startLowerCaseMatch);
//Fehlertypen
//~~~~~~~~~~~

const INFO : integer = 0;
const WARNING : integer = 1;
const ERROR : integer = 2;
const SERROR : integer = 3;

type import=class

		fDatev, fVerzeichnis			: TFileStream;
	  //neue Variablen zu Fehlerbehandlung - 16.7.
	  fehlertext                 : string;
	  fehlercap                  : string;
	  fehlermeldungen            : TStringList;
    templateDir                : String;
    saveDir                    : string;
	  errors, warnings, infos    : bool;
	  ShowErrorLog               : bool;
	  //Ende 16.7.
	  //Statuswert f¸r Elternprozeﬂ
	  fertig                     : bool;

	  //fdatev,fverzeichnis        : file;
	  nameexcel                  : string;
    const maxDatasetLength = 1024;
	  var datensatz                  : array[0..maxDatasetLength] of byte;
	  start                      : word;
	  endsumme,fexcellang        : longint;
	  zeile                      : integeR;

	  zeilenende                             : byte;
	  lesestring                             : string[255];
	  lesearray                              : array[1..30] of byte;
	  habe,sol,datu,gegenkont,kont,tex       : array[1..5] of byte;
    // enthalten die Nummern der Excel-Spalten, welche die jeweilige Information beinhalten.
	  bele1,bele2, skont, kos1, kos2       : array[1..5] of byte;
	  waehrun, kostmeng								: array[1..5] of byte;

    // required for the additional information
    const maxAdditionColIndex = 19;
    var additionCols          : array[0..maxAdditionColIndex] of byte;
    additionColCnt            : byte;
    additionNames             : array[0..maxAdditionColIndex] of String;

	  haben,soll,datum,gegenkonto,konto      : string[15];
	  beleg1,beleg2                          : string;
	  text                                   : string[255];
	  Kost1, Kost2											: string[8];
	  Skonto													: string[15];
		KostMenge											: string [15];
		Waehrung												: string [5];

     Ende                                     : boolean;
 	   startzeitraum,endezeitraum,beraternummer : string[6];
     bearbeiter,berater,mandant,vorlauf       : string[8];
     jahr                                     : string[3];
     daten         : a;
     verzeichnis   : b;
     handle        : integer;
     //VCL-Objects for communication with MS-Excel.
     client                 : TDDeClientConv;
     client1                : TDDeClientConv;
     quelle                 : string;
     spaltenanz,zeilanz     : integer;
     aktspalte              : integer;
     mstring                : string;
     zeilenstring           : string;

     // application-logic related functions
     procedure   box;
     procedure   fehler(text : string; typ : integer);
     procedure   readywarten;
     procedure   zellelesen;
     function    stringzahl(s : string) : word;
     procedure   speicherholen;
     procedure   oeffneeingabe;
     procedure   oeffneausgabe;
     procedure   speichergeben;
     procedure   schliesseausgabe;
     procedure   exportspeicherverzeichnis;
	   procedure   exportspeicherdaten;
     procedure   schreibedaten;
     procedure   schreibeverzeichnis;
     procedure   exportAdditions(additions : array of String);
     procedure   exportDataset;
     procedure   exportsumme;
	   procedure   rechnen;
     procedure   lesen;
     function    istdatum (s,typ : string; meldung : string) : string;
     function    istzahl(s,typ : string; meldung : string) : string;
     procedure   belegtesten(beleg : string);
     function    textkonvert(text : string) : string;
     procedure   fehlersuchen;
     procedure   loeschen;
     procedure   konvertsatz;
     procedure   readAndExportOneDataset;
     procedure   importsatz;
     procedure   fortschritt;
     constructor create(handl : integer;quell : string;clien,clien1 : TDDeClientConv;templateDir,saveDir,beraternr : string);virtual;
     procedure   execute;
  end;

implementation

{******************Abbruch bei Fehler*******************}
procedure import.Box;
begin;
 Application.MessageBox(pchar(fehlertext),pchar(fehlercap),MB_OK);
end;

procedure import.readywarten;
var s : string;
    i : integer;
begin;
 i:=0;s:='';
 while ((s<>'Ready')and(i<100)) do
  begin;
   inc(i);
	s:=client.RequestData('STATUS');
  end;
 //evtl. Exception auslˆsen
 if ((i=100)and(s<>'Ready')) then
  begin;
   fehler ('Excel ist nicht mehr verf¸gbar!', SERROR);
  end;
end;

procedure import.zellelesen;
begin;
 mstring:=client1.RequestData('Z'+inttostr(zeile)+'S'+inttostr(aktspalte));
 mstring:=copy(mstring,1,strlen(pchar(mstring))-2);
end;

procedure import.fehler(text : string; typ : integer);
var fehler : EConvertError;
    Entry  : string;
begin;
 //Fehlermeldung formatieren und in Stringliste aufnehmen
 Entry := '';
 if typ=INFO then
  begin;
     Entry := 'Info';
     infos:=true;
  end
 else if typ=WARNING then
   begin;
     Entry := 'Warnung';
     warnings:=true;
   end
 else if (typ=ERROR) then
   begin;
     Entry := 'Fehler';
	  errors:=true;
	end
 else
  if typ=SERROR then Entry := 'Fehler';
 Entry:=Entry + ' in Zeile ' + IntToStr(zeile) + ': ' + text;
 fehlermeldungen.Add(Entry);
 //bei Fehler mit dem nicht fortgefahren werden kann,
 //Bearbeitung abbrechen
 if typ=SERROR then
  begin;
   fehler:=EconvertError.create(text);
	  raise fehler;
  end;
end;
{********************String-Lesefunktionen*************************}
function import.stringzahl(s : string) : word;
var i : integer;
    w : word;
begin;
 w:=0;
 for i:=1 to length(s) do begin;
  w:=w*10+ord(s[i])-48;
 end;
 stringzahl:=w;
end;
{*****************init-Routinen************************}
{Speicher belegen}
procedure import.speicherholen;
var j : word;
begin;
 getmem(daten,64000);
 getmem(verzeichnis,1024);
 for j:=64000 downto 0 do begin;
  daten^[j]:=0;
 end;
end;
{Dateien oeffnen}
procedure import.oeffneeingabe;
var dataColIndex    : array[Low(dataColNames)..High(dataColNames)] of byte;
    dataColIter,excelColIter  : integer;
    s               : string;
begin;
 //******************Spalten auswerten*****************
 //Anzahl der Spalten bestimmen
 for dataColIter:=Low(dataColNames) to High(dataColNames) do begin;
  dataColIndex[dataColIter]:=0;
 end;
 excelColIter:=1;
 //nach betreffenden Spalten suchen
 while excelColIter<20 do
  begin;
   //in s spalteninhalt laden
   readywarten;
   zeile:=1;aktspalte:=excelColIter;
   zellelesen;
   s:=mstring;
   //nach Feld suchen (z.B. Text --> Spalte 2)
   for dataColIter:=Low(dataColNames) to High(dataColNames) do begin;
    case colHeadCmpCmd[dataColIter] of
      exactLowerCaseMatch:
        if (dataColNames[dataColIter] = s) then dataColIndex[dataColIter] := excelColIter;
      startLowerCaseMatch:
        if (Pos(LowerCase(dataColNames[dataColIter]), LowerCase(s)) = 1) then dataColIndex[dataColIter] := excelColIter;
    end;
   end;
   inc(excelColIter);
  end;
 //Auswertung, wieviele spalten gelesen werden m¸ssen
 spaltenanz:=0;
 //die letzten beiden Schl¸sselworte sind Kann-Felder
 for dataColIter:=Low(dataColNames) to High(dataColNames) do begin;
  if (dataColIndex[dataColIter] > spaltenanz) then spaltenanz := dataColIndex[dataColIter];
  //die letzten beiden Schl¸sselworte sind Kann-Felder (Beleg1 und 2)
  if ((dataColIndex[dataColIter] = 0) and (dataColIter < 7)) then fehler('Fehler im Tabellenkopf.', SERROR);
 end;
 //********zeilen auswerten*******************
 s:=' ';dataColIter:=1;
 while ((s<>'')and(s<>'EndeDatensatz')) do
  begin;
   Readywarten;
   zeile:=dataColIter;aktspalte:=1;
   zellelesen;
   s:=mstring;
   inc(dataColIter);
  end;
 if (s='') then fehler('kein EndeDatensatz oder leere Zelle.', SERROR);
 zeilanz:=dataColIter-1;
end;
{Datev-Datei oeffnen}
procedure import.oeffneausgabe;
var i : integer;
  cwd : string;
begin;
  // determine current working directory
  cwd := GetCurrentDir() + '\diskette';
	try
		fDatev := TFileStream.Create (templateDir + '\Ed00001', fmOpenRead);
		fDatev.Read(daten^[1], 256);
		for i :=$60 to 511 do begin;
			daten^[i] := $00;
		end;
		fDatev.Destroy;
	except
		on e : Exception do
		begin
			fehler('Fehler beim Zugriff auf die Datei ' + templateDir + '\Ed00001: ' + e.Message, SERROR);
			exit;
		end;
	end;
  {Init Startparameter}
  start:=$5F;
  {
  old, hopefully obsolete format
  start:=$5F;
  }
	try
		fVerzeichnis := TFileStream.Create (templateDir + '\Ev01', fmOpenReadWrite);
		fVerzeichnis.Read(verzeichnis^[1], 256);
		fVerzeichnis.Destroy;
	except
		on e : Exception do
		begin
			fehler('Fehler beim Zugriff auf die Datei ' + templateDir + '\Ev01: ' + e.Message, SERROR);
			exit;
		end;
	end;
	try
    fDatev:=nil;
    fVerzeichnis:=nil;
		fDatev := TFileStream.Create (saveDir + '\Ed00001', fmCreate);
		fDatev.Size := 0;
		fVerzeichnis := TFileStream.Create (saveDir + '\Ev01', fmCreate);
	except
		on e : Exception do
		begin
			fehler('Fehler beim Zugriff auf die Datei ' + saveDir + '\E...: ' + e.Message, SERROR);
			exit;
		end;
	end;
end;

{*************************exit-routinen***************************}
{Speicher freigeben}
procedure import.speichergeben;
begin;
 freemem(daten,64000);
 freemem(verzeichnis,1024);
end;
procedure import.schliesseausgabe;
begin;
  try
	  if (fDatev <> nil) then fDatev.Destroy;
	  if (fVerzeichnis <> nil) then fVerzeichnis.Destroy;
  except
    on e: Exception do
    begin
    end;
  end;
end;
{*****************************Ende schliessen**********************}
{*******************ExportSpeicherfunktionen*******************}
{******fÅr Datei mit verzeichn. dv01}
procedure import.exportspeicherverzeichnis;
var i : integer;
begin;
  // write worker
  for i := 1 to 2 do
    verzeichnis^[$89 + i - 1] := ord(bearbeiter[i]);

  // write consultant number
  for i := 1 to 6 do
  begin
    verzeichnis^[6 + i] := ord(beraternummer[i]);
    verzeichnis^[$8B + i] := ord(beraternummer[i]);
  end;

  // write client number (only lowest 5 bytes)
  while (length(mandant) < 5) do
    mandant := '0' + mandant;
  for i := 1 to 5 do
    verzeichnis^[$92 + i - 1] := ord(mandant[Length(mandant) - 5 + i]);

  // write starting date
  for i := 1 to 6 do
    verzeichnis^[$A0 + i] := ord(startzeitraum[i]);

  // write ending date
  for i := 1 to 6 do
    verzeichnis^[$A6 + i] := ord(endezeitraum[i]);

  // write vorlauf number
  while (length(vorlauf) < 6) do
    vorlauf := '0' + vorlauf;
  for i := 1 to 6 do
    verzeichnis^[$97 + i - 1] := ord(vorlauf[length(vorlauf) - 6 + i]);

{
 old, hopefully obsolete format
 for i:=1 to 6 do begin;
  if i<3 then verzeichnis^[i+$88]:=ord(bearbeiter[i]);
  if i<=5 then
	begin;
	 verzeichnis^[i+$91]:=ord(mandant[i]);
	end;
  verzeichnis^[i+$A0]:=ord(startzeitraum[i]);
  verzeichnis^[i+$A6]:=ord(endezeitraum[i]);//63
  verzeichnis^[i+$96]:=ord(vorlauf[i]);
 end;
 }
end;
{******for File containing accounting data 'de001'}
procedure import.exportspeicherdaten;
var i : integer;
begin;
  // write worker
  for i := 1 to 2 do
    daten^[$09 + i - 1] := ord(bearbeiter[i]);

  // write consultant number
  for i := 1 to 6 do
    daten^[$0C + i - 1] := ord(beraternummer[i]);

  // write client number (only lowest 5 bytes)
  for i := 1 to 5 do
    daten^[$12 + i - 1] := ord(mandant[Length(mandant) - 5 + i]);

  // write vorlauf number
  for i := 1 to 6 do
    daten^[$17 + i - 1] := ord(vorlauf[i]);

  // write starting date
  for i := 1 to 6 do
    daten^[$1D + i - 1] := ord(startzeitraum[i]);

  // write ending date
  for i := 1 to 6 do
    daten^[$23 + i - 1] := ord(endezeitraum[i]);


 {
 old, hopefully obsolete format.
 for i:=1 to 5 do begin;
  daten^[i+17]:=ord(mandant[i]);
  if i<3 then daten^[i+8]:=ord(bearbeiter[i]);
 end;
 for i:=1 to 6 do begin;
  daten^[i+28]:=ord(startzeitraum[i]);
  daten^[i+34]:=ord(endezeitraum[i]);
  daten^[i+22]:=ord(vorlauf[i]);
 end;
 }
end;

{*******************DateiSchreibefunktionen********************}
{ getestet fÅr start<64000=OK}
procedure import.schreibedaten;
var bloecke : integer;
begin;
 {$I-}
 bloecke:=2*(trunc(start/256)+1);
 fDatev.Write(daten^[1], bloecke * 128);
 {$I+}
end;
{getestet=OK}
procedure import.schreibeverzeichnis;
begin;
 {$I-}
 fVerzeichnis.Write(verzeichnis^[1], 256);
 //blockwrite(fverzeichnis,verzeichnis^[1],2,Result);
 //if Result<>2 then fehler('Kann nicht in Datei ''a:\dv01'' schreiben.', SERROR);
 {$I+}
end;


procedure import.exportAdditions(additions : array of String);
var additionsIter : integer;
	  cachedDataset : array[Low(datensatz)..High(Datensatz)] of byte;
    strIter       : integer;
begin
  // cache the current dataset
  move(datensatz, cachedDataset, Sizeof(datensatz));

  for additionsIter := Low(additions) to High(additions) do
  begin
    if (Length(additions[additionsIter]) > 0) then
    begin
      for strIter := 0 to Length(additions[additionsIter]) do
        datensatz[strIter] := ord(additions[additionsIter][strIter]);
      datensatz[strIter] := 0;
      exportDataset;
    end;
  end;

  // restore the current dataset
  move(cachedDataset, datensatz, Sizeof(cachedDataset));

end;

{*************************************************************}
{sbetrag,hbetrag,gegenkonto,datum,konto,text}
procedure import.exportDataset;
var laenge   : byte;
	 i         : integer;
   start_tmp : integer;
   currDatasetIndex : integer;
begin;
 laenge := 0;
 // determine the length of the dataset
 for i:=1 to maxDataSetLength do begin;
  if datensatz[i]=0 then
  begin
    laenge:=i;
    break;
  end;
 end;
 if (start+laenge+6>=64000) then
  begin
	  schreibedaten;
	  //Speicherbereich auf 0 zur¸cksetzen
	  for i := 0 to 64000 do
    begin
	   daten^[i]:=0;
	  end;
	  start:=0;
  end;
  currDatasetIndex := Low(datensatz) + 1;
  while true do
   begin
    if datensatz[currDatasetIndex]= 0 then
      break;
    // set last 6 characters at zero to signal buffer overflow
    if (start mod 256 >=251) then
    begin
      // fillup the remainder with zeros
      while (start mod 256 <> 1) do
      begin
        daten^[start] := $00;
        start := start + 1;
      end;
    end;
    // assign the data
    daten[start] := datensatz[currDatasetIndex];
    // increase positions (is there a ++ Operator in pascal ?)
    start := start + 1;
    currDatasetIndex :=currDatasetIndex + 1;
  end;    
 // old, hopefully obsolete, splitting mechanism.
 {
 // is the mem-buffer full and needs to be flushed?
 if (start+laenge-2+5>=64000) then
  begin;
	  schreibedaten;
	  //Speicherbereich auf 0 zur¸cksetzen
	  for i := 0 to 64000 do begin;
	   daten^[i]:=0;
	  end;
	  start:=0;
  end;

 daten^[start]:=datensatz[1];
 inc(start);
 if (trunc((start+laenge-2+5)/256))>trunc(start/256) then
	begin;
		start_tmp := start;
		start:=(trunc(start/256)+1)*256+1;
		for i := start_tmp to start do
		begin
			daten^[i] := 0;
		end;
	end;
	for i:=2 to laenge-1 do begin;
		daten^[start]:=datensatz[i];
		inc(start);
	end;}
end;

procedure import.exportsumme;
var i : integer;
    l : longint;
begin;
 l:=endsumme;
 datensatz[1]:=ord('y');
 if endsumme<0 then datensatz[2]:=ord('w');
 if endsumme>=0 then datensatz[2]:=ord('x');
 if endsumme<0 then endsumme:=endsumme-2*endsumme;
 for i:=14 downto  3 do begin;
  datensatz[i]:=((endsumme mod 10)+48);
  endsumme:=trunc(endsumme/10);
 end;
 datensatz[15]:=ord('y');
 datensatz[16]:=ord('z');
 exportDataset;
 endsumme:=l;
end;
{**************************neu********************************}
procedure import.rechnen;
var i    : integer;
    zahl : longint;
begin;
 zahl:=0;
 i:=2;
 repeat;
  inc(i);
  zahl:=10*zahl+(datensatz[i]-48);
 until ((datensatz[i+1]=ord('a')) or (datensatz[i+1]=ord('l')));
 if datensatz[2]=ord('-') then endsumme:=endsumme-zahl;
 if datensatz[2]=ord('+') then endsumme:=endsumme+zahl;
end;
{liest inhalt einer zelle der excel-datei in string}
procedure import.lesen;
var zeilform : string;
begin;
 //wird durch DDE-Befehle ersetzt
 zeilenende:=0;
 //Einlesen ¸ber DDE
 readywarten;
 //wenn erste Spalte diese Zeile aus Excel lesen
 if aktspalte=1 then
   begin;
    zeilform := 'Z'+inttostr(zeile)+'S'+inttostr(aktspalte) +
                ':Z'+inttostr(zeile)+'S'+inttostr(spaltenanz);
    zeilenstring := client1.RequestData(zeilform);
    Insert (chr(9), zeilenstring, Length(zeilenstring)-1);
   end;
 //erste Spalte aus zeilenstring auslesen (ohne Tabzeichen)
 mstring := Copy (zeilenstring, 1, Pos (chr(9),zeilenstring) - 1);
 //Spalte in Zeilenstring lˆschen
 zeilenstring := Copy (zeilenstring, Length (mstring) + 2, Length (zeilenstring) - Length (mstring) - 1);
 lesestring:=mstring;
 //wenn Zeilenende erreicht
 if aktspalte=spaltenanz then
  begin;
   aktspalte:=1;
   zeilenende:=1;
  end;
 if zeilenende<>1 then inc(aktspalte);
    //Leerspaten entfernen
    if Trim (lesestring) = '' then lesestring := '';
end;
{++++++++++++++++++++++++FEHLERSUCHE+++++++++++++++++++++++++++++++++++++++}
{testet Beleg}
procedure import.belegtesten(beleg : string);
begin;
 if Length(beleg)>6 then
  begin;
	fehler('Belege d¸rfen maximal 6stellig sein.', ERROR);
  end;
 try
  if (beleg <> '') then StrToInt(beleg);
 except
  on e : Exception do
  begin;
	fehler('Belege m¸ssen Zahlen sein.', ERROR);
	e.free;
  end;
 end;
end;
{konvertiert Datum mit Kommapunkten und Jahr in Zahlenfolge}
function import.istdatum (s,typ : string; meldung : string) : string;
var Counter : integer;
	 Position, OldPosition : integeR;
	 teilstring, Rueckgabe : string;
	 tmp                   : integer;
	 Monat, Tag            : integer;
begin;
		Counter := 0;
		OldPosition := 0;
		Rueckgabe := '';
		try
		while (true) do
				begin
						//n‰chsten Punkt suchen
						Position := Pos ('.',
									Copy (s, OldPosition + 1, Length (s) - OldPosition));
						if Position = 0 then break;

						teilstring := Copy (s, OldPosition + 1, Position - 1);

						//Teilstring auf Zahl pr¸fen
						tmp := StrToInt (teilstring);

						//Tage pr¸fen
						if (Counter = 0) then
							begin
								  if tmp > 31 then
									  fehler (meldung + ' Tage stimmen nicht.', ERROR);
							end
						//Monate pr¸fen
						else if (Counter = 1) then
							begin
								  if ((tmp > 12) or (tmp <= 0)) then
									  fehler (meldung + ' Monate stimmen nicht.', ERROR);
							end;

						//String anh‰ngen
						if Length (teilstring) = 1 then
							Rueckgabe := Rueckgabe + '0';
						Rueckgabe := Rueckgabe + teilstring;

						inc (Counter);
						OldPosition := OldPosition + Position;
				end;

		//Anzahl der Dezimalpunkte pr¸fen
		if Counter > 2 then
			fehler (meldung, ERROR);

		//es wurde kein Dezimalpunkt gefunden
		if (Counter = 0) then
			Rueckgabe := s;

		//datum logisch pr¸fen - vierstelliges Datum erwartet
		Monat := StrToInt (Copy (Rueckgabe, 3, 2));
		Tag := StrToInt (Copy (Rueckgabe, 1, 2));
		EncodeDate (StrToInt (Copy (jahr, 1, 2)), Monat, Tag);

		//Exception tritt ein, wenn einzelne Teilstrings keine Integer sind
		//oder das angegebene Datum nicht existiert
		except
				On EConvertError do fehler (meldung, ERROR);
		end;

		istdatum := Rueckgabe;
end;

{konvertiert zahl mit komma in ohne komma, erzeugt fehler falls keine zahl}
function import.istzahl(s,typ : string; meldung : string) : string;
var i        : integer;
	 w        : string[50];
	 fehlerhaft : boolean;
	 nachkommastellen : integer;
   r : TRegExpr;
   regexpStr : string;
begin;
	//leere Strings, die nur whitechars enthalten, abfangen
	if (Trim (s) = '') then
	begin;
		istzahl:='';
		exit;
	end;
	if (LowerCase(typ) = 'kost') then
	begin
		istzahl := s;
    // Integer-Check disabled, cause it are strings now.
		//try
		//	i := StrToInt (s);
		//except
		//	on e : EConvertError do
		//	begin;
		//		fehler ('Keine Zahl eingegeben', ERROR);
		//		fehlerhaft := false;
		//	end;
		//end; //Try
		if (Length (s) > 8) then
		begin;
			fehler ('Kostenstellennummer darf nicht l‰nger als sechs Zeichen sein.', ERROR);
		end;	// StrLen > 8
	end //if Kost
	else if (LowerCase (typ) = 'beleg') then
	begin
		istzahl := s;
		if (Length (s) > 12) then
		begin;
			fehler ('Belegnummer darf nicht l‰nger als 12 Zeichen sein.', ERROR);
		end;	// StrLen > 12
    // must only contain chars of type (0-9, a-z, A-Z, $ & % * + - /)
    r := TRegExpr.Create;
    regexpStr := '([^\$,^&,^%,^\+,^\-,^\*,^0-9,^A-Z,^/])';
    r.Expression := regexpStr;
    r.Exec(UpperCase(s));
    if (r.MatchPos[0] <> -1) then
    begin
      fehler ('Belege d¸rfen nur die Zeichen 0-9, a-z, A-Z, $, &, %, *, +, -, / beinhalten.', ERROR);
      fehler ('Fehlerhafter Beleg-Text ist:"' + s + '"', ERROR);
    end;
	end // if typ = beleg
	else
	begin;
		w:='';
		fehlerhaft:=false;
		nachkommastellen := 0;

		if Trim(s)='' then
		begin;
			istzahl:='';
			exit;
		end;
		for i:=1 to length(s) do begin;
			if ((ord(s[i])<ord('0')) or (ord(s[i])>ord('9'))) and (s[i]<>',') and
				(s[i]<>'.') then fehlerhaft:=true;
			//Stringzeichen anh‰ngen
			w[0]:=chr(i);
			w[i]:=s[i];

			//Stellen-Z‰hler erhˆhen
			if (s[i]=',') or (s[i]='.') then
			begin;
				  if (length(s) <> i+2) then fehlerhaft:=true;
				  w[i]:=s[i+1];
				  w[i+1]:=s[i+2];
				  w[0]:=chr(i+1);
				  if (ord(s[i+2])<ord('0')) or (ord(s[i+2])>ord('9')) or
					  (ord(s[i+1])<ord('0')) or (ord(s[i+1])>ord('9'))
					  then
					  	fehlerhaft:=true;
				  nachkommastellen := 2;
				  break;//i:=length(s);
			end;
		end;
		if (w='0') or (w='000') then w:='';

		//Test f¸r spezielle Formate
		if (not fehlerhaft) then
		begin
			if (LowerCase(typ) = 'betrag') and (nachkommastellen <> 2) then
			begin
				fehlerhaft := true;
				meldung := meldung + ' Es fehlen die Nachkommastellen.';
			end;
			if (LowerCase(typ) = 'gegenkonto') and (Length(w) = 0) then
			begin
  			fehler ('Gegenkonto ist 0 oder leer.', INFO);
			end;
		end;

		//Ausgabe der Fehlermedung
		if fehlerhaft=true then
		begin;
			if meldung <> '' then
				fehler(meldung, ERROR)
			else
				fehler('Keine Zahl eingegeben', ERROR);
		end;
		istzahl:=w;
	end;
end;
{Ñndert Umlaute excel -> datev}
function import.textkonvert(text : string) : string;
var i : integer;
	 w : string;
begin;
 w:='';
 for i:=1 to length(text) do begin;
  //w[0]:=chr(i);
  w:=w+text[i];
  {Ñ}if w[i]=chr($E4) then w[i]:='Ñ';
  {î}if w[i]=chr($F6) then w[i]:='î';
  {Å}if w[i]=chr($FC) then w[i]:='Å';
  {é}if w[i]=chr($C4) then w[i]:='é';
  {ô}if w[i]=chr($D6) then w[i]:='ô';
  {ö}if w[i]=chr($DC) then w[i]:='ö';
  {·}if w[i]=chr($DF) then w[i]:='·';
 end;
 if length(text)<>0 then textkonvert:=w else textkonvert:=text;
end;
{sieht nach, ob Satz fehlerfrei}
procedure import.fehlersuchen;
begin;
	text:=textkonvert(text);
	//triviale Tests f¸r Konten
	if Length (konto)>5 then
	begin;
		fehler ('Konto ist l‰nger als 5 Zeichen(hier darf kein Ust.-Schl¸ssel eingegeben werden).', ERROR);
	end;
	{if gegenkonto='9000' then
	begin;
		fehler ('Das Konto 9000 darf nur auf der Kontoseite stehen.', ERROR);
	end;}
	if (gegenkonto=konto) then
	begin;
		fehler ('Gegenkonto und Konto sind gleich.', ERROR);
	end;
	//Felder auf richtiges Format testen
	datum:=istdatum (datum, 'Datum', 'Datum hat falsches Format.');
	haben:=istzahl (haben,'betrag', 'Habenbetrag hat falsches Format.');
	soll:=istzahl (soll,'betrag', 'Sollbetrag hat falsches Format.');
	gegenkonto:=istzahl (gegenkonto,'Gegenkonto', 'Gegenkonto hat falsches Format.');
	konto:=istzahl (konto,'Konto', 'Konto hat falsches Format.');
	Skonto := istzahl (Skonto, 'Skonto', 'Skonto hat falsches Format');
	KostMenge := istzahl (KostMenge, 'Haben', 'Kostenmenge hat falsches Format');
	Kost1 := istzahl (Kost1, 'Kost', 'Kost1 hat falsches Format');
	Kost2 := istzahl (Kost2, 'Kost', 'Kost2 hat falsches Format');
	Beleg1 := istzahl (Beleg1, 'Beleg', 'Beleg1 hat falsches Format');
	Beleg2 := istzahl (Beleg2, 'Beleg', 'Beleg2 hat falsches Format');

	if (text = '') then
		fehler ('Buchungstext darf nicht leer sein.', ERROR);

 if (datum = '') then
  begin;
	fehler ('Datum fehlt.', ERROR);
  end;
 if (((haben<>'') and (soll<>'')) or ((haben='') and (soll=''))) then
  begin;
	fehler('Haben- und Soll-Betrag sind angegeben.', ERROR);
  end;
 // Empty value for field gegenkonto is not an error any more, datev now supports it.
 //if (gegenkonto='') then
 // begin;
 //	fehler('Gegenkonto ist nicht vorhanden.', ERROR);
 // end;
 if ((zeile=2) and (datum='')) or ((length(datum)<>0) and (length(datum)<>4)) then
  begin;
   fehler('Datum ist erforderlich oder hat falsches Format.', ERROR);
  end;
end;
{lîscht nach jedem importiertem Satz(auch'kein Datensatz' inhalte von satzstrings}
{++++++++++++++++++++FEHLERSUCHE ENDE+++++++++++++++++++++++++++++++++++++}
procedure import.loeschen;
VAR I : integer;
begin;
	soll:='';
	text:='';
	haben:='';
	gegenkonto:='';
	konto:='';
	datum:='';
	beleg1:='';
	beleg2:='';
	Kost1 := '';
	Kost2 := '';
	Skonto := '';
	Waehrung := '';
	KostMenge := '';
	for i:=1 to 255 do datensatz[i]:=0;
end;
{konvertiert einen importierten Datensatz}
procedure import.konvertsatz;
var i,position   : integeR;
begin;
 datensatz[1]:=ord('y');
 position:=2;
 if length(soll)<>0 then
  begin;
	datensatz[2]:=ord('+');
	for i:=1 to length(soll) do begin;
	 inc(position);
    datensatz[position]:=ord(soll[i]);
	end;
  end;
 if length(haben)<>0 then
  begin;
	datensatz[2]:=ord('-');
	for i:=1 to length(haben) do begin;
	 inc(position);
	 datensatz[position]:=ord(haben[i]);
	end;
  end;
 if (length(gegenkonto) > 5) then
 begin;
	inc(position);datensatz[position]:=ord('l');
	while (length(gegenkonto) > 5) do
		begin;
			inc(position);datensatz[position]:=ord(gegenkonto[1]);
			gegenkonto := copy(gegenkonto, 2, length(gegenkonto) - 1);
		end;
 end;
 inc(position);datensatz[position]:=ord('a');
 for i:=1 to length(gegenkonto) do begin;
  inc(position);
  datensatz[position]:=ord(gegenkonto[i]);
 end;

 //inc(position);datensatz[position]:=ord('b');
	if length(beleg1)<>0 then
	begin;
		inc(position);datensatz[position]:=$BD;
		for i:=1 to length(beleg1) do begin;
			inc(position);
			datensatz[position]:=ord(beleg1[i]);
		end;
		inc(position);datensatz[position]:=$1C;
	end;

	if length(beleg2)<>0 then
	begin;
		inc(position);datensatz[position]:=$BE;
		for i:=1 to length(beleg2) do begin;
			inc(position);
			datensatz[position]:=ord(beleg2[i]);
		end;
		inc(position);datensatz[position]:=$1C;
	end;


 if length(datum)<>0 then
  begin;
	inc(position);datensatz[position]:=ord('d');
	for i:=1 to length(datum) do begin;
	 inc(position);
	 datensatz[position]:=ord(datum[i]);
	end;
  end;
 if length(konto)<>0 then
  begin;
	inc(position);datensatz[position]:=ord('e');
	for i:=1 to length(konto) do begin;
	 inc(position);
	 datensatz[position]:=ord(konto[i]);
	end;
  end;

	//Kostenstelle1
	if length(Kost1) <> 0 then
	begin;
		inc(position);datensatz[position]:=$BB;
		for i:=1 to length(Kost1) do begin;
			inc(position);
			datensatz[position]:=ord(Kost1[i]);
		end;
		inc(position);datensatz[position]:=$1C;
	end;

	//Kostenstelle2
	if length(Kost2) <> 0 then
	begin;
		inc(position);datensatz[position]:=$BC;
		for i:=1 to length(Kost2) do begin;
			inc(position);
			datensatz[position]:=ord(Kost2[i]);
		end;
		inc(position);datensatz[position]:=$1C;
	end;

	//Skonto
	if length(Skonto) <> 0 then
	begin;
		inc(position);datensatz[position]:=ord('h');
		for i:=1 to length(Skonto) do begin;
			inc(position);
			datensatz[position]:=ord(Skonto[i]);
		end;
		inc(position);datensatz[position]:=$1C;
	end;

	//Kostenmenge
	if length(KostMenge) <> 0 then
	begin;
		inc(position);datensatz[position]:= $6B;
		for i:=1 to length(KostMenge) do begin;
			inc(position);
			datensatz[position]:=ord(KostMenge[i]);
		end;
		inc(position);datensatz[position]:=$1C;
	end;

	//Buchungstext
 if length(text)<>0 then
  begin;
	inc(position);datensatz[position]:=$1E;
	if length(text)>30 then text[0]:=chr(30);
	for i:=1 to length(text) do begin;
	 inc(position);
	 datensatz[position]:=ord(text[i]);
	end;
  end;

	//W‰hrung
	if length(waehrung) <> 0 then
	begin;
		inc(position);datensatz[position]:=$1C;
		inc(position);datensatz[position]:=$B3;
		for i:=1 to length(waehrung) do begin;
			inc(position);
			datensatz[position]:=ord(waehrung[i]);
		end;
	end;
	//1C B3

 inc(position);datensatz[position]:=$1C;
 datensatz[position+1]:=0;
end;
{importiert einen Datensatz}
const testerw : integer = 0;
procedure import.readAndExportOneDataset;
var i,i1  : integeR;
    intDatum, intStart, intEnde : integer;
    additionsIndex : integer;
    additions : array[0..maxAdditionColIndex] of string;
begin;
  loeschen;
  zeile:=zeile+1;
  if (zeile>=zeilanz) then
  begin;
	  exportsumme;
	  Ende:=true;
	  exit;
  end;

  // read one single line from the excel table.
  text:='';
  for i:=0 to maxAdditionColIndex do
  begin
    additions[i] := '';
  end;
  for i:=1 to 20 do begin;
    if (lesestring='kein Datensatz') or (lesestring='EndeDatensatz') then break;
    // get next col from cache.
    lesen;

    if (i = datu[1]) then datum:=lesestring;
    if (i = habe[1]) then haben:=lesestring;
    if (i = sol[1]) then soll:=lesestring;
    if (i = gegenkont[1]) then gegenkonto:=lesestring;
    if (i = kont[1]) then Konto:=lesestring;
    if (i = bele1[1]) then
    begin
      Beleg1:=lesestring;
    end;
    if (i = bele2[1]) then beleg2:=lesestring;
    if (i = kos1[1]) then Kost1:=lesestring;
    if (i = kos2[1]) then Kost2:=lesestring;
    if (i = skont[1]) then Skonto:=lesestring;
    if (i = waehrun[1]) then Waehrung := lesestring;
    if (i = KostMeng [1]) then KostMenge := lesestring;

    // read Text-Columns.
    i1:=0;
    repeat;
	    inc(i1);
	    if (tex[i1]=i) and (length(text)<30) then text:=text+lesestring;
    until (i1=5);

    // read additional information.
    for additionsIndex := 0 to additionColCnt do
    begin
      if (i = additionCols[additionsIndex]) then
        additions[additionsIndex] := 'y' + char($B7) + additionNames[additionsIndex]
               + char($1C) + char($B8) + lesestring + char($1C);
    end;

    // if the table has less then 20 columns we have to break here.
    if zeilenende=1 then break;
   end;
   if (UpperCase(lesestring) <> 'KEIN DATENSATZ')
	   and (UpperCase(lesestring) <> 'ENDEDATENSATZ') then
   begin;
	   fehlersuchen;

	//Test auf das Datum - muss erweitert werden
	intDatum := stringzahl(copy(datum,3,2)) * 100 + stringzahl(copy(datum, 1, 2));
	intStart := stringzahl(copy(startzeitraum,3,2)) * 100 + stringzahl(copy(startzeitraum, 1, 2));
	intEnde := stringzahl(copy(endezeitraum,3,2)) * 100 + stringzahl(copy(endezeitraum, 1, 2));
	if ((intDatum >= intStart)	and (intDatum <= intEnde)) then
    begin;
     if errors=false then
       begin;
        konvertsatz;
        rechnen;
        exportDataset;
        exportAdditions(additions);
       end;
    end
   else
    begin;
     fehler('Datum liegt nicht im Erfassungszeitraum. Beleg wird nicht mit exportiert.', WARNING);
    end;
  end;
  // if the excel-table has more then 20 cols, then we need to read and ignore the remaining cols here.
 repeat;
   if zeilenende<>1 then lesen;
 until zeilenende=1;
end;
{importiert tabellenkopf}
procedure import.importsatz;
var i,i1,additionsPos    : integer;
	 fehlerstring : string;
begin;
 {Werte fÅr importieren initialisieren}
 Ende:=false;
 Zeile:=1;
 //alle nicht notwendigen Arraywerte zur¸cksetzen
 bele1[1]:=0;
 bele2[1]:=0;
 for i:=1 to maxAdditionColIndex do tex[i]:=0;
 //Spalten suchen
 for i:=1 to spaltenanz do begin;
  lesen;
  if LowerCase(lesestring)='datum' then datu[1]:=i;
  if LowerCase(lesestring)='haben' then habe[1]:=i;
  if LowerCase(lesestring)='soll'  then sol[1]:=i;
  if LowerCase(lesestring)='gegenkonto' then gegenkont[1]:=i;
  if LowerCase(lesestring)='konto' then kont[1]:=i;
  if LowerCase(lesestring)='beleg1' then bele1[1]:=i;
  if LowerCase(lesestring)='beleg2' then bele2[1]:=i;
  if LowerCase(lesestring)='kost1' then Kos1[1]:=i;
  if LowerCase(lesestring)='kost2' then Kos2[1]:=i;
  if LowerCase(lesestring)='skonto' then Skont[1]:=i;
  if LowerCase(lesestring)='waehrung' then waehrun[1]:=i;
  if LowerCase(lesestring)='kostmenge' then KostMeng[1]:=i;
  if LowerCase(lesestring)='text'  then
	begin;
	 i1:=0;
	 repeat;
	  inc(i1);
	 until (tex[i1] = 0) or (i1=5);
	 if (tex[i1] = 0) then tex[i1]:=i;
	end;
  additionsPos := Pos(colAddionsPrefix, LowerCase(lesestring));
  if (additionsPos = 1) and (additionColCnt <= maxAdditionColIndex) then
	begin
    additionCols[additionColCnt] := i;
    additionNames[additionColCnt] := Copy(lesestring, Length(colAddionsPrefix) + 1, Length(lesestring) - Length(colAddionsPrefix));
    additionColCnt := additionColCnt + 1;
  end;
  if zeilenende = 1 then break;
 end;
 repeat;if zeilenende<>1 then lesen;until zeilenende=1;
 {FehlerÅberprÅfung}
 fehlerstring:='';
 if (datu[1]=0) then fehlerstring:='Datum';
 if (habe[1]=0) then fehlerstring:='Haben';
 if (sol[1]=0) then fehlerstring:='Soll';
 if (gegenkont[1]=0) then fehlerstring:='Gegenkonto';
 if (kont[1]=0) then fehlerstring:='Konto';
 if (tex[1]=0) then fehlerstring:='Text';
 if fehlerstring<>'' then fehler('Fehler im Tabellenkopf (Spalte '+fehlerstring+'nicht gefunden).', SERROR);
 repeat;
  inc(zeile);
  lesen;
  case zeile of
	2 : begin;
			mandant:=lesestring;
      // Check Mandant-Number for integer
			try
				StrToInt (mandant);
			except
				on EConvertError do
					fehler ('Mandantennummer ' + mandant + ' ist keine g¸ltige Zahl.', ERROR);
			end;
			while (Length(mandant) < 5) do
				mandant := '0' + mandant;
		 end;
	3 : begin
			bearbeiter:=lesestring;
			if (Length (bearbeiter) <> 2) then
				fehler('Die Bearbeiter-ID muss genau 2 Zeichen lang sein. ', SERROR);
		 end;
  end;
  repeat;lesen;until zeilenende=1;
 until zeile=3;
end;

procedure import.fortschritt;
begin;
 //Werte werden an Windowsfenster angepaﬂt und weitergeleitet
 PostMessage(Handle,55555,zeile,zeilanz);
end;
{***************************Hauptprogramm************************}
constructor import.create(handl : integer;quell : string;clien,clien1 : TDDeclientconv;templateDir,saveDir,beraternr : string);
begin;
 handle := handl;
 client := clien;
 client1 := clien1;
 quelle := quelle;
 beraterNummer := beraternr;
 fehlermeldungen := TStringList.Create;
 fehlertext := '';
 fehlercap := '';
 errors := false;
 warnings := false;
 infos := false;
 ShowErrorLog:=true;
 Self.templateDir := templateDir;
 Self.saveDir := saveDir;
  additionColCnt := 0;
end;

procedure import.execute;
var fehl       : boolean;
begin;
 //Statuswert f¸r Elternprozeﬂ
 fertig:=false;
 fehl:=false;
 try
	 {Speicher belegen}
	 speicherholen;
	 {Dateien oeffnen}
	 //TThread
	 oeffneeingabe;
	 oeffneausgabe;
	 {Startfenster oeffnen}
	 {Tabellenkopf importieren und in Speicher exportieren}
	 importsatz;
	 {Eingabe der Kopfdaten}
	 {Jahr zuschreiben}
	 startzeitraum:=startzeitraum+copy(jahr,1,2);
	 endezeitraum:=endezeitraum+copy(jahr,1,2);
	 vorlauf:=vorlauf+copy(jahr,1,2);
	 {Exportieren der Kopfdaten}
	 exportspeicherverzeichnis;
	 exportspeicherdaten;
	 repeat;
     readAndExportOneDataset;
     fortschritt;
	 until (Ende=true);
	 {Daten abspeichern}
	 //TThread
	 {synchronize(}schreibedaten;
	 {synchronize(}schreibeverzeichnis;
   except
    on fehler : EConvertError do
     begin;
      fehlertext:='Bei der Konvertierung sind Fehler augetreten.'+
                  'Bitte beheben Sie diese Fehler.';
		//TThread
		box;
		fehl:=true;
	  end;
	 on e : Exception do
	 begin
		fehlertext := 'Es ist ein ungekanter Fehler aufgetreten: ' + e.Message;
		box;
		fehl := true;
	 end;
 end;
 {Speicher freigeben}
 //speichergeben;
 {Dateien schliessen}
 schliesseausgabe;
 {Eingabebildschirm schliessen}
 PostMessage(Handle,55556,0,0);
 PostMessage(Handle,55555,0,zeilanz);
 //falls vorher schon Exception ausgelˆst wurde
 if (fehl=true) then
	begin;
	 fertig:=true;
	 //TThreadTerminate;
	 exit;
	end;
 fehlercap:='Fehler';
 if ((warnings=true)and(errors=true))then
	  fehlertext := 'Es sind Warnungen und Fehler aufgetreten. Bitte sehen Sie im Fehlerprotokoll nach.'
 else if (errors=true) then
	  fehlertext := 'Es sind Fehler aufgetreten. Bitte sehen Sie im Fehlerprotokoll nach.'
 else if (warnings=true) then
	  fehlertext := 'Es sind Warnungen aufgetreten. Bitte sehen Sie im Fehlerprotokoll nach.'
 else if (infos=true) then
	  fehlertext := 'Es sind Meldungen aufgetreten. Bitte sehen Sie im Protokoll nach.'
 else
  begin;
	fehlertext:='Konvertierung erfolgreich beendet.';
	fehlercap:='Fertig';
	ShowErrorLog:=false;
  end;
 box;
 fertig:=true;
end;
end.