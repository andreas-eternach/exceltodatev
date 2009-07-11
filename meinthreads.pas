unit meinthreads;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, DdeMan,nesy24b, ComCtrls;
type
MeinThread=class(TThread)
public
s : string;
client : TDDeClientConv;
k : import;
Handle : integer;
constructor create(Handle1 : integer;k1 : import;s1 : string;client1 : TDDeclientConv);virtual;
procedure execute;override;
private
end;


implementation

constructor MeinThread.create;
begin;
 s:=s1;
 client:=client1;
 Handle:=Handle1;
 k:=k1;
 inherited create(true);
 priority:=tpNormal;
 Resume;
end;

procedure MeinThread.execute;
begin;
 try
  k.starten(Handle,s,client);
 except
  on e : EConvertError do
   begin;
    Application.MessageBox(pchar(e.Message),'Fehler',IDOK);
    Terminate;
   end;
 end;
end;

end.
 