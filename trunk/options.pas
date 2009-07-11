unit options;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, FileCtrl;

type
  TOptionDialog = class(TForm)
    Label1: TLabel;
    Label2: TLabel;
    edTemplateDir: TEdit;
    edSaveDir: TEdit;
    btnSelectTemplateDir: TButton;
    btnSelectSaveDir: TButton;
    procedure FormCreate(Sender: TObject);
    procedure btnSelectSaveDirClick(Sender: TObject);
    procedure btnSelectTemplateDirClick(Sender: TObject);
    procedure selSaveDirCanClose(Sender: TObject; var CanClose: Boolean);
  private
    { Private-Deklarationen }
  public
    { Public-Deklarationen }
    function getTemplateDir : String;
    function getSaveDir : String;
  end;

var
  OptionDialog: TOptionDialog;

implementation

{$R *.dfm}

function TOptionDialog.getTemplateDir : String;
begin
  Result := edTemplateDir.Text;
end;

function TOptionDialog.getSaveDir : String;
begin
  Result := edSaveDir.Text;
end;

procedure TOptionDialog.selSaveDirCanClose(Sender: TObject;
  var CanClose: Boolean);
begin
  // check whether a directory is selected and contains the template files.
end;

procedure TOptionDialog.btnSelectTemplateDirClick(Sender: TObject);
var newPath:string;
begin
  newPath:= edTemplateDir.Text;

  // remove missing directory
  if ( not DirectoryExists(newPath)) then
  begin;
    newPath :='';
    edTemplateDir.Text:='a:\nesy';
  end;

  // open the select-dialog.
  if (SelectDirectory(newPath, [], 0)) then
    edTemplateDir.Text := newPath;

end;

procedure TOptionDialog.btnSelectSaveDirClick(Sender: TObject);
var newPath : string;
begin
  newPath:= edTemplateDir.Text;

  // remove missing directory
  if (not DirectoryExists(newPath)) then
  begin
    newPath :='';
    edSaveDir.Text:='a:\';
  end;

  // open the select-dialog.
  if (SelectDirectory(newPath, [], 0)) then
    edSaveDir.Text:=newPath;
end;

procedure TOptionDialog.FormCreate(Sender: TObject);
begin
  edTemplateDir.Text:= GetCurrentDir() + '\diskette\nesy';
end;

end.
