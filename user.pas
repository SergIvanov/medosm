unit user;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, DBCtrls, ExtCtrls, DateUtils;

type
  TfUser = class(TForm)
    Panel1: TPanel;
    Label1: TLabel;
    Button1: TButton;
    ComboBox1: TComboBox;
    lbl1: TLabel;
    edt1: TEdit;
    btnClose: TButton;
    procedure Button1Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure btnCloseClick(Sender: TObject);
    procedure edt1KeyPress(Sender: TObject; var Key: Char);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  fUser: TfUser;
  us:string;
implementation

uses DM, RepUsl22,FormEkonom;

{$R *.dfm}

procedure TfUser.btnCloseClick(Sender: TObject);
begin

 fUser.Close;

end;

Function AutentUser(Login : string;Pass : string) : BOOL;
begin
fDM.qryUser.Active:=False;
fDM.qryUser.Parameters.ParamByName('User').Value:=Login;
fDM.qryUser.Parameters.ParamByName('Pass').Value:=Pass;
fDM.qryUser.Active:=True;

if fDM.qryUser.RecordCount<>0 then Result := True
else
 Result := False;
end;

procedure TfUser.Button1Click(Sender: TObject);
var i:integer;
begin
  us:='';
  if AutentUser(ComboBox1.Text,edt1.Text)=False  then
  ShowMessage('�� ��������� �����, ���������� ��� ���!')
  //fUser.Close
  else
  begin
     if ComboBox1.Text='���������' then
      begin
      fUser.Hide;
//      RepUsl22.FormRepUsl22.ShowModal;
      FormEkonom.frmEkonom.ShowModal;
     end;

     if ComboBox1.Text='�������������' then
      begin
        us:=ComboBox1.Text;
        fDm.TRab.Active:=false;
        fDm.TRab.CommandText:='select * from rabotnik';
        fDm.TRab.Active:=true;
        fUser.Close;
      end else
      begin
      i:=ComboBox1.ItemIndex;
     us:=ComboBox1.Items[i];
     fDm.TRab.Active:=false;
     fDm.TRab.CommandText:='select * from rabotnik where user=('+quotedstr(us)+') AND (data_proh BETWEEN :date1 AND :date2)';
     fDM.TRab.Parameters.ParamByName('date1').Value:=StartOfTheMonth(Date);
     fDM.TRab.Parameters.ParamByName('date2').Value:=EndOfTheMonth(Date);
     fDm.TRab.Active:=true;
     fUser.Close;
     end;
  end;
end;
procedure TfUser.edt1KeyPress(Sender: TObject; var Key: Char);
begin

if Key = #13 then
Button1.Click;

end;

procedure TfUser.FormCreate(Sender: TObject);
var f: TextFile;
    str,dir: string;
    i:integer;
begin
 dir:=ExtractFileDir(ParamStr(0))+'\users.ini';
 AssignFile(f, dir);
 reset(f);
 i:=1;
while not eof(f) do begin
 readln(f, str);
 Combobox1.Items[i]:=str;
 inc(i);
end;
closefile(f);
end;

end.
