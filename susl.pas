unit susl;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, DBGridEhGrouping, GridsEh, DBGridEh, DBCtrls;

type
  TfSUsl = class(TForm)
    Panel1: TPanel;
    Panel2: TPanel;
    DBGridEh1: TDBGridEh;
    DBNavigator1: TDBNavigator;
    procedure FormShow(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  fSUsl: TfSUsl;

implementation

uses DM;

{$R *.dfm}

procedure TfSUsl.FormShow(Sender: TObject);
begin
 if not(fDM.TFYS.Active) then
 begin
   fDM.TFYS.Active:=true;
 end;

end;

end.
