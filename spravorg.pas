unit spravorg;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Mask, DBCtrls, ExtCtrls, Grids, DBGrids, DBGridEhGrouping,
  GridsEh, DBGridEh, ToolCtrlsEh, DBGridEhToolCtrls, DynVarsEh, EhLibVCL,
  DBAxisGridsEh,System.UITypes;

type
  TfSpravOrg = class(TForm)
    Panel1: TPanel;
    Button1: TButton;
    Button2: TButton;
    DBGridEh1: TDBGridEh;
    procedure DBGrid1DblClick(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
   // procedure Button3Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  fSpravOrg: TfSpravOrg;

implementation

uses DM, editor;

{$R *.dfm}

procedure TfSpravOrg.DBGrid1DblClick(Sender: TObject);
begin
 if MessageDlg('Изменить организацию?',mtConfirmation, [mbYes, mbNo], 0) = mrYes
 then
  begin
   fDM.TRab.Edit;
   fDM.TRab.FieldbyName('id_org').asinteger:=fDM.TOrgNew.FieldValues['id_org'];
   fDM.TRab.Post;
   fSpravOrg.Close;
  end;
end;

procedure TfSpravOrg.Button1Click(Sender: TObject);
begin
 if MessageDlg('Изменить организацию?',mtConfirmation, [mbYes, mbNo], 0) = mrYes
 then
  begin
   fDM.TRab.Edit;
   fDM.TRab.FieldbyName('id_org').asinteger:=fDM.TOrgNew.FieldValues['id_org'];
   fDM.TRab.Post;
   fSpravOrg.Close;
  end;
end;

procedure TfSpravOrg.Button2Click(Sender: TObject);
begin
 fSpravOrg.Close;
end;

end.
