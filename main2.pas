unit main;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls, Buttons, Menus, Grids, DBGrids, ComCtrls,
  Mask, DBCtrls, DBGridEhGrouping, GridsEh, DBGridEh, EhLibADO;

type
  TfMain = class(TForm)
    Panel1: TPanel;
    Panel2: TPanel;
    Bevel1: TBevel;
    Button3: TButton;
    Button4: TButton;
    MainMenu1: TMainMenu;
    N1: TMenuItem;
    N2: TMenuItem;
    N3: TMenuItem;
    N4: TMenuItem;
    Button1: TButton;
    N5: TMenuItem;
    N101: TMenuItem;
    N6: TMenuItem;
    N9: TMenuItem;
    DBGrid1: TDBGridEh;
    procedure Button4Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure DBGrid1DblClick(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure N4Click(Sender: TObject);
    procedure N5Click(Sender: TObject);
    procedure N9Click(Sender: TObject);
    procedure DBGrid1DrawColumnCell(Sender: TObject; const Rect: TRect;
      DataCol: Integer; Column: TColumn; State: TGridDrawState);
    procedure DBGrid1KeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure N6Click(Sender: TObject);
    procedure DBGridEh1DblClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  fMain: TfMain;

implementation

uses DM, editor, newsotr, edspravorg, pril9;

{$R *.dfm}

procedure TfMain.Button4Click(Sender: TObject);
begin
  fDM.TRab.Append;
  fEditor.ShowModal;
end;

procedure TfMain.Button3Click(Sender: TObject);
begin
 fEditor.ShowModal;
end;

procedure TfMain.DBGrid1DblClick(Sender: TObject);     //Удалить!!!!!!!!!!
begin
 fEditor.ShowModal;
end;

procedure TfMain.DBGrid1DrawColumnCell(Sender: TObject; const Rect: TRect;
  DataCol: Integer; Column: TColumn; State: TGridDrawState);
var
  s: String;
begin
  if Column.Index = 1 then begin
    if fDm.TRab['fam'] <> Null then
      s:= fDm.TRab['fam'] + ' ';
    if fDm.TRab['Imya'] <> Null then
      s:= s + Copy(fDm.TRab['Imya'], 1, 1) + '.';
    if fDm.TRab['fath'] <> Null then
      s:= s + Copy(fDm.TRab['fath'], 1, 1)+ '.';
   // DBGrid1.Canvas.TextOut(Rect.Left, Rect.Top, s);
  end;
end;


procedure TfMain.DBGrid1KeyDown(Sender: TObject; var Key: Word;    //Управление с клавиатуры
  Shift: TShiftState);
begin
if Key=46 Then   //delete
 begin
 if MessageDlg('Удалить запись ?',mtConfirmation, [mbYes, mbNo], 0) = mrYes then
  begin
   fDM.TRab.Delete;
 end;
 end;
if Key=13 Then   //edit
 begin
  fEditor.ShowModal;
 end;

end;


procedure TfMain.DBGridEh1DblClick(Sender: TObject);
begin
 fEditor.ShowModal;
end;


procedure TfMain.Button1Click(Sender: TObject);
var query:string;
begin
 if MessageDlg('Удалить запись? ',mtConfirmation, [mbYes, mbNo], 0) = mrYes then
  begin
   fDM.TRab.Delete;
 end;
end;

procedure TfMain.N4Click(Sender: TObject);
begin
fNewSotr.Showmodal;
end;

procedure TfMain.N5Click(Sender: TObject);
begin
 fEdSpravOrg.Showmodal;
end;

procedure TfMain.N6Click(Sender: TObject);
begin
 Close;
end;

procedure TfMain.N9Click(Sender: TObject);
begin
 fPril9.Showmodal;
end;

end.
