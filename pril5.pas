unit pril5;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls, Grids, DBGrids, Mask, DBCtrls, WordXP,
  OleServer, ComCtrls, Data.DB;

type
  TfPril3 = class(TForm)
    Panel1: TPanel;
    Label1: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    Label6: TLabel;
    Label7: TLabel;
    DBEdit1: TDBEdit;
    DBEdit2: TDBEdit;
    DBEdit3: TDBEdit;
    Label8: TLabel;
    Label9: TLabel;
    DBEdit4: TDBEdit;
    DBEdit5: TDBEdit;
    DBEdit6: TDBEdit;
    DBGrid1: TDBGrid;
    Label12: TLabel;
    Edit2: TEdit;
    Button1: TButton;
    Button2: TButton;
    DBEdit8: TDBEdit;
    WordApplication1: TWordApplication;
    WordDocument1: TWordDocument;
    Label10: TLabel;
    DBEdit9: TDBEdit;
    DBMemo1: TDBMemo;
    Button3: TButton;
    Button4: TButton;
    DBComboBox1: TDBComboBox;
    Label11: TLabel;
    Button5: TButton;
    Edit1: TEdit;
    DBMemo2: TDBMemo;
    ListBox1: TListBox;
    Label13: TLabel;
    Label2: TLabel;
    Label14: TLabel;
    procedure Edit2Change(Sender: TObject);
    procedure DBGrid1DblClick(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure Button4Click(Sender: TObject);
    procedure Button5Click(Sender: TObject);
    procedure Edit1Change(Sender: TObject);
    procedure ListBox1DblClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  fPril3: TfPril3;
  s, scode: string;
  mode: boolean;

implementation

uses DM, spravorg, editor, addmkb, factor;

{$R *.dfm}

procedure TfPril3.Edit1Change(Sender: TObject);
begin
  fDM.MKBLocate(Edit1.Text);
end;

procedure TfPril3.Edit2Change(Sender: TObject);
begin
  // ���� � ���� Edit1 ���� ���� ���� �����,
  if Edit2.Text <> '' then
  begin
    fDM.TMKB.Filtered := False; // ��������� ������
    ed := Edit2.Text; // �������� � fDM ����� �����
    fDM.TMKB.Filtered := True; // �������� ������
  end
  // ���� ���� ���, ���������� ���������:
  else
    fDM.TMKB.Filtered := False;
end;

procedure TfPril3.ListBox1DblClick(Sender: TObject);
var
  index: integer;
begin
  index := ListBox1.ItemIndex;
  fDM.TRab.Edit;
  fDM.TRab.FieldByName('rekom').AsString := fDM.TRab.FieldByName('rekom')
    .AsString + ' ' + ListBox1.Items[index];
end;

procedure TfPril3.DBGrid1DblClick(Sender: TObject);

begin
  if not(fDM.TRab.FieldValues['mkb'] = Null) then
  begin
    if MessageDlg('�������� �����������?', mtConfirmation, [mbYes, mbNo], 0)
      = mrYes then
    begin
      fDM.TRab.Edit;
      fDM.TRab.FieldByName('mkb').AsString := fDM.TMKB.FieldValues['NAIM'];
      fDM.TRab.FieldByName('mkb_code').AsString := fDM.TMKB.FieldValues['CODE'];
      fDM.TRab.Post;
    end;
  end
  Else
  begin
    fDM.TRab.Edit;
    fDM.TRab.FieldByName('mkb').AsString := fDM.TMKB.FieldValues['NAIM'];
    fDM.TRab.FieldByName('mkb_code').AsString := fDM.TMKB.FieldValues['CODE'];
    fDM.TRab.Post;
  end;
end;

procedure TfPril3.Button1Click(Sender: TObject);
begin
  fDM.TRab.Edit;
  fDM.TRab.Post;
  fPril3.Close;
end;

procedure TfPril3.Button2Click(Sender: TObject);
var
  Shablon: OleVariant;
  sfact, sfactn, sissl, svr:string;
begin
  Shablon := ExtractFileDir(ParamStr(0)) + '\�������\pril3.doc';
  WordApplication1.Connect;
  WordApplication1.Visible := True;
  WordApplication1.Documents.Add(Shablon, EmptyParam, EmptyParam, EmptyParam);
  WordDocument1.ConnectTo(WordApplication1.ActiveDocument);

  WordApplication1.Visible := True;

  fEditor.WordGotoBookmark('f');
  fEditor.WordInsertText(fDM.TRab.FieldByName('fam').AsString);

  fEditor.WordGotoBookmark('i');
  fEditor.WordInsertText(fDM.TRab.FieldByName('imya').AsString);

  fEditor.WordGotoBookmark('o');
  fEditor.WordInsertText(fDM.TRab.FieldByName('fath').AsString);

  fEditor.WordGotoBookmark('org');
  fEditor.WordInsertText(fDM.TOrg.FieldByName('imya_poln').AsString);

  fEditor.WordGotoBookmark('strukt_p');
  fEditor.WordInsertText(fDM.TRab.FieldByName('imya_strukt_podr').AsString);

  fEditor.WordGotoBookmark('prof');
  fEditor.WordInsertText(fDM.TRab.FieldByName('profes').AsString);

//  fFactor.MakeStr(sfact, sfactn, sissl, svr);
//  fEditor.WordGotoBookmark('vred_f');
//  fEditor.WordInsertText(sfact);
//
   fEditor.WordGotoBookmark('vred_f');
   fEditor.WordInsertText(fDM.TRab.FieldByName('vredn_fact').AsString);

  fEditor.WordGotoBookmark('mkb');
  fEditor.WordInsertText(fDM.TRab.FieldByName('mkb').AsString);

  fEditor.WordGotoBookmark('disp_gr');
  fEditor.WordInsertText(fDM.TRab.FieldByName('disp_gr').AsString);

  fEditor.WordGotoBookmark('rekom');
  fEditor.WordInsertText(fDM.TRab.FieldByName('rekom').AsString);

  WordApplication1.Disconnect;
end;

procedure TfPril3.Button3Click(Sender: TObject);
var
  query: string;
begin
  if MessageDlg('�������� ���10 � ����������?', mtConfirmation, [mbYes, mbNo],
    0) = mrYes then
  begin
    fDM.TRab.Edit;
    fDM.TRab.FieldByName('mkb').AsString := '';
    fDM.TRab.FieldByName('mkb_code').AsString := '';
    fDM.TRab.FieldByName('firstb').AsString := '';
    fDM.TRab.Post;
  end;
end;

procedure TfPril3.Button4Click(Sender: TObject);
begin
  Close;
end;

procedure TfPril3.Button5Click(Sender: TObject);
begin
  fAddMkb.ShowModal;
end;

end.

