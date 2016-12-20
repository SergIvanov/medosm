unit fys;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Mask, DBCtrlsEh, ExtCtrls, DBGridEh, DBLookupEh,
  DBGridEhGrouping, GridsEh, DBCtrls, ComCtrls;

type
  TfFYS = class(TForm)
    Panel1: TPanel;
    Panel2: TPanel;
    DBEditEh1: TDBEditEh;
    DBEditEh2: TDBEditEh;
    DBEditEh3: TDBEditEh;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    DBLookupComboboxEh1: TDBLookupComboboxEh;
    DBLookupComboboxEh2: TDBLookupComboboxEh;
    Label4: TLabel;
    Label5: TLabel;
    Edit1: TEdit;
    Edit2: TEdit;
    Label6: TLabel;
    Label7: TLabel;
    Label8: TLabel;
    DBGridEh1: TDBGridEh;
    Panel3: TPanel;
    Button1: TButton;
    DBNavigator1: TDBNavigator;
    Label9: TLabel;
    DBLookupComboBox1: TDBLookupComboBox;
    DBDateTimeEditEh1: TDBDateTimeEditEh;
    Label10: TLabel;
    Label11: TLabel;
    DBEditEh5: TDBEditEh;
    DateTimePicker1: TDateTimePicker;
    Button2: TButton;
    DateTimePicker2: TDateTimePicker;
    Label12: TLabel;
    procedure DBLookupComboboxEh1Change(Sender: TObject);
    procedure DBLookupComboboxEh2Change(Sender: TObject);
    procedure Edit1Change(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormShow(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure DateTimePicker2Exit(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  fFYS: TfFYS;

implementation

uses DM;

{$R *.dfm}

procedure TfFYS.Button1Click(Sender: TObject);
begin
  fDM.TUsl.Insert;
  fDM.TUsl.FieldByName('kod_usl').AsString :=
    fDM.TFYS.FieldByName('kod_usl').AsString;
  fDM.TUsl.FieldByName('naimusl').AsString :=
    fDM.TFYS.FieldByName('Name_usl').AsString;
  fDM.TUsl.FieldByName('kol').AsString := Edit1.Text;
  fDM.TUsl.FieldByName('cena').AsString := Edit2.Text;
  fDM.TUsl.FieldByName('data').AsString := FormatDateTime('dd.mm.yyyy',
    DateTimePicker1.Date);
  fDM.TUsl.FieldByName('kod_vr').AsString :=
    fDM.TFYS.FieldByName('kod_vr').AsString;
  fDM.TUsl.FieldByName('kod_ms').AsString :=
    fDM.TFYS.FieldByName('kod_ms').AsString;
  fDM.TUsl.FieldByName('mkb').AsString :=
    'Z00.0';
  fDM.TUsl.Post;
end;

procedure TfFYS.Button2Click(Sender: TObject);
begin
  fDM.Query2.Close;
  fDM.Query2.SQL.Clear;
  fDM.Query2.SQL.Add('SELECT ser, nom FROM polisdms WHERE (fam=' +
    quotedstr(fDM.TRab.FieldByName('fam').AsString) + ') AND (im=' +
    quotedstr(fDM.TRab.FieldByName('imya').AsString) + ') AND (ot=' +
    quotedstr(fDM.TRab.FieldByName('fath').AsString) + ') AND (dr=:date)');
  fDM.Query2.Parameters.ParamByName('date').Value :=
    FormatDateTime('dd.mm.yyyy', fDM.TRab.FieldByName('data_r').AsDateTime);
  fDM.Query2.Open;

  if not(fDM.TRab.Modified) then
    fDM.TRab.Edit;
  if fDM.Query2.RecordCount > 0 then
    fDM.TRab.FieldByName('polis').AsString := fDM.Query2.FieldByName('ser')
      .AsString + '�' + fDM.Query2.FieldByName('nom').AsString;
  fDM.TRab.Post;
end;

procedure TfFYS.DateTimePicker2Exit(Sender: TObject);
begin
  fDM.TRab.Edit;
  fDM.TRab.FieldByName('date_of').AsString := FormatDateTime('dd.mm.yyyy',
    DateTimePicker2.Date);
  fDM.TRab.Post;
end;

procedure TfFYS.DBLookupComboboxEh1Change(Sender: TObject);
begin
  fDM.TFYS.Active := false;
  fDM.TFYS.CommandText := 'SELECT * FROM FYS WHERE idotdel=' +
    fDM.TOtdel.FieldByName('id').AsString;
  fDM.TFYS.Active := true;
end;

procedure TfFYS.DBLookupComboboxEh2Change(Sender: TObject);
begin
  Edit2.Text := fDM.TFYS.FieldByName('Zena').AsString;
end;

procedure TfFYS.Edit1Change(Sender: TObject);
begin
  if Edit1.Text <> '' then
    Edit2.Text := IntToStr(StrToInt(Edit1.Text) * fDM.TFYS.FieldByName('Zena')
      .AsInteger);
end;

procedure TfFYS.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  fDM.TRab.Edit;
  fDM.TRab.FieldByName('date_of').AsString := FormatDateTime('dd.mm.yyyy',
    DateTimePicker2.Date);
  fDM.TRab.Post;

end;

procedure TfFYS.FormShow(Sender: TObject);
begin



  DateTimePicker1.Date := fDM.TRab.FieldByName('data_proh').AsDateTime;
  if fDM.TRab.FieldByName('date_of').AsString = '' then
    DateTimePicker2.Date := fDM.TRab.FieldByName('data_proh').AsDateTime
  else
    DateTimePicker2.Date := fDM.TRab.FieldByName('date_of').AsDateTime;
  fDM.Query2.Close;
  fDM.Query2.SQL.Clear;
  fDM.Query2.SQL.Add('SELECT ser, nom FROM polisdms WHERE (fam=' +
    quotedstr(fDM.TRab.FieldByName('fam').AsString) + ') AND (im=' +
    quotedstr(fDM.TRab.FieldByName('imya').AsString) + ') AND (ot=' +
    quotedstr(fDM.TRab.FieldByName('fath').AsString) + ') AND (dr=:date)');
  fDM.Query2.Parameters.ParamByName('date').Value :=
    FormatDateTime('dd.mm.yyyy', fDM.TRab.FieldByName('data_r').AsDateTime);
  fDM.Query2.Open;

  if not(fDM.TRab.Modified) then
    fDM.TRab.Edit;
  if fDM.Query2.RecordCount > 0 then
    fDM.TRab.FieldByName('polis').AsString := fDM.Query2.FieldByName('ser')
      .AsString + '�' + fDM.Query2.FieldByName('nom').AsString;
  fDM.TRab.Post;

end;

end.
