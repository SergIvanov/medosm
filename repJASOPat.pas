unit repJASOPat;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ComCtrls, ExtCtrls, CheckLst, DBGridEh, Mask, DBCtrlsEh,
  DBLookupEh, ShellApi, OleServer, ComObj;

type
  TfRepJASOPat = class(TForm)
    CheckListBox1: TCheckListBox;
    Panel1: TPanel;
    Panel2: TPanel;
    DateTimePicker1: TDateTimePicker;
    DateTimePicker2: TDateTimePicker;
    Label1: TLabel;
    до: TLabel;
    CheckBox1: TCheckBox;
    Button1: TButton;
    DBLookupComboboxEh1: TDBLookupComboboxEh;
    procedure CheckBox1Click(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    function getId(s: String): String;
    procedure FormCreate(Sender: TObject);

  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  fRepJASOPat: TfRepJASOPat;

implementation

uses DM,DateUtils;

{$R *.dfm}

function TfRepJASOPat.getId(s: String): String;
var
  i: integer;
  id, sId: string;
begin
  for i := 1 to length(s) do
  begin
    if (s[i] = ' ') then
    begin
      sId := Copy(s, 1, i - 1);
      break;
    end;
  end;
  Result := sId;
end;

procedure TfRepJASOPat.Button1Click(Sender: TObject);
var
  ExelTab: Variant;
  st_in, st, q, sId: string;
  i, k, sum, sumall, c: integer;
begin
  st_in := ExtractFilePath(Application.ExeName);
  ExelTab := CreateOleObject('Excel.Application');
  ExelTab.Workbooks.Open[st_in + '\Шаблоны\jaso.xls'];

  for i := 0 to CheckListBox1.Items.Count - 1 do
  begin

    If CheckListBox1.Checked[i] Then
    begin
      if sId = '' then
        sId := getId(CheckListBox1.Items.Strings[i])
      else
        sId := sId + ', ' + getId(CheckListBox1.Items.Strings[i]);
    end;
  end;

  fDM.QStacRep.SQL.Clear;
  q := 'SELECT rabotnik.data_proh,rabotnik.id,rabotnik.fam, rabotnik.imya, rabotnik.fath, rabotnik.data_r, rabotnik.polis '
    + 'FROM rabotnik WHERE rabotnik.id_org = ' + fDM.TOrg.FieldByName('id_org')
    .AsString + ' AND (date_of BETWEEN :date AND :date2) AND (rabotnik.id in(' +
    sId + '))';
  fDM.QStacRep.SQL.Add(q);
  fDM.QStacRep.Parameters.ParamByName('date').Value :=
    FormatDateTime('dd.mm.yyyy', DateTimePicker1.Date);
  fDM.QStacRep.Parameters.ParamByName('date2').Value :=
    FormatDateTime('dd.mm.yyyy', DateTimePicker2.Date);
  fDM.QStacRep.Open;

  i := 10;
  k := 0;
  c := 0;
  sum := 0;
  sumall := 0;
  fDM.QStacRep.First;
  while not(fDM.QStacRep.Eof) do
  begin
    inc(k);
    ExelTab.Cells.Item[i, 1] := k;
    ExelTab.Cells.Item[i, 2] := fDM.QStacRep.FieldByName('fam').AsString + ' ' +
      fDM.QStacRep.FieldByName('imya').AsString +' '+
      fDM.QStacRep.FieldByName('fath').AsString ;
    ExelTab.Cells.Item[i, 3] := fDM.QStacRep.FieldByName('data_r').AsString;
    ExelTab.Cells.Item[i, 4] := fDM.QStacRep.FieldByName('polis').AsString;
    ExelTab.Cells.Item[i, 13] := fDM.QStacRep.FieldByName('data_proh').AsString;
    ExelTab.Range[ExelTab.Cells[i, 1], ExelTab.Cells[i, 12]].Select;
    ExelTab.Selection.Borders.LineStyle := 1;

    fDM.Query.SQL.Clear;
    q := 'SELECT rabotnik.data_proh, usl.data, usl.kod_usl, usl.naimusl,usl.kol,usl.cena,usl.data,usl.kod_vr,usl.kod_ms,usl.mkb'
      + ' FROM usl INNER JOIN rabotnik ON rabotnik.id=usl.idrab WHERE (usl.idrab = '
      + fDM.QStacRep.FieldByName('id').AsString +
      ') AND(rabotnik.date_of BETWEEN :date AND :date2)';
    fDM.Query.SQL.Add(q);
    fDM.Query.Parameters.ParamByName('date').Value :=
      FormatDateTime('dd.mm.yyyy', DateTimePicker1.Date);
    fDM.Query.Parameters.ParamByName('date2').Value :=
      FormatDateTime('dd.mm.yyyy', DateTimePicker2.Date);
    fDM.Query.Open;
    c := i;
    fDM.Query.First;
    while not(fDM.Query.Eof) do
    begin
      ExelTab.Cells.Item[i, 5] := fDM.Query.FieldByName('kod_usl').AsString;
      ExelTab.Cells.Item[i, 6] := fDM.Query.FieldByName('naimusl').AsString;
      ExelTab.Cells.Item[i, 7] := fDM.Query.FieldByName('kol').AsString;
      ExelTab.Cells.Item[i, 8] := fDM.Query.FieldByName('cena').AsString;
      ExelTab.Cells.Item[i, 9] := fDM.Query.FieldByName('data').AsString;
      if fDM.Query.FieldByName('kod_vr').AsString = '' then
        ExelTab.Cells.Item[i, 10] := fDM.Query.FieldByName('kod_vr').AsString
      else
        ExelTab.Cells.Item[i, 10] := fDM.Query.FieldByName('kod_ms').AsString;
      ExelTab.Cells.Item[i, 11] := fDM.Query.FieldByName('mkb').AsString;
      ExelTab.Cells.Item[i, 12] := fDM.Query.FieldByName('cena').AsString;
      sum := sum + fDM.Query.FieldByName('cena').AsInteger;
      ExelTab.Range[ExelTab.Cells[i, 1], ExelTab.Cells[i, 12]].Select;
      ExelTab.Selection.Borders.LineStyle := 1;
      fDM.Query.Next;
      inc(i);
    end;
    if c = i then
      inc(i);
    ExelTab.Range[ExelTab.Cells[i, 1], ExelTab.Cells[i, 12]].Select;
    ExelTab.Selection.Borders.LineStyle := 1;
    ExelTab.Cells.Item[i, 2].Font.Bold := True;
    ExelTab.Cells.Item[i, 2].Value := 'Итого:';
    ExelTab.Cells.Item[i, 12].Font.Bold := True;
    ExelTab.Cells.Item[i, 12] := sum;
    sumall := sum + sumall;
    sum := 0;
    inc(i);
    fDM.QStacRep.Next;
  end;
  inc(i);
  ExelTab.Cells.Item[i, 2].Font.Bold := True;
  ExelTab.Cells.Item[i, 2] := 'Итого:';
  ExelTab.Cells.Item[i, 3].Font.Bold := True;
  ExelTab.Cells.Item[i, 3] := sumall;

  ExelTab.Cells.Item[i + 2, 2] := 'Главный врач';
  ExelTab.Cells.Item[i + 2, 5] := ' В. Н. Костина';

  ExelTab.Cells.Item[i + 4, 2] := 'Главный бухгалтер';
  ExelTab.Cells.Item[i + 4, 5] := ' И. Н. Пиягина';

  ExelTab.ActiveWorkbook.SaveAs(st_in + 'Журнал\экономист\ЖАСО\' +
    fDM.TOrg.FieldByName('imya_org').AsString + FormatDateTime('_yyyy_mm_dd',
    Date) + FormatDateTime('_hh_mm', time) + '.xls', EmptyParam, EmptyParam,
    EmptyParam, EmptyParam, EmptyParam, EmptyParam, false, EmptyParam,
    EmptyParam, EmptyParam);
  ExelTab.Quit;
  ShellExecute(Self.Handle, 'open', Pchar(st_in + '\Журнал\экономист\ЖАСО\'),
    nil, nil, SW_SHOWNORMAL);
end;

procedure TfRepJASOPat.CheckBox1Click(Sender: TObject);
var
  i: integer;
  q: String;
begin
  if CheckBox1.Checked then
  begin
    i := 0;
    CheckListBox1.Items.Clear;
    fDM.Query.SQL.Clear;
    fDM.Query.Close;
    q := 'SELECT rabotnik.id, rabotnik.fam, rabotnik.imya, rabotnik.fath, rabotnik.data_r, rabotnik.date_of '
      + 'FROM rabotnik WHERE rabotnik.id_org = ' + fDM.TOrg.FieldByName('id_org').AsString
       + ' AND (date_of BETWEEN :date AND :date2)';
    fDM.Query.SQL.Add(q);
    fDM.Query.Parameters.ParamByName('date').Value :=
      FormatDateTime('dd.mm.yyyy', DateTimePicker1.Date);
    fDM.Query.Parameters.ParamByName('date2').Value :=
      FormatDateTime('dd.mm.yyyy', DateTimePicker2.Date);
    fDM.Query.Open;

    fDM.Query.First;
    while not(fDM.Query.Eof) do
    begin
      CheckListBox1.Items.Add(fDM.Query.FieldByName('id').AsString + ' ' +
        fDM.Query.FieldByName('fam').AsString + ' ' + fDM.Query.FieldByName
        ('imya').AsString + ' ' + fDM.Query.FieldByName('fath').AsString + ' ' +
        fDM.Query.FieldByName('data_r').AsString+ ' Дата оформления: '+fDM.Query.FieldByName('date_of').AsString);
      fDM.Query.Next;
      inc(i);
    end;
  end
  else
    CheckListBox1.Items.Clear;

end;

procedure TfRepJASOPat.FormCreate(Sender: TObject);
begin
 DateTimePicker1.Date := StartOfTheMonth(Now);
  DateTimePicker2.Date := EndOfTheMonth(Now);
end;

end.
