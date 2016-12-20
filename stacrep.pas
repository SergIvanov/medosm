unit stacrep;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, DBGridEhGrouping, GridsEh, DBGridEh, DBCtrls, StdCtrls, ComCtrls,
  ExtCtrls, DBGridEhImpExp, OleServer, ComObj, Mask, DBCtrlsEh, DBLookupEh,
  ShellApi;

type
  TfStacRep = class(TForm)
    Panel1: TPanel;
    DBGridEh1: TDBGridEh;
    DateTimePicker1: TDateTimePicker;
    DateTimePicker2: TDateTimePicker;
    Label1: TLabel;
    Label2: TLabel;
    Button1: TButton;
    Button2: TButton;
    DBLookupComboboxEh1: TDBLookupComboboxEh;
    Button3: TButton;
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure Button3Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  fStacRep: TfStacRep;

implementation

uses DM;

{$R *.dfm}

procedure TfStacRep.Button1Click(Sender: TObject);
var
  q: string;
begin
  fDM.QStacRep.SQL.Clear;
  // fDM.QStacRep.SQL.Add('SELECT rabotnik.fam,rabotnik.imya,rabotnik.fath,rabotnik.data_r,rabotnik.profes,rabotnik.vredn_fact,rabotnik.staz_rab_fact,rabotnik.data_proh, organiz.imya FROM rabotnik LEFT JOIN organiz ON organiz.id_org=rabotnik.id_org organiz WHERE (data_proh BETWEEN :date AND :date2)');
  // SELECT rabotnik.fam, rabotnik.imya, rabotnik.fath, rabotnik.data_r, organiz.imya_org, rabotnik.profes, rabotnik.vredn_fact, rabotnik.staz_rab_fact, rabotnik.data_proh
  // FROM rabotnik INNER JOIN organiz ON rabotnik.id_org = organiz.id_org WHERE (data_proh BETWEEN :date AND :date2)
  q := 'SELECT rabotnik.fam, rabotnik.imya, rabotnik.fath, rabotnik.data_r, organiz.imya_org, rabotnik.profes, rabotnik.idusl, rabotnik.mkb_code, rabotnik.vredn_fact, rabotnik.staz_rab_fact, rabotnik.data_proh '
    + 'FROM rabotnik INNER JOIN organiz ON rabotnik.id_org = organiz.id_org WHERE (data_proh BETWEEN :date AND :date2) AND (idusl='
    + fDM.TSUsl.FieldByName('id').AsString + ')';
  fDM.QStacRep.SQL.Add(q);
  fDM.QStacRep.Parameters.ParamByName('date').Value :=
    FormatDateTime('dd.mm.yyyy', DateTimePicker1.Date);
  fDM.QStacRep.Parameters.ParamByName('date2').Value :=
    FormatDateTime('dd.mm.yyyy', DateTimePicker2.Date);
  fDM.QStacRep.Open;
end;

procedure TfStacRep.Button2Click(Sender: TObject);
var
  ExelTab: Variant;
  st_in, st: string;
  i, k, staj, stajd: integer;
begin
  st_in := ExtractFilePath(Application.ExeName);
  ExelTab := CreateOleObject('Excel.Application');
  ExelTab.Workbooks.Open[st_in + '\�������\profosm.xls'];
  i := 5;
  k := 0;
  fDM.QStacRep.First;
  ExelTab.Cells.Item[2, 6] := '� ' + DateToStr(DateTimePicker1.Date) +
    '   ��   ' + DateToStr(DateTimePicker2.Date);
  while not(fDM.QStacRep.Eof) do
  begin
    inc(k);
    ExelTab.Cells.Item[i, 1] := k;
    ExelTab.Cells.Item[i, 3] := fDM.QStacRep.FieldByName('data_proh').AsString;
    ExelTab.Cells.Item[i, 4] := fDM.QStacRep.FieldByName('fam').AsString + ' ' +
      fDM.QStacRep.FieldByName('imya').AsString + ' ' + fDM.QStacRep.FieldByName
      ('fath').AsString;
    ExelTab.Cells.Item[i, 5] := fDM.QStacRep.FieldByName('data_r').AsString;
    ExelTab.Cells.Item[i, 6] := fDM.QStacRep.FieldByName('imya_org').AsString;
    ExelTab.Cells.Item[i, 7] := fDM.QStacRep.FieldByName('profes').AsString;
    ExelTab.Cells.Item[i, 8] := fDM.QStacRep.FieldByName('vredn_fact').AsString;

    st := fDM.QStacRep.FieldByName('staz_rab_fact').AsString;

    try
      staj := strtoint(st);
      stajd := currentyear;
      staj := stajd - staj;
      st := inttostr(staj);
      ExelTab.Cells.Item[i, 9] := st;
    except
      ExelTab.Cells.Item[i, 9] := st;
      // showmessage('�������� ������ ����� ���������!');
    end;

    inc(i);
    fDM.QStacRep.Next;
  end;

  ExelTab.ActiveWorkbook.SaveAs(st_in + '������\���������_' +
    FormatDateTime('_yyyy_mm_dd', Date) + FormatDateTime('_hh_mm', time) +
    '.xls', EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam,
    EmptyParam, false, EmptyParam, EmptyParam, EmptyParam);
  ExelTab.Quit;
end;

procedure TfStacRep.Button3Click(Sender: TObject);
var
  ExelTab: Variant;
  st_in: string;
  i, k: integer;
begin
  st_in := ExtractFilePath(Application.ExeName);
  ExelTab := CreateOleObject('Excel.Application');
  ExelTab.Workbooks.Open[st_in + '\�������\profcentr.xls'];
  i := 12;
  k := 0;
  fDM.QStacRep.First;
  ExelTab.Cells.Item[8, 2] := '� ' + DateToStr(DateTimePicker1.Date) +
    '   ��   ' + DateToStr(DateTimePicker2.Date);
  while not(fDM.QStacRep.Eof) do
  begin
    inc(k);
    ExelTab.Cells.Item[i, 1] := k;
    ExelTab.Cells.Item[i, 2] := fDM.QStacRep.FieldByName('fam').AsString + ' ' +
      fDM.QStacRep.FieldByName('imya').AsString + ' ' + fDM.QStacRep.FieldByName
      ('fath').AsString;
    ExelTab.Cells.Item[i, 3] := fDM.QStacRep.FieldByName('data_r').AsString;
    ExelTab.Cells.Item[i, 4] := fDM.QStacRep.FieldByName('imya_org').AsString;
    ExelTab.Cells.Item[i, 12] := fDM.QStacRep.FieldByName('mkb_code').AsString;
    inc(i);
    fDM.QStacRep.Next;
  end;

  ExelTab.ActiveWorkbook.SaveAs(st_in + '������\���������_' +
    FormatDateTime('_yyyy_mm_dd', Date) + FormatDateTime('_hh_mm', time) +
    '.xls', EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam,
    EmptyParam, false, EmptyParam, EmptyParam, EmptyParam);
  ExelTab.Quit;
  ShellExecute(Self.Handle, 'open', Pchar(st_in + '\������'), nil, nil,
    SW_SHOWNORMAL);
end;

procedure TfStacRep.FormCreate(Sender: TObject);
begin
  DateTimePicker1.Date := Date;
  DateTimePicker2.Date := Date;
end;

procedure TfStacRep.FormShow(Sender: TObject);
begin
  DateTimePicker1.Date := Date;
  DateTimePicker2.Date := Date;
end;

end.