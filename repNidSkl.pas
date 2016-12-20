unit repNidSkl;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ComCtrls, DBCtrls, OleServer, ComObj;

type
  TfRepNidSKL = class(TForm)
    DateTimePicker1: TDateTimePicker;
    DateTimePicker2: TDateTimePicker;
    Label1: TLabel;
    Label2: TLabel;
    DBLookupComboBox1: TDBLookupComboBox;
    Button1: TButton;
    procedure Button1Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  fRepNidSKL: TfRepNidSKL;

implementation

uses DM;

{$R *.dfm}

procedure TfRepNidSKL.Button1Click(Sender: TObject);
var
  ExelTab: Variant;
  st_in: string;
  i,k: integer;
begin
  st_in := ExtractFilePath(Application.ExeName);
  ExelTab := CreateOleObject('Excel.Application');
  ExelTab.Workbooks.Open[st_in + '\�������\listskl.xls'];

  i:=6;

  fDM.QNeedSKL.SQL.Clear;
  fDM.QNeedSKL.SQL.Add
    ('SELECT fam, imya, fath, data_r, profes FROM rabotnik WHERE (data_proh BETWEEN :date AND :date2) AND (id_org=:org) AND (skl='+quotedstr('��')+')  AND (rabotnik.cat_osv not Like "�%") ORDER BY fam');
  fDM.QNeedSKL.Parameters.ParamByName('date').Value :=
    FormatDateTime('dd.mm.yyyy', DateTimePicker1.Date);
  fDM.QNeedSKL.Parameters.ParamByName('date2').Value :=
    FormatDateTime('dd.mm.yyyy', DateTimePicker2.Date);
  fDM.QNeedSKL.Parameters.ParamByName('org').Value :=
    fDM.TOrg.FieldByName('id_org').AsString;
  fDM.QNeedSKL.Active := true;

  fDM.QNeedSKL.First;
  k:=0;
  ExelTab.Cells.Item[1, 1] := fDM.TOrg.FieldByName('imya_poln').AsString;
  while not(fDM.QNeedSKL.Eof) do
  begin
    inc(k);
    ExelTab.Cells.Item[i, 1] := k;
    ExelTab.Cells.Item[i, 2] := fDM.QNeedSKL.FieldByName('fam').AsString;
    ExelTab.Cells.Item[i, 3] := fDM.QNeedSKL.FieldByName('imya').AsString;
    ExelTab.Cells.Item[i, 4] := fDM.QNeedSKL.FieldByName('fath').AsString;
    ExelTab.Cells.Item[i, 5] := fDM.QNeedSKL.FieldByName('data_r').AsString;
    ExelTab.Cells.Item[i, 6] := fDM.QNeedSKL.FieldByName('profes').AsString;
    inc(i);
    fDM.QNeedSKL.Next;
  end;
  ExelTab.ActiveWorkbook.SaveAs(st_in + '������\������\���\' +
    fDM.TOrg.FieldByName('imya_org').AsString + FormatDateTime('_yyyy_mm_dd',
    Date) + FormatDateTime('_hh_mm', time) + '.xls', EmptyParam, EmptyParam,
    EmptyParam, EmptyParam, EmptyParam, EmptyParam, false, EmptyParam,
    EmptyParam, EmptyParam);
  ExelTab.Quit;
end;

procedure TfRepNidSKL.FormCreate(Sender: TObject);
begin
  DateTimePicker1.Date := Date;
  DateTimePicker2.Date := Date;
end;

end.
