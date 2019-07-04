unit sumuslrep;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, DBGridEhGrouping, GridsEh, DBGridEh, DBCtrls, StdCtrls, ComCtrls,
  ExtCtrls, DBGridEhImpExp, OleServer, ComObj, ToolCtrlsEh, DBGridEhToolCtrls,
  DynVarsEh, EhLibVCL, DBAxisGridsEh;

type
  TfUsl = class(TForm)
    DateTimePicker1: TDateTimePicker;
    DateTimePicker2: TDateTimePicker;
    Button1: TButton;
    DBLookupComboBox1: TDBLookupComboBox;
    DBGridEh1: TDBGridEh;
    Panel1: TPanel;
    Button3: TButton;
    Button4: TButton;
    Button2: TButton;
    Button5: TButton;
    ProgressBar1: TProgressBar;
    procedure Button1Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure Button4Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure Button5Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  fUsl: TfUsl;

implementation

uses DM;

{$R *.dfm}

procedure TfUsl.Button1Click(Sender: TObject);
begin
  fDM.QSumUsl.SQL.Clear;
  fDM.QSumUsl.SQL.Add
    ('SELECT fam, imya, fath, sex, profes, vredn_fact,cat_osv, idusl, iduslold, extest,extestold FROM rabotnik WHERE (data_proh BETWEEN :date AND :date2) AND  ((UCase(cat_osv) Not Like "П%") or (IsNull(cat_osv))) AND (id_org=:org) ORDER BY fam');
  fDM.QSumUsl.Parameters.ParamByName('date').Value :=
    FormatDateTime('dd.mm.yyyy', DateTimePicker1.Date);
  fDM.QSumUsl.Parameters.ParamByName('date2').Value :=
    FormatDateTime('dd.mm.yyyy', DateTimePicker2.Date);
  fDM.QSumUsl.Parameters.ParamByName('org').Value :=
    fDM.TOrg.FieldByName('id_org').AsString;
  fDM.QSumUsl.Active := true;
  fDM.QSumUsl.First;
end;

procedure TfUsl.Button2Click(Sender: TObject); // Акт по услугам 6Ц
var
  ExelTab: Variant;
  st_in,s: string;
  i: integer;
begin
  st_in := ExtractFilePath(Application.ExeName);

  ExelTab := CreateOleObject('Excel.Application');
  if fDM.TOrg.FieldByName('dif_price').AsInteger <> 1 then
    ExelTab.Workbooks.Open[st_in + '\Шаблоны\akt6c.xls']
  else
    ExelTab.Workbooks.Open[st_in + '\Шаблоны\akt6cdif.xls'];
  i := 12;
  fDM.QSumUsl.First;
  while not(fDM.QSumUsl.Eof) do
  begin
    if (fDM.QSumUsl.FieldByName('idusl').AsInteger = 1000) or
      (fDM.QSumUsl.FieldByName('idusl').AsInteger = 1002) then
    begin
      ExelTab.Cells.Item[i, 2] := fDM.QSumUsl.FieldByName('fam').AsString;
      s := copy(fDM.QSumUsl.FieldByName('imya').AsString, 1, 1);
      ExelTab.Cells.Item[i, 3] := s;
      s := copy(fDM.QSumUsl.FieldByName('fath').AsString, 1, 1);
      ExelTab.Cells.Item[i, 4] := s;
      ExelTab.Cells.Item[i, 5] := fDM.QSumUsl.FieldByName('profes').AsString;
      ExelTab.Cells.Item[7, 5] := fDM.TOrg.FieldByName('imya_poln').AsString;
      ExelTab.Cells.Item[5, 4] := 'с ' + DateToStr(DateTimePicker1.Date) +
        '   по   ' + DateToStr(DateTimePicker2.Date);

      if (fDM.QSumUsl.FieldByName('extest').AsString = '1') then
        // Экспресс тест при вэк. Нет истории изменения!
        ExelTab.Cells.Item[i, 11] := 1;

      // Распределение по группам
      if (fDM.QSumUsl.FieldByName('cat_osv').AsString = 'Р1.3гр1') or
        (fDM.QSumUsl.FieldByName('cat_osv').AsString = 'Р1.3гр2') or
        (fDM.QSumUsl.FieldByName('cat_osv').AsString = 'Р1.3гр3') or
        (fDM.QSumUsl.FieldByName('cat_osv').AsString = 'Р1.3гр5') or
        (fDM.QSumUsl.FieldByName('cat_osv').AsString = 'Р1.3гр6') then
      begin
        if (fDM.QSumUsl.FieldByName('sex').AsString = 'муж') then
          ExelTab.Cells.Item[i, 6].Value := 1;
        if (fDM.QSumUsl.FieldByName('sex').AsString = 'жен') then
          ExelTab.Cells.Item[i, 7].Value := 1;
      end;

      if (fDM.QSumUsl.FieldByName('cat_osv').AsString = 'Р1гр1') or
        (fDM.QSumUsl.FieldByName('cat_osv').AsString = 'Р1гр2') or
        (fDM.QSumUsl.FieldByName('cat_osv').AsString = 'Р1гр3') or
        (fDM.QSumUsl.FieldByName('cat_osv').AsString = 'Р1гр5') or
        (fDM.QSumUsl.FieldByName('cat_osv').AsString = 'Р1гр6') then
      begin
        if (fDM.QSumUsl.FieldByName('sex').AsString = 'муж') then
          ExelTab.Cells.Item[i, 6].Value := 1;
        if (fDM.QSumUsl.FieldByName('sex').AsString = 'жен') then
          ExelTab.Cells.Item[i, 7].Value := 1;
      end;

      if (fDM.QSumUsl.FieldByName('cat_osv').AsString = 'Р1.3гр4') or
        (fDM.QSumUsl.FieldByName('cat_osv').AsString = 'Р1гр4') then
      begin
        if (fDM.QSumUsl.FieldByName('sex').AsString = 'муж') then
          ExelTab.Cells.Item[i, 8].Value := 1;
        if (fDM.QSumUsl.FieldByName('sex').AsString = 'жен') then
          ExelTab.Cells.Item[i, 9].Value := 1;
      end;
    end;
    // --------------------------
    // ------------2
    if (fDM.QSumUsl.FieldByName('iduslold').AsInteger = 1000) or
      (fDM.QSumUsl.FieldByName('iduslold').AsInteger = 1002) then
    begin
      ExelTab.Cells.Item[i, 2] := fDM.QSumUsl.FieldByName('fam').AsString;
      s := copy(fDM.QSumUsl.FieldByName('imya').AsString, 1, 1);
      ExelTab.Cells.Item[i, 3] := s;
      s := copy(fDM.QSumUsl.FieldByName('fath').AsString, 1, 1);
      ExelTab.Cells.Item[i, 4] := s;
      ExelTab.Cells.Item[i, 5] := fDM.QSumUsl.FieldByName('profes').AsString;
      ExelTab.Cells.Item[7, 5] := fDM.TOrg.FieldByName('imya_poln').AsString;
      ExelTab.Cells.Item[5, 4] := 'с ' + DateToStr(DateTimePicker1.Date) +
        '   по   ' + DateToStr(DateTimePicker2.Date);
      if (fDM.QSumUsl.FieldByName('extest').AsString = '1') then
        // Экспресс тест при вэк. Нет истории изменения!
        ExelTab.Cells.Item[i, 11] := 1;

      // Распределение по группам
      if (fDM.QSumUsl.FieldByName('cat_osv').AsString = 'Р1.3гр1') or
        (fDM.QSumUsl.FieldByName('cat_osv').AsString = 'Р1.3гр2') or
        (fDM.QSumUsl.FieldByName('cat_osv').AsString = 'Р1.3гр3') or
        (fDM.QSumUsl.FieldByName('cat_osv').AsString = 'Р1.3гр5') or
        (fDM.QSumUsl.FieldByName('cat_osv').AsString = 'Р1.3гр6') then
      begin
        if (fDM.QSumUsl.FieldByName('sex').AsString = 'муж') then
          ExelTab.Cells.Item[i, 6].Value := -1;
        if (fDM.QSumUsl.FieldByName('sex').AsString = 'жен') then
          ExelTab.Cells.Item[i, 7].Value := -1;
      end;

      if (fDM.QSumUsl.FieldByName('cat_osv').AsString = 'Р1гр1') or
        (fDM.QSumUsl.FieldByName('cat_osv').AsString = 'Р1гр2') or
        (fDM.QSumUsl.FieldByName('cat_osv').AsString = 'Р1гр3') or
        (fDM.QSumUsl.FieldByName('cat_osv').AsString = 'Р1гр5') or
        (fDM.QSumUsl.FieldByName('cat_osv').AsString = 'Р1гр6') then
      begin
        if (fDM.QSumUsl.FieldByName('sex').AsString = 'муж') then
          ExelTab.Cells.Item[i, 6].Value := -1;
        if (fDM.QSumUsl.FieldByName('sex').AsString = 'жен') then
          ExelTab.Cells.Item[i, 7].Value := -1;
      end;

      if (fDM.QSumUsl.FieldByName('cat_osv').AsString = 'Р1.3гр4') or
        (fDM.QSumUsl.FieldByName('cat_osv').AsString = 'Р1гр4') then
      begin
        if (fDM.QSumUsl.FieldByName('sex').AsString = 'муж') then
          ExelTab.Cells.Item[i, 8].Value := -1;
        if (fDM.QSumUsl.FieldByName('sex').AsString = 'жен') then
          ExelTab.Cells.Item[i, 9].Value := -1;
      end;
      // -------------
    end
    else
    begin
      if (fDM.QSumUsl.FieldByName('idusl').AsInteger = 1001) then
      begin
        ExelTab.Cells.Item[i, 2] := fDM.QSumUsl.FieldByName('fam').AsString;
        s := copy(fDM.QSumUsl.FieldByName('imya').AsString, 1, 1);
        ExelTab.Cells.Item[i, 3] := s;
        s := copy(fDM.QSumUsl.FieldByName('fath').AsString, 1, 1);
        ExelTab.Cells.Item[i, 4] := s;
        ExelTab.Cells.Item[i, 5] := fDM.QSumUsl.FieldByName('profes').AsString;
        ExelTab.Cells.Item[7, 5] := fDM.TOrg.FieldByName('imya_poln').AsString;
        ExelTab.Cells.Item[5, 4] := 'с ' + DateToStr(DateTimePicker1.Date) +
          '   по   ' + DateToStr(DateTimePicker2.Date);
        ExelTab.Cells.Item[i, 10].Value := 1;
        if (fDM.QSumUsl.FieldByName('extest').AsString = '1') then
          // Экспресс тест при вэк. Нет истории изменения!
          ExelTab.Cells.Item[i, 11] := 1;

        // ------
      end;
      // ----2
      if (fDM.QSumUsl.FieldByName('iduslold').AsInteger = 1001) then
      begin
       ExelTab.Cells.Item[i, 2] := fDM.QSumUsl.FieldByName('fam').AsString;
        s := copy(fDM.QSumUsl.FieldByName('imya').AsString, 1, 1);
        ExelTab.Cells.Item[i, 3] := s;
        s := copy(fDM.QSumUsl.FieldByName('fath').AsString, 1, 1);
        ExelTab.Cells.Item[i, 4] := s;
        ExelTab.Cells.Item[i, 5] := fDM.QSumUsl.FieldByName('profes').AsString;
        ExelTab.Cells.Item[7, 5] := fDM.TOrg.FieldByName('imya_poln').AsString;
        ExelTab.Cells.Item[5, 4] := 'с ' + DateToStr(DateTimePicker1.Date) +
          '   по   ' + DateToStr(DateTimePicker2.Date);
        ExelTab.Cells.Item[i, 10].Value := -1;
        if (fDM.QSumUsl.FieldByName('extest').AsString = '1') then
          // Экспресс тест при вэк. Нет истории изменения!
          ExelTab.Cells.Item[i, 11] := 1;
        // ------
      end;
      // --------
      // else
      // begin
      // fDM.QSumUsl.Next;
      // continue;
      // end;
    end;

    if (fDM.QSumUsl.FieldByName('idusl').AsInteger >= 1000) OR
      (fDM.QSumUsl.FieldByName('iduslold').AsInteger >= 1000) then
      inc(i);
    fDM.QSumUsl.Next;
  end;
  // ExelTab.Visible := true;
  ExelTab.ActiveWorkbook.SaveAs(st_in + 'Журнал\экономист\6C\' +
    fDM.TOrg.FieldByName('imya_org').AsString + FormatDateTime('_yyyy_mm_dd',
    Date) + FormatDateTime('_hh_mm', time) + '.xls', EmptyParam, EmptyParam,
    EmptyParam, EmptyParam, EmptyParam, EmptyParam, false, EmptyParam,
    EmptyParam, EmptyParam);
  ExelTab.Quit;
end;

procedure TfUsl.Button3Click(Sender: TObject);
var
  ExelTab: Variant;
  st_in: string;
  i: integer;
begin
  st_in := ExtractFilePath(Application.ExeName);
  ExelTab := CreateOleObject('Excel.Application');
  if fDM.TOrg.FieldByName('dif_price').AsInteger <> 1 then
    // Проверять в какой шаблон записывать данные
    ExelTab.Workbooks.Open[st_in + '\Шаблоны\akt.xls']
  else
    ExelTab.Workbooks.Open[st_in + '\Шаблоны\aktdif.xls'];
  i := 10;
  fDM.QSumUsl.First;
  while not(fDM.QSumUsl.Eof) do
  begin
    if (fDM.QSumUsl.FieldByName('idusl').AsInteger <> 1000) and
      (fDM.QSumUsl.FieldByName('idusl').AsInteger <> 1001) and
      (fDM.QSumUsl.FieldByName('idusl').AsInteger <> 1002) then
    begin
      ExelTab.Cells.Item[i, 2] := fDM.QSumUsl.FieldByName('fam').AsString + ' '
        + fDM.QSumUsl.FieldByName('imya').AsString + ' ' +
        fDM.QSumUsl.FieldByName('fath').AsString;
      ExelTab.Cells.Item[i, 3] := fDM.QSumUsl.FieldByName('profes').AsString;
      ExelTab.Cells.Item[5, 3] := fDM.TOrg.FieldByName('imya_poln').AsString;
      ExelTab.Cells.Item[4, 6] := 'с ' + DateToStr(DateTimePicker1.Date) +
        '   по   ' + DateToStr(DateTimePicker2.Date);
      // ExelTab.Cells.Item[4,7]:=DateToStr(DateTimePicker2.Date);

      // 1
      if (fDM.QSumUsl.FieldByName('idusl').AsInteger = 1) and
        (fDM.QSumUsl.FieldByName('sex').AsString = 'муж') then
        ExelTab.Cells.Item[i, 4] := 1;
      if (fDM.QSumUsl.FieldByName('idusl').AsInteger = 1) and
        (fDM.QSumUsl.FieldByName('sex').AsString = 'жен') then
        ExelTab.Cells.Item[i, 5] := 1;

      if (fDM.QSumUsl.FieldByName('iduslold').AsInteger = 1) and
        (fDM.QSumUsl.FieldByName('sex').AsString = 'муж') then
        ExelTab.Cells.Item[i, 4] := -1;
      if (fDM.QSumUsl.FieldByName('iduslold').AsInteger = 1) and
        (fDM.QSumUsl.FieldByName('sex').AsString = 'жен') then
        ExelTab.Cells.Item[i, 5] := -1;
      // ---------------------------------------------------------
      // 2
      if (fDM.QSumUsl.FieldByName('idusl').AsInteger = 2) and
        (fDM.QSumUsl.FieldByName('sex').AsString = 'муж') then
        ExelTab.Cells.Item[i, 6] := 1;
      if (fDM.QSumUsl.FieldByName('idusl').AsInteger = 2) and
        (fDM.QSumUsl.FieldByName('sex').AsString = 'жен') then
        ExelTab.Cells.Item[i, 7] := 1;

      if (fDM.QSumUsl.FieldByName('iduslold').AsInteger = 2) and
        (fDM.QSumUsl.FieldByName('sex').AsString = 'муж') then
        ExelTab.Cells.Item[i, 6] := -1;
      if (fDM.QSumUsl.FieldByName('iduslold').AsInteger = 2) and
        (fDM.QSumUsl.FieldByName('sex').AsString = 'жен') then
        ExelTab.Cells.Item[i, 7] := -1;
      // ----------------------------------------------------------
      // 3
      if (fDM.QSumUsl.FieldByName('idusl').AsInteger = 3) and
        (fDM.QSumUsl.FieldByName('sex').AsString = 'муж') then
        ExelTab.Cells.Item[i, 8] := 1;
      if (fDM.QSumUsl.FieldByName('idusl').AsInteger = 3) and
        (fDM.QSumUsl.FieldByName('sex').AsString = 'жен') then
        ExelTab.Cells.Item[i, 9] := 1;

      if (fDM.QSumUsl.FieldByName('iduslold').AsInteger = 3) and
        (fDM.QSumUsl.FieldByName('sex').AsString = 'муж') then
        ExelTab.Cells.Item[i, 8] := -1;
      if (fDM.QSumUsl.FieldByName('iduslold').AsInteger = 3) and
        (fDM.QSumUsl.FieldByName('sex').AsString = 'жен') then
        ExelTab.Cells.Item[i, 9] := -1;
      // ---------------------------------------------------------
      // 4
      if (fDM.QSumUsl.FieldByName('idusl').AsInteger = 4) and
        (fDM.QSumUsl.FieldByName('sex').AsString = 'муж') then
        ExelTab.Cells.Item[i, 10] := 1;
      if (fDM.QSumUsl.FieldByName('idusl').AsInteger = 4) and
        (fDM.QSumUsl.FieldByName('sex').AsString = 'жен') then
        ExelTab.Cells.Item[i, 11] := 1;

      if (fDM.QSumUsl.FieldByName('iduslold').AsInteger = 4) and
        (fDM.QSumUsl.FieldByName('sex').AsString = 'муж') then
        ExelTab.Cells.Item[i, 10] := -1;
      if (fDM.QSumUsl.FieldByName('iduslold').AsInteger = 4) and
        (fDM.QSumUsl.FieldByName('sex').AsString = 'жен') then
        ExelTab.Cells.Item[i, 11] := -1;
      // --------------------------------------------------------
      // 5
      if (fDM.QSumUsl.FieldByName('idusl').AsInteger = 5) and
        (fDM.QSumUsl.FieldByName('sex').AsString = 'муж') then
        ExelTab.Cells.Item[i, 12] := 1;
      if (fDM.QSumUsl.FieldByName('idusl').AsInteger = 5) and
        (fDM.QSumUsl.FieldByName('sex').AsString = 'жен') then
        ExelTab.Cells.Item[i, 13] := 1;

      if (fDM.QSumUsl.FieldByName('iduslold').AsInteger = 5) and
        (fDM.QSumUsl.FieldByName('sex').AsString = 'муж') then
        ExelTab.Cells.Item[i, 12] := -1;
      if (fDM.QSumUsl.FieldByName('iduslold').AsInteger = 5) and
        (fDM.QSumUsl.FieldByName('sex').AsString = 'жен') then
        ExelTab.Cells.Item[i, 13] := -1;
      // --------------------------------------------------------
      // 6
      if (fDM.QSumUsl.FieldByName('idusl').AsInteger = 6) and
        (fDM.QSumUsl.FieldByName('sex').AsString = 'муж') then
        ExelTab.Cells.Item[i, 14] := 1;
      if (fDM.QSumUsl.FieldByName('idusl').AsInteger = 6) and
        (fDM.QSumUsl.FieldByName('sex').AsString = 'жен') then
        ExelTab.Cells.Item[i, 15] := 1;

      if (fDM.QSumUsl.FieldByName('iduslold').AsInteger = 6) and
        (fDM.QSumUsl.FieldByName('sex').AsString = 'муж') then
        ExelTab.Cells.Item[i, 14] := -1;
      if (fDM.QSumUsl.FieldByName('iduslold').AsInteger = 6) and
        (fDM.QSumUsl.FieldByName('sex').AsString = 'жен') then
        ExelTab.Cells.Item[i, 15] := -1;
      // -----------------------------------------------------

      // 7
      if (fDM.QSumUsl.FieldByName('idusl').AsInteger = 7) and
        (fDM.QSumUsl.FieldByName('sex').AsString = 'муж') then
        ExelTab.Cells.Item[i, 16] := 1;
      if (fDM.QSumUsl.FieldByName('idusl').AsInteger = 7) and
        (fDM.QSumUsl.FieldByName('sex').AsString = 'жен') then
        ExelTab.Cells.Item[i, 17] := 1;

      if (fDM.QSumUsl.FieldByName('iduslold').AsInteger = 7) and
        (fDM.QSumUsl.FieldByName('sex').AsString = 'муж') then
        ExelTab.Cells.Item[i, 16] := -1;
      if (fDM.QSumUsl.FieldByName('iduslold').AsInteger = 7) and
        (fDM.QSumUsl.FieldByName('sex').AsString = 'жен') then
        ExelTab.Cells.Item[i, 17] := -1;
      // -----------------------------------------------------
      // 8
      if (fDM.QSumUsl.FieldByName('idusl').AsInteger = 8) and
        (fDM.QSumUsl.FieldByName('sex').AsString = 'муж') then
        ExelTab.Cells.Item[i, 18] := 1;
      if (fDM.QSumUsl.FieldByName('idusl').AsInteger = 8) and
        (fDM.QSumUsl.FieldByName('sex').AsString = 'жен') then
        ExelTab.Cells.Item[i, 19] := 1;

      if (fDM.QSumUsl.FieldByName('iduslold').AsInteger = 8) and
        (fDM.QSumUsl.FieldByName('sex').AsString = 'муж') then
        ExelTab.Cells.Item[i, 18] := -1;
      if (fDM.QSumUsl.FieldByName('iduslold').AsInteger = 8) and
        (fDM.QSumUsl.FieldByName('sex').AsString = 'жен') then
        ExelTab.Cells.Item[i, 19] := -1;
      // --------------------------------------------------------
      // 9
      if (fDM.QSumUsl.FieldByName('idusl').AsInteger = 9) then
        ExelTab.Cells.Item[i, 20] := 1;
      if (fDM.QSumUsl.FieldByName('idusl').AsInteger = 10) then
        ExelTab.Cells.Item[i, 21] := 1;
      if (fDM.QSumUsl.FieldByName('idusl').AsInteger = 11) then
        ExelTab.Cells.Item[i, 22] := 1;

      // if (fDM.QSumUsl.FieldByName('extest').AsString <> '') then
      // begin
      // if (fDM.QSumUsl.FieldByName('iduslold').AsInteger < 1000) and
      // (fDM.QSumUsl.FieldByName('iduslold').AsInteger > 0) and
      // (fDM.QSumUsl.FieldByName('iddusl').AsInteger < 1000) and
      // (fDM.QSumUsl.FieldByName('idusl').AsInteger > 0) then
      // ExelTab.Cells.Item[i, 23] := 1;
      // end;

      if (fDM.QSumUsl.FieldByName('extest').AsString = '1') then
        ExelTab.Cells.Item[i, 23] := 1;

      if (fDM.QSumUsl.FieldByName('iduslold').AsInteger = 9) then
        ExelTab.Cells.Item[i, 20] := -1;
      if (fDM.QSumUsl.FieldByName('iduslold').AsInteger = 10) then
        ExelTab.Cells.Item[i, 21] := -1;
      if (fDM.QSumUsl.FieldByName('iduslold').AsInteger = 11) then
        ExelTab.Cells.Item[i, 22] := -1;
      // -------------------------------------------------------------

      inc(i);
      fDM.QSumUsl.Next;
    end
    else
      fDM.QSumUsl.Next;
  end;
  // ExelTab.Workbooks.SaveAs (st_in+'\Журнал\akt'+formatdatetime('yyyy_mm_dd',date)+formatdatetime('_hh_mm',time)+'.xls');
  ExelTab.ActiveWorkbook.SaveAs(st_in + 'Журнал\экономист\302\' +
    fDM.TOrg.FieldByName('imya_org').AsString + FormatDateTime('_yyyy_mm_dd',
    Date) + FormatDateTime('_hh_mm', time) + '.xls', EmptyParam, EmptyParam,
    EmptyParam, EmptyParam, EmptyParam, EmptyParam, false, EmptyParam,
    EmptyParam, EmptyParam);
  ExelTab.Quit;
end;

procedure TfUsl.Button4Click(Sender: TObject);
var
  ExelTab: Variant;
  st_in: string;
  i, rcount, c: integer;
  k: boolean;
begin
  c := 0;
  fDM.TOrg.First;
  while not(fDM.TOrg.Eof) do
  begin
    rcount := fDM.TOrg.RecordCount;
    ProgressBar1.Max := rcount;
    inc(c);
    ProgressBar1.Position := round(c);

    fDM.QSumUsl.SQL.Clear;
    fDM.QSumUsl.SQL.Add
      ('SELECT fam, imya, fath, sex, profes, vredn_fact, cat_osv, idusl,iduslold, extest, extestold FROM rabotnik WHERE (data_proh BETWEEN :date AND :date2) AND ((UCase(cat_osv) Not Like "П%") or (IsNull(cat_osv))) AND (id_org=:org) ORDER BY fam');
    fDM.QSumUsl.Parameters.ParamByName('date').Value :=
      FormatDateTime('dd.mm.yyyy', DateTimePicker1.Date);
    fDM.QSumUsl.Parameters.ParamByName('date2').Value :=
      FormatDateTime('dd.mm.yyyy', DateTimePicker2.Date);
    fDM.QSumUsl.Parameters.ParamByName('org').Value :=
      fDM.TOrg.FieldByName('id_org').AsString;
    fDM.QSumUsl.Active := true;
    fDM.QSumUsl.First;
    // ---------
    st_in := ExtractFilePath(Application.ExeName);

    ExelTab := CreateOleObject('Excel.Application');
    if fDM.TOrg.FieldByName('dif_price').AsInteger <> 1 then
      ExelTab.Workbooks.Open[st_in + '\Шаблоны\akt.xls']
    else
      ExelTab.Workbooks.Open[st_in + '\Шаблоны\aktdif.xls'];
    i := 10;
    k := false;
    // fDM.QSumUsl.First;
    while not(fDM.QSumUsl.Eof) do
    begin
      if (fDM.QSumUsl.FieldByName('idusl').AsInteger <> 1000) and
        (fDM.QSumUsl.FieldByName('idusl').AsInteger <> 1001) and
        (fDM.QSumUsl.FieldByName('idusl').AsInteger <> 1002) then
      begin
        ExelTab.Cells.Item[i, 2] := fDM.QSumUsl.FieldByName('fam').AsString +
          ' ' + fDM.QSumUsl.FieldByName('imya').AsString + ' ' +
          fDM.QSumUsl.FieldByName('fath').AsString;
        ExelTab.Cells.Item[i, 3] := fDM.QSumUsl.FieldByName('profes').AsString;
        ExelTab.Cells.Item[5, 3] := fDM.TOrg.FieldByName('imya_poln').AsString;
        ExelTab.Cells.Item[4, 6] := 'с ' + DateToStr(DateTimePicker1.Date) +
          '   по   ' + DateToStr(DateTimePicker2.Date);

        if (fDM.QSumUsl.FieldByName('idusl').AsInteger = 1) and
          (fDM.QSumUsl.FieldByName('sex').AsString = 'муж') then
        begin
          k := true;
          ExelTab.Cells.Item[i, 4] := 1;
        end;
        if (fDM.QSumUsl.FieldByName('idusl').AsInteger = 1) and
          (fDM.QSumUsl.FieldByName('sex').AsString = 'жен') then
        begin
          k := true;
          ExelTab.Cells.Item[i, 5] := 1;
        end;
        if (fDM.QSumUsl.FieldByName('iduslold').AsInteger = 1) and
          (fDM.QSumUsl.FieldByName('sex').AsString = 'муж') then
        begin
          k := true;
          ExelTab.Cells.Item[i, 4] := -1;
        end;
        if (fDM.QSumUsl.FieldByName('iduslold').AsInteger = 1) and
          (fDM.QSumUsl.FieldByName('sex').AsString = 'жен') then
        begin
          k := true;
          ExelTab.Cells.Item[i, 5] := -1;
        end;

        if (fDM.QSumUsl.FieldByName('idusl').AsInteger = 2) and
          (fDM.QSumUsl.FieldByName('sex').AsString = 'муж') then
        begin
          k := true;
          ExelTab.Cells.Item[i, 6] := 1;
        end;
        if (fDM.QSumUsl.FieldByName('idusl').AsInteger = 2) and
          (fDM.QSumUsl.FieldByName('sex').AsString = 'жен') then
        begin
          k := true;
          ExelTab.Cells.Item[i, 7] := 1;
        end;
        if (fDM.QSumUsl.FieldByName('iduslold').AsInteger = 2) and
          (fDM.QSumUsl.FieldByName('sex').AsString = 'муж') then
        begin
          k := true;
          ExelTab.Cells.Item[i, 6] := -1;
        end;
        if (fDM.QSumUsl.FieldByName('iduslold').AsInteger = 2) and
          (fDM.QSumUsl.FieldByName('sex').AsString = 'жен') then
        begin
          k := true;
          ExelTab.Cells.Item[i, 7] := -1;
        end;

        if (fDM.QSumUsl.FieldByName('idusl').AsInteger = 3) and
          (fDM.QSumUsl.FieldByName('sex').AsString = 'муж') then
        begin
          k := true;
          ExelTab.Cells.Item[i, 8] := 1;
        end;
        if (fDM.QSumUsl.FieldByName('idusl').AsInteger = 3) and
          (fDM.QSumUsl.FieldByName('sex').AsString = 'жен') then
        begin
          k := true;
          ExelTab.Cells.Item[i, 9] := 1;
        end;
        if (fDM.QSumUsl.FieldByName('iduslold').AsInteger = 3) and
          (fDM.QSumUsl.FieldByName('sex').AsString = 'муж') then
        begin
          k := true;
          ExelTab.Cells.Item[i, 8] := -1;
        end;
        if (fDM.QSumUsl.FieldByName('iduslold').AsInteger = 3) and
          (fDM.QSumUsl.FieldByName('sex').AsString = 'жен') then
        begin
          k := true;
          ExelTab.Cells.Item[i, 9] := -1;
        end;

        if (fDM.QSumUsl.FieldByName('idusl').AsInteger = 4) and
          (fDM.QSumUsl.FieldByName('sex').AsString = 'муж') then
        begin
          k := true;
          ExelTab.Cells.Item[i, 10] := 1;
        end;
        if (fDM.QSumUsl.FieldByName('idusl').AsInteger = 4) and
          (fDM.QSumUsl.FieldByName('sex').AsString = 'жен') then
        begin
          k := true;
          ExelTab.Cells.Item[i, 11] := 1;
        end;
        if (fDM.QSumUsl.FieldByName('iduslold').AsInteger = 4) and
          (fDM.QSumUsl.FieldByName('sex').AsString = 'муж') then
        begin
          k := true;
          ExelTab.Cells.Item[i, 10] := -1;
        end;
        if (fDM.QSumUsl.FieldByName('iduslold').AsInteger = 4) and
          (fDM.QSumUsl.FieldByName('sex').AsString = 'жен') then
        begin
          k := true;
          ExelTab.Cells.Item[i, 11] := -1;
        end;

        if (fDM.QSumUsl.FieldByName('idusl').AsInteger = 5) and
          (fDM.QSumUsl.FieldByName('sex').AsString = 'муж') then
        begin
          k := true;
          ExelTab.Cells.Item[i, 12] := 1;
        end;
        if (fDM.QSumUsl.FieldByName('idusl').AsInteger = 5) and
          (fDM.QSumUsl.FieldByName('sex').AsString = 'жен') then
        begin
          k := true;
          ExelTab.Cells.Item[i, 13] := 1;
        end;
        if (fDM.QSumUsl.FieldByName('iduslold').AsInteger = 5) and
          (fDM.QSumUsl.FieldByName('sex').AsString = 'муж') then
        begin
          k := true;
          ExelTab.Cells.Item[i, 12] := -1;
        end;
        if (fDM.QSumUsl.FieldByName('iduslold').AsInteger = 5) and
          (fDM.QSumUsl.FieldByName('sex').AsString = 'жен') then
        begin
          k := true;
          ExelTab.Cells.Item[i, 13] := -1;
        end;

        if (fDM.QSumUsl.FieldByName('idusl').AsInteger = 6) and
          (fDM.QSumUsl.FieldByName('sex').AsString = 'муж') then
        begin
          k := true;
          ExelTab.Cells.Item[i, 14] := 1;
        end;
        if (fDM.QSumUsl.FieldByName('idusl').AsInteger = 6) and
          (fDM.QSumUsl.FieldByName('sex').AsString = 'жен') then
        begin
          k := true;
          ExelTab.Cells.Item[i, 15] := 1;
        end;
        if (fDM.QSumUsl.FieldByName('iduslold').AsInteger = 6) and
          (fDM.QSumUsl.FieldByName('sex').AsString = 'муж') then
        begin
          k := true;
          ExelTab.Cells.Item[i, 14] := -1;
        end;
        if (fDM.QSumUsl.FieldByName('iduslold').AsInteger = 6) and
          (fDM.QSumUsl.FieldByName('sex').AsString = 'жен') then
        begin
          k := true;
          ExelTab.Cells.Item[i, 15] := -1;
        end;

        if (fDM.QSumUsl.FieldByName('idusl').AsInteger = 7) and
          (fDM.QSumUsl.FieldByName('sex').AsString = 'муж') then
        begin
          k := true;
          ExelTab.Cells.Item[i, 16] := 1;
        end;
        if (fDM.QSumUsl.FieldByName('idusl').AsInteger = 7) and
          (fDM.QSumUsl.FieldByName('sex').AsString = 'жен') then
        begin
          k := true;
          ExelTab.Cells.Item[i, 17] := 1;
        end;
        if (fDM.QSumUsl.FieldByName('iduslold').AsInteger = 7) and
          (fDM.QSumUsl.FieldByName('sex').AsString = 'муж') then
        begin
          k := true;
          ExelTab.Cells.Item[i, 16] := -1;
        end;
        if (fDM.QSumUsl.FieldByName('iduslold').AsInteger = 7) and
          (fDM.QSumUsl.FieldByName('sex').AsString = 'жен') then
        begin
          k := true;
          ExelTab.Cells.Item[i, 17] := -1;
        end;

        if (fDM.QSumUsl.FieldByName('idusl').AsInteger = 8) and
          (fDM.QSumUsl.FieldByName('sex').AsString = 'муж') then
        begin
          k := true;
          ExelTab.Cells.Item[i, 18] := 1;
        end;

        if (fDM.QSumUsl.FieldByName('idusl').AsInteger = 8) and
          (fDM.QSumUsl.FieldByName('sex').AsString = 'жен') then
        begin
          k := true;
          ExelTab.Cells.Item[i, 19] := 1;
        end;
        if (fDM.QSumUsl.FieldByName('iduslold').AsInteger = 8) and
          (fDM.QSumUsl.FieldByName('sex').AsString = 'муж') then
        begin
          k := true;
          ExelTab.Cells.Item[i, 18] := -1;
        end;
        if (fDM.QSumUsl.FieldByName('iduslold').AsInteger = 8) and
          (fDM.QSumUsl.FieldByName('sex').AsString = 'жен') then
        begin
          k := true;
          ExelTab.Cells.Item[i, 19] := -1;
        end;

        if (fDM.QSumUsl.FieldByName('idusl').AsInteger = 9) then
        begin
          k := true;
          ExelTab.Cells.Item[i, 20] := 1;
        end;
        if (fDM.QSumUsl.FieldByName('idusl').AsInteger = 10) then
        begin
          k := true;
          ExelTab.Cells.Item[i, 21] := 1;
        end;
        if (fDM.QSumUsl.FieldByName('idusl').AsInteger = 11) then
        begin
          k := true;
          ExelTab.Cells.Item[i, 22] := 1;
        end;

        if (fDM.QSumUsl.FieldByName('extest').AsString = '1') then
        // Экспресс тест при вэк. Нет истории изменения!
        begin
          k := true;
          ExelTab.Cells.Item[i, 23] := 1;
        end;

        if (fDM.QSumUsl.FieldByName('iduslold').AsInteger = 9) then
        begin
          k := true;
          ExelTab.Cells.Item[i, 20] := -1;
        end;
        if (fDM.QSumUsl.FieldByName('iduslold').AsInteger = 10) then
        begin
          k := true;
          ExelTab.Cells.Item[i, 21] := -1;
        end;
        if (fDM.QSumUsl.FieldByName('iduslold').AsInteger = 11) then
        begin
          k := true;
          ExelTab.Cells.Item[i, 22] := -1;
        end;

        inc(i);
        fDM.QSumUsl.Next;
      end
      else
        fDM.QSumUsl.Next;
    end;
    if k then
    begin
      ExelTab.ActiveWorkbook.SaveAs(st_in + 'Журнал\экономист\302\' +
        fDM.TOrg.FieldByName('imya_org').AsString +
        FormatDateTime('_yyyy_mm_dd', Date) + FormatDateTime('_hh_mm', time) +
        '.xls', EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam,
        EmptyParam, false, EmptyParam, EmptyParam, EmptyParam);
      ExelTab.Quit;
      fDM.TOrg.Next;
    end
    else
    begin
      fDM.TOrg.Next;
      ExelTab.Quit;
    end;
    // -------
  end;
end;

procedure TfUsl.Button5Click(Sender: TObject);
var
  ExelTab: Variant;
  st_in, s: string;
  i, rcount, c: integer;
  k: boolean;
begin
c:=0;
  fDM.TOrg.First;
  while not(fDM.TOrg.Eof) do
  begin
    rcount := fDM.TOrg.RecordCount;
    ProgressBar1.Max := rcount;
    inc(c);
    ProgressBar1.Position := round(c);

    fDM.QSumUsl.SQL.Clear;
    fDM.QSumUsl.SQL.Add
      ('SELECT fam, imya, fath, sex, profes, vredn_fact, cat_osv, idusl,iduslold,extest, extestold FROM rabotnik WHERE (data_proh BETWEEN :date AND :date2) AND ((UCase(cat_osv) Not Like "П%") or (IsNull(cat_osv))) AND (id_org=:org) ORDER BY fam');
    fDM.QSumUsl.Parameters.ParamByName('date').Value :=
      FormatDateTime('dd.mm.yyyy', DateTimePicker1.Date);
    fDM.QSumUsl.Parameters.ParamByName('date2').Value :=
      FormatDateTime('dd.mm.yyyy', DateTimePicker2.Date);
    fDM.QSumUsl.Parameters.ParamByName('org').Value :=
      fDM.TOrg.FieldByName('id_org').AsString;
    fDM.QSumUsl.Active := true;
    fDM.QSumUsl.First;
    // ---------
    st_in := ExtractFilePath(Application.ExeName);

    ExelTab := CreateOleObject('Excel.Application');
    if fDM.TOrg.FieldByName('dif_price').AsInteger <> 1 then
      ExelTab.Workbooks.Open[st_in + '\Шаблоны\akt6c.xls']
    else
      ExelTab.Workbooks.Open[st_in + '\Шаблоны\akt6cdif.xls'];
    i := 12;
    k := false;
    // fDM.QSumUsl.First;
    while not(fDM.QSumUsl.Eof) do
    begin
      if (fDM.QSumUsl.FieldByName('idusl').AsInteger = 1001) then
      begin
        ExelTab.Cells.Item[i, 2] := fDM.QSumUsl.FieldByName('fam').AsString;
        s := copy(fDM.QSumUsl.FieldByName('imya').AsString, 1, 1);
        ExelTab.Cells.Item[i, 3] := s;
        s := copy(fDM.QSumUsl.FieldByName('fath').AsString, 1, 1);
        ExelTab.Cells.Item[i, 4] := s;
        ExelTab.Cells.Item[i, 5] := fDM.QSumUsl.FieldByName('profes').AsString;
        ExelTab.Cells.Item[7, 5] := fDM.TOrg.FieldByName('imya_poln').AsString;
        ExelTab.Cells.Item[5, 4] := 'с ' + DateToStr(DateTimePicker1.Date) +
          '   по   ' + DateToStr(DateTimePicker2.Date);
        ExelTab.Cells.Item[i, 10].Value := 1;
      end;

      if (fDM.QSumUsl.FieldByName('idusl').AsInteger = 1000) or
        (fDM.QSumUsl.FieldByName('idusl').AsInteger = 1002) then
      begin
        ExelTab.Cells.Item[i, 2] := fDM.QSumUsl.FieldByName('fam').AsString;
        s := copy(fDM.QSumUsl.FieldByName('imya').AsString, 1, 1);
        ExelTab.Cells.Item[i, 3] := s;
        s := copy(fDM.QSumUsl.FieldByName('fath').AsString, 1, 1);
        ExelTab.Cells.Item[i, 4] := s;
        ExelTab.Cells.Item[i, 5] := fDM.QSumUsl.FieldByName('profes').AsString;
        ExelTab.Cells.Item[7, 5] := fDM.TOrg.FieldByName('imya_poln').AsString;
        ExelTab.Cells.Item[5, 4] := 'с ' + DateToStr(DateTimePicker1.Date) +
          '   по   ' + DateToStr(DateTimePicker2.Date);

        if (fDM.QSumUsl.FieldByName('extest').AsString = '1') then
          ExelTab.Cells.Item[i, 11] := 1; // !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

        // Распределение по группам
        if (fDM.QSumUsl.FieldByName('cat_osv').AsString = 'Р1.3гр1') or
          (fDM.QSumUsl.FieldByName('cat_osv').AsString = 'Р1.3гр2') or
          (fDM.QSumUsl.FieldByName('cat_osv').AsString = 'Р1.3гр3') or
          (fDM.QSumUsl.FieldByName('cat_osv').AsString = 'Р1.3гр5') or
          (fDM.QSumUsl.FieldByName('cat_osv').AsString = 'Р1.3гр6') then
        begin
          k := true;
          if (fDM.QSumUsl.FieldByName('sex').AsString = 'муж') then
            ExelTab.Cells.Item[i, 6].Value := 1;
          if (fDM.QSumUsl.FieldByName('sex').AsString = 'жен') then
            ExelTab.Cells.Item[i, 7].Value := 1;
        end;

        if (fDM.QSumUsl.FieldByName('cat_osv').AsString = 'Р1гр1') or
          (fDM.QSumUsl.FieldByName('cat_osv').AsString = 'Р1гр2') or
          (fDM.QSumUsl.FieldByName('cat_osv').AsString = 'Р1гр3') or
          (fDM.QSumUsl.FieldByName('cat_osv').AsString = 'Р1гр5') or
          (fDM.QSumUsl.FieldByName('cat_osv').AsString = 'Р1гр6') then
        begin
          k := true;
          if (fDM.QSumUsl.FieldByName('sex').AsString = 'муж') then
            ExelTab.Cells.Item[i, 6].Value := 1;
          if (fDM.QSumUsl.FieldByName('sex').AsString = 'жен') then
            ExelTab.Cells.Item[i, 7].Value := 1;
        end;

        if (fDM.QSumUsl.FieldByName('cat_osv').AsString = 'Р1.3гр4') or
          (fDM.QSumUsl.FieldByName('cat_osv').AsString = 'Р1гр4') then
        begin
          k := true;
          if (fDM.QSumUsl.FieldByName('sex').AsString = 'муж') then
            ExelTab.Cells.Item[i, 8].Value := 1;
          if (fDM.QSumUsl.FieldByName('sex').AsString = 'жен') then
            ExelTab.Cells.Item[i, 9].Value := 1;
        end;

        if (fDM.QSumUsl.FieldByName('cat_osv').AsString <> '') then
        begin
          k := true;
        end;
        // --------------------------
      end;
      // ---2
      if (fDM.QSumUsl.FieldByName('iduslold').AsInteger = 1000) or
        (fDM.QSumUsl.FieldByName('iduslold').AsInteger = 1002) then
      begin
        ExelTab.Cells.Item[i, 2] := fDM.QSumUsl.FieldByName('fam').AsString;
        s := copy(fDM.QSumUsl.FieldByName('imya').AsString, 1, 1);
        ExelTab.Cells.Item[i, 3] := s;
        s := copy(fDM.QSumUsl.FieldByName('fath').AsString, 1, 1);
        ExelTab.Cells.Item[i, 4] := s;
        ExelTab.Cells.Item[i, 5] := fDM.QSumUsl.FieldByName('profes').AsString;
        ExelTab.Cells.Item[7, 5] := fDM.TOrg.FieldByName('imya_poln').AsString;
        ExelTab.Cells.Item[5, 4] := 'с ' + DateToStr(DateTimePicker1.Date) +
          '   по   ' + DateToStr(DateTimePicker2.Date);

        // Распределение по группам
        if (fDM.QSumUsl.FieldByName('cat_osv').AsString = 'Р1.3гр1') or
          (fDM.QSumUsl.FieldByName('cat_osv').AsString = 'Р1.3гр2') or
          (fDM.QSumUsl.FieldByName('cat_osv').AsString = 'Р1.3гр3') or
          (fDM.QSumUsl.FieldByName('cat_osv').AsString = 'Р1.3гр5') or
          (fDM.QSumUsl.FieldByName('cat_osv').AsString = 'Р1.3гр6') then
        begin
          k := true;
          if (fDM.QSumUsl.FieldByName('sex').AsString = 'муж') then
            ExelTab.Cells.Item[i, 6].Value := -1;
          if (fDM.QSumUsl.FieldByName('sex').AsString = 'жен') then
            ExelTab.Cells.Item[i, 7].Value := -1;
        end;

        if (fDM.QSumUsl.FieldByName('cat_osv').AsString = 'Р1гр1') or
          (fDM.QSumUsl.FieldByName('cat_osv').AsString = 'Р1гр2') or
          (fDM.QSumUsl.FieldByName('cat_osv').AsString = 'Р1гр3') or
          (fDM.QSumUsl.FieldByName('cat_osv').AsString = 'Р1гр5') or
          (fDM.QSumUsl.FieldByName('cat_osv').AsString = 'Р1гр6') then
        begin
          k := true;
          if (fDM.QSumUsl.FieldByName('sex').AsString = 'муж') then
            ExelTab.Cells.Item[i, 6].Value := -1;
          if (fDM.QSumUsl.FieldByName('sex').AsString = 'жен') then
            ExelTab.Cells.Item[i, 7].Value := -1;
        end;

        if (fDM.QSumUsl.FieldByName('cat_osv').AsString = 'Р1.3гр4') or
          (fDM.QSumUsl.FieldByName('cat_osv').AsString = 'Р1гр4') then
        begin
          k := true;
          if (fDM.QSumUsl.FieldByName('sex').AsString = 'муж') then
            ExelTab.Cells.Item[i, 8].Value := -1;
          if (fDM.QSumUsl.FieldByName('sex').AsString = 'жен') then
            ExelTab.Cells.Item[i, 9].Value := -1;
        end;

        if (fDM.QSumUsl.FieldByName('cat_osv').AsString <> '') then
        begin
          k := true;
        end;
        // -------
      end;
      // ----2
      if (fDM.QSumUsl.FieldByName('iduslold').AsInteger = 1001) then
      begin
        ExelTab.Cells.Item[i, 2] := fDM.QSumUsl.FieldByName('fam').AsString +
          ' ' + fDM.QSumUsl.FieldByName('imya').AsString + ' ' +
          fDM.QSumUsl.FieldByName('fath').AsString;
        ExelTab.Cells.Item[7, 5] := fDM.TOrg.FieldByName('imya_poln').AsString;
        ExelTab.Cells.Item[5, 4] := 'с ' + DateToStr(DateTimePicker1.Date) +
          '   по   ' + DateToStr(DateTimePicker2.Date);
        ExelTab.Cells.Item[i, 10].Value := -1;
        // ------
      end;
      // -----
      // end;
      if (fDM.QSumUsl.FieldByName('idusl').AsInteger >= 1000) OR
        (fDM.QSumUsl.FieldByName('iduslold').AsInteger >= 1000) then
        inc(i);
      fDM.QSumUsl.Next;
    end;

    if k then
    begin
      ExelTab.ActiveWorkbook.SaveAs(st_in + 'Журнал\экономист\6C\' +
        fDM.TOrg.FieldByName('imya_org').AsString +
        FormatDateTime('_yyyy_mm_dd', Date) + FormatDateTime('_hh_mm', time) +
        '.xls', EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam,
        EmptyParam, false, EmptyParam, EmptyParam, EmptyParam);
      ExelTab.Quit;
      fDM.TOrg.Next;
      k := false;
    end
    else
    begin
      fDM.TOrg.Next;
      ExelTab.Quit;
    end;
    // -------
  end;
end;

procedure TfUsl.FormCreate(Sender: TObject);
begin
  DateTimePicker1.Date := Date;
  DateTimePicker2.Date := Date;
end;

end.
