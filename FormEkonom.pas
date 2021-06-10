unit FormEkonom;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, DBGridEhGrouping, ToolCtrlsEh,
  DBGridEhToolCtrls, DynVarsEh, MemTableDataEh, Data.DB, MemDS, DBAccess,
  VirtualQuery, Data.Win.ADODB, MemTableEh, Vcl.StdCtrls, EhLibVCL, GridsEh,ShellApi,
  DBAxisGridsEh, DBGridEh, Vcl.ComCtrls, Vcl.DBCtrls, OleServer, ComObj,TlHelp32, ExcelXP,
  Vcl.ExtCtrls;


type
  TRabRecord = record
    Fam,imya,fath,Sex,profes,Codfactor,CodPrise,Naim,naim1:string;
    data_r:TDateTime;
    kol:Int8;
    cena:Real;
  end;

type
  TfrmEkonom = class(TForm)
    btn1: TButton;
    dblkcbb1: TDBLookupComboBox;
    dblkcbb2: TDBLookupComboBox;
    dtp1: TDateTimePicker;
    dtp2: TDateTimePicker;
    txt1: TStaticText;
    txt2: TStaticText;
    txt3: TStaticText;
    txt5: TStaticText;
    dbgrdh1: TDBGridEh;
    dbgrdh2: TDBGridEh;
    btn5: TButton;
    ds1: TDataSource;
    mtblh1: TMemTableEh;
    smlntfldmtblh1Idgroup: TSmallintField;
    strngfldmtblh1Fam: TStringField;
    strngfldmtblh1imya: TStringField;
    strngfldmtblh1fath: TStringField;
    strngfldmtblh1Sex: TStringField;
    dtfldmtblh1data_r: TDateField;
    strngfldmtblh1profes: TStringField;
    smlntfldmtblh1kol: TSmallintField;
    strngfldmtblh1Codfactor: TStringField;
    strngfldmtblh1CodPrise: TStringField;
    strngfldmtblh1Naim: TStringField;
    strngfldmtblh1naim1: TStringField;
    crncyfldmtblh1cena: TCurrencyField;
    qry1: TADOQuery;
    ds2: TDataSource;
    qryMax: TADOQuery;
    qry6c: TADOQuery;
    qryID: TADOQuery;
    vrtlqry1: TVirtualQuery;
    pnl1: TPanel;
    pgcTabPages: TPageControl;
    ts1: TTabSheet;
    ts2: TTabSheet;
    dbgrdh3: TDBGridEh;
    edt1: TEdit;
    dbgrdh4: TDBGridEh;
    ds3: TDataSource;
    ds4: TDataSource;
    ds5: TADODataSet;
    wdstrngfldds5CodPrise: TWideStringField;
    wdmfldds5Naim: TWideMemoField;
    bcdfldds5Cena: TBCDField;
    dbnvgr1: TDBNavigator;
    procedure FormCreate(Sender: TObject);
    procedure btn1Click(Sender: TObject);
    procedure PostMem(Idgroup:Int8;Fam:string;imya:string;fath:string;Sex:string;data_r:TDateTime;profes:string;kol:Int8;cena:Real;Codfactor:string;CodPrise:string;Naim:string;naim1:string);
    function Get302(id:Integer):TRabRecord;
    function Get6c(id:Integer):TRabRecord;
    procedure SetRowHeight(ExelTab: OleVariant;i:Integer;l:integer;k:integer);
    function KillTask(ExeFileName: string): integer;
    function GetProcessId: TProcessEntry32;
    procedure btn5Click(Sender: TObject);
    procedure    AktOBDP();
    procedure edt1Change(Sender: TObject);
    procedure dbgrdh4DblClick(Sender: TObject);

  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmEkonom: TfrmEkonom;

implementation

uses DM;

{$R *.dfm}


function TfrmEkonom.GetProcessId: TProcessEntry32;
const
  PROCESS_TERMINATE=$0001;
var
  ContinueLoop: BOOL;
  FSnapshotHandle: THandle;
  FProcessEntry32: TProcessEntry32;
begin
  FSnapshotHandle := CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0);
  FProcessEntry32.dwSize := Sizeof(FProcessEntry32);
  ContinueLoop := Process32First(FSnapshotHandle,
                                 FProcessEntry32);

  while integer(ContinueLoop) <> 0 do
  begin
    if ((UpperCase(ExtractFileName(FProcessEntry32.szExeFile)) =
         UpperCase('excel.exe'))
     or (UpperCase(FProcessEntry32.szExeFile) =
         UpperCase('excel.exe'))) then
      Result := FProcessEntry32;
    ContinueLoop := Process32Next(FSnapshotHandle,
                                  FProcessEntry32);
  end;

  CloseHandle(FSnapshotHandle);
end;

function TfrmEkonom.KillTask(ExeFileName: string): integer;

const

  PROCESS_TERMINATE=$0001;

var

  ContinueLoop: BOOL;

  FSnapshotHandle: THandle;

  FProcessEntry32: TProcessEntry32;

begin

  result := 0;



  FSnapshotHandle := CreateToolhelp32Snapshot

                     (TH32CS_SNAPPROCESS, 0);

  FProcessEntry32.dwSize := Sizeof(FProcessEntry32);

  ContinueLoop := Process32First(FSnapshotHandle,

                                 FProcessEntry32);



  while integer(ContinueLoop) <> 0 do

  begin

    if ((UpperCase(ExtractFileName(FProcessEntry32.szExeFile)) =

         UpperCase(ExeFileName))

     or (UpperCase(FProcessEntry32.szExeFile) =

         UpperCase(ExeFileName))) then

      Result := Integer(TerminateProcess(OpenProcess(

                        PROCESS_TERMINATE, BOOL(0),

                        FProcessEntry32.th32ProcessID), 0));

    ContinueLoop := Process32Next(FSnapshotHandle,

                                  FProcessEntry32);

  end;



  CloseHandle(FSnapshotHandle);

end;

procedure TfrmEkonom.SetRowHeight(ExelTab: OleVariant;i:Integer;l:integer;k:integer);
   begin

    if Frac(l/k)<1 then
    begin
   ExelTab.Rows[i].RowHeight:=(Trunc(l/k)+1)*15;
    end;
    if Frac(l/k)=0 then
    begin

     ExelTab.Rows[i].RowHeight:=(l/k)*15;
    end;


   end;

function TfrmEkonom.Get302(id:Integer):TRabRecord;
        var rec:TRabRecord;
        begin

        rec.cena :=0;

        qryMax.Active := False;

        qryMax.Parameters.ParamByName('date').Value :=  FormatDateTime('dd.mm.yyyy', dtp1.Date);
        qryMax.Parameters.ParamByName('date2').Value := FormatDateTime('dd.mm.yyyy', dtp2.Date);
        qryMax.Parameters.ParamByName('org').Value := DM.fDM.TOrg.FieldByName('id_org').AsString;
        qryMax.Parameters.ParamByName('id_dog').Value := dblkcbb2.KeyValue;
        qryMax.Parameters.ParamByName('id').Value := id;

        qryMax.Active := true;
        if   qryMax.RecordCount <> 0 then
        begin
        qryMax.First;

        rec.Fam :=qryMax.FieldByName('Fam').Value;
        rec.imya :=qryMax.FieldByName('Imya').Value;
        rec.fath :=qryMax.FieldByName('Fath').AsString;
        rec.Sex :=qryMax.FieldByName('Sex').Value;
        rec.data_r :=qryMax.FieldByName('data_r').Value;
        rec.profes :=qryMax.FieldByName('profes').Value;
        rec.Codfactor :=qryMax.FieldByName('code').Value;
        rec.CodPrise :=qryMax.FieldByName('kod_price').Value;
        rec.kol :=1;
        rec.cena := qryMax.FieldByName('MaxCena').Value;
        rec.Naim := qryMax.FieldByName('Naim').Value;
        rec.naim1 := qryMax.FieldByName('naim1').Value;
        Result:=rec;
        Exit;
        end
        else

        qryID.Active := False;

        qryID.Parameters.ParamByName('id').Value := id;
        qryID.Active := true;
        qryID.First;

        rec.Fam :=qryID.FieldByName('Fam').Value;
        rec.imya :=qryID.FieldByName('Imya').Value;
        rec.fath :=qryID.FieldByName('Fath').Value;
        rec.Naim := 'Проверить Факторы!!!';
        Result:=rec;
        Exit;
        begin

        end;



        end;

function TfrmEkonom.Get6c(id:Integer):TRabRecord;
        var rec:TRabRecord;

        begin
        rec.cena :=0;
        qry6c.Active := False;
        qry6c.Parameters.ParamByName('id').DataType := TFieldType.ftInteger;
        qry6c.Parameters.ParamByName('id').Value :=id;

        qry6c.Active := true;
        if   qry6c.RecordCount <> 0 then
        begin
        qry6c.First;

        rec.Fam :=qry6c.FieldByName('Fam').Value;
        rec.imya :=qry6c.FieldByName('Imya').Value;
        rec.fath :=qry6c.FieldByName('Fath').Value;
        rec.Sex :=qry6c.FieldByName('Sex').Value;
        rec.data_r :=qry6c.FieldByName('data_r').Value;
        rec.profes :=qry6c.FieldByName('profes').Value;
//        rec.Codfactor :=qry6c.FieldByName('code').Value;

        rec.kol :=1;

        if (qry6c.FieldByName('sex').Value = 'муж') and (qry6c.FieldByName('voz').Value < 40) then
        begin
        rec.Naim := 'Периодический медицинский осмотр (ВЭК) работников, обеспечивающих движение поездов (мужчины до 40 лет) - 6Ц';
        rec.naim1 := 'Периодический медицинский осмотр (ВЭК) работников, обеспечивающих движение поездов (мужчины до 40 лет) - 6Ц';
        rec.cena := 3986.2;
        rec.CodPrise :='2.1';
        end;

        if (qry6c.FieldByName('sex').Value = 'муж') and (qry6c.FieldByName('voz').Value >= 40) then
        begin
        rec.Naim := 'Периодический медицинский осмотр (ВЭК) работников, обеспечивающих движение поездов (мужчины после 40 лет) - 6Ц';
        rec.naim1 := 'Периодический медицинский осмотр (ВЭК) работников, обеспечивающих движение поездов (мужчины после 40 лет) - 6Ц';
        rec.cena := 3986.2;
        rec.CodPrise :='2.2';
        end;


        if (qry6c.FieldByName('sex').Value = 'жен') and (qry6c.FieldByName('voz').Value < 40) then
        begin
        rec.Naim := 'Периодический медицинский осмотр (ВЭК) работников, обеспечивающих движение поездов (женщины до 40 лет) - 6Ц';
        rec.naim1 := 'Периодический медицинский осмотр (ВЭК) работников, обеспечивающих движение поездов (женщины до 40 лет) - 6Ц';
        rec.cena := 4070.76;
        rec.CodPrise :='2.3';
        end;

        if (qry6c.FieldByName('sex').Value = 'жен') and (qry6c.FieldByName('voz').Value >= 40) then
        begin
        rec.Naim := 'Периодический медицинский осмотр (ВЭК) работников, обеспечивающих движение поездов (женщины после 40 лет) - 6Ц';
        rec.naim1 := 'Периодический медицинский осмотр (ВЭК) работников, обеспечивающих движение поездов (женщины после 40 лет) - 6Ц';
        rec.cena := 4528.18;
        rec.CodPrise :='2.4';
        end;

        end;

        Result:=rec;


        end;

procedure TfrmEkonom.PostMem(Idgroup:Int8;Fam:string;imya:string;fath:string;Sex:string;data_r:TDateTime;profes:string;kol:Int8;cena:Real;Codfactor:string;CodPrise:string;Naim:string;naim1:string);
begin

    with  ds1.DataSet do
    Begin


      Insert;

      FieldByName('Idgroup').Value := Idgroup;
      FieldByName('Fam').AsString := Fam;
      FieldByName('imya').AsString := imya;
      FieldByName('fath').AsString := fath;
      FieldByName('Sex').AsString := Sex;
      FieldByName('data_r').Value := data_r;
      FieldByName('profes').AsString := profes;
      FieldByName('kol').Value := kol;
      FieldByName('cena').value := cena;
      FieldByName('Codfactor').AsString := Codfactor;
      FieldByName('CodPrise').AsString := CodPrise;
      FieldByName('Naim').AsString := Naim;
      FieldByName('naim1').AsString := naim1;
      Post;
    End;

end;

procedure TfrmEkonom.btn1Click(Sender: TObject);
var rec,rec1:TRabRecord;
var h: integer;
begin
      h:=0;

      mtblh1.EmptyTable;

//Перебераем каждого сотрудника
//Составляем
//0) c6 повт (idusl=9) пропускаем и м\повт (idusl=1001) пропускаем и  категория освидетельствования cat_osv не содержит п
//1) выводим всех кто относиться к факторам
//2) потом следует гинекология
//3) потом 6с
//4) пофилактика



qry1.Active := False;
qry1.Parameters.ParamByName('date').Value :=  FormatDateTime('dd.mm.yyyy', dtp1.Date);
qry1.Parameters.ParamByName('date2').Value := FormatDateTime('dd.mm.yyyy', dtp2.Date);
qry1.Parameters.ParamByName('org').Value := DM.fDM.TOrg.FieldByName('id_org').AsString;
qry1.Active := true;


  qry1.First;
  while not (qry1.Eof) do
  begin
  // пропускаем

  if (qry1.FieldByName('idusl').Value = 9) or  (qry1.FieldByName('idusl').Value = 1001) then
        begin
         qry1.Next;
             Continue;
        end;


  if (qry1.FieldByName('idusl').Value >= 1) and  (qry1.FieldByName('idusl').Value <= 8) then
        begin
        rec := Get302(qry1.FieldByName('id').Value);
           PostMem(1,rec.Fam,rec.imya,rec.fath,rec.Sex,rec.data_r,rec.profes,rec.kol,rec.cena,rec.Codfactor,rec.CodPrise,rec.Naim,rec.naim1);
             if qry1.FieldByName('Sex').Value  = 'жен' then h:=h+1;
             qry1.Next;
             Continue;
        end;


    if (qry1.FieldByName('idusl').Value = 10) then
        begin

    PostMem(4,
            qry1.FieldByName('fam').AsString,
            qry1.FieldByName('imya').AsString,
            qry1.FieldByName('fath').AsString,
            qry1.FieldByName('Sex').AsString,
            qry1.FieldByName('data_r').Value,
            qry1.FieldByName('profes').AsString,
            1,
            4719,
            '',
            '6.2',
            'Профилактический медицинский осмотр 1 раз в 5 лет в амбулаторно-поликлинических условиях',
            'Профилактический медицинский осмотр 1 раз в 5 лет в амбулаторно-поликлинических условиях');
            //if qry1.FieldByName('Sex').Value  = 'жен' then h:=h+1;
            qry1.Next;
            Continue;
        end;

       if (qry1.FieldByName('idusl').Value = 11) then
        begin

    PostMem(4,
            qry1.FieldByName('fam').AsString,
            qry1.FieldByName('imya').AsString,
            qry1.FieldByName('fath').AsString,
            qry1.FieldByName('Sex').AsString,
            qry1.FieldByName('data_r').Value,
            qry1.FieldByName('profes').AsString,
            1,
            6540,
            '',
            '6.3',
            'Профилактический медицинский осмотр 1 раз в 5 лет в  условиях круглосуточного стационара',
            'Профилактический медицинский осмотр 1 раз в 5 лет в  условиях круглосуточного стационара');
           // if qry1.FieldByName('Sex').Value  = 'жен' then h:=h+1;
            qry1.Next;
            Continue;
        end;

        if (qry1.FieldByName('idusl').Value = 1002)  or (qry1.FieldByName('idusl').Value = 1000) then
        begin
             rec := Get302(qry1.FieldByName('id').Value);
             rec1 := Get6c(qry1.FieldByName('id').Value);


             if rec.cena > rec1.cena then
             begin
             PostMem(1,rec.Fam,rec.imya,rec.fath,rec.Sex,rec.data_r,rec.profes,rec.kol,rec.cena,rec.Codfactor,rec.CodPrise,rec.Naim,rec.naim1);
             if qry1.FieldByName('Sex').Value  = 'жен' then h:=h+1;
             qry1.Next;
             end;

             if rec.cena <= rec1.cena then
             begin
             PostMem(3,rec1.Fam,rec1.imya,rec1.fath,rec1.Sex,rec1.data_r,rec1.profes,rec1.kol,rec1.cena,rec1.Codfactor,rec1.CodPrise,rec1.Naim,rec1.naim1);
             //if qry1.FieldByName('Sex').Value  = 'жен' then h:=h+1;
             qry1.Next;
             end;
              Continue;
        end;



    qry1.Next;
  end;
  if h<>0 then
  begin
  PostMem(2,'','','','',StrToDate('01.01.2019'),'',h,226.87,'','6.1','Профилактический прием врача-акушера-гинеколога','Профилактический прием врача-акушера-гинеколога');
  end;




 mtblh1.SortByFields('Idgroup ASC,CodPrise ASC,CodFactor ASC,Fam ASC');

 vrtlqry1.Active := False;
 vrtlqry1.Active := True;
end;

procedure TfrmEkonom.AktOBDP();
var
 ExelTab: OleVariant;
 n: integer;
 st_in: string;
 sumall: Real;
   begin

    n:=8;
    sumall:=0;

    //vrtlqry1.Active := False;
    //vrtlqry1.Active := True;

    if vrtlqry1.RecordCount=0 then Exit;

    st_in := ExtractFilePath(Application.ExeName);

    ExelTab := CreateOleObject('Excel.Application');

    ExelTab.Workbooks.Open[st_in + '\Шаблоны\ReestrOBDP.xls'];


    ExelTab.Visible :=False;

    vrtlqry1.First;

   while not (vrtlqry1.Eof) do
    begin
      ExelTab.Cells.Item[n, 1].Value := vrtlqry1.FieldByName('CodPrise').AsString;
      ExelTab.Cells.Item[n, 2].Value := vrtlqry1.FieldByName('naim1').AsString;
      ExelTab.Cells.Item[n, 3].Value := vrtlqry1.FieldByName('kol').AsString;
      ExelTab.Cells.Item[n, 4].Value := vrtlqry1.FieldByName('cena').AsString;
      ExelTab.Cells.Item[n, 6].Value := vrtlqry1.FieldByName('sum').AsString;

      sumall:=sumall+vrtlqry1.FieldByName('sum').Value;
      vrtlqry1.Next;
      n:=n+1;
     end;

    ExelTab.Cells.Item[n, 5].Value := 'Итого:';
    ExelTab.Cells.Item[n, 6].Value := sumall;
    ExelTab.rows[n].font.bold:=True;

    ExelTab.Range[ExelTab.Cells[8, 1], ExelTab.Cells[n, 6]].Select;
    ExelTab.Selection.Borders.LineStyle := 1;

    ExelTab.Range[ExelTab.Cells[n+2, 1], ExelTab.Cells[n+2, 6]].MergeCells :=True;

    ExelTab.Cells.Item[n+2, 1].Value :='Итого: общая стоимость услуг составила: ' + FloatToStr(sumall) + ' руб.,в т.ч. НДС';

    n:=n+4;
    ExelTab.Range[ExelTab.Cells[n  , 1], ExelTab.Cells[n  , 2]].MergeCells :=True;
    ExelTab.Range[ExelTab.Cells[n+1, 1], ExelTab.Cells[n+1, 2]].MergeCells :=True;
    ExelTab.Range[ExelTab.Cells[n+2, 1], ExelTab.Cells[n+2, 2]].MergeCells :=True;
    ExelTab.Range[ExelTab.Cells[n+3, 1], ExelTab.Cells[n+3, 2]].MergeCells :=True;
    ExelTab.Range[ExelTab.Cells[n+4, 1], ExelTab.Cells[n+4, 2]].MergeCells :=True;

    ExelTab.rows[n].font.bold:=True;
    ExelTab.Cells.Item[n, 1].Value :='От Заказчика:';
    ExelTab.Cells.Item[n+1, 1].Value :='Начальник Южно-Уральской дирекции управления движением';
    ExelTab.Cells.Item[n+2, 1].Value :='структурного подразделения Центральной дирекции';
    ExelTab.Cells.Item[n+3, 1].Value :='управления движением-филиала ОАО"РЖД"';
    ExelTab.Cells.Item[n+4, 1].Value :='_______________________________П.А.Коваленко';
    ExelTab.Cells.Item[n+5, 1].Value :='М.П.';

    ExelTab.Range[ExelTab.Cells[n  , 3], ExelTab.Cells[n  , 6]].MergeCells :=True;
    ExelTab.Range[ExelTab.Cells[n+1, 3], ExelTab.Cells[n+1, 6]].MergeCells :=True;
    ExelTab.Range[ExelTab.Cells[n+2, 3], ExelTab.Cells[n+2, 6]].MergeCells :=True;
    ExelTab.Range[ExelTab.Cells[n+3, 3], ExelTab.Cells[n+3, 6]].MergeCells :=True;
    ExelTab.Range[ExelTab.Cells[n+4, 3], ExelTab.Cells[n+4, 6]].MergeCells :=True;

    ExelTab.Cells.Item[n, 3].Value :='От Заказчика:';
    ExelTab.Cells.Item[n+1, 3].Value :='Главный врач ЧУЗ "РЖД-Медицина" г.Курган"';
    ExelTab.Cells.Item[n+4, 3].Value :='_____________________________В.Н.Костина';
    ExelTab.Cells.Item[n+5, 3].Value :='М.П.';


    ExelTab.ActiveSheet.PageSetup.PrintArea := '$A$1:$F$'+IntToSTR(n+5);

    ExelTab.ActiveWorkbook.SaveAs(st_in + 'Журнал\экономист\акты\' + fDM.TOrg.FieldByName('imya_org').AsString + FormatDateTime('_yyyy_mm_dd', Date) + FormatDateTime('_hh_mm', time)+'_Реестр' + '.xls', EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam, false, EmptyParam, EmptyParam, EmptyParam);
    ExelTab.ActiveWorkbook.Close;
    ExelTab.Quit;

    ExelTab := Unassigned;

    //KillTask('excel.exe');

   TerminateProcess(OpenProcess(PROCESS_TERMINATE, BOOL(0), GetProcessId.th32ProcessID), 0);

    ShellExecute(Self.Handle, 'open', Pchar(st_in + '\Журнал\экономист\акты\'), nil, nil, SW_SHOWNORMAL);

   end;

procedure TfrmEkonom.btn5Click(Sender: TObject);
begin
 AktOBDP();
end;

procedure TfrmEkonom.dbgrdh4DblClick(Sender: TObject);
begin

 vrtlqry1.Append;

 vrtlqry1.FieldByName('CodPrise').AsString :=
        ds5.FieldByName('CodPrise').AsString;
 vrtlqry1.FieldByName('naim1').AsString :=
        ds5.FieldByName('Naim').AsString;
 vrtlqry1.FieldByName('cena').AsString :=
        ds5.FieldByName('Cena').AsString;
 vrtlqry1.FieldByName('kol').ReadOnly := False;
 vrtlqry1.FieldByName('kol').Value :=1;
 vrtlqry1.FieldByName('sum').ReadOnly := False;
 vrtlqry1.FieldByName('sum').AsString :=
        ds5.FieldByName('Cena').AsString;

vrtlqry1.Post;
 edt1.Clear;

end;

procedure TfrmEkonom.edt1Change(Sender: TObject);
begin
 if edt1.Text <> '' then
  begin
    ds5.Filtered := false;
    ds5.Filter :='CodPrise like %'+edt1.Text+'%';
    ds5.Filtered := true;
  end
  else
    ds5.Filtered := false;
end;

procedure TfrmEkonom.FormCreate(Sender: TObject);

begin
 dtp1.Date := Date;
 dtp2.Date := Date;

 //dblkcbb1.KeyValue := 27;
 dblkcbb2.KeyValue := 2;



end;





end.
