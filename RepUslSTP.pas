unit RepUslSTP;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.DBCtrls, Vcl.ComCtrls, Vcl.StdCtrls,
  MemTableDataEh, Data.DB, DBGridEhGrouping, ToolCtrlsEh, DBGridEhToolCtrls,
  DynVarsEh, EhLibVCL, GridsEh, DBAxisGridsEh, DBGridEh, MemTableEh,
  Data.Win.ADODB,ComObj,TlHelp32,Winapi.ShellAPI;

type
  TRepUslSTM = class(TForm)
    txt2: TStaticText;
    dtp1: TDateTimePicker;
    txt3: TStaticText;
    dtp2: TDateTimePicker;
    dbgrdh1: TDBGridEh;
    btn1: TButton;
    qry1: TADOQuery;
    ds1: TDataSource;
    btn2: TButton;
    btn3: TButton;
    procedure btn1Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure btn2Click(Sender: TObject);
    function GetProcessId: TProcessEntry32;
    procedure btn3Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  RepUslSTM: TRepUslSTM;
procedure FIO(var ExelTab: OleVariant;var i:Integer;var qry1:TADOQuery;r:Integer);

implementation
  uses DM;
{$R *.dfm}


function TRepUslSTM.GetProcessId: TProcessEntry32;
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

procedure TRepUslSTM.btn1Click(Sender: TObject);
begin

qry1.Active := False;
qry1.Parameters.ParamByName('date').Value :=  FormatDateTime('dd.mm.yyyy', dtp1.Date);
qry1.Parameters.ParamByName('date2').Value := FormatDateTime('dd.mm.yyyy', dtp2.Date);

qry1.Active := true;

end;

procedure FIO(var ExelTab: OleVariant;var i:Integer;var qry1:TADOQuery;r:Integer);
begin
  ExelTab.Cells.Item[i, r].Value := 1;
ExelTab.Cells.Item[i, 2].Value := qry1.FieldByName('fam').AsString+' '+qry1.FieldByName('imya').AsString+' '+qry1.FieldByName('fath').AsString;
ExelTab.Cells.Item[i, 3].Value := qry1.FieldByName('profes').AsString;
ExelTab.Cells.Item[i, 17].Value := 1;
i:=i+1;
end;


procedure TRepUslSTM.btn2Click(Sender: TObject);
var

ExelTab: OleVariant;
  st_in: string;
  i: integer;

begin

  if qry1.RecordCount=0 then Exit;


   st_in := ExtractFilePath(Application.ExeName);
   ExelTab := CreateOleObject('Excel.Application');
   ExelTab.Workbooks.Open[st_in + '\�������\aktdif2021.xls'];

   ExelTab.Visible :=False;
   i:=10;
   while not (qry1.Eof) do
   begin



   if qry1.FieldByName('sex').AsString = '���' then
    begin


   case qry1.FieldByName('idusl').AsInteger of
      16 : FIO(ExelTab,i,qry1,4);
      17 : FIO(ExelTab,i,qry1,6);
      18 : FIO(ExelTab,i,qry1,8);
      19 : FIO(ExelTab,i,qry1,10);
      20 : FIO(ExelTab,i,qry1,12);
      9 : FIO(ExelTab,i,qry1,14);
      10 : FIO(ExelTab,i,qry1,15);
      11 : FIO(ExelTab,i,qry1,16);

   end;


   end
    else
   begin

      case qry1.FieldByName('idusl').AsInteger of
      16 :FIO(ExelTab,i,qry1,5);
      17 : FIO(ExelTab,i,qry1,7);
      18 : FIO(ExelTab,i,qry1,9);
      19 : FIO(ExelTab,i,qry1,11);
      20 : FIO(ExelTab,i,qry1,12);
      9 : FIO(ExelTab,i,qry1,14);
      10 : FIO(ExelTab,i,qry1,15);
      11 : FIO(ExelTab,i,qry1,16);
   end;

   end;

    qry1.Next;
   end;

  //MyRange:=ExelTab.ActiveWorkbook.ActiveSheet.Rows('10');
   //EntireRow.Hidden := True;   IntToStr(i+1)+':1218'  .EntireRow.Hidden := True



  ExelTab.ActiveWorkbook.SaveAs(st_in + '������\���������\����\' + fDM.TOrg.FieldByName('imya_org').AsString + FormatDateTime('_yyyy_mm_dd', Date) + FormatDateTime('_hh_mm', time) + '.xls', EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam, false, EmptyParam, EmptyParam, EmptyParam);
  ExelTab.ActiveWorkbook.Close;
  ExelTab.Quit;

  ExelTab := Unassigned;

  //KillTask('excel.exe');

   TerminateProcess(OpenProcess(PROCESS_TERMINATE, BOOL(0), GetProcessId.th32ProcessID), 0);

  ShellExecute(Self.Handle, 'open', Pchar(st_in + '\������\���������\����\'), nil, nil, SW_SHOWNORMAL);

end;

procedure TRepUslSTM.btn3Click(Sender: TObject);
var

ExelTab: OleVariant;
  st_in: string;
  i: integer;
begin
   if qry1.RecordCount=0 then Exit;


   st_in := ExtractFilePath(Application.ExeName);
   ExelTab := CreateOleObject('Excel.Application');
   ExelTab.Workbooks.Open[st_in + '\�������\akt6cdif2021.xls'];

   ExelTab.Visible :=False;
   i:=10;
   while not (qry1.Eof) do
   begin







         if qry1.FieldByName('idusl').AsInteger = 1001 then
          begin
             ExelTab.Cells.Item[i, 7].Value := 1;ExelTab.Cells.Item[i, 8].Value := 1;

               ExelTab.Cells.Item[i, 2].Value := qry1.FieldByName('fam').AsString;
            ExelTab.Cells.Item[i, 3].Value := qry1.FieldByName('imya').AsString;
            ExelTab.Cells.Item[i, 4].Value := qry1.FieldByName('fath').AsString;

            ExelTab.Cells.Item[i, 5].Value := qry1.FieldByName('profes').AsString;
            i:=i+1;
          end;
        if qry1.FieldByName('idusl').AsInteger = 1002 then
          begin
            ExelTab.Cells.Item[i, 6].Value := 1;ExelTab.Cells.Item[i, 8].Value := 1;
            ExelTab.Cells.Item[i, 2].Value := qry1.FieldByName('fam').AsString;
            ExelTab.Cells.Item[i, 3].Value := qry1.FieldByName('imya').AsString;
            ExelTab.Cells.Item[i, 4].Value := qry1.FieldByName('fath').AsString;

            ExelTab.Cells.Item[i, 5].Value := qry1.FieldByName('profes').AsString;
            i:=i+1;
          end;

           if qry1.FieldByName('idusl').AsInteger = 12 then
          begin
            ExelTab.Cells.Item[i, 6].Value := 1;ExelTab.Cells.Item[i, 8].Value := 1;
            ExelTab.Cells.Item[i, 2].Value := qry1.FieldByName('fam').AsString;
            ExelTab.Cells.Item[i, 3].Value := qry1.FieldByName('imya').AsString;
            ExelTab.Cells.Item[i, 4].Value := qry1.FieldByName('fath').AsString;

            ExelTab.Cells.Item[i, 5].Value := qry1.FieldByName('profes').AsString;
            i:=i+1;
          end;



    qry1.Next;
   end;

  //MyRange:=ExelTab.ActiveWorkbook.ActiveSheet.Rows('10');
   //EntireRow.Hidden := True;   IntToStr(i+1)+':1218'  .EntireRow.Hidden := True



  ExelTab.ActiveWorkbook.SaveAs(st_in + '������\���������\����\' + fDM.TOrg.FieldByName('imya_org').AsString + FormatDateTime('_yyyy_mm_dd', Date) + FormatDateTime('_hh_mm', time) + '_6�'+'.xls', EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam, false, EmptyParam, EmptyParam, EmptyParam);
  ExelTab.ActiveWorkbook.Close;
  ExelTab.Quit;

  ExelTab := Unassigned;

  //KillTask('excel.exe');

   TerminateProcess(OpenProcess(PROCESS_TERMINATE, BOOL(0), GetProcessId.th32ProcessID), 0);

  ShellExecute(Self.Handle, 'open', Pchar(st_in + '\������\���������\����\'), nil, nil, SW_SHOWNORMAL);
end;

procedure TRepUslSTM.FormCreate(Sender: TObject);
begin
 dtp1.Date := Date;
 dtp2.Date := Date;
end;

end.
