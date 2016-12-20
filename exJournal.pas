unit exJournal;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, StdCtrls, ComCtrls, ComObj,Word2000, WordXP;

type
  TfExJournal = class(TForm)
    Panel2: TPanel;
    DateTimePicker1: TDateTimePicker;
    DateTimePicker2: TDateTimePicker;
    Label1: TLabel;
    Label2: TLabel;
    Button1: TButton;
    Button2: TButton;
    procedure Button1Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure Button2Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  fExJournal: TfExJournal;

implementation

uses DM, editor;

{$R *.dfm}

procedure TfExJournal.Button1Click(Sender: TObject);
var ExelTab:Variant;
    st_in,st,q:string;
    i,k:integer;
begin
 st_in:=ExtractFilePath(Application.ExeName);
 ExelTab:= CreateOleObject('Excel.Application');
 ExelTab.Workbooks.Open[st_in+'\�������\journal.xls'];

 try
 fDM.Query.SQL.Clear;
 q:='SELECT rabotnik.npp, rabotnik.data_proh, rabotnik.fam, rabotnik.imya, rabotnik.fath, rabotnik.id_org, organiz.imya_org, organiz.imya_poln, rabotnik.data_r, rabotnik.sex, rabotnik.profes, rabotnik.mkb_code, rabotnik.prkrab, rabotnik.dop_pri_usl,'+' rabotnik.lpu, rabotnik.slexp, rabotnik.otkl, rabotnik.def, rabotnik.result, rabotnik.datezaklMSE, rabotnik.datenaprMSE, rabotnik.zaklMSE, rabotnik.dopinf, rabotnik.sostexp, rabotnik.podpis FROM'+' rabotnik LEFT JOIN organiz ON rabotnik.id_org=organiz.id_org WHERE (prkrab<>"") AND (data_proh BETWEEN :date AND :date2) ORDER BY id ASC';
 fDM.Query.SQL.Add(q);
 fDM.Query.Parameters.ParamByName('date').Value:=FormatDateTime('dd.mm.yyyy',DateTimePicker1.Date);
 fDM.Query.Parameters.ParamByName('date2').Value:=FormatDateTime('dd.mm.yyyy',DateTimePicker2.Date);
 fDM.Query.Open;

 fDM.Query.First;
 i:=5;
 k:=0;
 while not(fDM.Query.Eof) do
 begin
  inc(k);
  ExelTab.Cells.Item[i,1]:=fDM.Query.FieldByName('npp').AsString;
  ExelTab.Cells.Item[i,2]:=fDM.Query.FieldByName('data_proh').AsString;
  ExelTab.Cells.Item[i,3]:=fDM.Query.FieldByName('lpu').AsString;
  ExelTab.Cells.Item[i,4]:=fDM.Query.FieldByName('fam').AsString+' '+fDM.Query.FieldByName('imya').AsString+' '+fDM.Query.FieldByName('fath').AsString;
  ExelTab.Cells.Item[i,5]:=fDM.Query.FieldByName('imya_org').AsString;
  ExelTab.Cells.Item[i,6]:=fDM.Query.FieldByName('data_r').AsString;
  ExelTab.Cells.Item[i,7]:=fDM.Query.FieldByName('sex').AsString;
  ExelTab.Cells.Item[i,8]:=fDM.Query.FieldByName('profes').AsString;
  ExelTab.Cells.Item[i,9]:=fDM.Query.FieldByName('mkb_code').AsString;
  ExelTab.Cells.Item[i,10]:='�� ����.';
  ExelTab.Cells.Item[i,11]:=fDM.Query.FieldByName('prkrab').AsString;
  ExelTab.Cells.Item[i,12]:=fDM.Query.FieldByName('otkl').AsString;
  ExelTab.Cells.Item[i,13]:=fDM.Query.FieldByName('def').AsString;
  ExelTab.Cells.Item[i,14]:=fDM.Query.FieldByName('result').AsString;
  ExelTab.Cells.Item[i,15]:=fDM.Query.FieldByName('dop_pri_usl').AsString; //��� � ��� �������
  ExelTab.Cells.Item[i,16]:=fDM.Query.FieldByName('datenaprMSE').AsString;
  ExelTab.Cells.Item[i,17]:=fDM.Query.FieldByName('zaklMSE').AsString;
  ExelTab.Cells.Item[i,18]:=fDM.Query.FieldByName('datezaklMSE').AsString;
  ExelTab.Cells.Item[i,19]:=fDM.Query.FieldByName('dopinf').AsString;
  ExelTab.Cells.Item[i,20]:=fDM.Query.FieldByName('sostexp').AsString;
  ExelTab.Cells.Item[i,21]:=fDM.Query.FieldByName('podpis').AsString;

  inc(i);
  fDM.Query.Next;
 end;
 ExelTab.Range[ExelTab.Cells[5, 1], ExelTab.Cells[i-1, 21]].Select;
 ExelTab.Selection.Borders.LineStyle:=1;

 ExelTab.Visible := true;
 except
 ExelTab.Quit;
 showmessage('������ ������������ �������!');
 end;
end;

procedure TfExJournal.Button2Click(Sender: TObject);
var Shablon,h,s,vschange:OleVariant;
    q:string;
    i:integer;
begin
 fDM.Query.SQL.Clear;
 q:='SELECT rabotnik.data_proh, rabotnik.fam, rabotnik.imya, rabotnik.fath, rabotnik.id_org, organiz.imya_org, organiz.imya_poln, rabotnik.data_r, rabotnik.sex, rabotnik.profes, rabotnik.mkb_code, rabotnik.prkrab,rabotnik.dop_pri_usl,'+' rabotnik.lpu, rabotnik.slexp, rabotnik.otkl, rabotnik.def, rabotnik.result, rabotnik.datezaklMSE, rabotnik.datenaprMSE, rabotnik.zaklMSE, rabotnik.dopinf, rabotnik.sostexp, rabotnik.podpis FROM'+' rabotnik LEFT JOIN organiz ON rabotnik.id_org=organiz.id_org WHERE (prkrab<>"") AND (data_proh BETWEEN :date AND :date2) ORDER BY data_proh ASC';
 fDM.Query.SQL.Add(q);
 fDM.Query.Parameters.ParamByName('date').Value:=FormatDateTime('dd.mm.yyyy',DateTimePicker1.Date);
 fDM.Query.Parameters.ParamByName('date2').Value:=FormatDateTime('dd.mm.yyyy',DateTimePicker2.Date);
 fDM.Query.Open;

 h:=$00000000;
 i:=1;
 fDM.Query.First;
 while not(fDM.Query.Eof) do
 begin
  Shablon:=ExtractFileDir(ParamStr(0))+'\�������\prot2.doc';
  fEditor.WordApplication1.Connect;
  fEditor.WordApplication1.Visible:=false;
  fEditor.WordApplication1.Documents.Add(Shablon,EmptyParam, EmptyParam,EmptyParam);
  fEditor.WordDocument1.ConnectTo(fEditor.WordApplication1.ActiveDocument);
 
  fEditor.WordGotoBookmark('date');
  fEditor.WordInsertText(fDM.Query.FieldByName('data_proh').AsString);

  fEditor.WordGotoBookmark('fio');
  fEditor.WordInsertText(fDM.Query.FieldByName('fam').AsString+' '+fDM.Query.FieldByName('imya').AsString+' '+fDM.Query.FieldByName('fath').AsString);

  fEditor.WordGotoBookmark('org');
  fEditor.WordInsertText(fDM.Query.FieldByName('imya_poln').AsString);

  fEditor.WordGotoBookmark('prkrab');
  fEditor.WordInsertText(fDM.Query.FieldByName('prkrab').AsString);

  fEditor.WordGotoBookmark('dop_pri_usl');
  fEditor.WordInsertText(fDM.Query.FieldByName('dop_pri_usl').AsString);

  fDM.Query.Next;

  fEditor.WordGotoBookmark('date2');
  fEditor.WordInsertText(fDM.Query.FieldByName('data_proh').AsString);

  fEditor.WordGotoBookmark('fio2');
  fEditor.WordInsertText(fDM.Query.FieldByName('fam').AsString+' '+fDM.Query.FieldByName('imya').AsString+' '+fDM.Query.FieldByName('fath').AsString);

  fEditor.WordGotoBookmark('org2');
  fEditor.WordInsertText(fDM.Query.FieldByName('imya_poln').AsString);

  fEditor.WordGotoBookmark('prkrab2');
  fEditor.WordInsertText(fDM.Query.FieldByName('prkrab').AsString);

  fEditor.WordGotoBookmark('dop_pri_usl2');
  fEditor.WordInsertText(fDM.Query.FieldByName('dop_pri_usl').AsString);

  s:=ExtractFileDir(ParamStr(0))+'\������\����������\'+inttostr(i)+'.doc';
  fEditor.WordDocument1.SaveAs(s,h);

  vschange:= wdDoNotSaveChanges;
  fEditor.WordDocument1.Close(vschange);
  fEditor.WordApplication1.Disconnect;
  
  fDM.Query.Next;
  inc(i);
 end;
 ShowMessage('��������� '+inttostr(i-1)+' ������! ����� ��������� � �������� '+ExtractFileDir(ParamStr(0))+'\������\����������\');
end;

procedure TfExJournal.FormCreate(Sender: TObject);
begin
 DateTimePicker2.Date:=date;
 DateTimePicker1.Date:=date;
end;

end.
