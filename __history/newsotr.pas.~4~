unit newsotr;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Grids, DBGrids, StdCtrls, ComObj;

type
  TfNewSotr = class(TForm)
    Edit1: TEdit;
    OD1: TOpenDialog;
    Button1: TButton;
    Button2: TButton;
    Button3: TButton;
    Edit2: TEdit;
    Label1: TLabel;
    procedure Button1Click(Sender: TObject);
    procedure ParseFIO(sfio:string; var sf,si,so:string);
    procedure Button2Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  fNewSotr: TfNewSotr;
  ExelTab:Variant;

implementation

uses DM;

{$R *.dfm}

procedure TfNewSotr.Button1Click(Sender: TObject);
begin
 if OD1.Execute then Edit1.Text:= OD1.FileName;
 end;

procedure TfNewSotr.ParseFIO(sfio:string; var sf, si, so: string);
 var i,k,lchar:integer;
begin
 k:=1;
 lchar:=1;
 for i:=1 to length(sfio) do
 begin
  if (sfio[i]=' ') and (k=1) then
  begin
   sf:=Copy(sfio, lchar, i-1);
   inc(k);
   lchar:=i+1;
   continue;
  end;

  if (sfio[i]=' ') and (k=2) then
  begin
   si:=Copy(sfio, lchar, i-lchar);
   inc(k);
   lchar:=i+1;
   continue;
  end;

  if i=length(sfio) then
  begin
   so:=Copy(sfio, lchar, i-lchar+1);
   inc(k);
   //lchar:=i+1;
   continue;
  end;
 end;

end;

procedure TfNewSotr.Button2Click(Sender: TObject);
var i,k:integer;
    ArNewSotr: Array [1..1000, 1..11] of string;
    sf,si,so,Q, id_org:string;
begin
  ExelTab:= CreateOleObject('Excel.Application');
  ExelTab.Workbooks.Open[Edit1.Text];

  //пока не последняя строка
  i:=2;
  while string(ExelTab.Cells.Item[i,1].Value)<>'' do
  begin
    ParseFIO(string(ExelTab.Cells.Item[i,3].Value),sf,si,so);
    for k:=1 to 11 do
     begin
      if k=1 then ArNewSotr[i,k]:=string(ExelTab.Cells.Item[i,2].Value); //номер цеха
      if k=2 then ArNewSotr[i,k]:=sf;
      if k=3 then ArNewSotr[i,k]:=si;
      if k=4 then ArNewSotr[i,k]:=so;
      if k=5 then ArNewSotr[i,k]:=string(ExelTab.Cells.Item[i,4].Value); //дата р
      if k=6 then ArNewSotr[i,k]:=string(ExelTab.Cells.Item[i,5].Value);  //пол
      if k=7 then ArNewSotr[i,k]:=string(ExelTab.Cells.Item[i,6].Value); //профессия
      if k=8 then ArNewSotr[i,k]:=string(ExelTab.Cells.Item[i,7].Value); //стаж
      if k=9 then ArNewSotr[i,k]:=string(ExelTab.Cells.Item[i,8].Value);  //опасные факторы
      if k=10 then ArNewSotr[i,k]:=string(ExelTab.Cells.Item[i,9].Value); //дата посл осмотра
      if k=11 then ArNewSotr[i,k]:=string(ExelTab.Cells.Item[i,10].Value); //срок осмотра
     end;
     inc(i);
  end;

i:=2;
id_org:=Edit2.Text;
 while string(ExelTab.Cells.Item[i,1].Value)<>'0' do
 begin
  Q:='INSERT INTO rabotnik (fam, imya, fath, sex, data_r, profes, vredn_fact,staz_rab_fact, data_proh, srok_proh, imya_strukt_podr,id_org) VALUES ('+quotedstr(ArNewSotr[i,2])+','+quotedstr(ArNewSotr[i,3])+','+quotedstr(ArNewSotr[i,4])+','+quotedstr(ArNewSotr[i,6])+','+quotedstr(ArNewSotr[i,5])+','+quotedstr(ArNewSotr[i,7])+','
  +quotedstr(ArNewSotr[i,9])+','+quotedstr(ArNewSotr[i,8])+','+quotedstr(ArNewSotr[i,10])+','+quotedstr(ArNewSotr[i,11])+','+quotedstr(ArNewSotr[i,1])+','+quotedstr(id_org)+')';

  fDM.QNewSotr.SQL.Clear;
  fDM.QNewSotr.SQL.Add(Q);

  try
  fDM.ADOConnection1.BeginTrans;
  fDM.QNewSotr.ExecSQL;
  fDM.ADOConnection1.CommitTrans;
  except
  fDM.ADOConnection1.RollbackTrans;
  end;

  inc(i);
 end;
 ExelTab.Quit;
  fDM.TRab.Active:=false;
  fDM.TRab.Active:=true;
end;

end.
