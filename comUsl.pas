unit comUsl;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ComCtrls, DBGridEhGrouping, GridsEh, DBGridEh, StdCtrls, OleServer,
  ComObj, DBGridEhImpExp;

type
  TfCommonUsl = class(TForm)
    DateTimePicker1: TDateTimePicker;
    DateTimePicker2: TDateTimePicker;
    Button1: TButton;
    DBGridEh1: TDBGridEh;
    Button2: TButton;
    Button3: TButton;
    Button4: TButton;
    procedure Button1Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure Button4Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  fCommonUsl: TfCommonUsl;

implementation

uses DM;

{$R *.dfm}

procedure TfCommonUsl.Button1Click(Sender: TObject);
begin
  fDM.QComUsl.SQL.Clear;
  fDM.QComUsl.SQL.Add
    ('SELECT rabotnik.disp_gr AS [� ��], Count(rabotnik.id) AS [���-��] FROM rabotnik WHERE (((rabotnik.id_org) In (1,2,3,4,5,6,7,8,9,11,12,13,14,15,16,17,19,20,21,22,23,24,25,27,28,31,32,34,35,36,38,40,43,44,56,57,63,71,76))'
    + ' AND (idusl<>1000) AND ((rabotnik.data_proh) Between :date And :date2)) AND (rabotnik.cat_osv not Like "�%") GROUP BY rabotnik.disp_gr');
  fDM.QComUsl.Parameters.ParamByName('date').Value :=
    FormatDateTime('dd.mm.yyyy', DateTimePicker1.Date);
  fDM.QComUsl.Parameters.ParamByName('date2').Value :=
    FormatDateTime('dd.mm.yyyy', DateTimePicker2.Date);
  fDM.QComUsl.Active := true;
  fDM.QComUsl.First;
end;

procedure TfCommonUsl.Button2Click(Sender: TObject);
begin
  fDM.QComUsl.SQL.Clear;
  fDM.QComUsl.SQL.Add
    ('SELECT rabotnik.disp_gr AS [� ��], Count(rabotnik.id) AS [���-��] FROM rabotnik WHERE (((rabotnik.id_org) In (1,2,3,4,5,6,7,8,9,11,12,13,14,15,16,17,19,20,21,22,23,24,25,27,28,31,32,34,35,36,38,40,43,44,56,57,63,71,76))'
    + ' AND (idusl=1000) AND ((rabotnik.data_proh) Between :date And :date2)) AND (rabotnik.cat_osv not Like "�%") GROUP BY rabotnik.disp_gr');
  fDM.QComUsl.Parameters.ParamByName('date').Value :=
    FormatDateTime('dd.mm.yyyy', DateTimePicker1.Date);
  fDM.QComUsl.Parameters.ParamByName('date2').Value :=
    FormatDateTime('dd.mm.yyyy', DateTimePicker2.Date);
  fDM.QComUsl.Active := true;
  fDM.QComUsl.First;
end;

procedure TfCommonUsl.Button3Click(Sender: TObject);
begin
  fDM.QComUsl.SQL.Clear;
  fDM.QComUsl.SQL.Add
    ('SELECT rabotnik.disp_gr AS [� ��], Count(rabotnik.id) AS [���-��] FROM rabotnik WHERE (((rabotnik.id_org) In (1,2,3,4,5,6,7,8,9,11,12,13,14,15,16,17,19,20,21,22,23,24,25,27,28,31,32,34,35,36,38,40,43,44,56,57,63,71,76))'
    + ' AND (firstb=+' + quotedstr('��') +
    ') AND ((rabotnik.data_proh) Between :date And :date2)) AND (rabotnik.cat_osv not Like "�%") GROUP BY rabotnik.disp_gr');
  fDM.QComUsl.Parameters.ParamByName('date').Value :=
    FormatDateTime('dd.mm.yyyy', DateTimePicker1.Date);
  fDM.QComUsl.Parameters.ParamByName('date2').Value :=
    FormatDateTime('dd.mm.yyyy', DateTimePicker2.Date);
  fDM.QComUsl.Active := true;
  fDM.QComUsl.First;
end;

procedure TfCommonUsl.Button4Click(Sender: TObject);
begin
  fDM.QComUsl.SQL.Clear;
  fDM.QComUsl.SQL.Add
    ('SELECT Count(rabotnik.id) AS [���-��] FROM rabotnik WHERE (((rabotnik.id_org) In (37,46,60,39,41,10)) AND ((rabotnik.data_proh) Between :date And :date2)) AND (rabotnik.cat_osv not Like "�%") ');
  fDM.QComUsl.Parameters.ParamByName('date').Value :=
    FormatDateTime('dd.mm.yyyy', DateTimePicker1.Date);
  fDM.QComUsl.Parameters.ParamByName('date2').Value :=
    FormatDateTime('dd.mm.yyyy', DateTimePicker2.Date);
  fDM.QComUsl.Active := true;
  fDM.QComUsl.First;
end;

procedure TfCommonUsl.FormCreate(Sender: TObject);
begin
  DateTimePicker1.Date := Date;
  DateTimePicker2.Date := Date;
end;

end.
