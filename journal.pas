unit journal;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, DBGridEhGrouping, Menus, GridsEh, DBGridEh, StdCtrls,
  ADOInt,
  ComCtrls, DB, ToolCtrlsEh, DBGridEhToolCtrls, DynVarsEh, EhLibVCL,
  DBAxisGridsEh;

type
  TfJournal = class(TForm)
    Panel1: TPanel;
    Panel2: TPanel;
    DBGridEh1: TDBGridEh;
    MainMenu1: TMainMenu;
    N1: TMenuItem;
    Excel1: TMenuItem;
    DateTimePicker1: TDateTimePicker;
    DateTimePicker2: TDateTimePicker;
    От: TLabel;
    до: TLabel;
    CheckBox1: TCheckBox;
    Label1: TLabel;
    CheckBox2: TCheckBox;
    procedure FormCreate(Sender: TObject);
    procedure CheckBox1Click(Sender: TObject);
    procedure Excel1Click(Sender: TObject);
    procedure DBGridEh1DblClick(Sender: TObject);
    procedure CheckBox2Click(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  fJournal: TfJournal;

implementation

uses DM, editor, exJournal, user;

{$R *.dfm}

procedure TfJournal.CheckBox1Click(Sender: TObject);
begin
  fDM.TJournal.Active := false;
  if (CheckBox1.Checked) then
  begin
    if user.us <> 'Администратор' then
      fDM.TJournal.CommandText :=
        'SELECT user, id, npp, data_proh, fam, imya, fath, id_org, data_r, sex, profes, mkb_code, prkrab, dop_pri_usl, lpu,'
        + ' slexp, otkl, def, result, datezaklMSE, datenaprMSE, zaklMSE, dopinf, sostexp, podpis,vredn_fact FROM rabotnik WHERE (prkrab<>"") AND (data_proh BETWEEN :date1 AND :date2)and (user='
        + quotedstr(us) + ')' + 'ORDER BY id ASC'
    else
      fDM.TJournal.CommandText :=
        'SELECT user, id, npp, data_proh, fam, imya, fath, id_org, data_r, sex, profes, mkb_code, prkrab, dop_pri_usl, lpu,'
        + ' slexp, otkl, def, result, datezaklMSE, datenaprMSE, zaklMSE, dopinf, sostexp, podpis,vredn_fact FROM rabotnik WHERE (prkrab<>"") AND (data_proh BETWEEN :date1 AND :date2) ORDER BY id ASC';
    // fDM.TJournal.CommandText:='SELECT user, id, npp, data_proh, fam, imya, fath, id_org, data_r, sex, profes, mkb_code, prkrab, dop_pri_usl, lpu,'+' slexp, otkl, def, result, datezaklMSE, datenaprMSE, zaklMSE, dopinf, sostexp, podpis FROM rabotnik WHERE (prkrab<>"") AND (data_proh BETWEEN :date1 AND :date2) ORDER BY id ASC';
    fDM.TJournal.Parameters.ParamByName('date1').Value :=
      FormatDateTime('dd.mm.yyyy', DateTimePicker1.Date);
    fDM.TJournal.Parameters.ParamByName('date2').Value :=
      FormatDateTime('dd.mm.yyyy', DateTimePicker2.Date);
    fDM.TJournal.Active := true;
  end
  else if user.us <> 'Администратор' then
    fDM.TJournal.CommandText :=
      'SELECT user, id, npp, data_proh, fam, imya, fath, id_org, data_r, sex, profes, mkb_code, prkrab, dop_pri_usl, lpu, slexp, otkl, def, result, datezaklMSE, datenaprMSE, zaklMSE, dopinf, sostexp, podpis,vredn_fact FROM rabotnik WHERE (prkrab<>"") and (user='
      + quotedstr(us) + ')' + 'ORDER BY id ASC'
  else
    fDM.TJournal.CommandText :=
      'SELECT user, id, npp, data_proh, fam, imya, fath, id_org, data_r, sex, profes, mkb_code, prkrab, dop_pri_usl, lpu,'
      +' slexp, otkl, def, result, datezaklMSE, datenaprMSE, zaklMSE, dopinf, sostexp, podpis,vredn_fact FROM rabotnik WHERE (prkrab<>"") ORDER BY id ASC';
  fDM.TJournal.Active := true;
end;

procedure TfJournal.CheckBox2Click(Sender: TObject);
begin
  fDM.TJournal.Active := false;
  if CheckBox2.Checked then
  begin
    if user.us <> 'Администратор' then
    begin
      fDM.TJournal.CommandText :=
        'SELECT user, id, npp, data_proh, fam, imya, fath, id_org, data_r, sex, profes, mkb_code, prkrab, dop_pri_usl, lpu,'
        + ' slexp, otkl, def, result, datezaklMSE, datenaprMSE, zaklMSE, dopinf, sostexp, podpis,vredn_fact FROM rabotnik WHERE (prkrab<>"") ORDER BY id ASC';
      DBGridEh1.ReadOnly := true;
      fDM.TJournal.Active := true;
    end;
  end
  else if user.us <> 'Администратор' then
  begin
    fDM.TJournal.CommandText :=
      'SELECT user, id, npp, data_proh, fam, imya, fath, id_org, data_r, sex, profes, mkb_code, prkrab, dop_pri_usl, lpu, slexp, otkl, def, result, datezaklMSE, datenaprMSE, zaklMSE, dopinf, sostexp, podpis,vredn_fact FROM rabotnik WHERE (prkrab<>"") and (user='
      + quotedstr(us) + ')' + 'ORDER BY id ASC';
    DBGridEh1.ReadOnly := false;
  end
  else
    fDM.TJournal.CommandText :=
      'SELECT user, id, npp, data_proh, fam, imya, fath, id_org, data_r, sex, profes, mkb_code, prkrab, dop_pri_usl, lpu,'
      +' slexp, otkl, def, result, datezaklMSE, datenaprMSE, zaklMSE, dopinf, sostexp, podpis,vredn_fact FROM rabotnik WHERE (prkrab<>"") ORDER BY id ASC';
  fDM.TJournal.Active := true;
end;

procedure TfJournal.DBGridEh1DblClick(Sender: TObject);
begin
  // fDM.TRab.Locate('id',fDM.TJournal.FieldByName('id').AsString,[]);
  // fEditor.ShowModal;
end;

procedure TfJournal.Excel1Click(Sender: TObject);
begin
  fExJournal.ShowModal;
end;

procedure TfJournal.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  CheckBox2.Checked := false;
  if fDM.TRab.RecordCount > 0 then
  begin
    fDM.TRab.UpdateCursorPos(); // Обновить одну строку!!!
    fDM.TRab.Recordset.Resync(adAffectCurrent, adResyncAllValues);
    fDM.TRab.Resync([rmExact]);
  end;
end;

procedure TfJournal.FormCreate(Sender: TObject);
begin
  DateTimePicker1.Date := Date;
  DateTimePicker2.Date := Date;
end;

end.
