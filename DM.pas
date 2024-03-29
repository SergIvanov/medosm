unit DM;

interface

uses
  SysUtils, Classes, DB, ADODB, MemTableDataEh, DataDriverEh, ADODataDriverEh,
  MemTableEh, DBXFirebird, SqlExpr, FireDAC.Stan.Intf, FireDAC.Stan.Option,
  FireDAC.Stan.Error, FireDAC.UI.Intf, FireDAC.Phys.Intf, FireDAC.Stan.Def,
  FireDAC.Stan.Pool, FireDAC.Stan.Async, FireDAC.Phys, FireDAC.Phys.FB,
  FireDAC.Phys.FBDef, FireDAC.VCLUI.Wait, FireDAC.Stan.Param, FireDAC.DatS,
  FireDAC.DApt.Intf, FireDAC.DApt, FireDAC.Comp.DataSet, FireDAC.Comp.Client,
  FireDAC.Phys.IBBase;

type
  TfDM = class(TDataModule)
    ADOConnection1: TADOConnection;
    DSRab: TDataSource;
    TOrg: TADOTable;
    DSOrg: TDataSource;
    TOrgimya_org: TWideStringField;
    TOrgform_sobstv: TWideStringField;
    DSorgnew: TDataSource;
    TOrgNew: TADOTable;
    TOrgNewid_org: TAutoIncField;
    TOrgNewimya_org: TWideStringField;
    TOrgNewform_sobstv: TWideStringField;
    DSQOrg: TDataSource;
    QOrg: TADOQuery;
    QNewSotr: TADOQuery;
    DSMKB: TDataSource;
    TMKB: TADOTable;
    QVSEGO: TADOQuery;
    Qmkbfirst: TADOQuery;
    mkbfirst: TADOQuery;
    TRab: TADODataSet;
   // ADODataDriverEh1: TADODataDriverEh;
    TRabid: TAutoIncField;
    TRabid_org: TIntegerField;
    TRabfam: TWideStringField;
    TRabimya: TWideStringField;
    TRabfath: TWideStringField;
    TRabsex: TWideStringField;
    TRabser_pasp: TWideStringField;
    TRabdata_vid_pasp: TWideStringField;
    TRabkem_vid_pasp: TWideStringField;
    TRabadres: TWideStringField;
    TRabphone: TWideStringField;
    TRabpolis: TWideStringField;
    TRabvid_ek_deat: TWideStringField;
    TRabvredn_fact: TWideMemoField;
    TRabstaz_rab_fact: TWideStringField;
    TRabmed_organ: TWideStringField;
    TRabplan_osm: TWideStringField;
    TRabsrok_proh: TWideStringField;
    TRabusl_tr: TWideStringField;
    TRabnom_pasp_zd: TWideStringField;
    TRabmkb: TWideStringField;
    TRabmkb_code: TWideStringField;
    TRabcat_osv: TWideStringField;
    TRabch_osm: TWordField;
    TRabgoden: TWideStringField;
    TRabdop_pri_usl: TWideStringField;
    TRabnom_st_protiv: TWideStringField;
    TRabnom_mps: TWideStringField;
    TRaborg: TStringField;
    TRabdisp_gr: TWideStringField;
    TRabfirstb: TWideStringField;
    TRabneedobslprof: TWideStringField;
    TRabneedambl: TWideStringField;
    TRabneedstcl: TWideStringField;
    TRabskl: TWideStringField;
    TRabneedlpit: TWideStringField;
    TRabneedmse: TWideStringField;
    TRabneedd: TWideStringField;
    TRabnom_pasp: TWideStringField;
    TOrgimya_poln: TWideMemoField;
    TOrgNewimya_poln: TWideMemoField;
    TRabimya_strukt_podr: TWideStringField;
    TRabnpp: TIntegerField;
    TRabuser: TWideStringField;
    TRabrekom: TWideMemoField;
    TSUslM: TADODataSet;
    DSSUslM: TDataSource;
    QSumUsl: TADOQuery;
    DSSumUsl: TDataSource;
    QSumUslfam: TWideStringField;
    QSumUslimya: TWideStringField;
    QSumUslfath: TWideStringField;
    QSumUslsex: TWideStringField;
    QSumUslvredn_fact: TWideMemoField;
    QSumUslprofes: TWideStringField;
    TRabidusl: TIntegerField;
    TRabUslM: TStringField;
    QSumUslidusl: TIntegerField;
    DSComUsl: TDataSource;
    QComUsl: TADOQuery;
    TRabprofes: TWideStringField;
    TFYS: TADODataSet;
    DSFYS: TDataSource;
    TFYSid: TAutoIncField;
    TFYSName_usl: TWideStringField;
    TFYSZena: TFloatField;
    TFYSidotdel: TIntegerField;
    TOtdel: TADODataSet;
    DSOtdel: TDataSource;
    TUsl: TADODataSet;
    DSUsl: TDataSource;
    QStacRep: TADOQuery;
    DSStacRep: TDataSource;
    TRabdata_proh: TDateTimeField;
    TSUslMid: TIntegerField;
    TSUslMnaim: TWideStringField;
    TSUslMcena: TIntegerField;
    TSUsl: TADODataSet;
    DSSUsl: TDataSource;
    TPeriodVigr: TADODataSet;
    DSPeriodVigr: TDataSource;
    TOrgid_org: TAutoIncField;
    TIssl: TADODataSet;
    DSIssl: TDataSource;
    TSIssl: TADODataSet;
    TIsslid: TAutoIncField;
    TIsslidrab: TIntegerField;
    TIsslidissl: TIntegerField;
    TIssldate1: TDateTimeField;
    TIsslznach1: TWideStringField;
    TIssldate2: TDateTimeField;
    TIsslznach2: TWideStringField;
    TIssldate3: TDateTimeField;
    TIsslznach3: TWideStringField;
    TIssllookissl: TStringField;
    TRabprkrab: TWideStringField;
    TJournal: TADODataSet;
    DSJournal: TDataSource;
    TJournaldata_proh: TDateTimeField;
    TJournalfam: TWideStringField;
    TJournalimya: TWideStringField;
    TJournalfath: TWideStringField;
    TJournalid_org: TIntegerField;
    TJournalsex: TWideStringField;
    TJournalprofes: TWideStringField;
    TJournalmkb_code: TWideStringField;
    TJournalprkrab: TWideStringField;
    TJournallookOrg: TStringField;
    TRablpu: TWideStringField;
    TRabslexp: TWideStringField;
    TRabotkl: TWideStringField;
    TRabdef: TWideStringField;
    TRabresult: TWideStringField;
    TRabdatenaprMSE: TDateTimeField;
    TRabzaklMSE: TWideStringField;
    TRabdatezaklMSE: TDateTimeField;
    TRabdopinf: TWideStringField;
    TRabsostexp: TWideStringField;
    TRabpodpis: TWideStringField;
    TJournallpu: TWideStringField;
    TJournalslexp: TWideStringField;
    TJournalotkl: TWideStringField;
    TJournaldef: TWideStringField;
    TJournalresult: TWideStringField;
    TJournaldatezaklMSE: TDateTimeField;
    TJournaldatenaprMSE: TDateTimeField;
    TJournalzaklMSE: TWideStringField;
    TJournaldopinf: TWideStringField;
    TJournalsostexp: TWideStringField;
    TJournalpodpis: TWideStringField;
    Query: TADOQuery;
    TJournaldop_pri_usl: TWideStringField;
    TJournalnpp: TIntegerField;
    TJournalid: TAutoIncField;
    TJournaluser: TWideStringField;
    QSumUslcat_osv: TWideStringField;
    TOrgdif_price: TWordField;
    TOrgNewdif_price: TWordField;
    QueryRab: TADOQuery;
    TFactor: TADODataSet;
    DSFactor: TDataSource;
    TFactorid: TAutoIncField;
    TFactorid_rab: TIntegerField;
    TFactorcode: TWideStringField;
    TFactornaim: TWideStringField;
    TFactorvrach: TWideStringField;
    TSFactor: TADODataSet;
    DSSFactor: TDataSource;
    TSFactorid: TAutoIncField;
    TSFactorcode: TWideStringField;
    TSFactornaim: TWideStringField;
    TSFactorvrach: TWideStringField;
    TMKBCODE: TWideStringField;
    TMKBNAIM: TWideStringField;
    TFactId: TADODataSet;
    DSFactId: TDataSource;
    TRabiduslold: TIntegerField;
    QSumUsliduslold: TIntegerField;
    TSFactorissl: TWideMemoField;
    TRabdata_r: TDateTimeField;
    TFactorissl: TWideMemoField;
    TJournaldata_r: TDateTimeField;
    QSumUslextest: TWideStringField;
    QSumUslextestold: TWideStringField;
    TRabextest: TWideStringField;
    TRabextestold: TWideStringField;
    TFYSkod_vr: TIntegerField;
    TFYSkod_ms: TIntegerField;
    TUslid: TAutoIncField;
    TUslidrab: TIntegerField;
    TUslnaimusl: TWideStringField;
    TUslkol: TWordField;
    TUslcena: TIntegerField;
    TUsldata: TDateTimeField;
    TUslkod_usl: TWideStringField;
    TFYSkod_usl: TWideStringField;
    TUslkod_vr: TIntegerField;
    TUslkod_ms: TIntegerField;
    TDMS: TADODataSet;
    DSDMS: TDataSource;
    TDMSid: TAutoIncField;
    TDMSfam: TWideStringField;
    TDMSim: TWideStringField;
    TDMSot: TWideStringField;
    TDMSdr: TDateTimeField;
    TDMSser: TWideStringField;
    TDMSnom: TWideStringField;
    Query2: TADOQuery;
    TRabdate_of: TDateTimeField;
    TUslmkb: TWideStringField;
    QNeedSKL: TADOQuery;
    WideStringField1: TWideStringField;
    WideStringField2: TWideStringField;
    WideStringField3: TWideStringField;
    WideStringField5: TWideStringField;
    QNeedSKLdata_r: TDateTimeField;
    TTempList: TADODataSet;
    DSTempList: TDataSource;
    TTempListfam: TWideStringField;
    TTempListim: TWideStringField;
    TTempListot: TWideStringField;
    TTempListdr: TDateTimeField;
    ProfPatolog: TADOTable;
    DSProfPatolog: TDataSource;
    TRabid_prof: TIntegerField;
    TRabprof: TStringField;
    qryUser: TADOQuery;
    blnfldTRabpsih: TBooleanField;
    wdmfldTRabvredn_fact2: TWideMemoField;
    smlntfldTRabProfSan: TSmallintField;
    wdstrngfldTRabPredDZ: TWideStringField;
    wdstrngfldTRabVUPZ: TWideStringField;
    tblProfSan: TADOTable;
    dsProfSan: TDataSource;
    strngfldTRabLookUpProfSan: TStringField;
    wdstrngfldTRabmr: TWideStringField;
    wdstrngfldTRabmrr: TWideStringField;
    wdstrngfldTRabmrg: TWideStringField;
    wdstrngfldTRabmrn: TWideStringField;
    wdstrngfldTRabmru: TWideStringField;
    wdstrngfldTRabmrd: TWideStringField;
    wdstrngfldTRabmrk: TWideStringField;
    wdstrngfldTRabmrkv: TWideStringField;
    wdmfldTJournalvredn_fact: TWideMemoField;
    tblTDogovor: TADOTable;
    atncfldTDogovorid: TAutoIncField;
    wdstrngfldTDogovorNaim: TWideStringField;
    dsDogovor: TDataSource;
    tblZak282: TADOTable;
    wrdfldTRabZak282: TWordField;
    qryJurnal282: TADOQuery;
    dsJurnal282: TDataSource;
    wdstrngfldTRabprkrab282p2: TWideStringField;
    dsZak282: TDataSource;
    blnfldTOrgzd: TBooleanField;
    procedure TMKBFilterRecord(DataSet: TDataSet; var Accept: Boolean);
    procedure TRabAfterInsert(DataSet: TDataSet);
    procedure TRabprkrabChange(Sender: TField);
    procedure TRabAfterEdit(DataSet: TDataSet);
    procedure TRabdata_rChange(Sender: TField);
    procedure TSFactorFilterRecord(DataSet: TDataSet; var Accept: Boolean);
    procedure DataModuleCreate(Sender: TObject);

  private
    { Private declarations }
  public
    { Public declarations }
    Procedure MKBLocate(s: string);
  end;

var
  fDM: TfDM;
  ed, prkrab: String; // ����� �� Edit1

implementation

uses user, main;

{$R *.dfm}

procedure TfDM.DataModuleCreate(Sender: TObject);
begin

//fdphysfbdrvrlnk1.VendorLib:= ExtractFileDir(ParamStr(0))+'\32\fbclient.dll';

TMKB.Active := True;
TOrgNew.Active := True;

end;

procedure TfDM.MKBLocate(s: string);
begin
  TMKB.Locate('CODE', s, [loPartialKey]);
end;

procedure TfDM.TMKBFilterRecord(DataSet: TDataSet; var Accept: Boolean);
var
  s: String; // ��� �������� ����
begin
  s := Copy(DataSet['NAIM'], 1, Length(ed));
  Accept := s = ed;
end;

procedure TfDM.TRabAfterEdit(DataSet: TDataSet);
begin
  prkrab := fDM.TRab.FieldByName('prkrab').AsString;

end;

procedure TfDM.TRabAfterInsert(DataSet: TDataSet);
var
  today: TDateTime;
  todaystr: string;
  I: Integer;
begin
  if user.us <> '�������������' then
  begin
    fDM.TRab.FieldByName('user').AsString := user.us;
    today := Date;
    todaystr := DateToStr(today);
    for I := 1 to Length(todaystr) do
      if todaystr[I] = '/' then
        Delete(todaystr, I, 1);
    fDM.TRab.FieldByName('data_proh').AsString := todaystr;
    fDM.TRab.FieldByName('sostexp').AsString :=
      '����������� �.�.�������� �.�. ������� �.�. �������� �.�. �������� �.�.';
  end;
   FDM.TRab.FieldByName('psih').Value:=False;
  { ���������� ������ � ���� mesyac }
end;



procedure TfDM.TRabdata_rChange(Sender: TField);
begin
  fDM.QueryRab.SQL.Clear;
  if (fDM.TRab.FieldByName('fam').AsString <> '') AND
    (fDM.TRab.FieldByName('imya').AsString <> '') then
  begin
    fDM.QueryRab.SQL.Text :=
      'Select * From SprRabotnik Where (fam=:fam) and (imya=:imya) and (fath=:fath) and (data_r=:dr)';
    fDM.QueryRab.Parameters.ParamByName('fam').Value :=
      fDM.TRab.FieldByName('fam').AsString;
    fDM.QueryRab.Parameters.ParamByName('imya').Value :=
      fDM.TRab.FieldByName('imya').AsString;
    if fDM.TRab.FieldByName('fath').AsString <> '' then
      fDM.QueryRab.Parameters.ParamByName('fath').Value :=
        fDM.TRab.FieldByName('fath').AsString
    else
      fDM.QueryRab.Parameters.ParamByName('fath').Value := '';
    fDM.QueryRab.Parameters.ParamByName('dr').Value :=
      fDM.TRab.FieldByName('data_r').AsString;

    fDM.QueryRab.Open;;
    fDM.TRab.Edit;
//    fDM.TRab.FieldByName('sex').AsString :=
//      fDM.QueryRab.FieldByName('sex').AsString;
//    fDM.TRab.FieldByName('vredn_fact').AsString :=
//      fDM.QueryRab.FieldByName('vredn_fact').AsString;
    fDM.TRab.FieldByName('id_org').AsString :=
      fDM.QueryRab.FieldByName('id_org').AsString;
    fDM.TRab.FieldByName('imya_strukt_podr').AsString :=
      fDM.QueryRab.FieldByName('imya_strukt_podr').AsString;
    fDM.TRab.FieldByName('profes').AsString :=
      fDM.QueryRab.FieldByName('profes').AsString;
    fDM.TRab.FieldByName('staz_rab_fact').AsString :=
      fDM.QueryRab.FieldByName('staz_rab_fact').AsString;
    fDM.TRab.FieldByName('cat_osv').AsString :=
      fDM.QueryRab.FieldByName('cat_osv').AsString;
//    fDM.TRab.FieldByName('idusl').AsString :=
//      fDM.QueryRab.FieldByName('idusl').AsString;
    fDM.TRab.Post;
    fDM.QueryRab.Close;
  end;

end;

procedure TfDM.TRabprkrabChange(Sender: TField);
var
  npp: Integer;
begin
  if prkrab = '' then
  begin
    fDM.Query.SQL.Clear;
    fDM.Query.SQL.Add('select max(npp) as maximum from rabotnik');
    fDM.Query.Open;
    npp := fDM.Query.FieldByName('maximum').AsInteger;
    fDM.TRab.FieldByName('npp').AsInteger := npp + 1;
  end;
end;

procedure TfDM.TSFactorFilterRecord(DataSet: TDataSet; var Accept: Boolean);
var
  s: String; // ��� �������� ����
begin
  s := Copy(DataSet['code'], 1, Length(ed));
  Accept := s = ed;
end;

end.

