object fDM: TfDM
  OldCreateOrder = False
  OnCreate = DataModuleCreate
  Height = 745
  Width = 1013
  object ADOConnection1: TADOConnection
    Connected = True
    ConnectionString = 
      'Provider=Microsoft.Jet.OLEDB.4.0;User ID=Admin;Data Source=meddb' +
      '.mdb;Mode=Share Deny None;Persist Security Info=False;Jet OLEDB:' +
      'System database="";Jet OLEDB:Registry Path="";Jet OLEDB:Database' +
      ' Password="";Jet OLEDB:Engine Type=5;Jet OLEDB:Database Locking ' +
      'Mode=1;Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Global Bulk' +
      ' Transactions=1;Jet OLEDB:New Database Password="";Jet OLEDB:Cre' +
      'ate System Database=False;Jet OLEDB:Encrypt Database=False;Jet O' +
      'LEDB:Don'#39't Copy Locale on Compact=False;Jet OLEDB:Compact Withou' +
      't Replica Repair=False;Jet OLEDB:SFP=False'
    CursorLocation = clUseServer
    LoginPrompt = False
    Mode = cmShareDenyNone
    Provider = 'Microsoft.Jet.OLEDB.4.0'
    Left = 88
    Top = 16
  end
  object DSRab: TDataSource
    DataSet = TRab
    Left = 104
    Top = 80
  end
  object TOrg: TADOTable
    Active = True
    Connection = ADOConnection1
    CursorType = ctStatic
    Filtered = True
    IndexFieldNames = 'imya_org'
    TableName = 'organiz'
    Left = 24
    Top = 136
    object TOrgid_org: TAutoIncField
      FieldName = 'id_org'
      ReadOnly = True
    end
    object TOrgimya_org: TWideStringField
      DisplayWidth = 59
      FieldName = 'imya_org'
      Size = 150
    end
    object TOrgform_sobstv: TWideStringField
      DisplayWidth = 28
      FieldName = 'form_sobstv'
      Size = 50
    end
    object TOrgimya_poln: TWideMemoField
      FieldName = 'imya_poln'
      BlobType = ftWideMemo
    end
    object TOrgdif_price: TWordField
      FieldName = 'dif_price'
    end
    object blnfldTOrgzd: TBooleanField
      FieldName = 'zd'
    end
  end
  object DSOrg: TDataSource
    DataSet = TOrg
    Left = 104
    Top = 136
  end
  object DSorgnew: TDataSource
    DataSet = TOrgNew
    Left = 104
    Top = 192
  end
  object TOrgNew: TADOTable
    Connection = ADOConnection1
    CursorType = ctStatic
    IndexFieldNames = 'imya_org'
    TableName = 'organiz'
    Left = 24
    Top = 192
    object TOrgNewid_org: TAutoIncField
      DisplayWidth = 12
      FieldName = 'id_org'
      ReadOnly = True
    end
    object TOrgNewimya_org: TWideStringField
      DisplayWidth = 52
      FieldName = 'imya_org'
      Size = 150
    end
    object TOrgNewform_sobstv: TWideStringField
      DisplayWidth = 27
      FieldName = 'form_sobstv'
      Size = 50
    end
    object TOrgNewimya_poln: TWideMemoField
      FieldName = 'imya_poln'
      BlobType = ftWideMemo
    end
    object TOrgNewdif_price: TWordField
      FieldName = 'dif_price'
    end
  end
  object DSQOrg: TDataSource
    DataSet = QOrg
    Left = 104
    Top = 256
  end
  object QOrg: TADOQuery
    Connection = ADOConnection1
    Parameters = <>
    Left = 24
    Top = 256
  end
  object QNewSotr: TADOQuery
    Connection = ADOConnection1
    DataSource = DSRab
    Parameters = <>
    Left = 24
    Top = 312
  end
  object DSMKB: TDataSource
    DataSet = TMKB
    Left = 96
    Top = 360
  end
  object TMKB: TADOTable
    Active = True
    Connection = ADOConnection1
    CursorType = ctStatic
    OnFilterRecord = TMKBFilterRecord
    TableName = 'mkb10'
    Left = 24
    Top = 360
    object TMKBCODE: TWideStringField
      FieldName = 'CODE'
      Size = 7
    end
    object TMKBNAIM: TWideStringField
      FieldName = 'NAIM'
      Size = 254
    end
  end
  object QVSEGO: TADOQuery
    Connection = ADOConnection1
    Parameters = <
      item
        Name = 'date1'
        DataType = ftDateTime
        Value = Null
      end
      item
        Name = 'date2'
        DataType = ftDateTime
        Value = Null
      end>
    Left = 24
    Top = 416
  end
  object Qmkbfirst: TADOQuery
    Connection = ADOConnection1
    Parameters = <>
    Left = 104
    Top = 416
  end
  object mkbfirst: TADOQuery
    Connection = ADOConnection1
    Parameters = <>
    Left = 176
    Top = 416
  end
  object TRab: TADODataSet
    Active = True
    Connection = ADOConnection1
    CursorType = ctStatic
    Filtered = True
    AfterInsert = TRabAfterInsert
    AfterEdit = TRabAfterEdit
    CommandText = 'select * from rabotnik'
    IndexFieldNames = 'id'
    Parameters = <
      item
        Name = 'date1'
        DataType = ftDateTime
        Value = Null
      end
      item
        Name = 'date2'
        DataType = ftDateTime
        Value = Null
      end>
    Left = 32
    Top = 88
    object TRabid: TAutoIncField
      FieldName = 'id'
      ReadOnly = True
    end
    object TRabfam: TWideStringField
      FieldName = 'fam'
      Size = 30
    end
    object TRabimya: TWideStringField
      FieldName = 'imya'
      Size = 15
    end
    object TRabfath: TWideStringField
      FieldName = 'fath'
    end
    object TRabdata_r: TDateTimeField
      FieldName = 'data_r'
    end
    object TRabsex: TWideStringField
      FieldName = 'sex'
      Size = 3
    end
    object TRabser_pasp: TWideStringField
      FieldName = 'ser_pasp'
      Size = 5
    end
    object TRabdata_vid_pasp: TWideStringField
      FieldName = 'data_vid_pasp'
      Size = 50
    end
    object TRabkem_vid_pasp: TWideStringField
      FieldName = 'kem_vid_pasp'
      Size = 100
    end
    object TRabadres: TWideStringField
      FieldName = 'adres'
      Size = 255
    end
    object TRabphone: TWideStringField
      FieldName = 'phone'
      Size = 25
    end
    object TRabpolis: TWideStringField
      FieldName = 'polis'
      Size = 35
    end
    object TRabprofes: TWideStringField
      FieldName = 'profes'
      Size = 255
    end
    object TRabvid_ek_deat: TWideStringField
      FieldName = 'vid_ek_deat'
      Size = 35
    end
    object TRabvredn_fact: TWideMemoField
      FieldName = 'vredn_fact'
      BlobType = ftWideMemo
    end
    object TRabimya_strukt_podr: TWideStringField
      FieldName = 'imya_strukt_podr'
      Size = 254
    end
    object TRabstaz_rab_fact: TWideStringField
      FieldName = 'staz_rab_fact'
      Size = 15
    end
    object TRabmed_organ: TWideStringField
      FieldName = 'med_organ'
      Size = 255
    end
    object TRabplan_osm: TWideStringField
      FieldName = 'plan_osm'
      Size = 255
    end
    object TRabfirstb: TWideStringField
      FieldName = 'firstb'
      Size = 3
    end
    object TRabsrok_proh: TWideStringField
      FieldName = 'srok_proh'
      Size = 50
    end
    object TRabusl_tr: TWideStringField
      FieldName = 'usl_tr'
      Size = 100
    end
    object TRabdate_of: TDateTimeField
      FieldName = 'date_of'
    end
    object TRabnom_pasp_zd: TWideStringField
      FieldName = 'nom_pasp_zd'
    end
    object TRabmkb: TWideStringField
      FieldName = 'mkb'
      Size = 255
    end
    object TRabmkb_code: TWideStringField
      FieldName = 'mkb_code'
      Size = 100
    end
    object TRabcat_osv: TWideStringField
      FieldName = 'cat_osv'
      Size = 150
    end
    object TRabch_osm: TWordField
      FieldName = 'ch_osm'
    end
    object TRabprkrab: TWideStringField
      FieldName = 'prkrab'
      OnChange = TRabprkrabChange
      Size = 255
    end
    object TRabgoden: TWideStringField
      FieldName = 'goden'
      Size = 30
    end
    object TRabdop_pri_usl: TWideStringField
      FieldName = 'dop_pri_usl'
      Size = 255
    end
    object TRabnom_st_protiv: TWideStringField
      FieldName = 'nom_st_protiv'
      Size = 50
    end
    object TRabnom_mps: TWideStringField
      FieldName = 'nom_mps'
      Size = 50
    end
    object TRabdisp_gr: TWideStringField
      FieldName = 'disp_gr'
      Size = 100
    end
    object TRabneedobslprof: TWideStringField
      FieldName = 'needobslprof'
      Size = 3
    end
    object TRabneedambl: TWideStringField
      FieldName = 'needambl'
      Size = 3
    end
    object TRabneedstcl: TWideStringField
      FieldName = 'needstcl'
      Size = 3
    end
    object TRabskl: TWideStringField
      FieldName = 'skl'
      Size = 3
    end
    object TRabneedlpit: TWideStringField
      FieldName = 'needlpit'
      Size = 3
    end
    object TRabneedmse: TWideStringField
      FieldName = 'needmse'
      Size = 3
    end
    object TRabneedd: TWideStringField
      FieldName = 'needd'
      Size = 55
    end
    object TRabnom_pasp: TWideStringField
      FieldName = 'nom_pasp'
      Size = 10
    end
    object TRabnpp: TIntegerField
      FieldName = 'npp'
    end
    object TRabuser: TWideStringField
      FieldName = 'user'
      Size = 50
    end
    object TRabrekom: TWideMemoField
      FieldName = 'rekom'
      BlobType = ftWideMemo
    end
    object TRabidusl: TIntegerField
      FieldName = 'idusl'
    end
    object TRabiduslold: TIntegerField
      FieldName = 'iduslold'
    end
    object TRabUslM: TStringField
      FieldKind = fkLookup
      FieldName = 'UslM'
      LookupDataSet = TSUslM
      LookupKeyFields = 'id'
      LookupResultField = 'naim'
      KeyFields = 'idusl'
      Lookup = True
    end
    object TRabextest: TWideStringField
      FieldName = 'extest'
      Size = 1
    end
    object TRabextestold: TWideStringField
      FieldName = 'extestold'
      Size = 1
    end
    object TRabdata_proh: TDateTimeField
      FieldName = 'data_proh'
    end
    object TRablpu: TWideStringField
      FieldName = 'lpu'
      Size = 255
    end
    object TRabslexp: TWideStringField
      FieldName = 'slexp'
      Size = 255
    end
    object TRabotkl: TWideStringField
      FieldName = 'otkl'
      Size = 255
    end
    object TRabdef: TWideStringField
      FieldName = 'def'
      Size = 255
    end
    object TRabresult: TWideStringField
      FieldName = 'result'
      Size = 255
    end
    object TRabdatenaprMSE: TDateTimeField
      FieldName = 'datenaprMSE'
    end
    object TRabzaklMSE: TWideStringField
      FieldName = 'zaklMSE'
      Size = 255
    end
    object TRabdatezaklMSE: TDateTimeField
      FieldName = 'datezaklMSE'
    end
    object TRabdopinf: TWideStringField
      FieldName = 'dopinf'
      Size = 255
    end
    object TRabsostexp: TWideStringField
      FieldName = 'sostexp'
      Size = 255
    end
    object TRabpodpis: TWideStringField
      FieldName = 'podpis'
      Size = 255
    end
    object TRabid_org: TIntegerField
      FieldName = 'id_org'
    end
    object TRabid_prof: TIntegerField
      FieldName = 'id_prof'
    end
    object TRaborg: TStringField
      FieldKind = fkLookup
      FieldName = 'org'
      LookupDataSet = TOrg
      LookupKeyFields = 'id_org'
      LookupResultField = 'imya_org'
      KeyFields = 'id_org'
      Lookup = True
    end
    object TRabprof: TStringField
      FieldKind = fkLookup
      FieldName = 'prof'
      LookupDataSet = ProfPatolog
      LookupKeyFields = 'id'
      LookupResultField = 'Naim'
      KeyFields = 'id_prof'
      Lookup = True
    end
    object blnfldTRabpsih: TBooleanField
      FieldName = 'psih'
    end
    object wdmfldTRabvredn_fact2: TWideMemoField
      FieldName = 'vredn_fact2'
      BlobType = ftWideMemo
    end
    object smlntfldTRabProfSan: TSmallintField
      FieldName = 'ProfSan'
    end
    object wdstrngfldTRabPredDZ: TWideStringField
      FieldName = 'PredDZ'
      Size = 7
    end
    object wdstrngfldTRabVUPZ: TWideStringField
      FieldName = 'VUPZ'
      Size = 7
    end
    object strngfldTRabLookUpProfSan: TStringField
      FieldKind = fkLookup
      FieldName = 'LookUpProfSan'
      LookupDataSet = tblProfSan
      LookupKeyFields = 'id'
      LookupResultField = 'Name'
      KeyFields = 'ProfSan'
      Lookup = True
    end
    object wdstrngfldTRabmr: TWideStringField
      FieldName = 'mr'
      Size = 150
    end
    object wdstrngfldTRabmrr: TWideStringField
      FieldName = 'mrr'
      Size = 150
    end
    object wdstrngfldTRabmrg: TWideStringField
      FieldName = 'mrg'
      Size = 100
    end
    object wdstrngfldTRabmrn: TWideStringField
      FieldName = 'mrn'
      Size = 150
    end
    object wdstrngfldTRabmru: TWideStringField
      FieldName = 'mru'
      Size = 150
    end
    object wdstrngfldTRabmrd: TWideStringField
      FieldName = 'mrd'
      Size = 10
    end
    object wdstrngfldTRabmrk: TWideStringField
      FieldName = 'mrk'
      Size = 10
    end
    object wdstrngfldTRabmrkv: TWideStringField
      FieldName = 'mrkv'
      Size = 10
    end
    object wrdfldTRabZak282: TWordField
      FieldName = 'Zak282'
    end
    object wdstrngfldTRabprkrab282p2: TWideStringField
      FieldName = 'prkrab282p2'
      Size = 255
    end
  end
  object TSUslM: TADODataSet
    Active = True
    Connection = ADOConnection1
    CursorType = ctStatic
    Filtered = True
    CommandText = 'select * from SUslM'
    IndexFieldNames = 'cena'
    Parameters = <>
    Left = 24
    Top = 480
    object TSUslMid: TIntegerField
      FieldName = 'id'
    end
    object TSUslMnaim: TWideStringField
      FieldName = 'naim'
      Size = 50
    end
    object TSUslMcena: TIntegerField
      FieldName = 'cena'
    end
  end
  object DSSUslM: TDataSource
    AutoEdit = False
    DataSet = TSUslM
    Left = 88
    Top = 480
  end
  object QSumUsl: TADOQuery
    Connection = ADOConnection1
    CursorType = ctStatic
    Parameters = <
      item
        Name = 'date'
        Attributes = [paNullable]
        DataType = ftWideString
        NumericScale = 255
        Precision = 255
        Size = 510
        Value = Null
      end
      item
        Name = 'date2'
        Attributes = [paNullable]
        DataType = ftWideString
        NumericScale = 255
        Precision = 255
        Size = 510
        Value = Null
      end
      item
        Name = 'org'
        Attributes = [paNullable]
        DataType = ftWideString
        NumericScale = 255
        Precision = 255
        Size = 510
        Value = Null
      end>
    SQL.Strings = (
      
        'SELECT fam, imya, fath, sex, profes, vredn_fact, cat_osv, idusl,' +
        ' iduslold, extest,extestold  FROM rabotnik WHERE (data_proh BETW' +
        'EEN :date AND :date2) AND (id_org=:org)')
    Left = 32
    Top = 544
    object QSumUslfam: TWideStringField
      FieldName = 'fam'
      Size = 30
    end
    object QSumUslimya: TWideStringField
      FieldName = 'imya'
      Size = 15
    end
    object QSumUslfath: TWideStringField
      FieldName = 'fath'
    end
    object QSumUslsex: TWideStringField
      FieldName = 'sex'
      Size = 3
    end
    object QSumUslvredn_fact: TWideMemoField
      FieldName = 'vredn_fact'
      BlobType = ftWideMemo
    end
    object QSumUslprofes: TWideStringField
      FieldName = 'profes'
      Size = 60
    end
    object QSumUslidusl: TIntegerField
      FieldName = 'idusl'
    end
    object QSumUslcat_osv: TWideStringField
      FieldName = 'cat_osv'
      Size = 150
    end
    object QSumUsliduslold: TIntegerField
      FieldName = 'iduslold'
    end
    object QSumUslextest: TWideStringField
      FieldName = 'extest'
      Size = 4
    end
    object QSumUslextestold: TWideStringField
      FieldName = 'extestold'
      Size = 1
    end
  end
  object DSSumUsl: TDataSource
    DataSet = QSumUsl
    Left = 96
    Top = 536
  end
  object DSComUsl: TDataSource
    DataSet = QComUsl
    Left = 96
    Top = 600
  end
  object QComUsl: TADOQuery
    Connection = ADOConnection1
    CursorType = ctStatic
    Parameters = <
      item
        Name = 'date'
        Attributes = [paNullable]
        DataType = ftWideString
        NumericScale = 255
        Precision = 255
        Size = 510
        Value = Null
      end
      item
        Name = 'date2'
        Attributes = [paNullable]
        DataType = ftWideString
        NumericScale = 255
        Precision = 255
        Size = 510
        Value = Null
      end>
    SQL.Strings = (
      
        'SELECT rabotnik.disp_gr AS ['#1044' '#1075#1088'], Count(rabotnik.id) AS ['#1050#1086#1083'-'#1074#1086 +
        '] FROM rabotnik WHERE (((rabotnik.id_org) In (28,4,1,7,8,20,21,2' +
        '2,17,2,19,3,5,16,15,6,25,31,13,35,27,43,12,34,40,44,29)) AND (id' +
        'usl<>1000) AND ((rabotnik.data_proh) Between :date And :date2))G' +
        'ROUP BY rabotnik.disp_gr')
    Left = 24
    Top = 600
  end
  object TFYS: TADODataSet
    Connection = ADOConnection1
    CursorType = ctStatic
    CommandText = 'select * from FYS'
    Parameters = <>
    Left = 336
    Top = 80
    object TFYSid: TAutoIncField
      FieldName = 'id'
      ReadOnly = True
    end
    object TFYSkod_usl: TWideStringField
      FieldName = 'kod_usl'
      Size = 255
    end
    object TFYSName_usl: TWideStringField
      FieldName = 'Name_usl'
      Size = 255
    end
    object TFYSZena: TFloatField
      FieldName = 'Zena'
    end
    object TFYSidotdel: TIntegerField
      FieldName = 'idotdel'
    end
    object TFYSkod_vr: TIntegerField
      FieldName = 'kod_vr'
    end
    object TFYSkod_ms: TIntegerField
      FieldName = 'kod_ms'
    end
  end
  object DSFYS: TDataSource
    DataSet = TFYS
    Left = 424
    Top = 80
  end
  object TOtdel: TADODataSet
    Active = True
    Connection = ADOConnection1
    CursorType = ctStatic
    CommandText = 'select * from otdel'
    Parameters = <>
    Left = 336
    Top = 130
  end
  object DSOtdel: TDataSource
    DataSet = TOtdel
    Left = 424
    Top = 130
  end
  object TUsl: TADODataSet
    Active = True
    Connection = ADOConnection1
    CursorType = ctStatic
    CommandText = 'select * from usl'
    DataSource = DSRab
    IndexFieldNames = 'idrab'
    MasterFields = 'id'
    Parameters = <>
    Left = 336
    Top = 184
    object TUslid: TAutoIncField
      FieldName = 'id'
      ReadOnly = True
    end
    object TUslidrab: TIntegerField
      FieldName = 'idrab'
    end
    object TUslnaimusl: TWideStringField
      FieldName = 'naimusl'
      Size = 255
    end
    object TUslkol: TWordField
      FieldName = 'kol'
    end
    object TUslcena: TIntegerField
      FieldName = 'cena'
    end
    object TUsldata: TDateTimeField
      FieldName = 'data'
    end
    object TUslkod_usl: TWideStringField
      DisplayWidth = 40
      FieldName = 'kod_usl'
      Size = 40
    end
    object TUslkod_vr: TIntegerField
      FieldName = 'kod_vr'
    end
    object TUslkod_ms: TIntegerField
      FieldName = 'kod_ms'
    end
    object TUslmkb: TWideStringField
      FieldName = 'mkb'
    end
  end
  object DSUsl: TDataSource
    DataSet = TUsl
    Left = 424
    Top = 184
  end
  object QStacRep: TADOQuery
    Connection = ADOConnection1
    CursorType = ctStatic
    Parameters = <
      item
        Name = 'date'
        Attributes = [paNullable]
        DataType = ftWideString
        NumericScale = 255
        Precision = 255
        Size = 510
        Value = Null
      end
      item
        Name = 'date2'
        Attributes = [paNullable]
        DataType = ftWideString
        NumericScale = 255
        Precision = 255
        Size = 510
        Value = Null
      end>
    SQL.Strings = (
      
        'SELECT rabotnik.fam, rabotnik.imya, rabotnik.fath, rabotnik.data' +
        '_r, organiz.imya_org, rabotnik.profes, rabotnik.idusl, rabotnik.' +
        'vredn_fact, rabotnik.staz_rab_fact, rabotnik.data_proh FROM rabo' +
        'tnik INNER JOIN organiz ON rabotnik.id_org = organiz.id_org WHER' +
        'E (data_proh BETWEEN :date AND :date2)')
    Left = 336
    Top = 241
  end
  object DSStacRep: TDataSource
    DataSet = QStacRep
    Left = 424
    Top = 241
  end
  object TSUsl: TADODataSet
    Connection = ADOConnection1
    CursorType = ctStatic
    CommandText = 'select * from SUslM'
    Parameters = <>
    Left = 336
    Top = 300
  end
  object DSSUsl: TDataSource
    DataSet = TSUsl
    Left = 416
    Top = 300
  end
  object TPeriodVigr: TADODataSet
    Connection = ADOConnection1
    CursorType = ctStatic
    CommandText = 'select * from periodvigr'
    Parameters = <>
    Left = 336
    Top = 355
  end
  object DSPeriodVigr: TDataSource
    DataSet = TPeriodVigr
    Left = 408
    Top = 355
  end
  object TIssl: TADODataSet
    Connection = ADOConnection1
    CursorType = ctStatic
    CommandText = 'select * from issl'
    DataSource = DSRab
    IndexFieldNames = 'idrab'
    MasterFields = 'id'
    Parameters = <>
    Left = 336
    Top = 417
    object TIsslid: TAutoIncField
      FieldName = 'id'
      ReadOnly = True
    end
    object TIsslidrab: TIntegerField
      FieldName = 'idrab'
    end
    object TIsslidissl: TIntegerField
      FieldName = 'idissl'
    end
    object TIssldate1: TDateTimeField
      FieldName = 'date1'
    end
    object TIsslznach1: TWideStringField
      FieldName = 'znach1'
      Size = 255
    end
    object TIssldate2: TDateTimeField
      FieldName = 'date2'
    end
    object TIsslznach2: TWideStringField
      FieldName = 'znach2'
      Size = 255
    end
    object TIssldate3: TDateTimeField
      FieldName = 'date3'
    end
    object TIsslznach3: TWideStringField
      FieldName = 'znach3'
      Size = 255
    end
    object TIssllookissl: TStringField
      FieldKind = fkLookup
      FieldName = 'lookissl'
      LookupDataSet = TSIssl
      LookupKeyFields = 'id'
      LookupResultField = 'naim'
      KeyFields = 'idissl'
      Lookup = True
    end
  end
  object DSIssl: TDataSource
    DataSet = TIssl
    Left = 408
    Top = 417
  end
  object TSIssl: TADODataSet
    Connection = ADOConnection1
    CursorType = ctStatic
    CommandText = 'select * from SpravIssl'
    Parameters = <>
    Left = 336
    Top = 478
  end
  object TJournal: TADODataSet
    Connection = ADOConnection1
    CursorType = ctStatic
    CommandText = 
      'select user,id, npp, data_proh, fam, imya, fath, id_org, data_r,' +
      ' sex, profes, mkb_code, prkrab, dop_pri_usl, lpu, slexp, otkl, d' +
      'ef, result,  datezaklMSE, datenaprMSE, zaklMSE, dopinf, sostexp,' +
      ' podpis, vredn_fact from rabotnik WHERE prkrab<>'#39#39' ORDER BY id A' +
      'SC'
    FieldDefs = <
      item
        Name = 'user'
        DataType = ftWideString
        Size = 50
      end
      item
        Name = 'id'
        Attributes = [faReadonly, faFixed]
        DataType = ftAutoInc
      end
      item
        Name = 'npp'
        Attributes = [faFixed]
        DataType = ftInteger
      end
      item
        Name = 'data_proh'
        Attributes = [faFixed]
        DataType = ftDateTime
      end
      item
        Name = 'fam'
        DataType = ftWideString
        Size = 30
      end
      item
        Name = 'imya'
        DataType = ftWideString
        Size = 15
      end
      item
        Name = 'fath'
        DataType = ftWideString
        Size = 20
      end
      item
        Name = 'id_org'
        Attributes = [faFixed]
        DataType = ftInteger
      end
      item
        Name = 'data_r'
        Attributes = [faFixed]
        DataType = ftDateTime
      end
      item
        Name = 'sex'
        DataType = ftWideString
        Size = 3
      end
      item
        Name = 'profes'
        DataType = ftWideString
        Size = 255
      end
      item
        Name = 'mkb_code'
        DataType = ftWideString
        Size = 100
      end
      item
        Name = 'prkrab'
        DataType = ftWideString
        Size = 255
      end
      item
        Name = 'dop_pri_usl'
        DataType = ftWideString
        Size = 255
      end
      item
        Name = 'lpu'
        DataType = ftWideString
        Size = 255
      end
      item
        Name = 'slexp'
        DataType = ftWideString
        Size = 255
      end
      item
        Name = 'otkl'
        DataType = ftWideString
        Size = 255
      end
      item
        Name = 'def'
        DataType = ftWideString
        Size = 255
      end
      item
        Name = 'result'
        DataType = ftWideString
        Size = 255
      end
      item
        Name = 'datezaklMSE'
        Attributes = [faFixed]
        DataType = ftDateTime
      end
      item
        Name = 'datenaprMSE'
        Attributes = [faFixed]
        DataType = ftDateTime
      end
      item
        Name = 'zaklMSE'
        DataType = ftWideString
        Size = 255
      end
      item
        Name = 'dopinf'
        DataType = ftWideString
        Size = 255
      end
      item
        Name = 'sostexp'
        DataType = ftWideString
        Size = 255
      end
      item
        Name = 'podpis'
        DataType = ftWideString
        Size = 255
      end
      item
        Name = 'vredn_fact'
        DataType = ftWideMemo
      end>
    IndexFieldNames = 'data_proh'
    Parameters = <>
    StoreDefs = True
    Left = 336
    Top = 535
    object TJournalnpp: TIntegerField
      FieldName = 'npp'
    end
    object TJournaldata_proh: TDateTimeField
      FieldName = 'data_proh'
    end
    object TJournalfam: TWideStringField
      FieldName = 'fam'
      Size = 30
    end
    object TJournalimya: TWideStringField
      FieldName = 'imya'
      Size = 15
    end
    object TJournalfath: TWideStringField
      FieldName = 'fath'
    end
    object TJournaldata_r: TDateTimeField
      FieldName = 'data_r'
    end
    object TJournalid_org: TIntegerField
      FieldName = 'id_org'
    end
    object TJournalsex: TWideStringField
      FieldName = 'sex'
      Size = 3
    end
    object TJournalprofes: TWideStringField
      FieldName = 'profes'
      Size = 255
    end
    object TJournalmkb_code: TWideStringField
      FieldName = 'mkb_code'
      Size = 100
    end
    object TJournalprkrab: TWideStringField
      FieldName = 'prkrab'
      Size = 255
    end
    object TJournallookOrg: TStringField
      FieldKind = fkLookup
      FieldName = 'lookOrg'
      LookupDataSet = TOrg
      LookupKeyFields = 'id_org'
      LookupResultField = 'imya_org'
      KeyFields = 'id_org'
      Lookup = True
    end
    object TJournallpu: TWideStringField
      FieldName = 'lpu'
      Size = 255
    end
    object TJournalslexp: TWideStringField
      FieldName = 'slexp'
      Size = 255
    end
    object TJournalotkl: TWideStringField
      FieldName = 'otkl'
      Size = 255
    end
    object TJournaldef: TWideStringField
      FieldName = 'def'
      Size = 255
    end
    object TJournalresult: TWideStringField
      FieldName = 'result'
      Size = 255
    end
    object TJournaldop_pri_usl: TWideStringField
      FieldName = 'dop_pri_usl'
      Size = 255
    end
    object TJournaldatezaklMSE: TDateTimeField
      FieldName = 'datezaklMSE'
    end
    object TJournaldatenaprMSE: TDateTimeField
      FieldName = 'datenaprMSE'
    end
    object TJournalzaklMSE: TWideStringField
      FieldName = 'zaklMSE'
      Size = 255
    end
    object TJournaldopinf: TWideStringField
      FieldName = 'dopinf'
      Size = 255
    end
    object TJournalsostexp: TWideStringField
      FieldName = 'sostexp'
      Size = 255
    end
    object TJournalpodpis: TWideStringField
      FieldName = 'podpis'
      Size = 255
    end
    object TJournalid: TAutoIncField
      FieldName = 'id'
      ReadOnly = True
    end
    object TJournaluser: TWideStringField
      FieldName = 'user'
      Size = 50
    end
    object wdmfldTJournalvredn_fact: TWideMemoField
      FieldName = 'vredn_fact'
      BlobType = ftWideMemo
    end
  end
  object DSJournal: TDataSource
    DataSet = TJournal
    Left = 408
    Top = 535
  end
  object Query: TADOQuery
    Connection = ADOConnection1
    Parameters = <
      item
        Name = 'date'
        DataType = ftDateTime
        Value = Null
      end
      item
        Name = 'date2'
        DataType = ftDateTime
        Value = Null
      end>
    Left = 176
    Top = 480
  end
  object QueryRab: TADOQuery
    Connection = ADOConnection1
    CursorLocation = clUseServer
    Parameters = <
      item
        Name = 'fam'
        DataType = ftString
        Size = -1
        Value = Null
      end
      item
        Name = 'imya'
        DataType = ftString
        Size = -1
        Value = Null
      end
      item
        Name = 'fath'
        DataType = ftString
        Size = -1
        Value = Null
      end
      item
        Name = 'dr'
        DataType = ftDateTime
        Value = Null
      end>
    Left = 337
    Top = 592
  end
  object TFactor: TADODataSet
    Active = True
    Connection = ADOConnection1
    CursorType = ctStatic
    CommandText = 'select * from factor'
    DataSource = DSRab
    IndexFieldNames = 'id_rab'
    MasterFields = 'id'
    Parameters = <>
    Left = 528
    Top = 80
    object TFactorid: TAutoIncField
      FieldName = 'id'
      ReadOnly = True
    end
    object TFactorid_rab: TIntegerField
      FieldName = 'id_rab'
    end
    object TFactorcode: TWideStringField
      FieldName = 'code'
      Size = 255
    end
    object TFactornaim: TWideStringField
      FieldName = 'naim'
      Size = 255
    end
    object TFactorvrach: TWideStringField
      FieldName = 'vrach'
      Size = 255
    end
    object TFactorissl: TWideMemoField
      FieldName = 'issl'
      BlobType = ftWideMemo
    end
  end
  object DSFactor: TDataSource
    DataSet = TFactor
    Left = 616
    Top = 80
  end
  object TSFactor: TADODataSet
    Active = True
    Connection = ADOConnection1
    CursorType = ctStatic
    LockType = ltReadOnly
    OnFilterRecord = TSFactorFilterRecord
    CommandText = 'select * from sfactor'
    IndexFieldNames = 'code'
    Parameters = <>
    Left = 512
    Top = 136
    object TSFactorid: TAutoIncField
      FieldName = 'id'
      ReadOnly = True
    end
    object TSFactorcode: TWideStringField
      FieldName = 'code'
      Size = 255
    end
    object TSFactornaim: TWideStringField
      FieldName = 'naim'
      Size = 255
    end
    object TSFactorvrach: TWideStringField
      FieldName = 'vrach'
      Size = 255
    end
    object TSFactorissl: TWideMemoField
      FieldName = 'issl'
      BlobType = ftWideMemo
    end
  end
  object DSSFactor: TDataSource
    DataSet = TSFactor
    Left = 616
    Top = 144
  end
  object TFactId: TADODataSet
    Connection = ADOConnection1
    Parameters = <>
    Left = 528
    Top = 208
  end
  object DSFactId: TDataSource
    DataSet = TFactId
    Left = 616
    Top = 208
  end
  object TDMS: TADODataSet
    Connection = ADOConnection1
    CursorType = ctStatic
    CommandText = 'select * from polisdms'
    Parameters = <>
    Left = 528
    Top = 280
    object TDMSid: TAutoIncField
      FieldName = 'id'
      ReadOnly = True
    end
    object TDMSfam: TWideStringField
      FieldName = 'fam'
      Size = 255
    end
    object TDMSim: TWideStringField
      FieldName = 'im'
      Size = 255
    end
    object TDMSot: TWideStringField
      FieldName = 'ot'
      Size = 255
    end
    object TDMSdr: TDateTimeField
      FieldName = 'dr'
    end
    object TDMSser: TWideStringField
      FieldName = 'ser'
      Size = 255
    end
    object TDMSnom: TWideStringField
      FieldName = 'nom'
      Size = 255
    end
  end
  object DSDMS: TDataSource
    DataSet = TDMS
    Left = 616
    Top = 280
  end
  object Query2: TADOQuery
    Connection = ADOConnection1
    Parameters = <>
    Left = 176
    Top = 536
  end
  object QNeedSKL: TADOQuery
    Connection = ADOConnection1
    CursorType = ctStatic
    Parameters = <
      item
        Name = 'date'
        Attributes = [paNullable]
        DataType = ftWideString
        NumericScale = 255
        Precision = 255
        Size = 510
        Value = Null
      end
      item
        Name = 'date2'
        Attributes = [paNullable]
        DataType = ftWideString
        NumericScale = 255
        Precision = 255
        Size = 510
        Value = Null
      end
      item
        Name = 'org'
        Attributes = [paNullable]
        DataType = ftWideString
        NumericScale = 255
        Precision = 255
        Size = 510
        Value = Null
      end>
    SQL.Strings = (
      
        'SELECT fam, imya, fath, data_r, profes  FROM rabotnik WHERE (dat' +
        'a_proh BETWEEN :date AND :date2) AND (id_org=:org)')
    Left = 416
    Top = 592
    object WideStringField1: TWideStringField
      FieldName = 'fam'
      Size = 30
    end
    object WideStringField2: TWideStringField
      FieldName = 'imya'
      Size = 15
    end
    object WideStringField3: TWideStringField
      FieldName = 'fath'
    end
    object WideStringField5: TWideStringField
      FieldName = 'profes'
      Size = 60
    end
    object QNeedSKLdata_r: TDateTimeField
      FieldName = 'data_r'
    end
  end
  object TTempList: TADODataSet
    Connection = ADOConnection1
    CursorType = ctStatic
    CommandText = 'select fam, im, ot, dr from TempList'
    Parameters = <>
    Left = 528
    Top = 352
    object TTempListfam: TWideStringField
      FieldName = 'fam'
      Size = 50
    end
    object TTempListim: TWideStringField
      FieldName = 'im'
      Size = 50
    end
    object TTempListot: TWideStringField
      FieldName = 'ot'
      Size = 50
    end
    object TTempListdr: TDateTimeField
      FieldName = 'dr'
    end
  end
  object DSTempList: TDataSource
    DataSet = TTempList
    Left = 616
    Top = 352
  end
  object ProfPatolog: TADOTable
    Active = True
    Connection = ADOConnection1
    CursorType = ctStatic
    Filtered = True
    IndexFieldNames = 'id'
    TableName = 'ProfPatolog'
    Left = 24
    Top = 656
  end
  object DSProfPatolog: TDataSource
    DataSet = ProfPatolog
    Left = 96
    Top = 656
  end
  object qryUser: TADOQuery
    Connection = ADOConnection1
    CursorType = ctStatic
    Parameters = <
      item
        Name = 'User'
        Attributes = [paNullable]
        DataType = ftWideString
        NumericScale = 255
        Precision = 255
        Size = 510
        Value = Null
      end
      item
        Name = 'Pass'
        Attributes = [paNullable]
        DataType = ftWideString
        NumericScale = 255
        Precision = 255
        Size = 510
        Value = Null
      end>
    Prepared = True
    SQL.Strings = (
      'SELECT User.*'
      'FROM [User]'
      'Where ( User.Login=:User) and ( User.Password=:Pass)'
      ';')
    Left = 200
    Top = 72
  end
  object tblProfSan: TADOTable
    Active = True
    Connection = ADOConnection1
    CursorType = ctStatic
    TableName = 'ProfSan'
    Left = 208
    Top = 224
  end
  object dsProfSan: TDataSource
    DataSet = tblProfSan
    Left = 224
    Top = 304
  end
  object tblTDogovor: TADOTable
    Active = True
    Connection = ADOConnection1
    CursorType = ctStatic
    Filtered = True
    TableName = 'Dogovor'
    Left = 792
    Top = 416
    object atncfldTDogovorid: TAutoIncField
      FieldName = 'id'
      ReadOnly = True
    end
    object wdstrngfldTDogovorNaim: TWideStringField
      FieldName = 'Naim'
      Size = 255
    end
  end
  object dsDogovor: TDataSource
    DataSet = tblTDogovor
    Left = 848
    Top = 416
  end
  object tblZak282: TADOTable
    Active = True
    Connection = ADOConnection1
    CursorType = ctStatic
    LockType = ltReadOnly
    IndexName = 'PrimaryKey'
    TableName = 'Zak282'
    Left = 584
    Top = 656
  end
  object qryJurnal282: TADOQuery
    Connection = ADOConnection1
    CursorType = ctStatic
    Parameters = <>
    SQL.Strings = (
      
        'select user,id, npp, data_proh, fam, imya, fath, id_org, data_r,' +
        ' sex, profes, mkb_code, prkrab, dop_pri_usl, lpu, slexp, otkl, d' +
        'ef, result,  datezaklMSE, datenaprMSE, zaklMSE, dopinf, sostexp,' +
        ' podpis, vredn_fact from rabotnik WHERE prkrab<>'#39#39' ORDER BY id A' +
        'SC')
    Left = 584
    Top = 600
  end
  object dsJurnal282: TDataSource
    DataSet = qryJurnal282
    Left = 656
    Top = 600
  end
  object dsZak282: TDataSource
    DataSet = tblZak282
    Left = 656
    Top = 664
  end
end
