object frmEkonom: TfrmEkonom
  Left = 0
  Top = 0
  Caption = #1060#1086#1088#1084#1072' '#1076#1083#1103' '#1101#1082#1086#1085#1086#1084#1080#1089#1090#1086#1074
  ClientHeight = 596
  ClientWidth = 978
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  OnCreate = FormCreate
  PixelsPerInch = 96
  TextHeight = 13
  object pgcTabPages: TPageControl
    Left = 0
    Top = 0
    Width = 978
    Height = 596
    ActivePage = ts1
    Align = alClient
    TabOrder = 0
    object ts1: TTabSheet
      Caption = #1054#1089#1085#1086#1074#1085#1086#1077
      object pnl1: TPanel
        Left = 0
        Top = 0
        Width = 970
        Height = 145
        Align = alTop
        TabOrder = 0
        object btn1: TButton
          Left = 392
          Top = 0
          Width = 89
          Height = 52
          Caption = #1057#1092#1086#1088#1084#1080#1088#1086#1074#1072#1090#1100
          TabOrder = 0
          OnClick = btn1Click
        end
        object dblkcbb1: TDBLookupComboBox
          Left = 95
          Top = 57
          Width = 290
          Height = 21
          KeyField = 'id_org'
          ListField = 'imya_org'
          ListSource = fDM.DSOrg
          TabOrder = 1
        end
        object dblkcbb2: TDBLookupComboBox
          Left = 95
          Top = 84
          Width = 290
          Height = 21
          KeyField = 'id'
          ListField = 'Naim'
          ListSource = fDM.dsDogovor
          TabOrder = 2
        end
        object dtp1: TDateTimePicker
          Left = 228
          Top = 16
          Width = 121
          Height = 21
          Date = 40933.426980474530000000
          Time = 40933.426980474530000000
          TabOrder = 3
        end
        object dtp2: TDateTimePicker
          Left = 87
          Top = 16
          Width = 113
          Height = 21
          Date = 40933.429906493060000000
          Time = 40933.429906493060000000
          TabOrder = 4
        end
        object txt1: TStaticText
          Left = 15
          Top = 61
          Width = 74
          Height = 17
          Caption = #1054#1088#1075#1072#1085#1080#1079#1072#1094#1080#1103':'
          TabOrder = 5
        end
        object txt2: TStaticText
          Left = 17
          Top = 20
          Width = 64
          Height = 17
          Caption = #1047#1072' '#1087#1077#1088#1080#1086#1076' '#1089
          TabOrder = 6
        end
        object txt3: TStaticText
          Left = 206
          Top = 20
          Width = 16
          Height = 17
          Caption = #1087#1086
          TabOrder = 7
        end
        object txt5: TStaticText
          Left = 15
          Top = 84
          Width = 51
          Height = 17
          Caption = #1044#1086#1075#1086#1074#1086#1088':'
          TabOrder = 8
        end
      end
      object dbgrdh2: TDBGridEh
        Left = 0
        Top = 344
        Width = 970
        Height = 224
        Align = alBottom
        DataSource = ds2
        DynProps = <>
        ReadOnly = True
        TabOrder = 1
        object RowDetailData: TRowDetailPanelControlEh
        end
      end
      object dbgrdh1: TDBGridEh
        Left = 0
        Top = 145
        Width = 970
        Height = 199
        Align = alClient
        DataSource = ds1
        DynProps = <>
        OptionsEh = [dghFixed3D, dghHighlightFocus, dghClearSelection, dghMultiSortMarking, dghDialogFind, dghColumnResize, dghColumnMove, dghExtendVertLines]
        ReadOnly = True
        TabOrder = 2
        object RowDetailData: TRowDetailPanelControlEh
        end
      end
    end
    object ts2: TTabSheet
      Caption = #1056#1077#1077#1089#1090#1088
      ImageIndex = 1
      object btn5: TButton
        Left = 0
        Top = 0
        Width = 970
        Height = 33
        Align = alTop
        Caption = #1056#1077#1077#1089#1090#1088' '#1091#1089#1083' '#1087#1086' '#1054#1041#1044#1055
        TabOrder = 0
        OnClick = btn5Click
      end
      object dbgrdh3: TDBGridEh
        Left = 0
        Top = 281
        Width = 970
        Height = 352
        Align = alTop
        DataSource = ds3
        DynProps = <>
        TabOrder = 1
        Columns = <
          item
            CellButtons = <>
            DynProps = <>
            EditButtons = <>
            FieldName = 'CodPrise'
            Footers = <>
          end
          item
            CellButtons = <>
            DynProps = <>
            EditButtons = <>
            FieldName = 'naim1'
            Footers = <>
            Width = 600
          end
          item
            CellButtons = <>
            DynProps = <>
            EditButtons = <>
            FieldName = 'cena'
            Footers = <>
          end
          item
            CellButtons = <>
            DynProps = <>
            EditButtons = <>
            FieldName = 'kol'
            Footers = <>
          end
          item
            CellButtons = <>
            DynProps = <>
            EditButtons = <>
            FieldName = 'sum'
            Footers = <>
          end>
        object RowDetailData: TRowDetailPanelControlEh
        end
      end
      object edt1: TEdit
        Left = 0
        Top = 33
        Width = 970
        Height = 21
        Align = alTop
        TabOrder = 2
        OnChange = edt1Change
      end
      object dbgrdh4: TDBGridEh
        Left = 0
        Top = 54
        Width = 970
        Height = 227
        Align = alTop
        DataSource = ds4
        DrawMemoText = True
        DynProps = <>
        ReadOnly = True
        TabOrder = 3
        OnDblClick = dbgrdh4DblClick
        Columns = <
          item
            CellButtons = <>
            DynProps = <>
            EditButtons = <>
            FieldName = 'CodPrise'
            Footers = <>
          end
          item
            CellButtons = <>
            DynProps = <>
            EditButtons = <>
            FieldName = 'Naim'
            Footers = <>
            Width = 713
          end
          item
            CellButtons = <>
            DynProps = <>
            EditButtons = <>
            FieldName = 'Cena'
            Footers = <>
          end>
        object RowDetailData: TRowDetailPanelControlEh
        end
      end
      object dbnvgr1: TDBNavigator
        Left = 608
        Top = 540
        Width = 240
        Height = 25
        DataSource = ds3
        TabOrder = 4
      end
    end
  end
  object ds1: TDataSource
    DataSet = mtblh1
    Left = 525
    Top = 88
  end
  object mtblh1: TMemTableEh
    Active = True
    FieldDefs = <
      item
        Name = 'Idgroup'
        DataType = ftSmallint
        Precision = 15
      end
      item
        Name = 'Fam'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'imya'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'fath'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'Sex'
        DataType = ftString
        Size = 10
      end
      item
        Name = 'data_r'
        DataType = ftDate
      end
      item
        Name = 'profes'
        DataType = ftString
        Size = 255
      end
      item
        Name = 'kol'
        DataType = ftSmallint
        Precision = 15
      end
      item
        Name = 'Codfactor'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'CodPrise'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'Naim'
        DataType = ftString
        Size = 255
      end
      item
        Name = 'naim1'
        DataType = ftString
        Size = 255
      end
      item
        Name = 'cena'
        DataType = ftCurrency
        Precision = 15
      end>
    IndexDefs = <>
    Params = <>
    StoreDefs = True
    Left = 477
    Top = 88
    object smlntfldmtblh1Idgroup: TSmallintField
      DisplayWidth = 10
      FieldName = 'Idgroup'
    end
    object strngfldmtblh1Fam: TStringField
      DisplayWidth = 20
      FieldName = 'Fam'
    end
    object strngfldmtblh1imya: TStringField
      DisplayWidth = 20
      FieldName = 'imya'
    end
    object strngfldmtblh1fath: TStringField
      DisplayWidth = 20
      FieldName = 'fath'
    end
    object strngfldmtblh1Sex: TStringField
      DisplayWidth = 4
      FieldName = 'Sex'
      Size = 10
    end
    object dtfldmtblh1data_r: TDateField
      DisplayWidth = 10
      FieldName = 'data_r'
    end
    object strngfldmtblh1profes: TStringField
      DisplayWidth = 42
      FieldName = 'profes'
      Size = 255
    end
    object smlntfldmtblh1kol: TSmallintField
      DisplayWidth = 10
      FieldName = 'kol'
    end
    object strngfldmtblh1Codfactor: TStringField
      DisplayWidth = 20
      FieldName = 'Codfactor'
    end
    object strngfldmtblh1CodPrise: TStringField
      DisplayWidth = 20
      FieldName = 'CodPrise'
    end
    object strngfldmtblh1Naim: TStringField
      DisplayWidth = 42
      FieldName = 'Naim'
      Size = 255
    end
    object strngfldmtblh1naim1: TStringField
      DisplayWidth = 255
      FieldName = 'naim1'
      Size = 255
    end
    object crncyfldmtblh1cena: TCurrencyField
      DisplayWidth = 10
      FieldName = 'cena'
    end
    object MemTableData: TMemTableDataEh
      object DataStruct: TMTDataStructEh
        object Idgroup: TMTNumericDataFieldEh
          FieldName = 'Idgroup'
          NumericDataType = fdtSmallintEh
          AutoIncrement = False
          currency = False
          Precision = 15
        end
        object Fam: TMTStringDataFieldEh
          FieldName = 'Fam'
          StringDataType = fdtStringEh
        end
        object imya: TMTStringDataFieldEh
          FieldName = 'imya'
          StringDataType = fdtStringEh
        end
        object fath: TMTStringDataFieldEh
          FieldName = 'fath'
          StringDataType = fdtStringEh
        end
        object Sex: TMTStringDataFieldEh
          FieldName = 'Sex'
          StringDataType = fdtStringEh
          Size = 10
        end
        object data_r: TMTDateTimeDataFieldEh
          FieldName = 'data_r'
          DateTimeDataType = fdtDateEh
        end
        object profes: TMTStringDataFieldEh
          FieldName = 'profes'
          StringDataType = fdtStringEh
          Size = 255
        end
        object kol: TMTNumericDataFieldEh
          FieldName = 'kol'
          NumericDataType = fdtSmallintEh
          AutoIncrement = False
          currency = False
          Precision = 15
        end
        object Codfactor: TMTStringDataFieldEh
          FieldName = 'Codfactor'
          StringDataType = fdtStringEh
        end
        object CodPrise: TMTStringDataFieldEh
          FieldName = 'CodPrise'
          StringDataType = fdtStringEh
        end
        object Naim: TMTStringDataFieldEh
          FieldName = 'Naim'
          StringDataType = fdtStringEh
          Size = 255
        end
        object naim1: TMTStringDataFieldEh
          FieldName = 'naim1'
          StringDataType = fdtStringEh
          Size = 255
        end
        object cena: TMTNumericDataFieldEh
          FieldName = 'cena'
          NumericDataType = fdtCurrencyEh
          AutoIncrement = False
          currency = False
          Precision = 15
        end
      end
      object RecordsList: TRecordsListEh
        Data = (
          (
            0
            ''
            ''
            ''
            ''
            0d
            ''
            0
            ''
            ''
            ''
            ''
            0c))
      end
    end
  end
  object qry1: TADOQuery
    Connection = fDM.ADOConnection1
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
      'SELECT *'
      'FROM rabotnik'
      
        'WHERE (data_proh BETWEEN :date AND :date2) and (rabotnik.id_org=' +
        ':org) and ((UCase(cat_osv) Not Like "'#1055'%") or (IsNull(cat_osv)))'
      ';')
    Left = 477
    Top = 128
  end
  object ds2: TDataSource
    DataSet = qry1
    Left = 525
    Top = 136
  end
  object qryMax: TADOQuery
    Connection = fDM.ADOConnection1
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
      end
      item
        Name = 'id_dog'
        Attributes = [paNullable]
        DataType = ftWideString
        NumericScale = 255
        Precision = 255
        Size = 510
        Value = Null
      end
      item
        Name = 'id'
        Attributes = [paNullable]
        DataType = ftWideString
        NumericScale = 255
        Precision = 255
        Size = 510
        Value = Null
      end>
    SQL.Strings = (
      
        'SELECT rabotnik.id,rabotnik.fam, rabotnik.imya, rabotnik.fath, r' +
        'abotnik.sex, rabotnik.data_r,rabotnik.profes, factor.code,factor' +
        '.Naim, SprPrice.kod_price, SprPrice.naim AS naim1, Max(SprPrice.' +
        'cena) AS MaxCena'
      
        'FROM (rabotnik LEFT JOIN factor ON rabotnik.id = factor.id_rab) ' +
        'LEFT JOIN SprPrice ON factor.code = SprPrice.kod_sfactor'
      
        'WHERE (data_proh BETWEEN :date AND :date2) and (rabotnik.id_org=' +
        ':org) and (SprPrice.id_dog=:id_dog) and (rabotnik.id=:id)'
      
        'GROUP BY rabotnik.id,rabotnik.fam, rabotnik.imya, rabotnik.fath,' +
        ' rabotnik.sex, rabotnik.data_r, rabotnik.profes, factor.code,fac' +
        'tor.Naim, SprPrice.kod_price,SprPrice.naim'
      'ORDER BY Max(SprPrice.cena) DESC;')
    Left = 629
    Top = 40
  end
  object qry6c: TADOQuery
    Connection = fDM.ADOConnection1
    CursorType = ctStatic
    Parameters = <
      item
        Name = 'id'
        Attributes = [paNullable]
        DataType = ftWideString
        NumericScale = 255
        Precision = 255
        Size = 510
        Value = Null
      end>
    SQL.Strings = (
      
        'SELECT rabotnik.id_org, rabotnik.fam, rabotnik.imya, rabotnik.fa' +
        'th, rabotnik.sex, rabotnik.data_r,rabotnik.profes, rabotnik.data' +
        '_proh, DateDiff("yyyy",rabotnik.data_r, rabotnik.data_proh) as V' +
        'oz,rabotnik.idusl'
      'FROM rabotnik'
      'where (rabotnik.id = :id) '
      ';')
    Left = 581
    Top = 40
  end
  object qryID: TADOQuery
    Connection = fDM.ADOConnection1
    Parameters = <
      item
        Name = 'id'
        Attributes = [paNullable]
        DataType = ftWideString
        NumericScale = 255
        Precision = 255
        Size = 510
        Value = Null
      end>
    SQL.Strings = (
      'SELECT rabotnik.fam, rabotnik.imya, rabotnik.fath'
      'FROM rabotnik'
      'WHERE rabotnik.id=:id'
      '')
    Left = 533
    Top = 40
  end
  object vrtlqry1: TVirtualQuery
    SourceDataSets = <
      item
        DataSet = mtblh1
      end>
    SQL.Strings = (
      'Select CodPrise,naim1,cena,SUM(kol) as kol,cena*SUM(kol) as sum'
      'from mtblh1'
      'Group by CodPrise,naim1,cena')
    Left = 781
    Top = 40
  end
  object ds3: TDataSource
    DataSet = vrtlqry1
    Left = 816
    Top = 40
  end
  object ds4: TDataSource
    DataSet = ds5
    Left = 872
    Top = 104
  end
  object ds5: TADODataSet
    Active = True
    Connection = fDM.ADOConnection1
    CursorType = ctStatic
    CommandText = 'select *  from SprPriceEk'#13#10'where nom =1'
    Parameters = <>
    Left = 836
    Top = 104
    object wdstrngfldds5CodPrise: TWideStringField
      DisplayWidth = 10
      FieldName = 'CodPrise'
      Size = 10
    end
    object wdmfldds5Naim: TWideMemoField
      DisplayWidth = 82
      FieldName = 'Naim'
      BlobType = ftWideMemo
    end
    object bcdfldds5Cena: TBCDField
      DisplayWidth = 20
      FieldName = 'Cena'
      Precision = 19
    end
  end
end
