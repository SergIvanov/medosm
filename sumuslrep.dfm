object fUsl: TfUsl
  Left = 0
  Top = 0
  Caption = 'fUsl'
  ClientHeight = 448
  ClientWidth = 759
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
  object DBGridEh1: TDBGridEh
    Left = 0
    Top = 113
    Width = 759
    Height = 335
    Align = alClient
    DataSource = fDM.DSSumUsl
    DrawMemoText = True
    DynProps = <>
    EditActions = [geaDeleteEh]
    FooterRowCount = 1
    FooterParams.Color = clWindow
    SumList.Active = True
    TabOrder = 0
    Columns = <
      item
        CellButtons = <>
        DynProps = <>
        EditButtons = <>
        FieldName = 'fam'
        Footers = <>
      end
      item
        CellButtons = <>
        DynProps = <>
        EditButtons = <>
        FieldName = 'imya'
        Footers = <>
      end
      item
        CellButtons = <>
        DynProps = <>
        EditButtons = <>
        FieldName = 'fath'
        Footers = <>
      end
      item
        CellButtons = <>
        DynProps = <>
        EditButtons = <>
        FieldName = 'sex'
        Footers = <>
        Width = 25
      end
      item
        CellButtons = <>
        DynProps = <>
        EditButtons = <>
        FieldName = 'profes'
        Footers = <>
        Width = 90
      end
      item
        CellButtons = <>
        DynProps = <>
        EditButtons = <>
        FieldName = 'vredn_fact'
        Footers = <>
      end
      item
        CellButtons = <>
        DynProps = <>
        EditButtons = <>
        FieldName = 'cat_osv'
        Footers = <>
        Width = 50
      end
      item
        CellButtons = <>
        DynProps = <>
        EditButtons = <>
        FieldName = 'idusl'
        Footers = <>
      end
      item
        CellButtons = <>
        DynProps = <>
        EditButtons = <>
        FieldName = 'iduslold'
        Footers = <>
      end
      item
        CellButtons = <>
        DynProps = <>
        EditButtons = <>
        FieldName = 'extest'
        Footers = <>
      end>
    object RowDetailData: TRowDetailPanelControlEh
    end
  end
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 759
    Height = 113
    Align = alTop
    TabOrder = 1
    object Button1: TButton
      Left = 488
      Top = 8
      Width = 89
      Height = 25
      Caption = #1057#1092#1086#1088#1084#1080#1088#1086#1074#1072#1090#1100
      TabOrder = 0
      OnClick = Button1Click
    end
    object DateTimePicker1: TDateTimePicker
      Left = 8
      Top = 8
      Width = 121
      Height = 21
      Date = 40933.426980474530000000
      Time = 40933.426980474530000000
      TabOrder = 1
    end
    object DateTimePicker2: TDateTimePicker
      Left = 168
      Top = 8
      Width = 113
      Height = 21
      Date = 40933.429906493060000000
      Time = 40933.429906493060000000
      TabOrder = 2
    end
    object DBLookupComboBox1: TDBLookupComboBox
      Left = 320
      Top = 8
      Width = 145
      Height = 21
      KeyField = 'id_org'
      ListField = 'imya_org'
      ListSource = fDM.DSOrg
      TabOrder = 3
    end
    object Button3: TButton
      Left = 488
      Top = 39
      Width = 105
      Height = 25
      Caption = #1042#1099#1075#1088#1091#1079#1080#1090#1100' '#1086#1090#1095#1077#1090
      TabOrder = 4
      OnClick = Button3Click
    end
    object Button4: TButton
      Left = 608
      Top = 39
      Width = 105
      Height = 25
      Caption = #1042#1099#1075#1088#1091#1079#1080#1090#1100' '#1074#1089#1077' '
      TabOrder = 5
      OnClick = Button4Click
    end
    object Button2: TButton
      Left = 488
      Top = 70
      Width = 105
      Height = 25
      Caption = #1042#1099#1075#1088#1079#1080#1090#1100' 6'#1062
      TabOrder = 6
      OnClick = Button2Click
    end
    object Button5: TButton
      Left = 608
      Top = 70
      Width = 105
      Height = 25
      Caption = #1042#1099#1075#1088#1091#1079#1080#1090#1100' '#1074#1089#1077' 6'#1062
      TabOrder = 7
      OnClick = Button5Click
    end
    object ProgressBar1: TProgressBar
      Left = 8
      Top = 70
      Width = 457
      Height = 20
      BarColor = clHighlight
      TabOrder = 8
    end
  end
end
