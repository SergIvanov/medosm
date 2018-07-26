object fFYS: TfFYS
  Left = 0
  Top = 0
  Caption = 'fFYS'
  ClientHeight = 501
  ClientWidth = 734
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poOwnerFormCenter
  OnClose = FormClose
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 734
    Height = 233
    Align = alTop
    TabOrder = 0
    object Label1: TLabel
      Left = 7
      Top = 20
      Width = 44
      Height = 13
      Caption = #1060#1072#1084#1080#1083#1080#1103
    end
    object Label2: TLabel
      Left = 7
      Top = 60
      Width = 19
      Height = 13
      Caption = #1048#1084#1103
    end
    object Label3: TLabel
      Left = 98
      Top = 62
      Width = 49
      Height = 13
      Caption = #1054#1090#1095#1077#1089#1090#1074#1086
    end
    object Label4: TLabel
      Left = 296
      Top = 56
      Width = 60
      Height = 13
      Caption = #1057#1087#1077#1094#1080#1072#1083#1080#1089#1090
    end
    object Label5: TLabel
      Left = 208
      Top = 100
      Width = 35
      Height = 13
      Caption = #1059#1089#1083#1091#1075#1072
    end
    object Label6: TLabel
      Left = 364
      Top = 145
      Width = 35
      Height = 13
      Caption = #1050#1086#1083'-'#1074#1086
    end
    object Label7: TLabel
      Left = 429
      Top = 145
      Width = 26
      Height = 13
      Caption = #1062#1077#1085#1072
    end
    object Label8: TLabel
      Left = 532
      Top = 145
      Width = 76
      Height = 13
      Caption = #1044#1072#1090#1072' '#1086#1082#1072#1079#1072#1085#1080#1103
    end
    object Label9: TLabel
      Left = 11
      Top = 180
      Width = 66
      Height = 13
      Caption = #1054#1088#1075#1072#1085#1080#1079#1072#1094#1080#1103
    end
    object Label10: TLabel
      Left = 7
      Top = 100
      Width = 80
      Height = 13
      Caption = #1044#1072#1090#1072' '#1088#1086#1078#1076#1077#1085#1080#1103
    end
    object Label11: TLabel
      Left = 9
      Top = 139
      Width = 30
      Height = 13
      Caption = #1055#1086#1083#1080#1089
    end
    object Label12: TLabel
      Left = 458
      Top = 10
      Width = 164
      Height = 13
      Caption = #1044#1072#1090#1072' '#1087#1088#1080#1105#1084#1072' '#1087#1088#1086#1092#1087#1072#1090#1086#1083#1086#1075#1072
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'Tahoma'
      Font.Style = [fsBold]
      ParentFont = False
    end
    object DBEditEh1: TDBEditEh
      Left = 7
      Top = 34
      Width = 180
      Height = 21
      DataField = 'fam'
      DataSource = fDM.DSRab
      DynProps = <>
      EditButtons = <>
      TabOrder = 1
      Visible = True
    end
    object DBEditEh2: TDBEditEh
      Left = 7
      Top = 75
      Width = 85
      Height = 21
      DataField = 'imya'
      DataSource = fDM.DSRab
      DynProps = <>
      EditButtons = <>
      TabOrder = 2
      Visible = True
    end
    object DBEditEh3: TDBEditEh
      Left = 98
      Top = 75
      Width = 89
      Height = 21
      DataField = 'fath'
      DataSource = fDM.DSRab
      DynProps = <>
      EditButtons = <>
      TabOrder = 3
      Visible = True
    end
    object DBLookupComboboxEh1: TDBLookupComboboxEh
      Left = 296
      Top = 75
      Width = 336
      Height = 21
      DynProps = <>
      DataField = ''
      EditButtons = <>
      KeyField = 'id'
      ListField = #1057#1087#1077#1094#1080#1072#1083#1100#1085#1086#1089#1090#1100
      ListSource = fDM.DSOtdel
      TabOrder = 4
      Visible = True
      OnChange = DBLookupComboboxEh1Change
    end
    object DBLookupComboboxEh2: TDBLookupComboboxEh
      Left = 144
      Top = 117
      Width = 488
      Height = 21
      DynProps = <>
      DataField = ''
      DropDownBox.Columns = <
        item
          AutoFitColWidth = False
          FieldName = 'kod_usl'
          Width = 100
        end
        item
          FieldName = 'Name_usl'
        end>
      EditButtons = <>
      KeyField = 'id'
      ListField = 'kod_usl'
      ListSource = fDM.DSFYS
      TabOrder = 5
      Visible = True
      OnChange = DBLookupComboboxEh2Change
    end
    object Edit1: TEdit
      Left = 364
      Top = 164
      Width = 49
      Height = 21
      TabOrder = 6
      Text = '1'
      OnChange = Edit1Change
    end
    object Edit2: TEdit
      Left = 429
      Top = 164
      Width = 80
      Height = 21
      TabOrder = 7
    end
    object Button1: TButton
      Left = 512
      Top = 194
      Width = 103
      Height = 25
      Caption = #1044#1086#1073#1072#1074#1080#1090#1100' '#1091#1089#1083#1091#1075#1091
      TabOrder = 8
      OnClick = Button1Click
    end
    object DBLookupComboBox1: TDBLookupComboBox
      Left = 10
      Top = 197
      Width = 145
      Height = 21
      DataField = 'id_org'
      DataSource = fDM.DSRab
      KeyField = 'id_org'
      ListField = 'imya_org'
      ListSource = fDM.DSOrg
      ReadOnly = True
      TabOrder = 9
    end
    object DBDateTimeEditEh1: TDBDateTimeEditEh
      Left = 7
      Top = 114
      Width = 98
      Height = 21
      DataField = 'data_r'
      DataSource = fDM.DSRab
      DynProps = <>
      EditButtons = <>
      Kind = dtkDateEh
      TabOrder = 10
      Visible = True
    end
    object DBEditEh5: TDBEditEh
      Left = 7
      Top = 153
      Width = 180
      Height = 21
      DataField = 'polis'
      DataSource = fDM.DSRab
      DynProps = <>
      EditButtons = <>
      TabOrder = 11
      Visible = True
    end
    object DateTimePicker1: TDateTimePicker
      Left = 532
      Top = 164
      Width = 100
      Height = 21
      Date = 41372.677508263890000000
      Time = 41372.677508263890000000
      TabOrder = 12
    end
    object Button2: TButton
      Left = 193
      Top = 34
      Width = 64
      Height = 25
      Caption = #1054#1073#1085#1086#1074#1080#1090#1100
      TabOrder = 13
      OnClick = Button2Click
    end
    object DateTimePicker2: TDateTimePicker
      Left = 458
      Top = 29
      Width = 141
      Height = 21
      Date = 41515.546215289350000000
      Time = 41515.546215289350000000
      TabOrder = 0
      OnExit = DateTimePicker2Exit
    end
  end
  object Panel2: TPanel
    Left = 0
    Top = 233
    Width = 734
    Height = 212
    Align = alClient
    TabOrder = 1
    object DBGridEh1: TDBGridEh
      Left = 1
      Top = 1
      Width = 732
      Height = 210
      Align = alClient
      DataSource = fDM.DSUsl
      DynProps = <>
      FooterParams.Color = clWindow
      IndicatorOptions = [gioShowRowIndicatorEh, gioShowRecNoEh]
      OptionsEh = [dghFixed3D, dghHighlightFocus, dghClearSelection, dghDialogFind, dghShowRecNo, dghColumnResize, dghColumnMove, dghExtendVertLines]
      TabOrder = 0
      Columns = <
        item
          CellButtons = <>
          DynProps = <>
          EditButtons = <>
          FieldName = 'kod_usl'
          Footers = <>
          Width = 150
        end
        item
          CellButtons = <>
          DynProps = <>
          EditButtons = <>
          FieldName = 'naimusl'
          Footers = <>
          Width = 276
        end
        item
          CellButtons = <>
          DynProps = <>
          EditButtons = <>
          FieldName = 'kol'
          Footers = <>
          Width = 35
        end
        item
          CellButtons = <>
          DynProps = <>
          EditButtons = <>
          FieldName = 'cena'
          Footer.ValueType = fvtSum
          Footers = <>
          Width = 38
        end
        item
          CellButtons = <>
          DynProps = <>
          EditButtons = <>
          FieldName = 'kod_vr'
          Footers = <>
          Width = 47
        end
        item
          CellButtons = <>
          DynProps = <>
          EditButtons = <>
          FieldName = 'kod_ms'
          Footers = <>
          Width = 46
        end
        item
          CellButtons = <>
          DynProps = <>
          EditButtons = <>
          FieldName = 'mkb'
          Footers = <>
          Width = 57
        end
        item
          CellButtons = <>
          DynProps = <>
          EditButtons = <>
          FieldName = 'data'
          Footers = <>
          Width = 67
        end>
      object RowDetailData: TRowDetailPanelControlEh
      end
    end
  end
  object Panel3: TPanel
    Left = 0
    Top = 445
    Width = 734
    Height = 56
    Align = alBottom
    TabOrder = 2
    object DBNavigator1: TDBNavigator
      Left = 385
      Top = 5
      Width = 240
      Height = 25
      DataSource = fDM.DSUsl
      TabOrder = 0
    end
  end
end
