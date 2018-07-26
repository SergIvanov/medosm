object fJournal: TfJournal
  Left = 0
  Top = 0
  Caption = #1046#1091#1088#1085#1072#1083' '#1091#1095#1105#1090#1072' '#1042#1050
  ClientHeight = 462
  ClientWidth = 915
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  Menu = MainMenu1
  OldCreateOrder = False
  Position = poOwnerFormCenter
  OnClose = FormClose
  OnCreate = FormCreate
  PixelsPerInch = 96
  TextHeight = 13
  object Panel1: TPanel
    Left = 0
    Top = 73
    Width = 915
    Height = 389
    Align = alClient
    TabOrder = 0
    object DBGridEh1: TDBGridEh
      Left = 1
      Top = 1
      Width = 913
      Height = 387
      Align = alClient
      Ctl3D = False
      DataSource = fDM.DSJournal
      DrawMemoText = True
      DynProps = <>
      Flat = True
      FooterRowCount = 1
      FooterParams.Color = clWindow
      IndicatorOptions = [gioShowRowIndicatorEh, gioShowRecNoEh]
      OptionsEh = [dghFixed3D, dghHighlightFocus, dghClearSelection, dghDialogFind, dghShowRecNo, dghColumnResize, dghColumnMove, dghExtendVertLines]
      ParentCtl3D = False
      STFilter.InstantApply = True
      STFilter.Local = True
      STFilter.Visible = True
      TabOrder = 0
      TitleParams.RowHeight = 60
      OnDblClick = DBGridEh1DblClick
      Columns = <
        item
          CellButtons = <>
          DynProps = <>
          EditButtons = <>
          FieldName = 'npp'
          Footers = <>
          Title.Caption = #8470' '#1087#1088#1086#1090#1086#1082#1086#1083#1072
        end
        item
          CellButtons = <>
          DynProps = <>
          EditButtons = <>
          FieldName = 'data_proh'
          Footers = <>
          Title.Caption = #1044#1072#1090#1072' '#1101#1082#1089#1087#1077#1088#1090#1080#1079#1099
        end
        item
          CellButtons = <>
          DynProps = <>
          EditButtons = <>
          FieldName = 'lpu'
          Footers = <>
          Title.Caption = #1053#1072#1080#1084#1077#1085#1086#1074#1072#1085#1080#1077' '#1051#1055#1059
          Width = 150
        end
        item
          CellButtons = <>
          DynProps = <>
          EditButtons = <>
          FieldName = 'fam'
          Footers = <>
          Title.Caption = #1060#1072#1084#1080#1083#1080#1103
          Width = 150
        end
        item
          CellButtons = <>
          DynProps = <>
          EditButtons = <>
          FieldName = 'imya'
          Footers = <>
          Title.Caption = #1048#1084#1103
        end
        item
          CellButtons = <>
          DynProps = <>
          EditButtons = <>
          FieldName = 'fath'
          Footers = <>
          Title.Caption = #1054#1090#1095#1077#1089#1090#1074#1086
        end
        item
          CellButtons = <>
          DynProps = <>
          EditButtons = <>
          FieldName = 'data_r'
          Footers = <>
          Title.Caption = #1044#1072#1090#1072' '#1088#1086#1078#1076#1077#1085#1080#1103
          Width = 70
        end
        item
          CellButtons = <>
          DynProps = <>
          EditButtons = <>
          FieldName = 'lookOrg'
          Footers = <>
          Title.Caption = #1054#1088#1075#1072#1085#1080#1079#1072#1094#1080#1103
          Width = 102
        end
        item
          CellButtons = <>
          DynProps = <>
          EditButtons = <>
          FieldName = 'sex'
          Footers = <>
          Title.Caption = #1055#1086#1083
          Width = 30
        end
        item
          CellButtons = <>
          DynProps = <>
          EditButtons = <>
          FieldName = 'profes'
          Footers = <>
          Title.Caption = #1055#1088#1086#1092#1077#1089#1089#1080#1103
          Width = 150
        end
        item
          CellButtons = <>
          DynProps = <>
          EditButtons = <>
          FieldName = 'mkb_code'
          Footers = <>
          Title.Caption = #1055#1088#1080#1095#1080#1085#1072' '#1086#1073#1088#1072#1097#1077#1085#1080#1103
          Width = 90
        end
        item
          CellButtons = <>
          DynProps = <>
          EditButtons = <>
          FieldName = 'slexp'
          Footers = <>
          Title.Caption = #1061#1072#1088'-'#1082#1072' '#1089#1083#1091#1095#1072#1103' '#1101#1082#1089#1087#1077#1088#1090#1080#1079#1099
          Width = 150
        end
        item
          CellButtons = <>
          DynProps = <>
          EditButtons = <>
          FieldName = 'prkrab'
          Footers = <>
          Title.Caption = #1042#1080#1076' '#1080' '#1087#1088#1077#1076#1084#1077#1090' '#1101#1082#1089#1087#1077#1088#1090#1080#1079#1099' ('#1055#1088#1086#1090#1080#1074#1086#1087#1086#1082#1072#1079#1072#1085#1080#1103' '#1082' '#1088#1072#1073#1086#1090#1077')'
          Width = 150
        end
        item
          CellButtons = <>
          DynProps = <>
          EditButtons = <>
          FieldName = 'otkl'
          Footers = <>
          Title.Caption = #1054#1090#1082#1083#1086#1085#1077#1085#1080#1077' '#1086#1090' '#1089#1090#1072#1085#1076#1072#1088#1090#1086#1074
          Width = 150
        end
        item
          CellButtons = <>
          DynProps = <>
          EditButtons = <>
          FieldName = 'def'
          Footers = <>
          Title.Caption = #1044#1077#1092#1077#1082#1090#1099', '#1085#1072#1088#1091#1096#1077#1085#1080#1103', '#1086#1096#1080#1073#1082#1080' '#1080' '#1076#1088'.'
          Width = 150
        end
        item
          CellButtons = <>
          DynProps = <>
          EditButtons = <>
          FieldName = 'result'
          Footers = <>
          Title.Caption = 
            #1044#1086#1089#1090#1080#1078#1077#1085#1080#1077' '#1088#1077#1079'-'#1090#1072' '#1101#1090#1072#1087#1072' '#1080#1083#1080' '#1080#1089#1093#1086#1076#1072' '#1083#1077#1095#1077#1073#1085#1086'-'#1087#1088#1086#1092#1080#1083#1072#1082#1090#1080#1095#1077#1089#1082#1086#1075#1086' '#1084#1077#1088 +
            '-'#1090#1080#1103
          Width = 150
        end
        item
          CellButtons = <>
          DynProps = <>
          EditButtons = <>
          FieldName = 'dop_pri_usl'
          Footers = <>
          Title.Caption = #1047#1072#1082#1083#1102#1095#1077#1085#1080#1077' '#1101#1082#1089#1087#1077#1088#1090#1086#1074
          Width = 150
        end
        item
          CellButtons = <>
          DynProps = <>
          EditButtons = <>
          FieldName = 'datenaprMSE'
          Footers = <>
          Title.Caption = #1044#1072#1090#1072' '#1085#1072#1087#1088#1072#1074#1083'-'#1103' '#1074' '#1052#1057#1069
        end
        item
          CellButtons = <>
          DynProps = <>
          EditButtons = <>
          FieldName = 'zaklMSE'
          Footers = <>
          Title.Caption = #1047#1072#1082#1083#1102#1095#1077#1085#1080#1077' '#1052#1057#1069
          Width = 150
        end
        item
          CellButtons = <>
          DynProps = <>
          EditButtons = <>
          FieldName = 'datezaklMSE'
          Footers = <>
          Title.Caption = #1044#1072#1090#1072' '#1087#1086#1083#1091#1095#1077#1085#1080#1103' '#1079#1072#1082#1083'-'#1103' '#1052#1057#1069
        end
        item
          CellButtons = <>
          DynProps = <>
          EditButtons = <>
          FieldName = 'dopinf'
          Footers = <>
          Title.Caption = #1044#1086#1087' '#1080#1085#1092#1086#1088#1084#1072#1094#1080#1103' '#1087#1086' '#1079#1072#1082#1083#1102#1095#1077#1085#1080#1102' '#1076#1088#1091#1075#1080#1093' '#1091#1095#1088#1077#1078#1076#1077#1085#1080#1081
          Width = 150
        end
        item
          CellButtons = <>
          DynProps = <>
          EditButtons = <>
          FieldName = 'sostexp'
          Footers = <>
          Title.Caption = #1054#1089#1085#1086#1074#1085#1086#1081' '#1089#1086#1089#1090#1072#1074' '#1101#1082#1089#1087#1077#1088#1090#1086#1074
          Width = 150
        end
        item
          CellButtons = <>
          DynProps = <>
          EditButtons = <>
          FieldName = 'podpis'
          Footers = <>
          Title.Caption = #1055#1086#1076#1087#1080#1089#1080' '#1101#1082#1089#1087#1077#1088#1090#1086#1074
          Width = 150
        end
        item
          CellButtons = <>
          DynProps = <>
          EditButtons = <>
          FieldName = 'vredn_fact'
          Footers = <>
        end>
      object RowDetailData: TRowDetailPanelControlEh
      end
    end
  end
  object Panel2: TPanel
    Left = 0
    Top = 0
    Width = 915
    Height = 73
    Align = alTop
    TabOrder = 1
    object От: TLabel
      Left = 19
      Top = 32
      Width = 12
      Height = 13
      Caption = #1086#1090
    end
    object до: TLabel
      Left = 140
      Top = 32
      Width = 13
      Height = 13
      Caption = #1076#1086
    end
    object Label1: TLabel
      Left = 19
      Top = 5
      Width = 85
      Height = 13
      Caption = #1060#1080#1083#1100#1090#1088' '#1087#1086' '#1076#1072#1090#1077':'
    end
    object DateTimePicker1: TDateTimePicker
      Left = 37
      Top = 24
      Width = 97
      Height = 21
      Date = 41117.386894953720000000
      Time = 41117.386894953720000000
      TabOrder = 0
    end
    object DateTimePicker2: TDateTimePicker
      Left = 159
      Top = 24
      Width = 98
      Height = 21
      Date = 41117.387007916670000000
      Time = 41117.387007916670000000
      TabOrder = 1
    end
    object CheckBox1: TCheckBox
      Left = 19
      Top = 51
      Width = 78
      Height = 17
      Caption = #1055#1088#1080#1084#1077#1085#1080#1090#1100
      TabOrder = 2
      OnClick = CheckBox1Click
    end
    object CheckBox2: TCheckBox
      Left = 308
      Top = 24
      Width = 153
      Height = 17
      Caption = #1055#1088#1086#1089#1084#1086#1090#1088#1077#1090#1100' '#1074#1089#1077' '#1079#1072#1087#1080#1089#1080
      TabOrder = 3
      OnClick = CheckBox2Click
    end
  end
  object MainMenu1: TMainMenu
    Left = 704
    Top = 8
    object N1: TMenuItem
      Caption = #1060#1072#1081#1083
      object Excel1: TMenuItem
        Caption = #1069#1082#1089#1087#1086#1088#1090' '#1079#1072#1087#1080#1089#1077#1081
        OnClick = Excel1Click
      end
    end
  end
end
