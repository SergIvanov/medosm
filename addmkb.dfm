object fAddMKB: TfAddMKB
  Left = 0
  Top = 0
  Caption = #1044#1086#1073#1072#1074#1080#1090#1100' '#1077#1097#1077' '#1086#1076#1085#1086' '#1079#1072#1073#1086#1083#1077#1074#1072#1085#1080#1077
  ClientHeight = 434
  ClientWidth = 635
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poOwnerFormCenter
  PixelsPerInch = 96
  TextHeight = 13
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 635
    Height = 434
    Align = alClient
    TabOrder = 0
    object Label1: TLabel
      Left = 24
      Top = 8
      Width = 48
      Height = 13
      Caption = #1060#1072#1084#1080#1083#1080#1103':'
    end
    object Label3: TLabel
      Left = 24
      Top = 48
      Width = 149
      Height = 13
      Caption = #1054#1088#1075#1072#1085#1080#1079#1072#1094#1080#1103' ('#1055#1088#1077#1076#1087#1088#1080#1103#1090#1080#1077'):'
    end
    object Label7: TLabel
      Left = 24
      Top = 91
      Width = 145
      Height = 13
      Caption = #1053#1072#1080#1084#1077#1085#1086#1074#1072#1085#1080#1077' '#1079#1072#1073#1086#1083#1077#1074#1072#1085#1080#1103':'
    end
    object Label8: TLabel
      Left = 160
      Top = 8
      Width = 23
      Height = 13
      Caption = #1048#1084#1103':'
    end
    object Label9: TLabel
      Left = 312
      Top = 8
      Width = 53
      Height = 13
      Caption = #1054#1090#1095#1077#1089#1090#1074#1086':'
    end
    object Label12: TLabel
      Left = 18
      Top = 213
      Width = 42
      Height = 13
      Caption = #1060#1080#1083#1100#1090#1088':'
    end
    object Label10: TLabel
      Left = 24
      Top = 137
      Width = 36
      Height = 13
      Caption = #1052#1050#1041' 10'
    end
    object Label11: TLabel
      Left = 24
      Top = 183
      Width = 97
      Height = 13
      Caption = #1042#1087#1077#1088#1074#1099#1077' '#1074#1099#1103#1074#1083#1077#1085#1086
    end
    object Label13: TLabel
      Left = 384
      Top = 183
      Width = 223
      Height = 16
      Caption = #1044#1086#1073#1072#1074#1080#1090#1100' '#1077#1097#1077' '#1086#1076#1085#1086' '#1079#1072#1073#1086#1083#1077#1074#1072#1085#1080#1077
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -13
      Font.Name = 'Tahoma'
      Font.Style = [fsBold, fsItalic]
      ParentFont = False
    end
    object Label2: TLabel
      Left = 72
      Top = 213
      Width = 17
      Height = 13
      Caption = #1044'/'#1079
    end
    object Label4: TLabel
      Left = 302
      Top = 213
      Width = 36
      Height = 13
      Caption = #1052#1050#1041' 10'
    end
    object DBEdit1: TDBEdit
      Left = 24
      Top = 24
      Width = 121
      Height = 21
      DataField = 'fam'
      DataSource = fDM.DSRab
      ReadOnly = True
      TabOrder = 0
    end
    object DBEdit2: TDBEdit
      Left = 160
      Top = 24
      Width = 129
      Height = 21
      DataField = 'imya'
      DataSource = fDM.DSRab
      ReadOnly = True
      TabOrder = 1
    end
    object DBEdit3: TDBEdit
      Left = 312
      Top = 24
      Width = 121
      Height = 21
      DataField = 'fath'
      DataSource = fDM.DSRab
      ReadOnly = True
      TabOrder = 2
    end
    object DBEdit4: TDBEdit
      Left = 24
      Top = 64
      Width = 265
      Height = 21
      DataField = 'imya_org'
      DataSource = fDM.DSOrg
      ReadOnly = True
      TabOrder = 3
    end
    object DBGrid1: TDBGrid
      Left = 24
      Top = 250
      Width = 577
      Height = 140
      DataSource = fDM.DSMKB
      TabOrder = 4
      TitleFont.Charset = DEFAULT_CHARSET
      TitleFont.Color = clWindowText
      TitleFont.Height = -11
      TitleFont.Name = 'Tahoma'
      TitleFont.Style = []
      OnDblClick = DBGrid1DblClick
    end
    object Edit2: TEdit
      Left = 96
      Top = 210
      Width = 193
      Height = 21
      TabOrder = 5
      OnChange = Edit2Change
    end
    object Button1: TButton
      Left = 24
      Top = 396
      Width = 121
      Height = 25
      Caption = #1057#1086#1093#1088#1072#1085#1080#1090#1100
      TabOrder = 6
      OnClick = Button1Click
    end
    object DBEdit8: TDBEdit
      Left = 24
      Top = 110
      Width = 561
      Height = 21
      DataField = 'mkb'
      DataSource = fDM.DSRab
      TabOrder = 7
    end
    object DBEdit9: TDBEdit
      Left = 24
      Top = 156
      Width = 369
      Height = 21
      DataField = 'mkb_code'
      DataSource = fDM.DSRab
      TabOrder = 8
    end
    object Button4: TButton
      Left = 510
      Top = 396
      Width = 75
      Height = 25
      Caption = #1042#1099#1093#1086#1076
      TabOrder = 9
      OnClick = Button4Click
    end
    object DBComboBox1: TDBComboBox
      Left = 139
      Top = 183
      Width = 89
      Height = 21
      DataField = 'firstb'
      DataSource = fDM.DSRab
      Items.Strings = (
        #1044#1072
        #1053#1077#1090)
      TabOrder = 10
    end
    object Edit1: TEdit
      Left = 344
      Top = 210
      Width = 121
      Height = 21
      TabOrder = 11
      OnChange = Edit1Change
    end
  end
end
