﻿object Form1: TForm1
  Left = 0
  Top = 0
  Caption = #1057#1086#1075#1072#1079
  ClientHeight = 304
  ClientWidth = 643
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
  object pnl1: TPanel
    Left = 0
    Top = 0
    Width = 643
    Height = 75
    Align = alTop
    TabOrder = 0
    object lbl1: TLabel
      Left = 8
      Top = 24
      Width = 12
      Height = 13
      Caption = #1086#1090
    end
    object lblдо: TLabel
      Left = 135
      Top = 24
      Width = 23
      Height = 13
      Caption = 'lbl'#1076#1086
    end
    object dtp1: TDateTimePicker
      Left = 32
      Top = 16
      Width = 97
      Height = 21
      Date = 41603.665542847220000000
      Time = 41603.665542847220000000
      TabOrder = 0
    end
    object dtp2: TDateTimePicker
      Left = 154
      Top = 16
      Width = 97
      Height = 21
      Date = 41603.665572638890000000
      Time = 41603.665572638890000000
      TabOrder = 1
    end
    object chk1: TCheckBox
      Left = 8
      Top = 43
      Width = 64
      Height = 17
      Caption = #1060#1080#1083#1100#1090#1088
      TabOrder = 2
      OnClick = chk1Click
    end
    object btn1: TButton
      Left = 544
      Top = 44
      Width = 75
      Height = 25
      Caption = #1054#1090#1095#1077#1090
      TabOrder = 3
      OnClick = btn1Click
    end
    object cbb1: TDBLookupComboboxEh
      Left = 272
      Top = 16
      Width = 253
      Height = 21
      EditButtons = <>
      KeyField = 'id_org'
      ListField = 'imya_org'
      ListSource = fDM.DSOrg
      TabOrder = 4
      Visible = True
    end
    object btn2: TButton
      Left = 96
      Top = 45
      Width = 75
      Height = 25
      Caption = #1042#1089#1077
      TabOrder = 5
      OnClick = btn2Click
    end
  end
  object pnl2: TPanel
    Left = 0
    Top = 75
    Width = 643
    Height = 229
    Align = alClient
    TabOrder = 1
    object chklst1: TCheckListBox
      Left = 1
      Top = 1
      Width = 641
      Height = 227
      Align = alClient
      ItemHeight = 13
      TabOrder = 0
    end
  end
end
