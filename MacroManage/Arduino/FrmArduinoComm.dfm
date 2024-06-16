object Form1: TForm1
  Left = 0
  Top = 0
  Caption = 'Form1'
  ClientHeight = 206
  ClientWidth = 410
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poMainFormCenter
  OnClose = FormClose
  OnCreate = FormCreate
  PixelsPerInch = 96
  TextHeight = 13
  object Label1: TLabel
    Left = 88
    Top = 33
    Width = 54
    Height = 13
    Caption = #53685#49888#54252#53944' : '
  end
  object BitBtn1: TBitBtn
    Left = 80
    Top = 56
    Width = 113
    Height = 41
    Caption = #48372#46300#50672#44208
    TabOrder = 0
    OnClick = BitBtn1Click
  end
  object BitBtn2: TBitBtn
    Left = 199
    Top = 56
    Width = 113
    Height = 41
    Caption = #50672#44208#54644#51228
    TabOrder = 1
  end
  object BitBtn3: TBitBtn
    Left = 80
    Top = 103
    Width = 113
    Height = 41
    Caption = #53685#49888#49884#51089
    TabOrder = 2
  end
  object BitBtn4: TBitBtn
    Left = 199
    Top = 103
    Width = 113
    Height = 41
    Caption = #53685#49888#51473#51648
    TabOrder = 3
  end
  object ComComboBox1: TComComboBox
    Left = 152
    Top = 29
    Width = 145
    Height = 21
    Text = ''
    Style = csDropDownList
    ItemIndex = -1
    TabOrder = 4
  end
end
