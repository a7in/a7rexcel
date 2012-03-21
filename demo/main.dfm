object MainForm: TMainForm
  Left = 486
  Top = 272
  Width = 284
  Height = 279
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  Menu = MainMenu
  OldCreateOrder = False
  PixelsPerInch = 96
  TextHeight = 13
  object MainMenu: TMainMenu
    Left = 24
    Top = 64
    object SimplereportMenu: TMenuItem
      Caption = 'Simple report'
      OnClick = SimplereportMenuClick
    end
  end
  object Rep: TA7Rep
    Left = 24
    Top = 120
  end
end
