object FPalavrasComp: TFPalavrasComp
  Left = 0
  Top = 0
  BorderStyle = bsToolWindow
  Caption = 'Palavras Compat'#237'veis'
  ClientHeight = 365
  ClientWidth = 594
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 594
    Height = 105
    Align = alTop
    TabOrder = 0
    object Label1: TLabel
      Left = 8
      Top = 8
      Width = 73
      Height = 13
      Caption = 'Palavra Usada:'
    end
    object Label2: TLabel
      Left = 8
      Top = 48
      Width = 96
      Height = 13
      Caption = 'Palavra Compat'#237'vel:'
    end
    object Label3: TLabel
      Left = 256
      Top = 42
      Width = 28
      Height = 13
      Caption = 'Busca'
    end
    object vPalavra: TEdit
      Left = 8
      Top = 22
      Width = 129
      Height = 21
      CharCase = ecUpperCase
      TabOrder = 0
    end
    object vComp: TComboBox
      Left = 8
      Top = 61
      Width = 129
      Height = 21
      CharCase = ecUpperCase
      TabOrder = 1
    end
    object Button1: TButton
      Left = 152
      Top = 59
      Width = 89
      Height = 25
      Caption = 'Cadastrar'
      TabOrder = 2
      OnClick = Button1Click
    end
    object vBusca: TEdit
      Left = 256
      Top = 61
      Width = 305
      Height = 21
      CharCase = ecUpperCase
      TabOrder = 3
      OnChange = vBuscaChange
    end
  end
  object DBGrid1: TDBGrid
    Left = 0
    Top = 105
    Width = 594
    Height = 260
    Align = alClient
    DataSource = FDados.DSQryAux1
    PopupMenu = PopupMenu1
    TabOrder = 1
    TitleFont.Charset = DEFAULT_CHARSET
    TitleFont.Color = clWindowText
    TitleFont.Height = -11
    TitleFont.Name = 'Tahoma'
    TitleFont.Style = []
    Columns = <
      item
        Expanded = False
        FieldName = 'TEXTO'
        Width = 200
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'COMPATIVEL'
        Width = 200
        Visible = True
      end>
  end
  object PopupMenu1: TPopupMenu
    Left = 472
    Top = 160
    object ExcluirPalavraCompatvel1: TMenuItem
      Caption = 'Excluir Palavra Compat'#237'vel'
      OnClick = ExcluirPalavraCompatvel1Click
    end
  end
end
