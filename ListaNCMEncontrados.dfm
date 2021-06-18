object FListaNCMEncontrados: TFListaNCMEncontrados
  Left = 0
  Top = 0
  Caption = 'Lista de Produtos Encontrados '
  ClientHeight = 364
  ClientWidth = 723
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  PixelsPerInch = 96
  TextHeight = 13
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 723
    Height = 89
    Align = alTop
    TabOrder = 0
    object Button1: TButton
      Left = 8
      Top = 8
      Width = 185
      Height = 75
      Caption = 'Salvar Dados Encontrados com Tabela de NCM'
      TabOrder = 0
      WordWrap = True
      OnClick = Button1Click
    end
    object bTudo: TButton
      Left = 199
      Top = 8
      Width = 185
      Height = 41
      Caption = 'Listar Tudo'
      TabOrder = 1
      WordWrap = True
      OnClick = bTudoClick
    end
  end
  object Lista: TDBGrid
    Left = 0
    Top = 89
    Width = 723
    Height = 240
    Align = alClient
    DataSource = FDados.DSQryAux1
    TabOrder = 1
    TitleFont.Charset = DEFAULT_CHARSET
    TitleFont.Color = clWindowText
    TitleFont.Height = -11
    TitleFont.Name = 'Tahoma'
    TitleFont.Style = []
    Columns = <
      item
        Alignment = taCenter
        Color = clCream
        Expanded = False
        FieldName = 'NCM'
        Font.Charset = ANSI_CHARSET
        Font.Color = clWindowText
        Font.Height = -12
        Font.Name = 'Arial Black'
        Font.Style = [fsBold]
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'PRODUTO'
        Font.Charset = ANSI_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'Arial'
        Font.Style = []
        Title.Caption = 'Nome do Produto'
        Width = 400
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'DESCRICAO'
        Font.Charset = ANSI_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'Arial'
        Font.Style = []
        Title.Caption = 'Descri'#231#227'o Oficial (Tabela de NCM))'
        Width = 500
        Visible = True
      end>
  end
  object Panel3: TPanel
    Left = 0
    Top = 329
    Width = 723
    Height = 35
    Align = alBottom
    Color = clBlack
    ParentBackground = False
    TabOrder = 2
    object Barra: TGauge
      Left = 1
      Top = 24
      Width = 721
      Height = 10
      Align = alBottom
      ForeColor = clTeal
      Progress = 0
      ShowText = False
      ExplicitLeft = 2
      ExplicitWidth = 645
    end
    object Status: TLabel
      Left = 1
      Top = 1
      Width = 721
      Height = 23
      Align = alClient
      Font.Charset = ANSI_CHARSET
      Font.Color = clWhite
      Font.Height = -11
      Font.Name = 'Arial'
      Font.Style = [fsBold]
      ParentFont = False
      Layout = tlCenter
      ExplicitWidth = 3
      ExplicitHeight = 14
    end
  end
end
