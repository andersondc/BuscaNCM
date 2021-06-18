object FFinal: TFFinal
  Left = 0
  Top = 0
  Caption = 'Busca Finalizada'
  ClientHeight = 286
  ClientWidth = 538
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object PFinal: TPanel
    Left = 0
    Top = 0
    Width = 538
    Height = 286
    Align = alClient
    Color = clWhite
    ParentBackground = False
    TabOrder = 0
    ExplicitLeft = 72
    ExplicitTop = 120
    ExplicitWidth = 185
    ExplicitHeight = 41
    object Grafico: TChart
      Left = 64
      Top = 18
      Width = 400
      Height = 250
      Legend.Alignment = laBottom
      Legend.CheckBoxesStyle = cbsRadio
      Legend.TextSymbolGap = 12
      Legend.VertSpacing = -4
      Title.Text.Strings = (
        'Percentual de Eficacia na Busca')
      Chart3DPercent = 29
      TopAxis.SubAxes = <
        item
          Grid.Visible = False
          Horizontal = True
          OtherSide = True
        end>
      View3DOptions.Elevation = 315
      View3DOptions.Orthogonal = False
      View3DOptions.Perspective = 0
      View3DOptions.Rotation = 360
      TabOrder = 0
      DefaultCanvas = 'TGDIPlusCanvas'
      PrintMargins = (
        15
        18
        15
        18)
      ColorPaletteIndex = 13
      object Series1: TPieSeries
        XValues.Order = loAscending
        YValues.Name = 'Pie'
        YValues.Order = loNone
        Frame.InnerBrush.BackColor = clRed
        Frame.InnerBrush.Gradient.EndColor = clGray
        Frame.InnerBrush.Gradient.MidColor = clWhite
        Frame.InnerBrush.Gradient.StartColor = 4210752
        Frame.InnerBrush.Gradient.Visible = True
        Frame.MiddleBrush.BackColor = clYellow
        Frame.MiddleBrush.Gradient.EndColor = 8553090
        Frame.MiddleBrush.Gradient.MidColor = clWhite
        Frame.MiddleBrush.Gradient.StartColor = clGray
        Frame.MiddleBrush.Gradient.Visible = True
        Frame.OuterBrush.BackColor = clGreen
        Frame.OuterBrush.Gradient.EndColor = 4210752
        Frame.OuterBrush.Gradient.MidColor = clWhite
        Frame.OuterBrush.Gradient.StartColor = clSilver
        Frame.OuterBrush.Gradient.Visible = True
        Frame.Width = 4
        OtherSlice.Legend.Visible = False
        RotationAngle = 35
        Data = {
          04020000000000000000388E40FF020000004F4B0000000000807B40FF030000
          004EE36F}
      end
    end
  end
end
