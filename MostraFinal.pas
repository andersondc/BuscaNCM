unit MostraFinal;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.ExtCtrls, VclTee.TeeGDIPlus,
  VCLTee.TeEngine, VCLTee.Series, VCLTee.TeeProcs, VCLTee.Chart;

type
  TFFinal = class(TForm)
    PFinal: TPanel;
    Grafico: TChart;
    Series1: TPieSeries;
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  FFinal: TFFinal;

implementation

{$R *.dfm}

end.
