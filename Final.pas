unit Final;

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
    procedure FormShow(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  FFinal: TFFinal;

implementation

{$R *.dfm}

procedure TFFinal.FormShow(Sender: TObject);
begin
  Grafico.Series[0].ValuesList[0].Value:=80;
  Grafico.Series[0].ValuesList[1].Value:=20;
end;

end.
