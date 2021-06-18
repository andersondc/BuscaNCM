program BuscaNCM;

uses
  Vcl.Forms,
  Principal in 'Principal.pas' {FPrincipal},
  Dados in 'Dados.pas' {FDados: TDataModule},
  PalavrasComp in 'PalavrasComp.pas' {FPalavrasComp},
  Combinatoria in 'Combinatoria.pas' {frmGenerator},
  Vcl.Themes,
  Vcl.Styles,
  ListaProdutos in 'ListaProdutos.pas' {FListaProdutos},
  MostraFinal in 'MostraFinal.pas' {FFinal},
  ListaNCMEncontrados in 'ListaNCMEncontrados.pas' {FListaNCMEncontrados};

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TFPrincipal, FPrincipal);
  Application.CreateForm(TFDados, FDados);
  Application.CreateForm(TFPalavrasComp, FPalavrasComp);
  Application.CreateForm(TfrmGenerator, frmGenerator);
  Application.CreateForm(TFListaProdutos, FListaProdutos);
  Application.CreateForm(TFFinal, FFinal);
  Application.CreateForm(TFListaNCMEncontrados, FListaNCMEncontrados);
  Application.Run;
end.
