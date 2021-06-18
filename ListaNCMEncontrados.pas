unit ListaNCMEncontrados;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Data.DB, Vcl.Grids, Vcl.DBGrids,
  Vcl.ExtCtrls, Vcl.StdCtrls, Vcl.Samples.Gauges;

type
  TFListaNCMEncontrados = class(TForm)
    Panel1: TPanel;
    Lista: TDBGrid;
    Button1: TButton;
    bTudo: TButton;
    Panel3: TPanel;
    Barra: TGauge;
    Status: TLabel;
    procedure bTudoClick(Sender: TObject);
    procedure Button1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  FListaNCMEncontrados: TFListaNCMEncontrados;

implementation

{$R *.dfm}

uses Dados;

procedure TFListaNCMEncontrados.bTudoClick(Sender: TObject);
begin
  with Dados.FDados do
  begin
    QryAux1.Close;
    QryAux1.SQL.Text:=' select * from ncmEncontrados';
    QryAux1.Open;
  end;
end;

procedure TFListaNCMEncontrados.Button1Click(Sender: TObject);
var
  vdescricao: string;
  vncm: string;
begin
  with Dados.FDados do
  begin
    // Dados de Padroes para ncmEncontrados
    QryAux2.Close;
    QryAux2.SQL.Text:='select * from padroes';
    QryAux2.Open;

    QryAux2.Last;
    Barra.MaxValue:=QryAux2.RecordCount;
    Barra.Progress:=0;

    QryAux2.First;
    while not QryAux2.Eof do
    begin
      Barra.Progress:=Barra.Progress+1;
      Status.Caption:=' Cadastrando Produto: '+QryAux2.FieldByName('texto').AsString;

      QryAux.Close;
      QryAux.SQL.Text:=' select * from ncms where'+
                       ' ncm='+QryAux2.FieldByName('ncm').AsString;
      QryAux.Open;

      vncm:=trim(QryAux2.FieldByName('ncm').AsString);

      if length(vncm)<8 then
        vncm:=StringOfChar('0',8-length(vncm))+vncm;

      if QryAux.RecordCount=0 then
        vdescricao:=''
      else
        vdescricao:=QryAux.FieldByName('descricao').AsString;

      QryAux1.Close;
      QryAux1.SQL.Text:=' update or insert into ncmEncontrados ('+
                        ' ncm, produto, descricao, data_alteracao) values ('+
                        QuotedStr(vncm)+', '+
                        QuotedStr(QryAux2.FieldByName('texto').AsString)+', '+
                        QuotedStr(vdescricao)+', '+QuotedStr(DateToStr(date))+')'+
                        ' matching (produto)';
      QryAux1.ExecSQL;

      QryAux2.Next;

      lista.Update;
      lista.Repaint;

      Application.ProcessMessages;

    end;
  end;
end;

end.
