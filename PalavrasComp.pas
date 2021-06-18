unit PalavrasComp;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Data.DB, Vcl.Grids, Vcl.DBGrids,
  Vcl.ExtCtrls, Vcl.StdCtrls, Vcl.Menus;

type
  TFPalavrasComp = class(TForm)
    Panel1: TPanel;
    DBGrid1: TDBGrid;
    Label1: TLabel;
    vPalavra: TEdit;
    Label2: TLabel;
    vComp: TComboBox;
    Button1: TButton;
    vBusca: TEdit;
    Label3: TLabel;
    PopupMenu1: TPopupMenu;
    ExcluirPalavraCompatvel1: TMenuItem;
    procedure Button1Click(Sender: TObject);
    procedure vBuscaChange(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure ExcluirPalavraCompatvel1Click(Sender: TObject);
  private
    { Private declarations }
  public
    procedure CarregaNomes;
  end;

var
  FPalavrasComp: TFPalavrasComp;

implementation

{$R *.dfm}

uses Dados;

procedure TFPalavrasComp.Button1Click(Sender: TObject);
begin
  with Dados.FDados do
  begin
    QryAux1.Close;
    QryAux1.SQL.Text:=' insert into palavras (texto, compativel) values ('+
                      QuotedStr(vPalavra.Text)+', '+
                      QuotedStr(vComp.Text)+')';
    QryAux1.ExecSQL;
  end;

  vPalavra.Clear;
  vComp.Text:='';
  vPalavra.SetFocus;
end;

procedure TFPalavrasComp.vBuscaChange(Sender: TObject);
begin
  with Dados.FDados do
  begin
    QryAux1.Close;

    if Length(vBusca.Text)=0 then
      QryAux1.SQL.Text:=' select * from palavras'
    else
      QryAux1.SQL.Text:=' select * from palavras'+
                        ' where texto like +'+QuotedStr('%'+vBusca.Text+'%')+
                        ' or compativel like +'+QuotedStr('%'+vBusca.Text+'%');
    QryAux1.Open;
  end;
end;

procedure TFPalavrasComp.CarregaNomes;
begin
  vCOmp.Clear;

  with Dados.FDados do
  begin
    QryAux1.Close;
    QryAux1.SQL.Text:=' select * from palavras';
    QryAux1.Open;

    QryAux.First;
    while not QryAux1.Eof do
    begin
      if vComp.Items.IndexOf(QryAux1.FieldByName('compativel').AsString)<0 then
         vComp.Items.Add(QryAux1.FieldByName('compativel').AsString);
      QryAux1.Next;
    end;
  end;
end;

procedure TFPalavrasComp.ExcluirPalavraCompatvel1Click(Sender: TObject);
begin
  with Dados.FDados do
  begin
    QryAux2.Close;
    QryAux2.SQL.Text:=' delete from palavras'+
                      ' where texto='+
                      QuotedStr(QryAux1.FieldByName('texto').AsString)+
                      ' and compativel='+
                      QuotedStr(QryAux1.FieldByName('compativel').AsString);
    QryAux2.ExecSQL;
  end;
end;

procedure TFPalavrasComp.FormShow(Sender: TObject);
begin
  CarregaNomes;
end;

end.
