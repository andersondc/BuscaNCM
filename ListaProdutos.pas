unit ListaProdutos;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.ExtCtrls, Vcl.StdCtrls, Vcl.ComCtrls, StrUtils,
  Vcl.Grids, Vcl.Samples.Gauges, Data.DB, Vcl.DBGrids, Vcl.WinXCtrls, idHttp, IdSSLOpenSSL,
  IdIOHandler, IdIOHandlerSocket, IdIOHandlerStack, IdSSL, IdBaseComponent, IdURI, DateUtils,
  IdComponent, IdTCPConnection, IdTCPClient, Vcl.Imaging.pngimage;

type
  TFListaProdutos = class(TForm)
    POpcoes: TPanel;
    vArquivo: TEdit;
    Label1: TLabel;
    bBuscaArquivo: TButton;
    bTrataCSV: TButton;
    Panel2: TPanel;
    Panel3: TPanel;
    Pags: TPageControl;
    TabSheet1: TTabSheet;
    TabSheet2: TTabSheet;
    ListaCSV: TStringGrid;
    Status: TLabel;
    Barra: TGauge;
    TabSheet3: TTabSheet;
    Lista: TListBox;
    dbLista: TDBGrid;
    Edit1: TEdit;
    Abrir: TOpenDialog;
    vCampo: TComboBox;
    Button2: TButton;
    Label2: TLabel;
    bInicia: TButton;
    Label3: TLabel;
    lncmEncontrado: TLabel;
    Button3: TButton;
    Button4: TButton;
    bWeb: TButton;
    IdHTTP: TIdHTTP;
    IdSSLIOHandlerSocketOpenSSL1: TIdSSLIOHandlerSocketOpenSSL;
    RespM: TMemo;
    Label4: TLabel;
    nRegsNCM: TLabel;
    vCampoNCM: TComboBox;
    Label5: TLabel;
    Button5: TButton;
    Label6: TLabel;
    HoraIni: TLabel;
    HoraFim: TLabel;
    Label9: TLabel;
    ListBox1: TListBox;
    Image1: TImage;
    Label14: TLabel;
    vSemNCM: TLabel;
    Button6: TButton;
    Salva: TSaveDialog;
    PBuscar: TPanel;
    Image2: TImage;
    Label7: TLabel;
    Label10: TLabel;
    Label11: TLabel;
    Label13: TLabel;
    contaPadroes1: TLabel;
    contaPadroes2: TLabel;
    contaNCMs: TLabel;
    contaGoogle: TLabel;
    Label15: TLabel;
    contaPalavras: TLabel;
    Label12: TLabel;
    contaBing: TLabel;
    Label8: TLabel;
    contaPadroes3: TLabel;
    ativo1: TActivityIndicator;
    procedure bTrataCSVClick(Sender: TObject);
    procedure bBuscaArquivoClick(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure bIniciaClick(Sender: TObject);
    procedure dbListaDrawColumnCell(Sender: TObject; const Rect: TRect;
      DataCol: Integer; Column: TColumn; State: TGridDrawState);
    procedure Button3Click(Sender: TObject);
    procedure Button4Click(Sender: TObject);
    procedure bWebClick(Sender: TObject);
    procedure Button5Click(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure Button6Click(Sender: TObject);
  private
    { Private declarations }
  public
    procedure CarregaCSV;
    function CountPos(const subtext: string; Text: string): Integer;
    procedure ApagaTabelaProdutoTemp;
    procedure BuscaPadrao;
    procedure CadastraNCMBasico;
    function TiraAcentos(texto: string): string;
    function SemNumeros(texto: string): string;
    procedure BuscaDescricaoPadrao(texto: string);
    Procedure BuscaPadroes;
    Procedure BuscaPalavrasPadroes;
    Procedure BuscaPalavrasContagemPadrao;
    procedure SQLLista(param: integer);
    procedure BuscaInfoWeb;
    procedure BuscaInfoWeb1;
    function SomenteNumeros(const S: string): string;
    function BuscaNCM(valor: integer): integer;
    function LocalizarInfo(texto: string): string;
    function LocalizarInfo1(texto: string): string;
    function LocalizarInfoBing(texto: string): string;
    procedure CadastraMaisResultados(descricao: string; texto: string; qnt: integer);
    procedure ApagaTabelaResultadoNCM;
    procedure VerificaResultadosMais;
    procedure CadastraNCMdeTabela;
    function OccurrencesOfChar(const S: string; const C: char): integer;
    procedure ExcluiItemEncontrado(texto: string);
    procedure BuscaInfoWebBing;
    function TotalDeProdutos: integer;
    function MontaSQLPadroes: string;
  end;

var
  FListaProdutos: TFListaProdutos;
  ContaEncontrados: integer;
  NumeroRegistros: integer;
  NumeroRegistrosMais: integer;
  TotalRegistros: string;

implementation

{$R *.dfm}

uses Dados, Principal;

procedure TFListaProdutos.bBuscaArquivoClick(Sender: TObject);
begin
  Abrir.Execute;
  vArquivo.Text:=Abrir.FileName;
  RespM.Clear;
  RespM.Lines.Text:=vArquivo.Text;
  RespM.Lines.SaveToFile(Principal.vCaminho+'ArqCSV.txt');
  RespM.Clear;
end;

procedure TFListaProdutos.bTrataCSVClick(Sender: TObject);
begin
  if Length(vArquivo.Text)>0 then
  begin
    if FileExists(vArquivo.Text) then
    begin
      Status.Caption:='Carregando CSV...';
      Status.Repaint;
      Barra.Progress:=0;

      Lista.Items.LoadFromFile(vArquivo.Text);
      Barra.MaxValue:=Lista.Items.Count;

      CarregaCSV;
    end
    else
      ShowMessage('Arquivo Não Encontrado!');
  end
  else
    ShowMessage('Localize o Arquivo Antes de Carregar suas Informações!');
end;

procedure TFListaProdutos.Button2Click(Sender: TObject);
var
  i: integer;
begin
  ApagaTabelaProdutoTemp;

  with Dados.FDados do
  begin
    for i:=1 to ListaCSV.RowCount-1 do
    begin
      if ListaCSV.Cells[StrToInt(vCampo.Text),i]<>'' then
      begin
        if TiraAcentos(upperCase(ListaCSV.Cells[StrToInt(vCampo.Text),i]))<>'DESCRICAO' then
        begin
          QryProdutos.Close;
          QryProdutos.SQL.Text:=' insert into produtostemp (descricao)'+
                                ' values ('+QuotedStr(
                                            ListaCSV.Cells[StrToInt(vCampo.Text),i])+')';
          QryProdutos.ExecSQL;
        end;
      end;
    end;

    ShowMessage('Dados Salvos com Sucesso!');
    Pags.ActivePageIndex:=1;

    QryProdutos.Close;
    QryProdutos.SQL.Text:=' select * from produtostemp';
    QryProdutos.open;
  end;

end;

procedure TFListaProdutos.Button3Click(Sender: TObject);
begin
  SQLLista(1); // Sem NCM
end;

procedure TFListaProdutos.Button4Click(Sender: TObject);
begin
  SQLLista(0); // Sem NCM
end;

procedure TFListaProdutos.Button5Click(Sender: TObject);
begin
  CadastraNCMdeTabela;
end;

procedure TFListaProdutos.Button6Click(Sender: TObject);
var
  i: integer;
  j: integer;
  vncm: string;
  TextoLista: string;
begin
  if Length(vArquivo.Text)=0 then exit;

  if vCampo.Text='' then
  begin
    vCampo.SetFocus;
    exit;
  end;

  // Reabre Arquivo Original
  BTrataCSV.Click;

  // Localiza Dados no Banco e Cria Novo CSV
  Lista.Clear;
  SQLLista(0);

  Barra.Progress:=0;
  Barra.MaxValue:=ListaCSV.RowCount;

  with Dados.FDados do
  begin
    for i:=1 to ListaCSV.RowCount-1 do
    begin
      Status.Caption:='Tratando Dados do Produto '+
             ListaCSV.Cells[StrTOInt(vCampo.Text),i];
      try
        QryProdutos.Close;
        QryProdutos.SQL.Text:=' select ncm from produtostemp'+
                              ' where descricao='+
                              QuotedStr(ListaCSV.Cells[vCampo.ItemIndex,i]);
        QryProdutos.Open;
      finally
        vNCM:=QryProdutos.FieldByName('ncm').AsString;
      end;

      TextoLista:='';

      for j:=0 to ListaCSV.ColCount-1 do
      begin
        Status.Caption:='Tratando Dados: '+TextoLista;

        TextoLista:=TextoLista+ListaCSV.Cells[j,i]+';';
      end;

      TextoLista:=TextoLista+vNCM;
      Lista.Items.Add(TextoLista);

      Barra.Progress:=Barra.Progress+1;

      Application.ProcessMessages;
    end;
  end;

  Salva.Execute;
  Lista.Items.SaveToFile(Salva.FileName);

  ShowMessage('Arquivo com NCMs Encontrados Salvo em '+Salva.FileName);
end;

procedure TFListaProdutos.bWebClick(Sender: TObject);
begin
  if pBuscar.Visible=false then
     pBuscar.Visible:=true;

  bInicia.Caption:='Cancelar';

  Principal.FPrincipal.Anim.Animate:=true;
  BuscaInfoWeb;     // Google (NCM/Descricao)
//  BuscaInfoWebBing; // Bing (NCM)
  Principal.FPrincipal.Anim.Animate:=false;

//  ListBox1.Items.SaveToFile('C:\Users\DevDelphi\Desktop\Anderson D C\Estudos Delphi\NCM\WebBusca.txt');
end;

procedure TFListaProdutos.bIniciaClick(Sender: TObject);
var
  Eficacia: integer;
begin
  ContaEncontrados:=0;
  HoraFim.Caption:='-';


  Principal.FPrincipal.Anim.Animate:=true;
  Principal.FPrincipal.PListaBusca.Visible:=false;

  PBuscar.Left:=FListaProdutos.Width-PBuscar.Width-80;
  PBuscar.Top:=POpcoes.Height+50;
  PBuscar.Visible:=true;
  Ativo1.Animate:=true;

  contaPadroes1.Caption:='0 Encontrados';
  contaPadroes2.Caption:='0 Encontrados';
  contaPadroes3.Caption:='0 Encontrados';
  contancms.Caption:='0 Encontrados';
  contagoogle.Caption:='0 Encontrados';
  contabing.Caption:='0 Encontrados';

  HoraIni.Caption:=TimeTOStr(time);

  ApagaTabelaResultadoNCM;

  TotalRegistros:=IntToStr(TotalDeProdutos);


  if bInicia.Caption='Iniciar Busca' then
  begin
    bInicia.Caption:='Cancelar';
    BuscaPadroes;
    BuscaPalavrasContagemPadrao;
    BuscaPadrao;

  end
  else
  begin
    bInicia.Caption:='Iniciar Busca';
    Principal.vParar:=true;
    ShowMessage('Busca Cancelar!');
  end;

  bInicia.Caption:='Iniciar Busca';

  Principal.FPrincipal.Anim.Animate:=false;

  Eficacia:=trunc((100*ContaEncontrados)/StrToInt(TotalRegistros));

  HoraFim.Caption:=TimeTOStr(time);
  ShowMessage('Busca Concluída! Eficacia de '+IntTOStr(Eficacia)+'%');

  //  PBuscar.Visible:=false;
  Ativo1.Animate:=false;
end;

function TFListaProdutos.TotalDeProdutos: integer;
begin
  SQLLista(1);

  with Dados.FDados do
  begin
    QryProdutos.Last;
    result:=QryProdutos.RecordCount;
  end;
end;

procedure TFListaProdutos.VerificaResultadosMais;
begin
  with Dados.FDados do
  begin
    QryBusca.Close;
    QryBusca.SQL.Text:=' select * from resultadoncm';
    QryBusca.Open;

    if QryBusca.RecordCount>0 then
    begin
      Principal.FPrincipal.PListaBusca.Visible:=true;
      Principal.FPrincipal.PListaBusca.Height:=200;
    end;
  end;
end;

procedure TFListaProdutos.ApagaTabelaResultadoNCM;
begin
  with Dados.FDados do
  begin
    QryBusca.Close;
    QryBusca.SQL.Text:=' delete from resultadoncm';
    QryBusca.ExecSQL;
  end;
end;

// Funcao para Retirar Acentuação
function TFListaProdutos.TiraAcentos(texto: string): string;
begin
  texto:=UpperCase(texto);

  texto:=StringReplace(texto, 'Á', 'A',
  [rfReplaceAll, rfIgnoreCase]);
  texto:=StringReplace(texto, 'À', 'A',
  [rfReplaceAll, rfIgnoreCase]);
  texto:=StringReplace(texto, 'Ã', 'A',
  [rfReplaceAll, rfIgnoreCase]);
  texto:=StringReplace(texto, 'Â', 'A',
  [rfReplaceAll, rfIgnoreCase]);
  texto:=StringReplace(texto, 'É', 'E',
  [rfReplaceAll, rfIgnoreCase]);
  texto:=StringReplace(texto, 'È', 'E',
  [rfReplaceAll, rfIgnoreCase]);
  texto:=StringReplace(texto, 'Ê', 'E',
  [rfReplaceAll, rfIgnoreCase]);
  texto:=StringReplace(texto, 'Ç', 'C',
  [rfReplaceAll, rfIgnoreCase]);
  texto:=StringReplace(texto, 'Ú', 'U',
  [rfReplaceAll, rfIgnoreCase]);
  texto:=StringReplace(texto, 'Í', 'I',
  [rfReplaceAll, rfIgnoreCase]);
  texto:=StringReplace(texto, 'Î', 'I',
  [rfReplaceAll, rfIgnoreCase]);
  texto:=StringReplace(texto, 'Ô', 'O',
  [rfReplaceAll, rfIgnoreCase]);
  texto:=StringReplace(texto, 'Ñ', 'N',
  [rfReplaceAll, rfIgnoreCase]);
  texto:=StringReplace(texto, '''', '',
  [rfReplaceAll, rfIgnoreCase]);
  texto:=StringReplace(texto, '_', ' ',
  [rfReplaceAll, rfIgnoreCase]);
  texto:=StringReplace(texto, '<', '',
  [rfReplaceAll, rfIgnoreCase]);
  texto:=StringReplace(texto, '/>', '',
  [rfReplaceAll, rfIgnoreCase]);
  texto:=StringReplace(texto, '>', '',
  [rfReplaceAll, rfIgnoreCase]);
  texto:=StringReplace(texto, '/', '',
  [rfReplaceAll, rfIgnoreCase]);
  texto:=StringReplace(texto, '.', ' ',
  [rfReplaceAll, rfIgnoreCase]);

  result:= texto;
end;

function TFListaProdutos.SemNumeros(texto: string): string;
begin
  texto:=UpperCase(texto);

  texto:=StringReplace(texto, '0', '',
  [rfReplaceAll, rfIgnoreCase]);
  texto:=StringReplace(texto, '1', '',
  [rfReplaceAll, rfIgnoreCase]);
  texto:=StringReplace(texto, '2', '',
  [rfReplaceAll, rfIgnoreCase]);
  texto:=StringReplace(texto, '3', '',
  [rfReplaceAll, rfIgnoreCase]);
  texto:=StringReplace(texto, '4', '',
  [rfReplaceAll, rfIgnoreCase]);
  texto:=StringReplace(texto, '5', '',
  [rfReplaceAll, rfIgnoreCase]);
  texto:=StringReplace(texto, '6', '',
  [rfReplaceAll, rfIgnoreCase]);
  texto:=StringReplace(texto, '7', '',
  [rfReplaceAll, rfIgnoreCase]);
  texto:=StringReplace(texto, '8', '',
  [rfReplaceAll, rfIgnoreCase]);
  texto:=StringReplace(texto, '9', '',
  [rfReplaceAll, rfIgnoreCase]);
  texto:=StringReplace(texto, '.', ' ',
  [rfReplaceAll, rfIgnoreCase]);

  result:= texto;
end;

procedure TFListaProdutos.SQLLista(param: integer);
begin
  // param= 0: tudo, 1: sem ncm

  with Dados.FDados do
  begin
    QryProdutos.Close;

    if param=0 then
      QryProdutos.SQL.Text:=' select * from produtostemp'
    else
      QryProdutos.SQL.Text:=' select * from produtostemp where ncm is null';

    QryProdutos.open;
    QryProdutos.Last;
    if param=1 then vSemNCM.Caption:=QryProdutos.RecordCount.ToString;
    QryProdutos.First;
  end;
end;

procedure TFListaProdutos.CadastraMaisResultados
                          (descricao: string; texto: string; qnt: integer);
begin
  with Dados.FDados do
  begin
    QryBusca.Close;
    QryBusca.SQL.Text:=' insert into resultadoncm'+
                       ' (nome_original, texto, qnt) values ('+
                       QuotedStr(descricao)+', '+
                       QuotedStr(texto)+', '+
                       IntTOStr(qnt)+')';
    QryBusca.ExecSQL;

    NumeroRegistrosMais:=NumeroRegistrosMais+1;
    nRegsNCM.Caption:=IntToStr(NumeroRegistrosMais);
  end;
end;

procedure TFListaProdutos.CadastraNCMdeTabela;
var
  i: integer;
  contador: integer;

  ncmAux: string;
begin
  Barra.MaxValue:=ListaCSV.RowCount;
  Barra.Progress:=0;

  with Dados.FDados do
  begin
    for i:=1 to ListaCSV.RowCount-1 do
    begin
      if ListaCSV.Cells[StrToInt(vCampo.Text),i]<>'' then
      begin
       // Insere NCM Padrao baseado na pesquisa
        ncmAux:=SomenteNumeros(ListaCSV.Cells[StrToInt(vCampoNCM.Text),i]);

        if (length(trim(ncmAux))>0) and (trim(ncmAux)<>'0') then
        begin

          QryAux2.Close;
          QryAux2.SQL.Text:=' insert into padroes (ncm, texto) values ('+
                            ncmAux+', '+
                            QuotedStr(trim(TiraAcentos(UpperCase(ListaCSV.Cells[StrToInt(vCampo.Text),i]))))+')';
          QryAux2.ExecSQL;

          Status.Caption:='Cadastrando: '+ListaCSV.Cells[StrToInt(vCampo.Text),i];
          contador:=contador+1;
          Barra.Progress:=Barra.Progress+1;
          Application.ProcessMessages;

        end;
      end;
    end;

    ShowMessage('Dados Padrões Salvos com Sucesso!');

  end;
end;

procedure TFListaProdutos.BuscaPadrao;
var
  x: integer;
begin
  x:=0;

  with Dados.FDados do
  begin
//    ContaEncontrados:=0;

    SQLLista(1); // (sem NCM)
    ativo1.Top:=105;

    Barra.Progress:=0;

    QryProdutos.First;

    while not QryProdutos.Eof do
    begin
      if Principal.vParar=True then
      begin
        Principal.vParar:=false;
        exit;
      end;

      Barra.MaxValue:=(StrToInt(vSemNCM.Caption));
      NumeroRegistros:=Barra.MaxValue;

      Application.ProcessMessages;

//      BuscaDescricaoPadrao(QryProdutos.FieldByName('descricao').AsString);

//      if QryAux.RecordCount=1 then
//        CadastraNCMBasico
//      else
//      begin
        Principal.FPrincipal.CamNCM.Text:=
                  TiraAcentos(QryProdutos.FieldByName('descricao').AsString);

        Principal.FPrincipal.buscaProduto;

        STatus.Caption:='Buscando Informações - '+QryProdutos.FieldByName('descricao').AsString;

        if QryAUx.RecordCount=1 then
        begin
          x:=x+1;
          CadastraNCMBasico;
        end
        else
        begin
          try
            Principal.FPrincipal.bBuscar.Click;
          finally
            if QryAUx.RecordCount=1 then
            begin
              if Principal.FPrincipal.ListaPalavras.Count>1 then
              begin
                x:=x+1;
                CadastraNCMBasico;
              end;
            end;
          end;
        end;

        lncmEncontrado.Caption:=IntToStr(ContaEncontrados)+'/'+TotalRegistros;

        Application.ProcessMessages;

        Barra.Progress:=Barra.Progress+1;
//      end;

      contancms.Caption:=IntToStr(x)+' Encontrado(s)';
      QryProdutos.Next;
    end;

  end;

    // Localizando na Web Produtos que não encontrou (Google)
  bWeb.Click;


  // Verifica se Há Mais Resultados
  VerificaResultadosMais;
end;

Procedure TFListaProdutos.BuscaPadroes;
var
  x: integer;
begin
  x:=0;

  SQLLista(1); // Sem NCM
  ativo1.Top:=10;

  Barra.Progress:=0;

  with Dados.FDados do
  begin
    while not QryProdutos.Eof do
    begin
      if Principal.vParar=True then
      begin
        Principal.vParar:=false;
        exit;
      end;

      Barra.MaxValue:=(StrToInt(vSemNCM.Caption));
      NumeroRegistros:=Barra.MaxValue;

      STatus.Caption:='Buscando Informações em Padrões - '+QryProdutos.FieldByName('descricao').AsString;
      Barra.Progress:=Barra.Progress+1;

      Application.ProcessMessages;

      QryAux2.Close;
      QryAux2.SQL.Text:=' select first 1 * from padroes'+
                        ' where upper(texto) like '+
                        QuotedStr(TiraAcentos(UpperCase(QryProdutos.FieldByName('descricao').AsString)+'%'));
      QryAux2.Open;

//      if QryAux2.RecordCount=0 then
//      begin
//        QryAux2.Close;
//        QryAux2.SQL.Text:=' select first 1 * from padroes'+
//                          ' where upper(texto) like '+
//                          QuotedStr(TiraAcentos(UpperCase(QryProdutos.FieldByName('descricao').AsString)+'%'));
//        QryAux2.Open;
//      end;

      if QryAux2.RecordCount=1 then
      begin
        x:=x+1;

        QryAux1.Close;
        QryAux1.SQL.Text:=' update produtostemp set'+
                          ' ncm='+QryAux2.FieldByName('ncm').AsString+
                          ' where descricao='+
                          QuotedStr(trim(QryProdutos.FieldByName('descricao').AsString));
        QryAux1.ExecSQL;

        ContaEncontrados:=ContaEncontrados+1;
        lncmEncontrado.Caption:=IntToStr(ContaEncontrados)+'/'+TotalRegistros;
      end;

      contaPadroes1.Caption:=IntToStr(x)+' Encontrado(s)';

      QryProdutos.Next;
    end;
  end;

end;

Procedure TFListaProdutos.BuscaPalavrasContagemPadrao;
var
  i: integer;
  Qnt: integer;
  Texto: string;
  vSql: string;
  x: integer;
begin
  SQLLista(1); // Sem NCM

  ativo1.Top:=42;

  Qnt:=0;
  Texto:='';
  x:=0;

  with principal.FPrincipal do
  begin
    PMontaBusca.Height:=300;
    PMontaBusca.Visible:=true;
  end;

  Barra.Progress:=0;

  with Dados.FDados do
  begin
    while not QryProdutos.Eof do
    begin
      Qnt:=0;

      if Principal.vParar=True then
      begin
        Principal.vParar:=false;
        exit;
      end;

      Barra.MaxValue:=(StrToInt(vSemNCM.Caption));
      NumeroRegistros:=Barra.MaxValue;

      Barra.Progress:=Barra.Progress+1;

      Application.ProcessMessages;

      Principal.FPrincipal.explodeBusca(tiraAcentos(QryProdutos.FieldByName('descricao').AsString));
      Principal.FPrincipal.explodeBusca(Principal.FPrincipal.EliminaPalavrasPadroes);

      for i:=0 to Principal.FPrincipal.ListaCombinacoes.Count-1 do
      begin
        Principal.FPrincipal.explodeBusca(
                  trim(Principal.FPrincipal.ListaCombinacoes.Items.Strings[i]));

        if Principal.FPrincipal.ListaPalavras.Count>1 then
        begin
          if length(trim(Principal.FPrincipal.ListaCombinacoes.Items.Strings[i]))>4 then
          begin
            STatus.Caption:='Buscando Palavras em Padrões por Quantidade de Ocorrências - '+
                             Principal.FPrincipal.ListaCombinacoes.Items.Strings[i];

            QryAux.Close;
            QryAux.SQL.Text:=Principal.FPrincipal.MontaSQLPalavrasPadroes;
//            QryAux.SQL.Text:=' select first 1 p.ncm, count(p.ncm) as total'+
//                             ' from padroes p where p.texto like '+
//                             QuotedStr('%'+Principal.FPrincipal.ListaPalavras.Items.Strings[i]+'%')+
//                            ' group by p.ncm order by total desc';
            QryAux.Open;

            Application.ProcessMessages;

            if QryAux.FieldByName('total').AsInteger>Qnt then
            begin
              Qnt:=QryAux.FieldByName('total').AsInteger;
              Texto:=Principal.FPrincipal.ListaCombinacoes.Items.Strings[i];
              vSQL:=QryAux.SQL.Text;
            end;
          end;
        end;
      end;

      if qnt>0 then
      begin
        try
          QryAux.Close;
          QryAux.SQL.Text:=vSQL;
          QryAux.Open;
        finally
          x:=x+1;
          CadastraNCMBasico;
          lncmEncontrado.Caption:=IntToStr(ContaEncontrados)+'/'+TotalRegistros;
        end;
      end;

      if ativo1.Top=169 then
        contaPalavras.Caption:=IntToStr(x)+' Encontrado(s)'
      else
        contaPadroes2.Caption:=IntToStr(x)+' Encontrado(s)';

      QryProdutos.Next;
    end;
  end;

  Principal.FPrincipal.PMontaBusca.Visible:=false;
end;

Procedure TFListaProdutos.BuscaPalavrasPadroes;
var
  x: integer;
begin
  SQLLista(1); // Sem NCM
  ativo1.Top:=94;
  x:=0;

  with principal.FPrincipal do
  begin
    PMontaBusca.Height:=300;
    PMontaBusca.Visible:=true;
  end;

  Barra.Progress:=0;

  with Dados.FDados do
  begin
    while not QryProdutos.Eof do
    begin
      if Principal.vParar=True then
      begin
        Principal.vParar:=false;
        exit;
      end;

      Barra.MaxValue:=(StrToInt(vSemNCM.Caption));
      NumeroRegistros:=Barra.MaxValue;

      STatus.Caption:='Buscando Palavras em Padrões - '+QryProdutos.FieldByName('descricao').AsString;
      Barra.Progress:=Barra.Progress+1;

      Application.ProcessMessages;

      // Explode Busca para Procurar no Banco de Padroes
      try
        Principal.FPrincipal.explodeBusca(
                  TiraAcentos(
                  upperCase(
                  QryProdutos.FieldByName('descricao').AsString)));
      finally
        Principal.FPrincipal.PesquisaPadroes(
                  Principal.FPrincipal.EliminaPalavrasPadroes);
      end;

      if QryAux.RecordCount=1 then
      begin
        if QryAux.FieldByName('ncm').AsInteger<>0 then
        begin
          x:=x+1;
          QryAux1.Close;
          QryAux1.SQL.Text:=' update produtostemp set'+
                            ' ncm='+QryAux.FieldByName('ncm').AsString+
                            ' where descricao='+
                            QuotedStr(QryProdutos.FieldByName('descricao').AsString);
          QryAux1.ExecSQL;

          // Insere NCM Padrao baseado na pesquisa
          QryAux2.Close;
          QryAux2.SQL.Text:=' insert into padroes (ncm, texto) values ('+
                            QryAux.FieldByName('ncm').AsString+', '+
                            QuotedStr(TiraAcentos(UpperCase(QryProdutos.FieldByName('descricao').AsString)))+')';
          QryAux2.ExecSQL;

          ContaEncontrados:=ContaEncontrados+1;

          lncmEncontrado.Caption:=IntToStr(ContaEncontrados)+'/'+TotalRegistros;
        end;
      end;

      contaPadroes3.Caption:=IntToStr(x)+' Encontrado(s)';
      QryProdutos.Next;
    end;
  end;

  Principal.FPrincipal.PMontaBusca.Visible:=false;
end;

function TFListaProdutos.MontaSQLPadroes: string;
var
  auxSQL: string;
  auxLike: string;
  i: integer;

  auxPadrao: string;
  auxNCM: integer;
begin
  auxPadrao:='';

  auxSQL:='select ncm from padroes where ';

  auxLike:='';
  for i:=0 to Principal.FPrincipal.ListaPalavras.Count-1 do
  begin
    auxLike:=auxLike+' (texto LIKE '+
                      QuotedStr('%'+Principal.FPrincipal.ListaPalavras.Items.Strings[i]+'%')+
                     ') AND';
    auxPadrao:=auxPadrao+' '+Principal.FPrincipal.ListaPalavras.Items.Strings[i];
  end;

  auxLike:=LeftStr(auxLike, Length(auxLike)-3);

  result:=auxSQL+auxLike+' group by ncm';
end;

procedure TFListaProdutos.BuscaDescricaoPadrao(texto: string);
begin
  exit;
  with Dados.FDados do
  begin
    QryAux.Close;
    QryAux.SQL.Text:=' select first 1 * from ncms n'+
                     ' inner join padroes p on p.texto like '+QuotedStr('%'+texto+'%')+
                     ' where n.ncm=p.ncm';
    QryAux.Open;
  end;
end;

procedure TFListaProdutos.BuscaInfoWeb;
var
  IdW: TIDHttp;
  auxTexto: string;
  parteNumerica: string;
  parteTexto: string;
  RespWeb: string;
  tentativa: integer;
  auxTexto1: string;

  x: integer;
begin
  // Lista somente sem NCM
  SQLLista(1);
  ativo1.Top:=141;
  x:=0;

  with Dados.FDados do
  begin

    QryProdutos.First;
    Barra.Progress:=0;

    while not QryProdutos.Eof do
    begin
      if Principal.vParar=True then
      begin
        Principal.vParar:=false;
        exit;
      end;

      Barra.MaxValue:=(StrTOInt(vSEMNCM.Caption));

      NumeroRegistros:=barra.MaxValue;

      auxTexto:=QryProdutos.FieldByName('descricao').AsString;

      auxTexto:=TIdUri.URLEncode('https://www.google.com/search?q=ncm+categoria+'+
                                 auxTexto);
      status.Caption:='Buscando Informacoes na Web (GOOGLE) do Produto: '+
                    QryProdutos.FieldByName('descricao').AsString+' - '+auxTexto;

//      vArquivo.Text:=auxTexto;

      RespM.Clear;

      IdW:=TIDHttp.Create(nil);

      lncmEncontrado.Caption:=IntToStr(ContaEncontrados)+'/'+TotalRegistros;

      Application.ProcessMessages;

      ListBox1.Items.Add('https://www.google.com/search?q=ncm+categoria+'+auxTexto);

      try
        tentativa:=0;
        while tentativa<5 do
        begin
          RespM.Lines.Text:=IdW.Get(auxTexto);
//          RespM.Lines.SaveToFile('C:\Desenvolvimento\XMLNCM\'+
//                                  QryProdutos.FieldByName('descricao').AsString+'_Google.xml');

          auxTexto1:=LocalizarInfo(RespM.Lines.Text);
          if LeftStr(UpperCase(auxTexto),3)<>'TML' then
            tentativa:=10;

          tentativa:=tentativa+1;

          Application.ProcessMessages;
        end;
      except
        auxTexto1:='';
      end;

      auxTexto:=auxTexto1;
      ListBox1.Items.Add(auxTexto);

      Application.ProcessMessages;

      if leftStr(trim(auxTexto),3)<>'tml' then
      begin
        parteNumerica:=SomenteNumeros(LeftStr(auxTexto,15));
        parteTexto:=StringReplace(RightStr(auxTexto,length(auxTexto)-10),'Outros - Pagina', '',
                    [rfReplaceAll, rfIgnoreCase]);
        parteTexto:=StringReplace(parteTexto,'- Cosmos<', '',
                    [rfReplaceAll, rfIgnoreCase]);
        parteTexto:=TiraAcentos(parteTexto);
        parteTexto:=SemNumeros(parteTexto);
        parteTexto:=StringReplace(parteTexto,'DIVH', '',
                    [rfReplaceAll, rfIgnoreCase]);
        parteTexto:=StringReplace(parteTexto,'LANG="PT-BR"HEADMETA CHARSET="UTF-"', '',
                    [rfReplaceAll, rfIgnoreCase]);

        if Length(trim(parteNumerica))>8 then parteNumerica:=leftStr(parteNumerica,8);

        if parteNumerica='' then parteNumerica:='0';

        status.Caption:='Pesquisando no Banco as Informações: '+
                         parteNumerica+' e '+parteTexto;

        ListBox1.Items.Add(parteNumerica+' / '+parteTexto);
        ListBox1.Items.SaveToFile('C:\Users\DevDelphi\Desktop\Anderson D C\Estudos Delphi\NCM\ResultWeb\Google.txt');

        // Busca Primeiro no Banco NCMs pela parte numerica encontrada
        if BuscaNCM(StrToInt(parteNumerica))=1 then
        begin
          CadastraNCMBasico;
          x:=x+1;
        end
        else
        begin   // Caso não encontre parte numerica, procura pelo texto
          Principal.FPrincipal.CamNCM.Text:=parteTexto;
          Principal.FPrincipal.BuscaProduto;

          if QryAux.RecordCount=1 then
          begin
            CadastraNCMBasico;
            x:=x+1;
          end
          else
          begin
            try
              Principal.FPrincipal.bBuscar.Click;
            finally
              if QryAux.RecordCount=1 then
              begin
                CadastraNCMBasico;
                x:=x+1;
              end;
             end;
           end;
        end;
      end;

      Application.ProcessMessages;

      idW.Free;

      barra.Progress:=barra.Progress+1;
      contaGoogle.Caption:=IntToStr(x)+' Encontrado(s)';

      QryProdutos.Next;
    end;
  end;
end;

procedure TFListaProdutos.BuscaInfoWeb1;
var
  IdW: TIDHttp;
  auxTexto: string;
  parteNumerica: string;
  parteTexto: string;
  RespWeb: string;
  tentativa: integer;
  auxTexto1: string;

  x: integer;
begin
  // Lista somente sem NCM
  SQLLista(1);
  ativo1.Top:=169;
  x:=0;

  with Dados.FDados do
  begin

    QryProdutos.First;
    Barra.Progress:=0;

    while not QryProdutos.Eof do
    begin
      if Principal.vParar=True then
      begin
        Principal.vParar:=false;
        exit;
      end;

      Barra.MaxValue:=StrToInt(TotalRegistros);

      NumeroRegistros:=barra.MaxValue;

      auxTexto:=QryProdutos.FieldByName('descricao').AsString;

      status.Caption:='Buscando Informacoes na Web (GOOGLE) - 2º Opção - do Produto: '+
                    QryProdutos.FieldByName('descricao').AsString+' - '+auxTexto;

      RespM.Clear;

      IdW:=TIDHttp.Create(nil);

      lncmEncontrado.Caption:=IntToStr(ContaEncontrados)+'/'+TotalRegistros;

      Application.ProcessMessages;

      try
        tentativa:=0;
        while tentativa<2 do
        begin
          RespM.Lines.Text:=IdW.Get(auxTexto);

          auxTexto1:=LocalizarInfo1(RespM.Lines.Text);
          if pos('lang',auxTexto)<0 then tentativa:=10
          else tentativa:=tentativa+1;

          Application.ProcessMessages;
        end;
      except
        auxTexto1:='';
      end;

      auxTexto:=auxTexto1;
      ListBox1.Items.Add(auxTexto);

      Application.ProcessMessages;

      if leftStr(trim(auxTexto),4)<>'lang' then
      begin
        parteNumerica:=SomenteNumeros(LeftStr(auxTexto,10));
        parteTexto:=StringReplace(RightStr(auxTexto,length(auxTexto)-10),'Outros - Pagina', '',
                    [rfReplaceAll, rfIgnoreCase]);
        parteTexto:=StringReplace(parteTexto,'- Cosmos<', '',
                    [rfReplaceAll, rfIgnoreCase]);

        if Length(trim(parteNumerica))>8 then parteNumerica:=leftStr(parteNumerica,8);

        if parteNumerica='' then parteNumerica:='0';

        status.Caption:='Pesquisando no Banco as Informações: '+
                         parteNumerica+' e '+parteTexto;


        // Busca Primeiro no Banco NCMs pela parte numerica encontrada
        if BuscaNCM(StrToInt(parteNumerica))=1 then
        begin
          CadastraNCMBasico;
          x:=x+1;
        end
        else
        begin   // Caso não encontre parte numerica, procura pelo texto
          Principal.FPrincipal.CamNCM.Text:=parteTexto;
          Principal.FPrincipal.BuscaProduto;

          if QryAux.RecordCount=1 then
          begin
            CadastraNCMBasico;
            x:=x+1;
          end
          else
          begin
            try
              Principal.FPrincipal.bBuscar.Click;
            finally
              if QryAux.RecordCount=1 then
              begin
                CadastraNCMBasico;
                x:=x+1;
              end;
             end;
           end;
        end;
      end;

      Application.ProcessMessages;

      idW.Free;

      barra.Progress:=barra.Progress+1;

      QryProdutos.Next;
    end;
  end;
end;

procedure TFListaProdutos.BuscaInfoWebBing;
var
  IdW: TIDHttp;
  auxTexto: string;
  auxTexto1: string;
  parteNumerica: string;
  RespWeb: string;
  Tentativa: integer;
  x: integer;
begin
  // Lista somente sem NCM
  SQLLista(1);
  ativo1.Top:=201;
  x:=0;

  with Dados.FDados do
  begin
    QryProdutos.First;
    Barra.Progress:=0;

    while not QryProdutos.Eof do
    begin
      if Principal.vParar=True then
      begin
        Principal.vParar:=false;
        exit;
      end;

      Barra.MaxValue:=(StrTOInt(vSEMNCM.Caption));

      NumeroRegistros:=barra.MaxValue;

      auxTexto:=QryProdutos.FieldByName('descricao').AsString;

      auxTexto:=TIdUri.URLEncode('https://www.bing.com/search?q=ncm+'+
                                 trim(auxTexto));
      status.Caption:='Buscando Informacoes na Web (BING) do Produto: '+
                    QryProdutos.FieldByName('descricao').AsString+' - '+auxTexto;

      ListBox1.Items.Add('https://www.bing.com/search?q=ncm+'+
                                 trim(auxTexto));

      RespM.Clear;

      IdW:=TIDHttp.Create(nil);

      lncmEncontrado.Caption:=IntToStr(ContaEncontrados)+'/'+TotalRegistros;

      Application.ProcessMessages;

      try
        tentativa:=0;
        while tentativa<3 do
        begin
          auxTexto1:=LocalizarInfoBing(IdW.Get(auxTexto));

          if pos('xml',auxTexto)<0 then tentativa:=10
          else tentativa:=tentativa+1;

          Application.ProcessMessages;

        end;
      except
        auxTexto1:='';
      end;

      auxTexto:=auxTexto1;

      ListBox1.Items.Add(auxTexto);

      parteNumerica:=SomenteNumeros(LeftStr(auxTexto,10));

      if Length(trim(parteNumerica))>8 then parteNumerica:=RightStr(parteNumerica,8);
      if parteNumerica='' then parteNumerica:='0';

      status.Caption:='Pesquisando no Banco as Informações: NCM '+
                       parteNumerica;

      ListBox1.Items.Add(parteNumerica);
      ListBox1.Items.SaveToFile('C:\Users\DevDelphi\Desktop\Anderson D C\Estudos Delphi\NCM\ResultWeb\Bing.txt');

      Application.ProcessMessages;

      // Busca no Banco NCMs pela parte numerica encontrada
      if BuscaNCM(StrToInt(parteNumerica))=1 then
      begin
        CadastraNCMBasico;
        x:=x+1;
      end;

      idW.Free;

      barra.Progress:=barra.Progress+1;
      contaBing.Caption:=IntToStr(x)+' Encontrado(s)';

      QryProdutos.Next;
    end;
  end;

end;

function TFListaProdutos.LocalizarInfo(texto: string): string;
var
  S : string;
  Startpos : integer;
  FSelPos: integer;

  PalavraChave: string;
  Localizado: string;
begin
  RespM.Lines.Text:=texto;
  result:=Copy(RespM.Lines.Text,Pos('Produtos da NCM',texto)+12,50);
end;

function TFListaProdutos.LocalizarInfo1(texto: string): string;
var
  S : string;
  Startpos : integer;
  FSelPos: integer;

  PalavraChave: string;
  Localizado: string;
begin
  RespM.Lines.Text:=texto;      //  '12345678901NCM:'
  result:=Copy(RespM.Lines.Text,Pos('NCM:',texto)+33,40);
end;

function TFListaProdutos.LocalizarInfoBing(texto: string): string;
var
  S : string;
  Startpos : integer;
  FSelPos: integer;

  PalavraChave: string;
  Localizado: string;
begin
  RespM.Lines.Text:=texto;
  result:=Copy(RespM.Lines.Text,Pos('<strong>NCM</strong>&#32',texto)+38,10);
end;

function TFListaProdutos.BuscaNCM(valor: integer): integer;
begin
  with Dados.FDados do
  begin
    QryAux.Close;
    QryAux.SQL.Text:='select * from ncms where ncm like '+QuotedStr('%'+IntToStr(valor)+'%');
    QryAux.Open;

    result:=QryAUx.RecordCount;
  end;
end;

procedure TFListaProdutos.ExcluiItemEncontrado(texto: string);
var
  ValAnt: integer;
begin
  with Dados.FDados do
  begin
    QryAux.Close;
    QryAux.SQL.Text:=' select * from resultadoncm';
    QryAux.Open;

    ValAnt:=QryAux.RecordCount;

    QryAux.Close;
    QryAux.SQL.Text:=' delete from resultadoncm'+
                     ' where nome_original='+QuotedStr('%'+trim(texto)+'%');
    QryAux.ExecSQL;

    QryAux.Close;
    QryAux.SQL.Text:=' select * from resultadoncm';
    QryAux.Open;

    if QryAux.RecordCount<ValAnt then
    begin
      NumeroRegistrosMais:=NumeroRegistrosMais-1;
      nRegsNCM.Caption:=IntToStr(NumeroRegistrosMais);
    end;
  end;
end;

procedure TFListaProdutos.FormShow(Sender: TObject);
begin
  Self.Left:=Screen.Monitors[0].Width+100;
  Self.WindowState:=wsMaximized;
end;

function TFListaProdutos.SomenteNumeros(const S: string): string;
var
  vText : PChar;
begin
  vText := PChar(S);
  Result := '';

  while (vText^ <> #0) do
  begin
    {$IFDEF UNICODE}
    if CharInSet(vText^, ['0'..'9']) then
    {$ELSE}
    if vText^ in ['0'..'9'] then
    {$ENDIF}
      Result := Result + vText^;

    Inc(vText);
  end;
end;

procedure TFListaProdutos.CadastraNCMBasico;
begin
  with Dados.FDados do
  begin
    if QryAux.FieldByName('ncm').AsInteger<>0 then
    begin
      QryAux2.Close;
      QryAux2.SQL.Text:=' update produtostemp set'+
                        ' ncm='+QryAux.FieldByName('ncm').AsString+
                        ' where descricao='+
                        QuotedStr(QryProdutos.FieldByName('descricao').AsString);
      QryAux2.ExecSQL;

      // Insere NCM Padrao baseado na pesquisa
      QryAux2.Close;
      QryAux2.SQL.Text:=' insert into padroes (ncm, texto) values ('+
                        QryAux.FieldByName('ncm').AsString+', '+
                        QuotedStr(TiraAcentos(UpperCase(QryProdutos.FieldByName('descricao').AsString)))+')';
      QryAux2.ExecSQL;

      ContaEncontrados:=ContaEncontrados+1;
      ExcluiItemEncontrado(QryProdutos.FieldByName('descricao').AsString);
    end;
  end;
end;

function TFListaProdutos.OccurrencesOfChar(const S: string; const C: char): integer;
var
  i: Integer;
begin
  result := 0;
  for i := 1 to Length(S) do
    if S[i] = C then
      inc(result);
end;

procedure TFListaProdutos.CarregaCSV;
var
  i: integer;
  j: integer;

  totalCampos: integer;

  valor: string;

  vNCM: string;
  vClassificacao: string;
  vDescricao: string;

  vAntCampo: integer;
  vAntNCM: integer;
begin
  vAntCampo:=vCampo.ItemIndex;
  vAntNCM:=vCampoNCM.ItemIndex;

  ListaCSV.RowCount:=1;
  vCampo.Clear;

//  totalCampos:=OccurrencesOfChar(';',edit1.Text);

  totalCampos:=CountPos(';',Lista.Items.Strings[1])+1;
  ListaCSV.ColCount:=totalCampos;

  for j:=0 to totalCampos-1 do
  begin
    ListaCSV.Cells[j,0]:=IntTOStr(j);

    if vCampo.Items.IndexOf(IntTOStr(j))<0 then
      vCampo.Items.Add(IntTOStr(j));
  end;

  vCampoNCM.Items:=vCampo.Items;

  for i:=0 to lista.Items.Count-1 do
  begin
    Barra.Progress:=i;

    edit1.Text:=Lista.Items.Strings[i]+';';
    Status.Caption:='Tratando: '+edit1.Text;
    Status.Repaint;

//    totalCampos:=CountPos(';',edit1.Text);
//    ListaCSV.ColCount:=totalCampos;

    for j:=0 to totalCampos-1 do
    begin
      valor:=LeftStr(edit1.Text,Pos(';',edit1.Text));
      edit1.Text:=RightStr(edit1.Text,Length(edit1.Text)-Pos(';',edit1.Text));

      valor:=StringReplace(valor,'"', '',[rfReplaceAll, rfIgnoreCase]);
      valor:=LeftStr(valor,Length(valor)-1);

      ListaCSV.Cells[j,ListaCSV.RowCount]:=valor;
    end;

    ListaCSV.RowCount:=ListaCSV.RowCount+1;

    Application.ProcessMessages;
  end;

  vCampo.ItemIndex:=vAntCampo;
  vCampoNCM.ItemIndex:=vAntNCM;


  Status.Caption:='';
  Barra.Progress:=0;

  Pags.ActivePageIndex:=0;

  edit1.Clear;
end;

procedure TFListaProdutos.ApagaTabelaProdutoTemp;
begin
  with Dados.FDados do
  begin
    QryProdutos.Close;
    QryProdutos.SQL.Text:=' delete from produtostemp';
    QryProdutos.ExecSQL;
  end;
end;

function TFListaProdutos.CountPos(const subtext: string; Text: string): Integer;
begin
  if (Length(subtext) = 0) or (Length(Text) = 0) or (Pos(subtext, Text) = 0)
  then
    Result := 0
  else
    Result := (Length(Text) - Length(StringReplace(Text, subtext, '',
      [rfReplaceAll]))) div Length(subtext);
end;

procedure TFListaProdutos.dbListaDrawColumnCell(Sender: TObject;
  const Rect: TRect; DataCol: Integer; Column: TColumn; State: TGridDrawState);
begin
  with Dados.FDados do
  begin
    If QryProdutos.FieldByName('ncm').AsInteger>0 then
      dbLista.Canvas.Brush.Color:=clSilver;

    dbLista.DefaultDrawDataCell(Rect, dbLista.columns[datacol].field, State);
  end;
end;

end.
