unit Principal;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Data.DB, Vcl.StdCtrls, Vcl.Grids, StrUtils, PalavrasComp,
  Vcl.DBGrids, Vcl.ExtCtrls, Vcl.Samples.Gauges, System.Types, Vcl.Menus, Combinatoria,
  Vcl.WinXCtrls, ListaProdutos, Vcl.Imaging.pngimage, MostraFinal, ListaNCMEncontrados;

type
  TFPrincipal = class(TForm)
    Lista: TListBox;
    pPrincipal: TPanel;
    bRecNCM: TButton;
    Panel2: TPanel;
    ListaNCM: TStringGrid;
    Panel4: TPanel;
    dbLista: TDBGrid;
    Panel3: TPanel;
    Barra: TGauge;
    Status: TLabel;
    ListaPalavras: TListBox;
    MainMenu1: TMainMenu;
    Cadastro1: TMenuItem;
    PalavrasCompatveis1: TMenuItem;
    NCMsIgnorados1: TMenuItem;
    N1: TMenuItem;
    RecarregarNCMs1: TMenuItem;
    PalavrasAux: TListBox;
    Edit1: TEdit;
    PopupMenu1: TPopupMenu;
    AdicionarcomoNCMPadroparaessaPesquisa1: TMenuItem;
    N2: TMenuItem;
    AdicionarcomoNCMnoaceito1: TMenuItem;
    PMontaBusca: TPanel;
    ListSequencia: TListBox;
    ListaCombinacoes: TListBox;
    AuxPalavras2: TListBox;
    CarregarListadeProdutos1: TMenuItem;
    CarregarProdutosparaBusca1: TMenuItem;
    PListaBusca: TPanel;
    Panel5: TPanel;
    Label1: TLabel;
    dbListaBusca: TDBGrid;
    Label2: TLabel;
    PNCMManual: TPanel;
    Label3: TLabel;
    vNCM: TEdit;
    bSalvarNCM: TButton;
    Button2: TButton;
    Button3: TButton;
    Panel6: TPanel;
    CamNCM: TEdit;
    bBuscar: TButton;
    Image1: TImage;
    Image2: TImage;
    Image3: TImage;
    Anim: TActivityIndicator;
    AdicionarNCMnaListadeProdutosItemSelecionado1: TMenuItem;
    Button1: TButton;
    Button4: TButton;
    procedure bRecNCMClick(Sender: TObject);
    procedure CamNCMChange(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure PalavrasCompatveis1Click(Sender: TObject);
    procedure bBuscarClick(Sender: TObject);
    procedure CamNCMKeyPress(Sender: TObject; var Key: Char);
    procedure RecarregarNCMs1Click(Sender: TObject);
    procedure AdicionarcomoNCMPadroparaessaPesquisa1Click(Sender: TObject);
    procedure bSalvarNCMClick(Sender: TObject);
    procedure CarregarProdutosparaBusca1Click(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure dbListaBuscaDblClick(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure vNCMKeyPress(Sender: TObject; var Key: Char);
    procedure Image2Click(Sender: TObject);
    procedure Image1Click(Sender: TObject);
    procedure FormResize(Sender: TObject);
    procedure Image3Click(Sender: TObject);
    procedure AdicionarNCMnaListadeProdutosItemSelecionado1Click(
      Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure Button4Click(Sender: TObject);
  private
    { Private declarations }
  public
    procedure CarregaListaNCM;
    procedure explodeBusca(valor: string);
    function MontaSQL: string;
    function MontaSQLPalavrasPadroes: string;
    function PalavraCompativel(texto: string): string;
    function EliminaPalavras: string;
    function EliminaPalavrasPadroes: string;
    function EliminaPalavrasIndiv: string;
    function NumRegs(texto: string): integer;
    function NumRegsPadroes(texto: string): integer;
    procedure Pesquisa(texto: string);
    procedure PesquisaPadroes(texto: string);
    procedure PesquisaPalavrasPadroes(texto: string);
    function AjustaTexto(texto: string): string;
    procedure SelecionaNCMPadrao(ncm: integer);
    function BuscaNCMPadrao(texto: string): integer;
    function VerificaCombinacoes: string;
    function VerificaCombinacoesPadroes: string;
    function SubDividir1(texto: string; indice: integer): string;
    procedure BuscaProduto;
    procedure ApagaTabelaResultadoNCM;
  end;

var
  FPrincipal: TFPrincipal;
  vOpcoes: integer;
  vTextoSalva: string;
  vParar: boolean;
  vCaminho: string;

implementation

{$R *.dfm}

uses Dados;

procedure TFPrincipal.bBuscarClick(Sender: TObject);
var
  TextoPadrao: string;
begin
  PListaBusca.Visible:=false;

  PMontaBusca.Height:=300;
  PMontaBusca.Visible:=true;

  if Dados.FDados.QryAux.RecordCount=0 then
  begin
    Application.ProcessMessages;

    TextoPadrao:=EliminaPalavras;
    Pesquisa(' '+TextoPadrao);
  end;

  Barra.Progress:=0;
  Status.Caption:='';

  PMontaBusca.Visible:=false;
end;

procedure TFPrincipal.bRecNCMClick(Sender: TObject);
begin
  Status.Caption:='Carregando Lista Inicial de NCMs';
  Status.Repaint;
  Barra.Progress:=0;
  Lista.Items.LoadFromFile('C:\Users\DevDelphi\Desktop\Anderson D C\Estudos Delphi\NCM\Tabela-NCM.csv');
  Barra.MaxValue:=Lista.Items.Count;

  CarregaListaNCM;
end;

procedure TFPrincipal.SelecionaNCMPadrao(ncm: integer);
var
  i: integer;
  auxTexto: string;
  auxNCM: string;
begin
  // Verifica Palavra de Busca
//  for i:=0 to ListaPalavras.Items.Count-1 do
//  begin
//    auxTexto:=auxTexto+' '+ListaPalavras.Items.Strings[i];
//  end;

  with Dados.FDados do
  begin
    // Insere NCM Padrao baseado na pesquisa
//    QryAux2.Close;
//    QryAux2.SQL.Text:=' insert into padroes (ncm, texto) values ('+
//                      IntToStr(ncm)+', '+
//                      QuotedStr(auxTexto)+')';
//    QryAux2.ExecSQL;

    QryAux2.Close;
    QryAux2.SQL.Text:=' insert into padroes (ncm, texto) values ('+
                      IntToStr(ncm)+', '+
                      QuotedStr(CamNCM.Text)+')';
    QryAux2.ExecSQL;

    // Cadastra no Banco Padroes e ProdutosTemp caso haja NCM selecionado
    if QryBusca.RecordCount>0 then
    begin
      QryAux2.Close;
      QryAux2.SQL.Text:=' insert into padroes (ncm, texto) values ('+
                        IntToStr(ncm)+', '+
                        QuotedStr(QryBusca.FieldByName('nome_original').AsString)+')';
      QryAux2.ExecSQL;

      QryAux2.Close;
      QryAux2.SQL.Text:=' update produtostemp set'+
                        ' ncm='+QryAux.FieldByName('ncm').AsString+
                        ' where descricao='+
                        QuotedStr(QryBusca.FieldByName('nome_original').AsString);
      QryAux2.ExecSQL;
    end;

//    // Insere outros NCMs da Pesquisa como NAOACEITOS
//    QryAux.First;
//    while not QryAUx.Eof do
//    begin
//      auxNCM:=auxNCM+QryAux.FieldByName('ncm').AsString+',';
//      QryAux.Next;
//    end;

//    QryAux2.Close;
//    QryAux2.SQL.Text:=' insert into naoaceitos (texto, ncm) values ('+
//                      QuotedStr(auxTexto)+', '+
//                      QuotedStr(auxNCM)+')';
//    QryAux2.ExecSQL;
  end;
end;

procedure TFPrincipal.explodeBusca(valor: string);
var
  i: integer;
  auxPalavra: string;
  Palavras: TStringDynArray;

  auxCar: string;
begin
  auxCar:=' ';

  listaPalavras.Clear;

  auxPalavra:=StringReplace(valor, '+', ' ', [rfReplaceAll]);
  Palavras:=SplitString(auxPalavra, auxCar);

  for i:=0 to length(Palavras)-1 do
  begin
    try
      if length(trim(Palavras[i]))>0 then
      begin
        if ListaPalavras.Items.IndexOf(Palavras[i])<0 then
        begin
//          ListaPalavras.Items.Add(Palavras[i]);
          PalavraCompativel(Palavras[i]);
        end;
      end;
    except
      exit;
    end;
  end;
end;

procedure TFPrincipal.FormCreate(Sender: TObject);
begin
  vOpcoes:=3;
//  vCaminho:='C:\Users\DevDelphi\Desktop\Anderson D C\Estudos Delphi\NCM\';
  vCaminho:=ExtractFileDir(GetCurrentDir);
end;

procedure TFPrincipal.FormResize(Sender: TObject);
begin
  Anim.Left:=pPrincipal.Width-Anim.Width-30;
end;

procedure TFPrincipal.FormShow(Sender: TObject);
begin
  vParar:=false;

  with Dados.FDados do
  begin
    QryAux.Close;
    QryAux.SQL.Text:='select * from ncms';
    QryAux.Open;
  end;

  if FileExists(Principal.vCaminho+'ArqCSV.txt') then
  begin
    ListaProdutos.FListaProdutos.RespM.Lines.
                  LoadFromFile(Principal.vCaminho+'ArqCSV.txt');
    ListaProdutos.FListaProdutos.vArquivo.Text:=
         ListaProdutos.FListaProdutos.RespM.Lines.Text;
    ListaProdutos.FListaProdutos.RespM.Clear;
  end;
end;

procedure TFPrincipal.Image1Click(Sender: TObject);
begin
  PalavrasComp.FPalavrasComp.ShowModal;
end;

procedure TFPrincipal.Image2Click(Sender: TObject);
begin
  ListaProdutos.FListaProdutos.Show;
end;

procedure TFPrincipal.Image3Click(Sender: TObject);
begin
  MostraFInal.FFinal.ShowModal;
end;

function TFPrincipal.PalavraCompativel(texto: string): string;
begin
  result:=texto;

  texto:=StringReplace(texto,'Ô', 'O',[rfReplaceAll, rfIgnoreCase]);
  texto:=StringReplace(texto,'-', '',[rfReplaceAll, rfIgnoreCase]);
  texto:=StringReplace(texto,'(', '',[rfReplaceAll, rfIgnoreCase]);
  texto:=StringReplace(texto,')', '',[rfReplaceAll, rfIgnoreCase]);
  texto:=StringReplace(texto,'COM GAS', 'C/GAS',[rfReplaceAll, rfIgnoreCase]);
  texto:=StringReplace(texto,'SEM GAS', 'S/GAS',[rfReplaceAll, rfIgnoreCase]);

  if Length(trim(texto))<3 then
  begin
    result:='';
    exit;
  end;

  if length(texto)>0 then
  begin
    with Dados.FDados do
    begin
      QryAux1.Close;
      QryAux1.SQL.Text:=' select * from palavras where'+
                        ' texto ='+QuotedStr(trim(texto));
      QryAux1.Open;

      if QryAux1.RecordCount>0 then
      begin
        QryAux1.First;
        while not QryAux1.Eof do
        begin
          result:=QryAux1.FieldByName('compativel').AsString;
          if ListaPalavras.Items.IndexOf(QryAux1.FieldByName('compativel').AsString)<0 then
             ListaPalavras.Items.Add(QryAux1.FieldByName('compativel').AsString);
          QryAux1.Next;
        end;
      end
      else
        ListaPalavras.Items.Add(texto);
    end;
  end;
end;

procedure TFPrincipal.CamNCMChange(Sender: TObject);
begin
  BuscaProduto;
end;

procedure TFPrincipal.BuscaProduto;
begin
  Pesquisa(' '+AjustaTexto(CamNCM.Text));    //+' '+CamNCM.Text);
end;

procedure TFPrincipal.AdicionarcomoNCMPadroparaessaPesquisa1Click(
  Sender: TObject);
begin
  with Dados.FDados do
  begin
    SelecionaNCMPadrao(QryAux.FieldByName('ncm').AsInteger);
  end;
end;

procedure TFPrincipal.AdicionarNCMnaListadeProdutosItemSelecionado1Click(
  Sender: TObject);
begin
  with Dados.FDados do
  begin
    QryAux2.Close;
    QryAux2.SQL.Text:=' update produtostemp set'+
                      ' ncm='+QryAux.FieldByName('ncm').AsString+
                      ' where descricao='+
                      QuotedStr(QryProdutos.FieldByName('descricao').AsString);
    QryAux2.ExecSQL;

    QryProdutos.Refresh;

    QryAux2.Close;
    QryAux2.SQL.Text:=' insert into padroes (ncm, texto) values ('+
                      QryAux.FieldByName('ncm').AsString+', '+
                      QuotedStr(ListaProdutos.FListaProdutos.TiraAcentos
                               (UpperCase(QryProdutos.FieldByName('descricao').
                               AsString)))+')';
    QryAux2.ExecSQL;

  end;
end;

function TFPrincipal.AjustaTexto(texto: string): string;
var
  i: integer;
  auxTexto: string;
begin
  texto:=ListaProdutos.FListaProdutos.TiraAcentos(UpperCase(texto));

  explodeBusca(texto);

  PalavrasAux.Clear;

  for i:=0 to listaPalavras.Count-1 do
  begin
    auxTexto:=auxTexto+' '+LeftStr(ListaPalavras.Items.Strings[i],6);
  end;

  ListaPalavras.Clear;
  result:=auxTexto;
end;

procedure TFPrincipal.CamNCMKeyPress(Sender: TObject; var Key: Char);
begin
  if (Key = #13) then
  begin
    bBuscar.Click;
  end;
end;

procedure TFPrincipal.Pesquisa(texto: string);
var
  vSQL: string;
begin
  with Dados.FDados do
  begin
    explodeBusca(texto);

    try
      if Length(trim(CamNCM.Text))=0 then
         vSQL:=' select * from ncms'
       else
         vSQL:=MontaSQL;

      QryAux.Close;
      QryAux.SQL.Text:=vSQL;
      QryAux.Open;

    except
      //
    end;
  end;
end;

procedure TFPrincipal.PesquisaPalavrasPadroes(texto: string);
var
  vSQL: string;
begin
  with Dados.FDados do
  begin
    explodeBusca(texto);

    vSQL:=MontaSQL;

    QryAux.Close;
    QryAux.SQL.Text:=vSQL;
    QryAux.Open;

  end;
end;

procedure TFPrincipal.PesquisaPadroes(texto: string);
var
  vSQL: string;
begin
  with Dados.FDados do
  begin
    explodeBusca(texto);

    try
      if Length(trim(CamNCM.Text))=0 then
         vSQL:=' select ncm from padroes'
       else
         vSQL:=ListaProdutos.FListaProdutos.MontaSQLPadroes;

      QryAux.Close;
      QryAux.SQL.Text:=vSQL;
      QryAux.Open;

    except
      //
    end;
  end;
end;

procedure TFPrincipal.RecarregarNCMs1Click(Sender: TObject);
begin
  bRecNCM.Click;
end;

function TFPrincipal.NumRegs(texto: string): integer;
begin
  with Dados.FDados do
  begin
    camNCM.Text:=texto;
    buscaProduto;

    if QryAUx.RecordCount=1 then
      result:=1
    else
    begin
      explodeBusca(texto);

      QryAux.Close;
      QryAux.SQL.Text:=MontaSQL;
      QryAux.Open;

      Application.ProcessMessages;

      result:=QryAUx.RecordCount;
    end;
  end;
end;

function TFPrincipal.NumRegsPadroes(texto: string): integer;
begin
  with Dados.FDados do
  begin
    camNCM.Text:=texto;
    explodeBusca(texto);

    QryAux.Close;
    QryAux.SQL.Text:=ListaProdutos.FListaProdutos.MontaSQLPadroes;
    QryAux.Open;

    Application.ProcessMessages;

    result:=QryAUx.RecordCount;
  end;
end;

function TFPrincipal.MontaSQL: string;
var
  auxSQL: string;
  auxLike: string;
  i: integer;

  auxPadrao: string;
  auxNCM: integer;
begin
  auxPadrao:='';

  auxSQL:='select * from ncms where ';

  auxLike:='';
  for i:=0 to ListaPalavras.Count-1 do
  begin
    auxLike:=auxLike+' (descricao LIKE '+
                      QuotedStr('%'+ListaPalavras.Items.Strings[i]+'%')+
                     ' or classificacao like '+
                      QuotedStr('%'+ListaPalavras.Items.Strings[i]+'%')+') AND';
    auxPadrao:=auxPadrao+' '+ListaPalavras.Items.Strings[i];
  end;

//  if ListaPalavras.Count>1 then
    auxLike:=LeftStr(auxLike, Length(auxLike)-3);
//  else
//    auxLike:='1=2';

  result:=auxSQL+auxLike;
end;

function TFPrincipal.MontaSQLPalavrasPadroes: string;
var
  auxSQL: string;
  auxLike: string;
  i: integer;

  auxPadrao: string;
  auxNCM: integer;
begin
  auxPadrao:='';

  auxSQL:=' select first 1 p.ncm, count(p.ncm) as total'+
          ' from padroes p where ';

  auxLike:='';
  for i:=0 to ListaPalavras.Count-1 do
  begin
    auxLike:=auxLike+' (p.texto LIKE '+
                      QuotedStr('%'+ListaPalavras.Items.Strings[i]+'%')+') AND';
    auxPadrao:=auxPadrao+' '+ListaPalavras.Items.Strings[i];
  end;

  auxLike:=LeftStr(auxLike, Length(auxLike)-3);

  result:=auxSQL+auxLike+
          ' group by p.ncm order by total desc';
end;

function TFPrincipal.BuscaNCMPadrao(texto: string): integer;
begin
  with Dados.FDados do
  begin
    QryAux2.Close;
    QryAux2.SQL.Text:=' select first 1 * from padroes where texto='+QuotedStr(texto);
    QryAux2.Open;

    if QryAux2.RecordCount>0 then
      result:=QryAux2.FieldByName('ncm').AsInteger
    else
      result:=0;
  end;
end;

procedure TFPrincipal.bSalvarNCMClick(Sender: TObject);
begin
  with Dados.FDados do
  begin
    QryAux2.Close;
    QryAux2.SQL.Text:=' update produtostemp set'+
                      ' ncm='+trim(vNCM.Text)+
                      ' where descricao='+
                      QuotedStr(QryBusca.FieldByName('nome_original').AsString);
    QryAux2.ExecSQL;

    // Insere NCM Padrao baseado na pesquisa
    QryAux2.Close;
    QryAux2.SQL.Text:=' insert into padroes (ncm, texto) values ('+
                      trim(vNCM.Text)+', '+
                      QuotedStr(ListaProdutos.FListaProdutos.
                                TiraAcentos(UpperCase(
                                QryBusca.FieldByName('nome_original').AsString)))+')';
    QryAux2.ExecSQL;
  end;

  PNCMManual.Visible:=false;
end;

procedure TFPrincipal.Button1Click(Sender: TObject);
begin
  with Dados.FDados do
  begin
    QryAux.Close;
    QryAux.SQL.Text:=' delete from padroes';
    QryAux.ExecSQL;
  end;
end;

procedure TFPrincipal.Button2Click(Sender: TObject);
begin
  PNCMManual.Visible:=false;
end;

procedure TFPrincipal.Button3Click(Sender: TObject);
begin
  PNCMManual.Visible:=true;
end;

procedure TFPrincipal.Button4Click(Sender: TObject);
begin
  ListaNCMEncontrados.FListaNCMEncontrados.ShowModal;
end;

procedure TFPrincipal.PalavrasCompatveis1Click(Sender: TObject);
begin
  PalavrasComp.FPalavrasComp.ShowModal;
end;

function TFPrincipal.EliminaPalavras: string;
var
  i: integer;
  j: integer;
  k: integer;
  l: integer;
  m: integer;
  texto: string;
  textoPadrao: string;
  total: integer;
  qntCalc: integer;
  qntAux: integer;
begin
  status.Caption:='Montando Combinações...';

  PalavrasAux.Clear;
  PalavrasAux.Items:=ListaPalavras.Items;
  ListaPalavras.Clear;

  Barra.Progress:=0;

  AuxPalavras2.Clear;
  AuxPalavras2.Items:=ListaPalavras.Items;

  ListaCombinacoes.Clear;
  ListaCombinacoes.Items:=AuxPalavras2.Items;

  ListSequencia.Clear;

  total:=0;

  for i:=0 to PalavrasAux.Count-1 do
  begin
    texto:=texto+' '+PalavrasAux.Items.Strings[i];
  end;

  ListSequencia.Items.Add('Palavra completa: '+texto);
  ListSequencia.Items.Add('---------------------------------------------------------------------------------');

  texto:=trim(texto);
  if ListaCombinacoes.Items.IndexOf(texto)<0 then ListaCombinacoes.Items.Add(texto);

  // Verifica retirando uma palavra por pesquisa
  for j:=0 to PalavrasAux.Count-1 do
  begin
    texto:='';

    for i:=0 to PalavrasAux.Count-1 do
    begin
      if j<>i then
          texto:=texto+' '+PalavrasAux.Items.Strings[i];
    end;

    texto:=trim(texto);
    ListSequencia.Items.Add('1 palavra a menos: '+texto);

    if ListaCombinacoes.Items.IndexOf(texto)<0 then ListaCombinacoes.Items.Add(texto);
  end;

  ListSequencia.Items.Add('---------------------------------------------------------------------------------');

  // Verifica retirando outras palavra por pesquisa
  m:=0;

  for m:=0 to PalavrasAux.Count-1 do
  begin
    for k:=0 to PalavrasAux.Count-1 do
    begin
      for j:=0 to PalavrasAux.Count-1 do
      begin
        texto:='';

        for i:=0 to PalavrasAux.Count-1 do
        begin
          if j<>i then
          begin
            Barra.MaxValue:=PalavrasAux.Count;
            Barra.Progress:=i;

            if (i<>k) and (i<>m) then
              texto:=texto+' '+PalavrasAux.Items.Strings[i];

            Application.ProcessMessages;
          end;
        end;

        texto:=trim(texto);
//        ListSequencia.Items.Add('x palavras menos: '+texto);
        if ListaCombinacoes.Items.IndexOf(texto)<0 then ListaCombinacoes.Items.Add(texto);

        texto:=SubDividir1(texto, j);
        texto:=trim(texto);
        ListSequencia.Items.Add(IntTOStr(j)+' - Ajustando Palavra: '+texto);
        if ListaCombinacoes.Items.IndexOf(texto)<0 then ListaCombinacoes.Items.Add(texto);

      end;
      ListSequencia.Items.Add('- Indice da 1º Palavra: '+IntTOStr(k)+' / Indice da 2º Palavra: '+IntToStr(m));

      AuxPalavras2.Clear;

      for i:=0 to PalavrasAux.Items.Count-1 do
      begin
        if i<>k then
          if AuxPalavras2.Items.IndexOf(PalavrasAux.Items.Strings[i])<0 then
           AuxPalavras2.Items.Add(PalavrasAux.Items.Strings[i]);
      end;

      PalavrasAux.Clear;
      PalavrasAux.Items:=AuxPalavras2.Items;

      ListSequencia.Items.Add('- Qnt de Palavras à Combinar: '+PalavrasAux.Items.Count.ToString);
      ListSequencia.Items.Add('---------------------------------------------------------------------------------');

    end;

    Application.ProcessMessages;
  end;

  ListSequencia.Items.Add('---------------------------------------------------------------------------------');

  texto:=VerificaCombinacoes;
  qntAux:=Dados.FDados.QryAux.RecordCount;

  if qntAux>1 then
  begin
    ListaProdutos.FListaProdutos.CadastraMaisResultados(
                  Dados.FDados.QryProdutos.FieldByName('descricao').AsString,
                  texto, qntAux);
  end;

  if texto<>'' then
    result:=texto
  else
    result:=EliminaPalavrasIndiv;
end;

function TFPrincipal.EliminaPalavrasPadroes: string;
var
  i: integer;
  j: integer;
  k: integer;
  l: integer;
  m: integer;
  texto: string;
  textoPadrao: string;
  total: integer;
  qntCalc: integer;
  qntAux: integer;
begin
  status.Caption:='Montando Combinações...';

  PalavrasAux.Clear;
  PalavrasAux.Items:=ListaPalavras.Items;
  ListaPalavras.Clear;

  Barra.Progress:=0;

  AuxPalavras2.Clear;
  AuxPalavras2.Items:=ListaPalavras.Items;

  ListaCombinacoes.Clear;
  ListaCombinacoes.Items:=AuxPalavras2.Items;

  ListSequencia.Clear;

  total:=0;

  for i:=0 to PalavrasAux.Count-1 do
  begin
    texto:=texto+' '+PalavrasAux.Items.Strings[i];
  end;

  ListSequencia.Items.Add('Palavra completa: '+texto);
  ListSequencia.Items.Add('---------------------------------------------------------------------------------');

  texto:=trim(texto);
  if ListaCombinacoes.Items.IndexOf(texto)<0 then ListaCombinacoes.Items.Add(texto);

  // Verifica retirando uma palavra por pesquisa
  for j:=0 to PalavrasAux.Count-1 do
  begin
    texto:='';

    for i:=0 to PalavrasAux.Count-1 do
    begin
      if j<>i then
          texto:=texto+' '+PalavrasAux.Items.Strings[i];
    end;

    texto:=trim(texto);
    ListSequencia.Items.Add('1 palavra a menos: '+texto);

    if ListaCombinacoes.Items.IndexOf(texto)<0 then ListaCombinacoes.Items.Add(texto);
  end;

  ListSequencia.Items.Add('---------------------------------------------------------------------------------');

  // Verifica retirando outras palavra por pesquisa
  m:=0;

  for m:=0 to PalavrasAux.Count-1 do
  begin
    for k:=0 to PalavrasAux.Count-1 do
    begin
      for j:=0 to PalavrasAux.Count-1 do
      begin
        texto:='';

        for i:=0 to PalavrasAux.Count-1 do
        begin
          if j<>i then
          begin
            Barra.MaxValue:=PalavrasAux.Count;
            Barra.Progress:=i;

            if (i<>k) and (i<>m) then
              texto:=texto+' '+PalavrasAux.Items.Strings[i];

            Application.ProcessMessages;
          end;
        end;

        texto:=trim(texto);
        if ListaCombinacoes.Items.IndexOf(texto)<0 then ListaCombinacoes.Items.Add(texto);

        texto:=SubDividir1(texto, j);
        texto:=trim(texto);
        ListSequencia.Items.Add(IntTOStr(j)+' - Ajustando Palavra: '+texto);
        if ListaCombinacoes.Items.IndexOf(texto)<0 then ListaCombinacoes.Items.Add(texto);

      end;
      ListSequencia.Items.Add('- Indice da 1º Palavra: '+IntTOStr(k)+' / Indice da 2º Palavra: '+IntToStr(m));

      AuxPalavras2.Clear;

      for i:=0 to PalavrasAux.Items.Count-1 do
      begin
        if i<>k then
          if AuxPalavras2.Items.IndexOf(PalavrasAux.Items.Strings[i])<0 then
           AuxPalavras2.Items.Add(PalavrasAux.Items.Strings[i]);
      end;

      PalavrasAux.Clear;
      PalavrasAux.Items:=AuxPalavras2.Items;

      ListSequencia.Items.Add('- Qnt de Palavras à Combinar: '+PalavrasAux.Items.Count.ToString);
      ListSequencia.Items.Add('---------------------------------------------------------------------------------');

    end;

    Application.ProcessMessages;
  end;

  ListSequencia.Items.Add('---------------------------------------------------------------------------------');

  texto:=VerificaCombinacoesPadroes;
  qntAux:=Dados.FDados.QryAux.RecordCount;

  if texto<>'' then
    result:=texto
  else
    result:=Dados.FDados.QryProdutos.FieldByName('descricao').AsString;
end;

procedure TFPrincipal.ApagaTabelaResultadoNCM;
begin
  with Dados.FDados do
  begin
    QryBusca.Close;
    QryBusca.SQL.Text:=' delete from resultadoncm';
    QryBusca.ExecSQL;
  end;
end;

function TFPrincipal.SubDividir1(texto: string; indice: integer): string;
var
  i: integer;
begin
  ListaPalavras.Clear;
  ExplodeBusca(texto);

  texto:='';

  for i:=0 to ListaPalavras.Items.Count-1 do
  begin
    if i<>indice then
    begin
      texto:=texto+' '+ListaPalavras.Items.Strings[i];
    end;
  end;

  result:=texto;
end;

function TFPrincipal.VerificaCombinacoes: string;
var
  i: integer;
  qntCalc: integer;
  texto: string;
  auxValor: integer;
begin
  qntCalc:=0;
  text:='';

  Barra.Progress:=0;

  AuxPalavras2.Clear;
  for i:=0 to listaCombinacoes.Items.Count-1 do
  begin
    if length(listaCombinacoes.Items.Strings[i])>0 then
      AuxPalavras2.Items.Add(listaCombinacoes.Items.Strings[i]);
  end;

  listaCombinacoes.Clear;
  listaCombinacoes.Items:=AuxPalavras2.Items;

  Barra.MaxValue:=listaCombinacoes.Items.Count;

  for i:=0 to listaCombinacoes.Items.Count-1 do
  begin
    Barra.Progress:=i;
    if ListaCombinacoes.Items.Strings[i]<>'' then
    begin
      ListaCombinacoes.ItemIndex:=i;

      auxValor:=NumRegs(ListaCombinacoes.Items.Strings[i]);

      status.Caption:='Consultando Palavra(s) Chave(s): '+ListaCombinacoes.Items.Strings[i];

      if auxValor>0 then
      begin
        if qntCalc=0 then qntCalc:=auxValor;

        if auxValor<=qntCalc then
        begin
          texto:=ListaCombinacoes.Items.Strings[i];
          qntCalc:=auxValor;

        end;

        if qntCalc=1 then
        begin
          result:=texto;
          exit;
        end;

      end;
    end;
  end;

  result:=texto;
end;

function TFPrincipal.VerificaCombinacoesPadroes: string;
var
  i: integer;
  qntCalc: integer;
  texto: string;
  auxValor: integer;
begin
  qntCalc:=0;
  text:='';

  Barra.Progress:=0;

  AuxPalavras2.Clear;
  for i:=0 to listaCombinacoes.Items.Count-1 do
  begin
    if length(listaCombinacoes.Items.Strings[i])>0 then
      AuxPalavras2.Items.Add(listaCombinacoes.Items.Strings[i]);
  end;

  listaCombinacoes.Clear;
  listaCombinacoes.Items:=AuxPalavras2.Items;

  Barra.MaxValue:=listaCombinacoes.Items.Count;

  for i:=0 to listaCombinacoes.Items.Count-1 do
  begin
    Barra.Progress:=i;
    if ListaCombinacoes.Items.Strings[i]<>'' then
    begin
      ListaCombinacoes.ItemIndex:=i;

      auxValor:=NumRegsPadroes(ListaCombinacoes.Items.Strings[i]);

      status.Caption:='Consultando Palavra(s) Chave(s): '+ListaCombinacoes.Items.Strings[i];

      if auxValor>0 then
      begin
        if qntCalc=0 then qntCalc:=auxValor;

        if auxValor<=qntCalc then
        begin
          texto:=ListaCombinacoes.Items.Strings[i];
          qntCalc:=auxValor;
        end;

        if qntCalc=1 then
        begin
          result:=texto;
          exit;
        end;

      end;
    end;
  end;

  result:=texto;
end;

procedure TFPrincipal.vNCMKeyPress(Sender: TObject; var Key: Char);
begin
  if (Key = #13) then
  begin
    bSalvarNCM.Click;
  end;
end;

function TFPrincipal.EliminaPalavrasIndiv: string;
var
  i: integer;
  texto: string;
  textoPadrao: string;
  total: integer;
  qntCalc: integer;
begin
///  PalavrasAux.Clear;
//  PalavrasAux.Items:=ListaPalavras.Items;
  ListaPalavras.Clear;

  total:=0;

  for i:=0 to PalavrasAux.Count-1 do
  begin
    texto:=PalavrasAux.Items.Strings[i];

    qntCalc:=NumRegs(texto);
    textoPadrao:=texto;
//    total:=qntCalc;

    texto:=trim(texto);
    ListSequencia.Items.Add('Individual: '+texto);
    if ListaCombinacoes.Items.IndexOf(texto)<0 then ListaCombinacoes.Items.Add(texto);

    if qntCalc>total then
    begin
      textoPadrao:=texto;
      total:=qntCalc;
    end;

    Application.ProcessMessages;
  end;

  ListSequencia.Items.Add('---------------------------');

  result:=textoPadrao;
end;

procedure TFPrincipal.CarregaListaNCM;
var
  i: integer;
  j: integer;

  valor: string;

  vNCM: string;
  vClassificacao: string;
  vDescricao: string;
begin
  edit1.Clear;

  for i:=0 to lista.Items.Count-1 do
  begin
    Barra.Progress:=i;

    edit1.Text:=Lista.Items.Strings[i];
    Status.Caption:='Tratando: '+edit1.Text;
    Status.Repaint;

    for j:=0 to 2 do
    begin
      valor:=LeftStr(edit1.Text,Pos(';',edit1.Text));
      edit1.Text:=RightStr(edit1.Text,Length(edit1.Text)-Pos(';',edit1.Text));

      valor:=StringReplace(valor,'"', '',[rfReplaceAll, rfIgnoreCase]);
      valor:=LeftStr(valor,Length(valor)-1);

      case j of
        0: vNCM:=valor;
        1: vClassificacao:=valor;
        2: vDescricao:=RightStr(valor,Length(valor)-11);
      end;
    end;

    Dados.FDados.QryAux.Close;
    Dados.FDados.QryAux.SQL.Text:=' update or insert into ncms (ncm, classificacao, descricao)'+
                                  ' values('+
                                  vncm+', '+
                                  QuotedStr(leftStr(vClassificacao,200))+', '+
                                  QuotedStr(leftStr(vDescricao,400))+')'+
                                  ' matching (ncm)';
    Dados.FDados.QryAux.ExecSQL;

    Application.ProcessMessages;
  end;

  Status.Caption:='';
  Barra.Progress:=0;

  edit1.Clear;
end;


procedure TFPrincipal.CarregarProdutosparaBusca1Click(Sender: TObject);
begin
  ListaProdutos.FListaProdutos.Show;
end;

procedure TFPrincipal.dbListaBuscaDblClick(Sender: TObject);
begin
  CamNCM.Text:=Dados.FDados.QryBusca.FieldByName('texto').AsString;
  BuscaProduto;
end;

end.
