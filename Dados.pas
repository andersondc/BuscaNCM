unit Dados;

interface

uses
  System.SysUtils, System.Classes, Data.DB, Datasnap.DBClient,
  FireDAC.Phys.FBDef, FireDAC.Stan.Intf, FireDAC.Stan.Option,
  FireDAC.Stan.Error, FireDAC.UI.Intf, FireDAC.Phys.Intf, FireDAC.Stan.Def,
  FireDAC.Stan.Pool, FireDAC.Stan.Async, FireDAC.Phys, FireDAC.Phys.FB,
  FireDAC.VCLUI.Wait, FireDAC.Stan.Param, FireDAC.DatS, FireDAC.DApt.Intf,
  FireDAC.DApt, FireDAC.Comp.DataSet, FireDAC.Comp.Client, FireDAC.Phys.IBBase;

type
  TFDados = class(TDataModule)
    DataSourceDS: TDataSource;
    FDPhysFBDriverLink: TFDPhysFBDriverLink;
    FDTransaction: TFDTransaction;
    Conexao: TFDConnection;
    QryAux: TFDQuery;
    QryAux1: TFDQuery;
    QryAux2: TFDQuery;
    DSQryAux1: TDataSource;
    QryProdutos: TFDQuery;
    dsProdutos: TDataSource;
    QryBusca: TFDQuery;
    dsBusca: TDataSource;
    procedure DataModuleCreate(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  FDados: TFDados;

implementation

{%CLASSGROUP 'Vcl.Controls.TControl'}

{$R *.dfm}



procedure TFDados.DataModuleCreate(Sender: TObject);
begin
  Conexao.Params.Database:=ExtractFileDir(GetCurrentDir)+'\BuscaNCM.FDB';
end;

end.
