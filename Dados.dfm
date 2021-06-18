object FDados: TFDados
  OldCreateOrder = False
  OnCreate = DataModuleCreate
  Height = 416
  Width = 498
  object DataSourceDS: TDataSource
    DataSet = QryAux
    Left = 200
    Top = 32
  end
  object FDPhysFBDriverLink: TFDPhysFBDriverLink
    VendorLib = 'C:\WebControl\Etiquetas\fbclient.dll'
    Left = 40
    Top = 80
  end
  object FDTransaction: TFDTransaction
    Connection = Conexao
    Left = 40
    Top = 144
  end
  object Conexao: TFDConnection
    Params.Strings = (
      'User_Name=sysdba'
      'Password=masterkey'
      
        'Database=C:\Users\DevDelphi\Desktop\Anderson D C\Estudos Delphi\' +
        'NCM\BuscaNCM.FDB'
      'CharacterSet=WIN1254'
      'DriverID=FB'
      'POOL_MaximumItems=1000')
    Connected = True
    LoginPrompt = False
    Left = 40
    Top = 24
  end
  object QryAux: TFDQuery
    Connection = Conexao
    SQL.Strings = (
      'select * from ncms')
    Left = 120
    Top = 32
  end
  object QryAux1: TFDQuery
    Active = True
    Connection = Conexao
    SQL.Strings = (
      'select * from ncmencontrados')
    Left = 120
    Top = 96
  end
  object QryAux2: TFDQuery
    Connection = Conexao
    Left = 120
    Top = 168
  end
  object DSQryAux1: TDataSource
    DataSet = QryAux1
    Left = 200
    Top = 104
  end
  object QryProdutos: TFDQuery
    Active = True
    Connection = Conexao
    SQL.Strings = (
      'select * from produtostemp')
    Left = 288
    Top = 32
  end
  object dsProdutos: TDataSource
    DataSet = QryProdutos
    Left = 368
    Top = 32
  end
  object QryBusca: TFDQuery
    Active = True
    Connection = Conexao
    SQL.Strings = (
      'select * from resultadoncm')
    Left = 288
    Top = 112
  end
  object dsBusca: TDataSource
    DataSet = QryBusca
    Left = 368
    Top = 112
  end
end
