unit Combinatoria;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls;

type
  TfrmGenerator = class(TForm)
    Edit1: TEdit;
    Edit2: TEdit;
    Edit3: TEdit;
    Button1: TButton;
    procedure Button1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmGenerator: TfrmGenerator;

  N, K: Integer;

  values: Array [1..100] of string;

  max_values: array[1..100] of Integer;

  combination: array[1..100] of Integer;



procedure generate (N, K: Integer);



implementation

{$R *.dfm}

//------------------------------------------------------------------------------

procedure TfrmGenerator.Button1Click(Sender: TObject);

var s: String;

    i: Integer;

begin

  N := StrToInt (Edit1.Text);

  K := StrToInt (Edit3.Text);

  s := Edit2.Text;

  s := s + ',';

  for i := 1 to N do begin

      values[i] := Copy (s, 1, Pos(',', s)-1);

      delete (s, 1, Pos(',', s));

  end;

  for i := 1 to K do begin

      max_values[K-i+1] := N-i+1;

      combination[i] := i;

  end;

  generate (N, K);

end;

//------------------------------------------------------------------------------

procedure next_combination (pos: integer);

begin

  if (pos = 0) then Exit;

  inc (combination[pos]);

  if (combination[pos] > max_values[pos]) then begin

     next_combination (pos - 1);

     combination[pos] := combination[pos - 1] + 1;

  end;

end;

//------------------------------------------------------------------------------

function fat (n: Integer):Integer;

var i: Integer;

begin

  Result := 1;

  for i := 2 to n do

      Result := Result * i;

end;

//------------------------------------------------------------------------------

procedure generate (N, K: Integer);

var n_comb, i, j: Integer;

    f: TextFile;

    s: String;

begin

  n_comb := fat (N) div (fat(K) * fat(N-K));

  ShowMessage ('Number of Combinations: ' + IntToStr (n_comb));

  Screen.Cursor := crHourGlass;

  AssignFile (f, ExtractFilePath (ParamStr (0)) + 'out.csv');

  ReWrite (f);

  for i := 1 to n_comb do begin

      s := '';

      for j := 1 to K do

          s := s + values[combination[j]] + ';';

      WriteLn (f, s);

      next_combination (K);

  end;

  CloseFile (f);

  Screen.Cursor := crDefault;

end;


end.
