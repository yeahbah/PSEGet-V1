program PSEGet;

{$APPTYPE CONSOLE}

{$R *.res}

uses
  SysUtils,
  uPseParser in 'uPseParser.pas';

const
  HelpStr =   'PSEGet v1.1 Copyright 2005-2007(c) Arnold Diaz \m/(^_^)' +#10#13+
              'This program does not provide any warranty.'+
              'Usage:' +#10#13+
              '    PSEGet [-u quoteurl] [-f format] [-d dateformat] [-o output]' +#10#13+
              '      -p pse report     PSE Quotation Report in text format.' +#10#13+
              ''+#10#13+
              '      -f format         Output format. e.g. S;D;O;H;L;C;V;I' +#10#13+
              '                        Where S - Stock Symbol' +#10#13+
              '                              D - Trading Date' +#10#13+
              '                              O - Open' +#10#13+
              '                              H - High' +#10#13+
              '                              L - Low' +#10#13+
              '                              C - Close' +#10#13+
              '                              V - Volume' +#10#13+
              '                              I - Open Interest'+#10#13+
              '                        if no output format is defined, PSEGet will use'+#10#13+
              '                        the default format which is S,D,O,H,L,C,V,I' +#10#13+
              ''+#10#13+
              '      -d dateformat     Trading date format. e.g. MM-dd-yyyy.' +#10#13+
              '                        Default trading date format is MM/dd/yyyy' +#10#13;

var
  PSE: TPSECollection;
  ParamNo, SwitchNo: Integer;
  Param: String;
  PDFFile, OutputFile: String;
  DateFormat: String = 'MM/dd/yyyy';
  OutputFormat: String = 'S,D,O,H,L,C,V,I';

procedure ShowHelpStr;
begin
  Writeln(HelpStr);
  Halt;
end;

begin
  ParamNo := 1;
  if ParamCount = 0 then
    ShowHelpStr;
  while ParamNo <= ParamCount do
  begin
    Param := ParamStr(ParamNo);
    if (Length(Param) > 1) and (Param[1] = '-') then
    begin
      for SwitchNo := 2 to Length(Param) do
        case Param[SwitchNo] of
          'p': begin
                 inc(ParamNo);
                 PDFFile := ParamStr(ParamNo);
                 if PDFFile = '' then
                   ShowHelpStr;
               end;
          'o': begin
                 inc(ParamNo);
                 OutputFile := ParamStr(ParamNo);
               end;
          'f': begin
                 inc(ParamNo);
                 OutputFormat := ParamStr(ParamNo);
               end;
          'd': begin
                 inc(ParamNo);
                 DateFormat := ParamStr(ParamNo);
               end;
        else
          ShowHelpStr;
        end;

    end else
      ShowHelpStr;
    inc(ParamNo);
  end;

//  ShortDateFormat := 'mm/dd/yyyy';
  PSE := TPSECollection.Create;
  try
    try
      PSE.QuoteDocument.LoadFromFile(PDFFile);
      if PSE.QuoteDocument.Count < 0 then
      begin
        WriteLn('Report is empty :(');
        Halt;
      end;

      OutputFile := ChangeFileExt(PDFFile, '.csv');
      PSE.ParseDocument;
      PSE.SaveToFile(OutputFile, OutputFormat, DateFormat);
    except
      on E: Exception do
        WriteLn(E.Message);
    end;
  finally
    PSE.Free;
  end;
end.
