{ *************************************************************************** }
{                                                                             }
{ PSEGet Quotation Report Parser v2.0                                   }
{                                                                             }
{ Copyright (c) 2005 Arnold Diaz                                              }
{                                                                             }
{ *************************************************************************** }

unit uPseParser;

interface

uses
  Classes, SysUtils;

type

  TStockItem = class(TCollectionItem)
  public
    Symbol: String;
    StockName: String;
    SubSectorName: String;
    Bid: Double;
    Ask: Double;
    Open: Double;
    High: Double;
    Low: Double;
    Close: Double;
    Volume: Int64;
    Value: Double;
    NetForeignBuy: Int64;
  end;

  TSectorCollection = class(TCollection)
  private
    FSectorName: String;
    FCode: String;
    FOpen: Double;
    FHigh: Double;
    FLow: Double;
    FClose: Double;
    FChange: Double;
    FVolume: Int64;
    FValue: Double;
    FForeignBuy: Double;
    FForeignSell: Double;
    function GetItems(Index: Integer): TStockItem;
    function GetTotalVolume: Int64;
    function GetTotalValue: Double;
    function GetForeignBuy: Double;
    function GetForeignSell: Double;
  public
    constructor Create;
    property Items[Index: Integer]: TStockItem read GetItems;
    function Add: TStockItem;
    property SectorName: String read FSectorName write FSectorName;
    property Code: String read FCode write FCode;
    property Open: Double read FOpen write FOpen;
    property High: Double read FHigh write FHigh;
    property Low: Double read FLow write FLow;
    property Close: Double read FClose write FClose;
    property Change: Double read FChange write FChange;
    property Volume: Int64 read FVolume write FVolume;
    property Value: Double read FValue write FValue;
    property TotalVolume: Int64 read GetTotalVolume;
    property TotalValue: Double read GetTotalValue;
    property ForeignBuy: Double read GetForeignBuy write FForeignBuy;
    property ForeignSell: Double read GetForeignSell write FForeignSell;
  end;

  TSectorItems = class(TCollectionItem)
  public
    SectorCollection: TSectorCollection;
    destructor Destroy; override;
  end;

  PTextFile = ^TextFile;
  TPSECollection = class(TCollection)
  private
    FQuoteDocument: TStringList;
    FTradeDate: System.TDateTime;
    FNumAdvances: ShortInt;
    FNumDeclines: ShortInt;
    FNumUnchanged: ShortInt;
    FNumTraded: SmallInt;
    FNumTrades: SmallInt;
    FOddLotVolume: Int64;
    FOddLotValue: Double;
    FMainCrossVolume: Int64;
    FMainCrossValue: Int64;
    FBondsVolume: Integer;
    FBondsValue: Double;
    FExchangeNotice: TStringList;
    FTotalForeignBuying: Double;
    FTotalForeignSelling: Double;
    function GetItems(Index: Integer): TSectorItems;
  public
    constructor Create;
    destructor Destroy; override;
    function Add: TSectorItems;
    procedure ParseDocument; overload;
    {$unsafecode on}
    procedure SaveToFile(AFileName: TFileName; DataFormat: String;
                         DateFormat: String; IndexValueAsVolume: Boolean = true;
                         MetastockAsc: Boolean = false);

    procedure WriteStockLine(ATextFile: PTextFile; DataFormat, DateFormat: String;
                             APersistent: TPersistent; IndexValueAsVolume: Boolean;
                             AParam: Integer = -1);
    property Items[Index: Integer]: TSectorItems read GetItems;
    property QuoteDocument: TStringList read FQuoteDocument write FQuoteDocument;
    property TradeDate: TDateTime read FTradeDate write FTradeDate;
    property NumAdvances: ShortInt read FNumAdvances;
    property NumDeclines: ShortInt read FNumDeclines;
    property NumUnchanged: ShortInt read FNumUnchanged;
    property NumTraded: SmallInt read FNumTraded;
    property NumTrades: SmallInt read FNumTrades;
    property OddLotVolume: Int64 read FOddLotVolume;
    property OddLotValue: Double read FOddLotValue;
    property MainCrossVolume: Int64 read FMainCrossVolume;
    property MainCrossValue: Int64 read FMainCrossValue;
    property BondsVolume: Integer read FBondsVolume;
    property BondsValue: Double read FBondsValue;
    property ExchangeNotice: TStringList read FExchangeNotice;
    property TotalForeignBuying: Double read FTotalForeignBuying write FTotalForeignBuying;
    property TotalForeignSelling: Double read FTotalForeignSelling write FTotalForeignSelling;
  end;

  TStrArray = array of String;
  TPSEReportFormat = (psePDF, pseBPI, pseMBO);

function Explode(const Separator: String; AText: String): TStrArray;
function Implode(const Separator: String; const StrArray: TStrArray): String;
function CleanupIntegerValue(Value: String): Int64;
function CleanupFloatValue(Value: String): Double;
function CleanupNumberValue(Value: String): String;

implementation

{******************************************************************************}
{   PSE Stock Quotation Report Format (Manila Bulletin Online, www.mb.com.ph)
{   old report format applicable to reports prior to Jan 2, 2006
{
{   <Header> ::= "PHILIPPINE STOCK EXCHANGE"
{                "DAILY QUOTATIONS REPORT"
{                <DATE>
{         <DATE> ::= DDDD M/DD/YYYY
{
{   <Column Title> ::= "NAME"   "SYMBOL"    "BID"   "ASK"   "OPEN"    "HIGH"    "LOW"   "CLOSE"   "VOLUME"    "VALUE"   "NET FOREIGN TRADE (Peso) BUYING (SELLING)"
{
{   <Stock Details> ::= <SECTOR>
{                       <SUBSECTOR>
{                       <STOCK>
{       <SECTOR> ::= "BANKS & FINANCIAL SERVICES" | "COMMERCIAL & INDUSTRIAL" | "PROPERTY" | "MINING" | "OIL" | "SME"  <-- NOTE: SPACES IN BETWEEN CHARATERS
{       <SUBSECTOR> ::= "****" "BANKS" | "FINANCIAL SERVICES" | "COMMUNICATION" | "POWER, ENERGY & OTHER UTILITIES" | "CONSTRUCTION & OTHER RELATED PRODUCTS"
{                              "FOOD, BEVERRAGES & TOBACCO" | "MANUFACTURING, DISTRIBUTION & TRADING" | "HOTEL, RECREATION & OTHER SERVICES" |
{                              "BONDS, PREFERRED & WARRANTS" | "OTHERS" | "PROPERTY" | "MINING" | "OIL" "****"
{       <STOCK> ::= NAME   SYMBOL    BID   ASK   OPEN    HIGH    LOW   CLOSE   VOLUME    VALUE   NETFOREIGN
{
{   <Market Indices> ::= "STOCK PRICE INDICES" <-- NOTE SPACES IN BETWEEN CHARACTERS
{                         <SECTORINDICES>
{                         <ADVANCEDECLINE>
{                         <SECTOR_OHL>
{       <SECTORINDICES> ::= SECTOR  CLOSE   <STATUS>    CHANGE    VOLUME    VALUE
{            <STATUS> ::=  UP | DOWN | UNCH
{       <ADVANCEDECLINE> ::= "NO. OF ADVANCES:"      ADVANCES      "NO. OF DECLINES:" DECLINES   "NO. OF UNCHANGED:"   UNCHANGED
{                            "NO. OF TRADED ISSUES:" TRADEISSUES
{                            "NO. OF TRADES:"        TRADES
{       <SECTOR_OHL> ::= SECTOR   OPEN    HIGH    LOW
{
{   <Miscelleneous Information> ::= "ODD LOTS VOLUME :"   ODDLOTVOLUME
{                                   "ODD LOTS VALUE :"    ODDLOTVALUE
{                                   "MAIN BOARD CROSS VOLUME :"   CROSSVOLUME
{                                   "BONDS VOLUME :"    BONDSVOLUME
{                                   "BONDS VALUE :"     BONDSVALUE
{                                   "TOTAL FOREIGN BUYING : P"    TOTALFOREIGNBUY
{                                   "TOTAL FOREIGN SELLING  : P"  TOTALFOREIGNSELL
{
{   <Exchange Notice> ::= TEXT
{
{   NOTE: Stock Quotation whose closing price is zero will be ignored.
{******************************************************************************}

type
  EUnsupportedFormat = class(Exception);
const
  SUnsupportedFormat = 'Unsupported PSE quotation report';

function Explode(const Separator: String; AText: String): TStrArray;
var
  i, j: Integer;
begin
  SetLength(Result, 0);
  i := Pos(Separator, AText);
  case i of
    0: begin
         SetLength(Result, 1);
         Result[0] := AText;
         Exit;
       end;
    1: Delete(AText, 1, Length(Separator)); //delimiter is at the front
  end;

  j := 0;
  while AnsiPos(Separator, AText) > 0 do
  begin
    i := AnsiPos(Separator, AText);
    if i > 1 then
    begin
      SetLength(Result, Length(Result) + 1);
      Result[j] := Copy(AText, 1, i - 1);
      Delete(AText, 1, i - 1);
      inc(j);
    end;
    Delete(AText, 1, Length(Separator));
  end;

  if AText <> '' then
  begin
    SetLength(Result, Length(Result) + 1);
    Result[High(Result)] := AText;
  end;
end;

function Implode(const Separator: String; const StrArray: TStrArray): String;
var
  i: Integer;
begin
  Result := '';
  for i := 0 to High(StrArray) do
    if i = 0 then
      Result := StrArray[i]
    else
      Result := Result + Separator + StrArray[i];
end;

function CleanupNumberValue(Value: String): String;
var
  i: Integer;
begin
  Result := '';
  Value := Trim(Value);
  if Value = '' then
    Exit
  else if Value[1] = '.' then
    Value := '0'+Value;

  for i := 1 to Length(Value) do
  begin
    if (Value[i] in ['0'..'9']) or
       ((Value[i] = '.') and (Result <> '')) then
      Result := Result + Value[i];
  end;
  if Result = '' then
    Result := '0';
end;

function CleanupIntegerValue(Value: String): Int64;
var
  i: Integer;
begin
  Value := CleanupNumberValue(Value);

  //ignore any character after .
  i := AnsiPos('.', Value);
  if i > 0 then
  begin
    System.Delete(Value, i, (Length(Value) - i) + 1);
  end;

  if not TryStrToInt64(Value, Result) then
    Result := 0

end;

function CleanupFloatValue(Value: String): Double;
begin
  Value := CleanupNumberValue(Value);
  if not TryStrToFloat(Value, Result) then
    Result := 0;
end;

{ TSectorCollection }

function TSectorCollection.Add: TStockItem;
begin
  Result := TStockItem(inherited Add);
end;

constructor TSectorCollection.Create;
begin
  inherited Create(TStockItem);
end;

function TSectorCollection.GetForeignBuy: Double;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do
    if Items[i].NetForeignBuy > 0 then
      Result := Result + Items[i].NetForeignBuy;
end;

function TSectorCollection.GetForeignSell: Double;
var
  i: Integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do
    if Items[i].NetForeignBuy < 0 then
      Result := Result + Items[i].NetForeignBuy;
  Result := Abs(Result);
end;

function TSectorCollection.GetItems(Index: Integer): TStockItem;
begin
  Result := TStockItem(inherited Items[Index]);
end;

function TSectorCollection.GetTotalValue: Double;
begin

end;

function TSectorCollection.GetTotalVolume: Int64;
begin

end;

{ TPSECollection }

function TPSECollection.Add: TSectorItems;
begin
  Result := TSectorItems(inherited Add);
end;

constructor TPSECollection.Create;
begin
  inherited Create(TSectorItems);
  FQuoteDocument := TStringList.Create;
  FExchangeNotice := TStringList.Create;
end;

destructor TPSECollection.Destroy;
begin
  FExchangeNotice.Free;
  FQuoteDocument.Free;
  inherited;
end;

function TPSECollection.GetItems(Index: Integer): TSectorItems;
begin
  Result := TSectorItems(inherited Items[Index]);
end;

procedure TPSECollection.ParseDocument;
var
  i, j, k: Integer;
  StrArray, StrArray1: TStrArray;
  SectorItems: TSectorItems;
  StockItem: TStockItem;
  CurrentSector: Byte;
  SectorNames: TStringList;
  TmpFloat: Double;
  TmpInt: Integer;
  SubSectorName: String;
  RptFormat: TPSEReportFormat;

const
  NewRptFormatDate = '01/02/2006';

  function IsInSectorList(ASectorName: String): Boolean;
  var
    l: ShortInt;
  begin
    Result := false;
    for l := 0 to SectorNames.Count - 1 do
      if Pos(SectorNames.ValueFromIndex[l], ASectorName) > 0  then
      begin
        Result := true;
        break;
      end;
  end;

  procedure CleanupReport;
  var
    TmpStr: TStringList;
    m, p: Integer;
    S: String;
  begin
    TmpStr := TStringList.Create;
    try
      //for m := 0 to QuoteDocument.Count - 1 do
      m := 0;
      while (m < QuoteDocument.Count) do
      begin
        // 10 is just an assumption just to skip the first page
        // that holds the trading date
        QuoteDocument[m] := Trim(QuoteDocument[m]);
        if (m > 10) and (QuoteDocument[m] <> '') then
        begin
          // try locate " and ! then append space after it
          S := QuoteDocument[m];
          p := Pos('"', S);
          if p > 0 then
            System.Insert(' ', S, p + 1);

          p := Pos('!', S);
          if p > 0 then
            System.Insert(' ', S, p + 1);

          QuoteDocument[m] := S;

          StrArray := Explode(' ', QuoteDocument[m]);
          if (LowerCase(StrArray[0]) = 'report') then
          begin
            inc(m, 6);
            Continue;
          end;
        end;

        if Trim(QuoteDocument[m]) <> '' then
          TmpStr.Add(Trim(QuoteDocument[m]));
        inc(m)
      end;
      QuoteDocument.Assign(TmpStr);
    finally
      TmpStr.Free;
    end;
  end;

  procedure InitIndeces;
  var
    l: SmallInt;
  begin
    if FTradeDate < StrToDate(NewRptFormatDate) then
    begin
      // old PSE report format
      if RptFormat in [pseMBO, pseBPI] then
      begin
        SectorNames.Add('^FINANCIAL=B A N K S  &  F I N A N C I A L   S E R V I C E S');
        SectorNames.Add('^COMMERCIAL=C O M M E R C I A L  &  I N D U S T R I A L');
      end else if RptFormat = psePDF then
      begin
        SectorNames.Add('^FINANCIAL=B A N K S & F I N A N C I A L S E R V I C E S');
        SectorNames.Add('^COMMERCIAL=C O M M E R C I A L & I N D U S T R I A L');
      end else
        raise EUnsupportedFormat.Create(SUnsupportedFormat);

      SectorNames.Add('^PROPERTY=P R O P E R T Y');
      SectorNames.Add('^MINING=M I N I N G');
      SectorNames.Add('^OIL=O I L');
      SectorNames.Add('^SME=S M E');
    end else
    begin
      // new PSE report format
      SectorNames.Add('^FINANCIAL=F I N A N C I A L');
      SectorNames.Add('^INDUSTRIAL=I N D U S T R I A L');

      if QuoteDocument.IndexOf('H O L D I N G F I R M S') > 0 then
        SectorNames.Add('^HOLDING=H O L D I N G F I R M S')
      else
        SectorNames.Add('^HOLDING=H O L D I N G    F I R M S');


      SectorNames.Add('^PROPERTY=P R O P E R T Y');
      SectorNames.Add('^SERVICE=S E R V I C E');

      if QuoteDocument.IndexOf('M I N I N G & O I L') > 0 then
        SectorNames.Add('^MINING-OIL=M I N I N G & O I L')
      else
        SectorNames.Add('^MINING-OIL=M I N I N G    &   O I L');

      SectorNames.Add('^PREFERRED=P R E F E R R E D');
      SectorNames.Add('^WARRANT=WARRANTS, PHIL. DEPOSIT RECEIPTS, ETC.');
      SectorNames.Add('^SME=SMALL AND MEDIUM ENTERPRISES');
    end;

    // initialize sectors
    for l := 0 to SectorNames.Count - 1 do
    begin
      SectorItems := Add;
      SectorItems.SectorCollection := TSectorCollection.Create;
      SectorItems.SectorCollection.SectorName := SectorNames.ValueFromIndex[l];
      SectorItems.SectorCollection.Code := SectorNames.Names[l];
    end;

  end;


begin
  //DateSeparator := '/';
  //ShortDateFormat := 'mm/dd/yyyy';
  if QuoteDocument.Count = 0 then
    Exit;

  CleanupReport;

  FExchangeNotice.Clear;
  SectorNames := TStringList.Create;
  try
    // get the trading date
    StrArray := Explode(' ', QuoteDocument[1]);
    if High(StrArray) = 8 then
    begin
      FTradeDate := StrToDateTime(StrArray[3]); // from pdf
      StrArray := Explode(' ', QuoteDocument[2]);
      if StrArray[0] = 'DAILY_QUOTE_REP_PSE' then
        RptFormat := pseBPI
      else
        RptFormat := psePDF;
    end else
    begin
      StrArray := Explode(' ', QuoteDocument[2]);
      if High(StrArray) = 1 then
        FTradeDate := StrToDateTime(StrArray[1]) // text format
      else
        raise Exception.Create('Unsupported PSE quotation report');
      RptFormat := pseMBO;
    end;

    InitIndeces;

    { top-bottom parsing }
    { parse sector block }
    for CurrentSector := 0 to SectorNames.Count - 1 do
    begin
      if (FTradeDate = StrToDateTime(NewRptFormatDate)) and (CurrentSector = 4) then
      begin
        i := QuoteDocument.IndexOf('S E R V I C E S');
      end else
        i := QuoteDocument.IndexOf(SectorNames.ValueFromIndex[CurrentSector]);
      if i < 0 then
        raise Exception.CreateFmt('Unable to locate %s', [SectorNames.Names[CurrentSector]]);
      inc(i);

      //DetectPageChange;
      
      while not IsInSectorList(Trim(QuoteDocument[i])) do
      begin
        //DetectPageChange;
        StrArray := Explode(' ', QuoteDocument[i]);
        if (QuoteDocument[i] = 'B O N D S') or
           (StrArray[0] = 'Financial') then
          break;

        if Copy(QuoteDocument[i], 1, 4) = '****' then
        begin
          SubSectorName := StrArray[1];
          inc(i);
        end;

        //DetectPageChange;
        if IsInSectorList(Trim(QuoteDocument[i])) then
          break;
        StrArray := Explode(' ', QuoteDocument[i]);
        StockItem := Items[CurrentSector].SectorCollection.Add;
        StockItem.SubSectorName := SubSectorName;

        // start from last index
        j := High(StrArray);
        if Trim(StrArray[j])[Length(StrArray[j])] = ')' then
          StockItem.NetForeignBuy := CleanupIntegerValue(StrArray[j]) * -1
        else
          StockItem.NetForeignBuy := CleanupIntegerValue(StrArray[j]);
        if Trim(StrArray[j - 1]) = '(' then
          j := j - 1;
        StockItem.Value := CleanupFloatValue(StrArray[j - 1]);
        StockItem.Volume := CleanupIntegerValue(StrArray[j - 2]);
        StockItem.Close := CleanupFloatValue(StrArray[j - 3]);
        StockItem.Low := CleanupFloatValue(StrArray[j - 4]);
        StockItem.High := CleanupFloatValue(StrArray[j - 5]);
        StockItem.Open := CleanupFloatValue(StrArray[j - 6]);
        StockItem.Ask := CleanupFloatValue(StrArray[j - 7]);
        StockItem.Bid := CleanupFloatValue(StrArray[j - 8]);
        StockItem.Symbol := StrArray[j - 9];

        //get the name and symbol of the stock
        for j := 0 to (j - 9) - 1 do
          StockItem.StockName := StockItem.StockName + ' '+ StrArray[j];

        inc(i);
      end; // while

    end; //for

    { parse stock price indices block
      Add Composite and All shares index }

    with Add do
    begin
      SectorCollection := TSectorCollection.Create;
      SectorCollection.SectorName := 'ALL SHARES';
      SectorCollection.Code := '^ALLSHARES';
    end;

    with Add do
    begin
      SectorCollection := TSectorCollection.Create;
      SectorCollection.SectorName := 'COMPOSITE';
      SectorCollection.Code := '^PSEi';
    end;

    dec(i);
    repeat
      inc(i);
      if i = QuoteDocument.Count then
        raise EUnsupportedFormat.Create(SUnsupportedFormat);
      StrArray := Explode(' ', QuoteDocument[i]);
    until (StrArray[0] = 'Banks') or (StrArray[0] = 'Financial');

    for CurrentSector := 0 to SectorNames.Count + 1 do
    begin
      StrArray := Explode(' ', QuoteDocument[i]);

      with Items[CurrentSector] do
      begin
        // skip these indeces
        if SectorNames.Count > 6 then
          if (SectorCollection.Code = '^PREFERRED') or
             (SectorCollection.Code = '^WARRANT') or
             (SectorCollection.Code = '^SME') then
             Continue;

        j := High(StrArray);
                   
        if (SectorCollection.Code = '^ALLSHARES') or
           (SectorCollection.Code = '^PSEi')  then
        begin
          SectorCollection.Change := CleanupFloatValue(StrArray[j]);
          SectorCollection.Close := CleanupFloatValue(StrArray[j - 2]);

          if SectorCollection.Code = '^PSEi' then
          begin
            StrArray := Explode(' ', QuoteDocument[i + 1]); // phisix volume
            SectorCollection.Volume := CleanupIntegerValue(StrArray[High(StrArray) - 1]);
            SectorCollection.Value := CleanupIntegerValue(StrArray[High(StrArray)]);
          end;

        end else
        begin
          SectorCollection.Value := CleanupFloatValue(StrArray[j]);
          SectorCollection.Volume := CleanupIntegerValue(StrArray[j - 1]);
          SectorCollection.Change := CleanupFloatValue(StrArray[j - 2]);
          SectorCollection.Close := CleanupFloatValue(StrArray[j - 4]);
        end;
        SectorCollection.ForeignBuy := 0;
        SectorCollection.ForeignSell := 0;
      end;
      inc(i);
    end;

    { Advance/Decline block }
    repeat
      StrArray := Explode(' ', QuoteDocument[i]);
      inc(i);
      if i = QuoteDocument.Count then
        raise EUnsupportedFormat.Create(SUnsupportedFormat);
    until StrArray[0] = 'NO.';

    FNumAdvances := CleanupIntegerValue(StrArray[3]);
    FNumDeclines := CleanupIntegerValue(StrArray[7]);
    FNumUnchanged := CleanupIntegerValue(StrArray[11]);

    StrArray := Explode(' ', QuoteDocument[i]);
    FNumTraded := CleanupIntegerValue(StrArray[4]);

    inc(i);
    StrArray := Explode(' ', QuoteDocument[i]);
    FNumTrades := CleanupIntegerValue(StrArray[3]);

    { Sector OHL block }
    inc(i, 2);
    CurrentSector := 0;
    while CurrentSector <= SectorNames.Count + 1 do
    begin
      StrArray := Explode(' ', QuoteDocument[i]);
      //if DetectPageChange then
      //  Continue;

      j := High(StrArray);
      with Items[CurrentSector] do
      begin
        if SectorNames.Count > 6 then
          if (SectorCollection.Code = '^PREFERRED') or
             (SectorCollection.Code = '^WARRANT') or
             (SectorCollection.Code = '^SME') then
          begin
             inc(CurrentSector);
             Continue;
          end;

        SectorCollection.Low := CleanupFloatValue(StrArray[j]);
        SectorCollection.High := CleanupFloatValue(StrArray[j - 1]);
        SectorCollection.Open := CleanupFloatValue(StrArray[j - 2]);
      end;
      inc(CurrentSector);
      inc(i);
    end;

    { miscleneous block }
    StrArray := Explode(' ', QuoteDocument[i]);
    //DetectPageChange;

    FOddLotVolume := CleanupIntegerValue(StrArray[4]);

    inc(i);
    StrArray := Explode(' ', QuoteDocument[i]);
    FOddLotValue := CleanupFloatValue(StrArray[4]);

    { cross volume/value}
    repeat
      StrArray := Explode(' ', QuoteDocument[i]);
      inc(i);
      if i = QuoteDocument.Count then
        raise EUnsupportedFormat.Create(SUnsupportedFormat);
    until StrArray[0] = 'MAIN';
    FMainCrossVolume := CleanupIntegerValue(StrArray[High(StrArray)]);

    repeat
      StrArray := Explode(' ', QuoteDocument[i]);
      inc(i);
      if i = QuoteDocument.Count then
        raise EUnsupportedFormat.Create(SUnsupportedFormat);
    until StrArray[0] = 'MAIN';
    FMainCrossValue := CleanupIntegerValue(StrArray[High(StrArray)]);

    { net foreign buy/sell }
    repeat
      StrArray := Explode(' ', QuoteDocument[i]);
      inc(i);
      if i = QuoteDocument.Count then
        raise EUnsupportedFormat.Create(SUnsupportedFormat);
    until StrArray[0] = 'TOTAL';
    FTotalForeignBuying := CleanupFloatValue(StrArray[High(StrArray)]);

    StrArray := Explode(' ', QuoteDocument[i]);
    FTotalForeignSelling := CleanupFloatValue(StrArray[High(StrArray)]);

    i := QuoteDocument.IndexOf('EXCHANGE NOTICE:');
    if i > 0 then
    begin
      while i < QuoteDocument.Count do
      begin
        FExchangeNotice.Add(QuoteDocument[i]);
        inc(i);
      end;
      FExchangeNotice.Delete(FExchangeNotice.Count - 1);
    end;
  finally
    SectorNames.Free;
  end;
end;

procedure TPSECollection.SaveToFile(AFileName: TFileName; DataFormat,
  DateFormat: String; IndexValueAsVolume: Boolean = true;
  MetastockAsc: Boolean = false);
var
  i, j, l: Integer;
  F: TextFile;
  StockItem: TStockItem;
  Sector: TSectorCollection;
begin
  if MetastockAsc then
  begin
    DataFormat := 'S,M,D,O,H,L,C,V,I';
    DateFormat := 'yyyymmdd';
  end;
  AssignFile(F, AFileName);
  try
    Rewrite(F);
    for i := 0 to Count - 1 do
    begin
      Sector := Items[i].SectorCollection;
      if FTradeDate < StrToDateTime('01/02/2006') then
        WriteStockLine(@F, DataFormat, DateFormat, Sector, IndexValueAsVolume, i)
      else
      begin
        if not (i in [6, 7, 8]) then
          WriteStockLine(@F, DataFormat, DateFormat, Sector, IndexValueAsVolume, i)
      end;

      for j := 0 to Sector.Count - 1 do
      begin
        StockItem := Sector.Items[j];
        if StockItem.Close = 0 then
          Continue;
        WriteStockLine(@F, DataFormat, DateFormat, StockItem, IndexValueAsVolume, i);
      end;
    end;
  finally
    Flush(F);
    CloseFile(F);
  end;
end;

procedure TPSECollection.WriteStockLine(ATextFile: PTextFile;
  DataFormat, DateFormat: String; APersistent: TPersistent;
  IndexValueAsVolume: Boolean; AParam: Integer = -1);
var
  S, S1: String;
  k: Integer;
begin
  S := DataFormat;
  DateFormat := DateFormat;
  S1 := '';
  if AnsiPos('\t', S) > 0 then
    S := StringReplace(S, '\t', #9, [rfReplaceAll]); // replace with tab

  if AnsiPos('D', S) > 0 then
    S := StringReplace(S, 'D', FormatDateTime(DateFormat, FTradeDate), [rfReplaceAll]);
  if AnsiPos('M', S) > 0 then
    S := StringReplace(S, 'M', 'D', [rfReplaceAll]);    
  for k := 1 to Length(S) do
  begin
    case S[k] of
      'S': begin
             if APersistent is TSectorCollection then
               S1 := S1 + TSectorCollection(APersistent).Code
             else
               S1 := S1 + TStockItem(APersistent).Symbol;
           end;
      'O': if APersistent is TSectorCollection then
             S1 := S1 + FloatToStr(TSectorCollection(APersistent).Open)
           else
             S1 := S1 + FloatToStr(TStockItem(APersistent).Open);
      'H': if APersistent is TSectorCollection then
             S1 := S1 + FloatToStr(TSectorCollection(APersistent).High)
           else
             S1 := S1 + FloatToStr(TStockItem(APersistent).High);
      'L': if APersistent is TSectorCollection then
             S1 := S1 + FloatToStr(TSectorCollection(APersistent).Low)
           else
             S1 := S1 + FloatToStr(TStockItem(APersistent).Low);
      'C': if APersistent is TSectorCollection then
             S1 := S1 + FloatToStr(TSectorCollection(APersistent).Close)
           else
             S1 := S1 + FloatToStr(TStockItem(APersistent).Close);
      'V': if APersistent is TSectorCollection then
           begin
             if IndexValueAsVolume then
               S1 := S1 + IntToStr(Trunc(TSectorCollection(APersistent).Value / 1000))
             else
               S1 := S1 + IntToStr(Trunc(TStockItem(APersistent).Volume / 1000));
           end else
             S1 := S1 + IntToStr(TStockItem(APersistent).Volume);
      'I': if APersistent is TSectorCollection then
           begin
             if AParam = Count - 1 then
               S1 := S1 + IntToStr(Trunc(TotalForeignBuying - TotalForeignSelling) div 1000)
             else
               S1 := S1 + IntToStr(
                             Trunc(
                                TSectorCollection(APersistent).ForeignBuy -
                                TSectorCollection(APersistent).ForeignSell
                              ) div 1000
                          )
           end else
             S1 := S1 + FloatToStr(TStockItem(APersistent).NetForeignBuy);

      else
        S1 := S1 + S[k];
    end;
  end;

  Writeln(ATextFile^, S1);
end;

{ TSectorItems }

destructor TSectorItems.Destroy;
begin
  if SectorCollection <> nil then
  begin
    SectorCollection.Clear;
    FreeAndNil(SectorCollection);
  end;
  inherited;
end;


end.
