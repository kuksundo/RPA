unit UnitOLDragDropRecord;

interface

uses
  System.Types,
  DragDrop,
  DragDropFormats,
  DropTarget,
  DropSource,
  Classes,
  ActiveX;

type
  TOLEmail_DragSourceRec = record
    FTaskID: string[5];
    FHullNo: string[10];
    FClaimNo: string[20];
    FOrderNo: string[20];
    FSubject: string[100];
//    FEPItem: TEngineParameterItemRecord;
//    FDragDataType: TDragDropDataType; //DragDrop시에 record 가 복수개인지 여부(TDragDropDataType)
    //Drag시 누른 키 값은 HiMECS_Watch2(다른 Propcess이므로)에 전달되지 않기 때문에 따로 전해줌
    //Ctrl + MouseDrag 이벤트가 안 잡혀서 실패함
    FShiftState: TShiftState; //DragDrop시에 Ctrl/Shift/Alt 키 상태 여부
    FSourceHandle: integer; //Drag Source Window Handle
  end;

  // TEngineParameterClipboardFormat defines our custom clipboard format.
  TOLEmail_DragSourceClipboardFormat = class(TCustomSimpleClipboardFormat)
  private
    FGotData: boolean;
    FDataKind: integer;//1: SensorRouteData
    FOLED: TOLEmail_DragSourceRec;
  protected
    function ReadData(Value: pointer; Size: integer): boolean; override;
    function WriteData(Value: pointer; Size: integer): boolean; override;
    function GetSize: integer; override;
    procedure SetOLED(const Value: TOLEmail_DragSourceRec);
  public
    function GetClipboardFormat: TClipFormat; override;
    procedure Clear; override;
    function HasData: boolean; override;
    property OLED: TOLEmail_DragSourceRec read FOLED write SetOLED;
    property DataKind: integer read FDataKind write FDataKind;
  end;

  // TEngineParameterDataFormat defines our custom data format.
  // In this case the data format is identical to the clipboard format, but we
  // need a data format class anyway.
  TOLEmail_DragSourceDataFormat = class(TCustomDataFormat)
  private
    FOLED: TOLEmail_DragSourceRec;
    FDataKind: integer;//1: SensorRouteData
    FGotData: boolean;
  protected
    class procedure RegisterCompatibleFormats; override;
    procedure SetOLED(const Value: TOLEmail_DragSourceRec);
  public
    function Assign(Source: TClipboardFormat): boolean; override;
    function AssignTo(Dest: TClipboardFormat): boolean; override;
    procedure Clear; override;
    function HasData: boolean; override;
    function NeedsData: boolean; override;
    property OLED: TOLEmail_DragSourceRec read FOLED write SetOLED;
    property DataKind: integer read FDataKind write FDataKind;
  end;

const
  sOLEMail_DragSource = 'TOLEmail_DragSourceRec';

implementation

uses
  Windows,
  SysUtils;

{ TOLEmail_DragSourceClipboardFormat }

procedure TOLEmail_DragSourceClipboardFormat.Clear;
begin
  FillChar(FOLED, SizeOf(FOLED), 0);
  FGotData := False;
end;

var
  CF_TOD: TClipFormat = 0;

function TOLEmail_DragSourceClipboardFormat.GetClipboardFormat: TClipFormat;
begin
  if (CF_TOD = 0) then
    CF_TOD := RegisterClipboardFormat(sOLEMail_DragSource);
  Result := CF_TOD;
end;

function TOLEmail_DragSourceClipboardFormat.GetSize: integer;
begin
  Result := SizeOf(TOLEmail_DragSourceRec);
end;

function TOLEmail_DragSourceClipboardFormat.HasData: boolean;
begin
  Result := FGotData;
end;

function TOLEmail_DragSourceClipboardFormat.ReadData(Value: pointer;
  Size: integer): boolean;
begin
  // Copy data from buffer into our structure.
  Move(Value^, FOLED, Size);

  FGotData := True;
  Result := True;
end;

procedure TOLEmail_DragSourceClipboardFormat.SetOLED(const Value: TOLEmail_DragSourceRec);
begin
  FOLED := Value;
  FGotData := True;
end;

function TOLEmail_DragSourceClipboardFormat.WriteData(Value: pointer;
  Size: integer): boolean;
begin
  Result := (Size = SizeOf(TOLEmail_DragSourceRec));
  if (Result) then
    // Copy data from our structure into buffer.
    Move(FOLED, Value^, Size);
end;

{ TOLEmail_DragSourceDataFormat }

function TOLEmail_DragSourceDataFormat.Assign(Source: TClipboardFormat): boolean;
begin
  Result := True;

  if (Source is TOLEmail_DragSourceClipboardFormat) then
    OLED := TOLEmail_DragSourceClipboardFormat(Source).OLED
  else
    Result := inherited Assign(Source);

  FGotData := Result;
end;

function TOLEmail_DragSourceDataFormat.AssignTo(Dest: TClipboardFormat): boolean;
begin
  Result := True;

  if (Dest is TOLEmail_DragSourceClipboardFormat) then
    TOLEmail_DragSourceDataFormat(Dest).OLED := FOLED
  else
    Result := inherited AssignTo(Dest);
end;

procedure TOLEmail_DragSourceDataFormat.Clear;
begin
  Changing;
  FillChar(FOLED, SizeOf(FOLED), 0);
  FGotData := False;
end;

function TOLEmail_DragSourceDataFormat.HasData: boolean;
begin
  Result := FGotData;
end;

function TOLEmail_DragSourceDataFormat.NeedsData: boolean;
begin
  Result := not FGotData;
end;

class procedure TOLEmail_DragSourceDataFormat.RegisterCompatibleFormats;
begin
  inherited RegisterCompatibleFormats;

  RegisterDataConversion(TOLEmail_DragSourceClipboardFormat);
end;

procedure TOLEmail_DragSourceDataFormat.SetOLED(const Value: TOLEmail_DragSourceRec);
begin
  Changing;
  FOLED := Value;
  FGotData := True;
end;

initialization
  // Data format registration
  TOLEmail_DragSourceDataFormat.RegisterDataFormat;
  // Clipboard format registration
  TOLEmail_DragSourceClipboardFormat.RegisterFormat;

finalization
end.
