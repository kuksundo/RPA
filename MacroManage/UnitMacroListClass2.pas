unit UnitMacroListClass2;

interface

uses classes, SysUtils, System.Contnrs, Vcl.Dialogs, Vcl.Controls,
  TypInfo, UnitIniAttriPersist, UnitIniConfigBase,
  mormot.core.json, mormot.core.base, mormot.core.data, mormot.core.unicode,
  mormot.core.text, mormot.orm.base, mormot.core.collections, mormot.core.variants,
  JHP.BaseConfigCollect, thundax.lib.actions_pjh;//System.Generics.Collections, , Generics.Legacy

const
  MACRO_START = 'MacroStart';
  MACRO_ONE_STEP = 'Macro One-Step';
  MACRO_STOP = 'Macro Stop';
  MACRO_MOUSE_POS = 'Macro Mouse position';

type
  TMacroSignalEventRec = record
    FHandleKind: integer;
    FTimeout_ms: cardinal;
    Fidx: integer;
  end;

  TMacroItem = class(TSynAutoCreateFields)
  private
    FItemName,
    FItemValue: string;
  published
    property ItemName: string read FItemName write FItemName;
    property ItemValue: string read FItemValue write FItemValue;
  end;

type
  TActionItem = class//(TSynAutoCreateFields)
  private
    FActionCode,
    FActionDesc,
    FCustomDesc: string;
    FActionType: TActionType;
    FExecMode: TExecuteMode;
    FXPos, FYPos, FWaitSec, FGridIndex, FVKExtendKey: integer;
    FInputText: string;
    function ToString: string;
  public
    class function AddActionItem2List(var AList: IList<IAction>;
      AItem: TActionItem; ADesc: string = ''): IAction;
    procedure AssignTo(ADest: TActionItem);
    procedure CopyActionList(ADest: TActionList);
  published
    property ActionCode: string read FActionCode write FActionCode;
    property ActionDesc: string read FActionDesc write FActionDesc;
    property CustomDesc: string read FCustomDesc write FCustomDesc;
    property ActionType: TActionType read FActionType write FActionType;
    property ExecMode: TExecuteMode read FExecMode write FExecMode;
    property XPos: integer read FXPos write FXPos;
    property YPos: integer read FYPos write FYPos;
    property WaitSec: integer read FWaitSec write FWaitSec;
    property GridIndex: integer read FGridIndex write FGridIndex; //1부터 시작함
    property VKExtendKey: integer read FVKExtendKey write FVKExtendKey; //Mouse Drag시 누른 확장 KeyCode 저장
    property InputText: string read FInputText write FInputText;
  end;

type
  //CHANGE CLASS TYPE HERE FOR TCollectionItemAutoCreateFields
  TMacros = class(TCollectionItemAutoCreateFields)
  private
    FMacroItem : TMacroItem;
    FMacroItemDescription: String;
  published
    property MacroItem : TMacroItem read FMacroItem;
    property MacroItemDescription : String read FMacroItemDescription;
  end;

//  TMacroCollect<T: TMacros> = class(Generics.Legacy.TCollection<T>);

  //THIS IS NEW CLASS FOR MANAGE COLLECTION AND ACCESS LIKE AN ARRAY
type
  TMacroCollection = class(TInterfacedCollection)
  private
    function GetCollItem(aIndex: Integer): TMacros;
  public
    class function GetClass: TCollectionItemClass; override;
    function Add: TMacros;
    property Item[aIndex: Integer]: TMacros read GetCollItem; default;
  end;

  TActions = class(TCollectionItemAutoCreateFields)
  private
    FActionItem: TActionItem;
  public
    procedure AssignActionItem(ASource: TActions);
    procedure AssignActionItem2(ASource: TActionItem);
  published
    property ActionItem : TActionItem read FActionItem;
  end;

  TActionCollection = class(TInterfacedCollection)
  private
    function GetCollItem(aIndex: Integer): TActions;
  protected
    class function GetClass: TCollectionItemClass; override;
  public
    procedure AssignCollect(ASource: TActionCollection);
    function Add: TActions;
    property Item[aIndex: Integer]: TActions read GetCollItem; default;
  end;

  TMacroArray = array of TMacroCollection;

  TMacroManagement = class//(TSynAutoCreateFields)
  private
    FRepeatCount : integer;
    FIsExecute,
    FIsDisplayCustomDesc //True: ActionDesc 대신 CustomDesc를 표시함
    : Boolean;
    FActionDesc: string;
    FCommIniFileName,
    FMacroName,
    FMacroDesc: string;
    FActItemListJson: string;

//    FMacroCollection: TMacroCollection;
    FMacroArray: TMacroArray;
    FRepeatPos: integer;//ActionList의 현재 실행 단계
    FActionStepEnable: Boolean;
    FBreakExecute: Boolean;

    procedure Clear;
  public
    FActionList: IList<IAction>;
    FActionItemList: IList<TActionItem>;//TActionCollection;

    destructor Destroy; override;
    function MacroArrayAdd: TMacros;
    procedure SetActionColl2ActionList;
    procedure SetActionItemList2ActionList;
    procedure ChangeMacroName(AMacroName: string);
    procedure CopyActionItemList(ASrc: IList<TActionItem>; var ADest: IList<TActionItem>);
    //AMsg를 Typing하는 TActionItem을 생성하여 FActionItemList에 추가함
    procedure AddTypeMsgMacro2ActItemList(AMsg: string);

    procedure ExecuteActionList();
    procedure ExecuteActItemList();
    procedure Action2HW(Action: IAction);

    property RepeatPos : integer read FRepeatPos write FRepeatPos;
    property ActionStepEnable : Boolean read FActionStepEnable write FActionStepEnable;
    property BreakExecute : Boolean read FBreakExecute write FBreakExecute;
  published
    [JHPIni('Macro','CommIniFileName','','CommIniFileName', tkString)]
    property CommIniFileName: string read FCommIniFileName write FCommIniFileName;
    [JHPIni('Macro','MacroName','','MacroName', tkString)]
    property MacroName: string read FMacroName write FMacroName;
    [JHPIni('Macro','MacroDesc','','MacroDesc', tkString)]
    property MacroDesc: string read FMacroDesc write FMacroDesc;
    [JHPIni('Macro','RepeatCount','1','RepeatCount', tkInteger)]
    property RepeatCount : integer read FRepeatCount write FRepeatCount;
    [JHPIni('Macro','IsExecute','True','IsExecute', tkEnumeration)]
    property IsExecute : Boolean read FIsExecute write FIsExecute;
    [JHPIni('Macro','IsDisplayCustomDesc','False','IsDisplayCustomDesc', tkEnumeration)]
    property IsDisplayCustomDesc : Boolean read FIsDisplayCustomDesc write FIsDisplayCustomDesc;
    [JHPIni('Macro','ActionDesc','','ActionDesc', tkString)]
    property ActionDesc: string read FActionDesc write FActionDesc;

    property MacroArray: TMacroArray read FMacroArray;
    property ActItemListJson: string read FActItemListJson write FActItemListJson;
  end;

  TMacroManagements = class//(TObjectList)
  public
    FMacroManageList: IList<TMacroManagement>;

    constructor Create();
//    destructor Destroy; override;

    function IsExistMacroName(AName: string): boolean;
    function AddMacro2ListWithName(AName: string=''): integer;
    function AddMacro2List(AMacro: TMacroManagement): integer;
    procedure DeleteMacroFromListWithIndex(AIdx: integer);
    procedure ChangeMacroNameFromIndex(AIdx: integer; AMacroName: string);

    function LoadFromJson(AJson: string): integer;
    function GetBase64FromMacroManageList: string;
    procedure GetMacroManageListFromBase64(ABase64: string);

    procedure ClearObject;
    function LoadFromJSONFile(AFileName: string; APassPhrase: string=''; AIsEncrypt: Boolean=False): integer; virtual;
    function SaveToJSONFile(AFileName: string; APassPhrase: string=''; AIsEncrypt: Boolean=False): integer; virtual;
    function LoadFromStream(AStream: TStream; APassPhrase: string=''; AIsEncrypt: Boolean=False): integer;
    function SaveToStream(AStream: TStream; APassPhrase: string=''; AIsEncrypt: Boolean=False): integer;

    //Json 파일에서 Macro Load하여 ARootMacro.FMacroManageList에 추가
    procedure AddMacro2RootFromJsonFile(AFileName: string; ARootMacro: TMacroManagements=nil);
  end;

  procedure CopyActionCollect(ASrc, ADest: TActionCollection);
  procedure CopyActionColl2ActionList(ASrcColl: TActionCollection; ADestActionList: TActionList);

implementation

uses UnitEncrypt2, UnitRttiUtil2, UnitBase64Util2;

procedure CopyActionCollect(ASrc, ADest: TActionCollection);
var
  i: integer;
begin
  ADest.Clear;

  for i := 0 to ASrc.Count - 1 do
  begin
//    ADest.Add.ActionItem.Assign(ASrc.Item[i].ActionItem);
  end;

//  ASrc.AssignTo(ADest);
end;

procedure CopyActionColl2ActionList(ASrcColl: TActionCollection; ADestActionList: TActionList);
var
  i: integer;
begin
  for i := 0 to ASrcColl.Count - 1 do
  begin
//    TActionItem.AddActionItem2List(ADestActionList, ASrcColl.Item[i].ActionItem);
  end;
end;

{ TMacroCollection }

function TMacroCollection.Add: TMacros;
begin
  Result := TMacros(inherited Add);
end;

class function TMacroCollection.GetClass: TCollectionItemClass;
begin
  Result := TMacros;
end;

function TMacroCollection.GetCollItem(aIndex: Integer): TMacros;
begin
  Result := TMacros(GetItem(aIndex));
end;

{ TActionCollection }

function TActionCollection.Add: TActions;
begin
  Result := TActions(inherited Add);
end;

procedure TActionCollection.AssignCollect(ASource: TActionCollection);
var
  i: integer;
begin
  for i := 0 to ASource.Count - 1 do
  begin
    Add.AssignActionItem(TActions(ASource.Items[i]));
  end;
end;

class function TActionCollection.GetClass: TCollectionItemClass;
begin
  Result := TActions;
end;

function TActionCollection.GetCollItem(aIndex: Integer): TActions;
begin
  Result := TActions(GetItem(aIndex));
end;

{ TMacroManagements }

function TMacroManagements.AddMacro2List(AMacro: TMacroManagement): integer;
var
  LMacroManagement: TMacroManagement;
  LMacroItem: TMacroItem;
  LActionItem: TActionItem;
begin
  LMacroManagement := TMacroManagement.Create;
  LMacroManagement.CommIniFileName := AMacro.CommIniFileName;
  LMacroManagement.MacroName := AMacro.MacroName;
  LMacroManagement.RepeatCount := AMacro.RepeatCount;
  LMacroManagement.IsExecute := AMacro.IsExecute;
  LMacroManagement.FActionList := Collections.NewList<IAction>;// TActionList.Create;
  LMacroManagement.FActionItemList := Collections.NewList<TActionItem>;//TActionList.Create;
  LMacroManagement.CopyActionItemList(AMacro.FActionItemList, LMacroManagement.FActionItemList);
//  LMacroManagement.FActionItemList.Data.LoadFromJson(LActItemListJson);

//  LMacroManagement.FActionCollection.AssignCollect(AMacro.ActionCollect);

  FMacroManageList.Add(LMacroManagement);
end;

function TMacroManagements.AddMacro2ListWithName(AName: string): integer;
var
  LMacroManagement: TMacroManagement;
  LMacroItem: TMacroItem;
  LActionItem: TActionItem;
begin
  LMacroManagement := TMacroManagement.Create;
  LMacroManagement.CommIniFileName := '';
  LMacroManagement.MacroName := AName;
  LMacroManagement.RepeatCount := 1;
  LMacroManagement.IsExecute := True;
  LMacroManagement.FActionList := Collections.NewList<IAction>;//TActionList.Create;
  LMacroManagement.FActionItemList := Collections.NewList<TActionItem>;//TActionList.Create;

  Result := FMacroManageList.Add(LMacroManagement);
end;

procedure TMacroManagements.AddMacro2RootFromJsonFile(AFileName: string;
  ARootMacro: TMacroManagements);
var
  i: integer;
begin
  if not Assigned(ARootMacro) then
    ARootMacro := Self;

  LoadFromJSONFile(AFileName);

  for i := 0 to FMacroManageList.Count - 1 do
  begin
    ARootMacro.AddMacro2List(FMacroManageList.Items[i]);
  end;
end;

procedure TMacroManagements.ChangeMacroNameFromIndex(AIdx: integer; AMacroName: string);
var
  LMacroManagement: TMacroManagement;
begin
  if AIdx < FMacroManageList.Count then
  begin
    LMacroManagement := FMacroManageList.Items[AIdx] as TMacroManagement;
    LMacroManagement.ChangeMacroName(AMacroName);
  end;
end;

procedure TMacroManagements.ClearObject;
var
  i: integer;
begin
  for i := FMacroManageList.Count - 1 downto 0 do
  begin
    TMacroManagement(FMacroManageList.Items[i]).Clear;

//    TMacroManagement(Self.Items[i]).Free;  ==> 이거살리면 Self.Clear할떄 에러남
  end;
end;

constructor TMacroManagements.Create;
begin
  FMacroManageList := Collections.NewList<TMacroManagement>;
end;

procedure TMacroManagements.DeleteMacroFromListWithIndex(AIdx: integer);
var
  LMacroManagement: TMacroManagement;
begin
  if MessageDlg('Are you sure to delete selected Macro?', mtConfirmation, mbOKCancel, 0) = mrOK then
  begin
    LMacroManagement := FMacroManageList.Items[AIdx] as TMacroManagement;
//    LMacroManagement.MacroCollect.Free;
    LMacroManagement.Clear;
    LMacroManagement.Free;
    FMacroManageList.Delete(AIdx);
  end;
end;

function TMacroManagements.GetBase64FromMacroManageList: string;
var
  LUtf8: RawUtf8;
begin
  LUtf8 := FMacroManageList.Data.SaveToJson();
  Result := Utf8ToString(MakeRawUTF8ToBin64(LUtf8));
end;

procedure TMacroManagements.GetMacroManageListFromBase64(ABase64: string);
var
  LUtf8: RawUtf8;
begin
  LUtf8 := MakeBase64ToUTF8(StringToUtf8(ABase64));
  LoadFromJson(Utf8ToString(LUtf8));
//  FMacroManageList.Clear;
//  FMacroManageList.Data.LoadFromJson(LUtf8);
end;

//destructor TMacroManagements.Destroy;
//begin

//  inherited;
//end;

function TMacroManagements.IsExistMacroName(AName: string): boolean;
var
  i: integer;
  LMacroManagement: TMacroManagement;
begin
  Result := False;

  if AName = '' then
    exit;

  for i := 0 to FMacroManageList.Count - 1 do
  begin
    LMacroManagement := FMacroManageList.Items[i] as TMacroManagement;

    if LMacroManagement.MacroName = AName then
    begin
      Result := True;
      break;
    end;
  end;
end;

function TMacroManagements.LoadFromJson(AJson: string): integer;
var
  LActItemListJson: RawUTF8;
  LDocList: IDocList; //LDocList4ActItem
  LDocDict: IDocDict;
  LMacroManagement: TMacroManagement;
  i: integer;
begin
  LActItemListJson := StringToUTF8(AJson);
  LDocList := DocList(LActItemListJson);

//    for LDocDict in LDocList.Objects do
//    begin
//      LActItemListJson := LDocDict['ActItemListJson'];
//    LDocList4ActItem := Collections.NewList<TActionItem>;
//    LDocList4ActItem := DocList(LActItemListJson);
//      LDocDict['ActItemListJson'] := '[]';
//    end;

//    LString := LDocList.Json;

  FMacroManageList.Clear;
  FMacroManageList.Data.LoadFromJson(LActItemListJson);

  i := 0;

  for LDocDict in LDocList.Objects do
  begin
    LMacroManagement := FMacroManageList.Items[i];
    LMacroManagement.FActionList := Collections.NewList<IAction>;
    LMacroManagement.FActionItemList := Collections.NewList<TActionItem>;
    LActItemListJson := LDocDict['ActItemListJson'];
    LMacroManagement.FActionItemList.Data.LoadFromJson(LActItemListJson);
    Inc(i);
  end;
end;

function TMacroManagements.LoadFromJSONFile(AFileName, APassPhrase: string;
  AIsEncrypt: Boolean): integer;
var
  LStrList: TStringList;
  LValid: Boolean;
  Fs: TFileStream;
  LMemStream: TMemoryStream;
  LString: string;
begin
  LStrList := TStringList.Create;
  try
    if AIsEncrypt then
    begin
      LMemStream := TMemoryStream.Create;
      Fs := TFileStream.Create(AFileName, fmOpenRead);
      try
        DecryptStream(Fs, LMemStream, APassphrase);
        LMemStream.Position := 0;
        LStrList.LoadFromStream(LMemStream);
      finally
        LMemStream.Free;
        Fs.Free;
      end;

    end
    else
    begin
      LStrList.LoadFromFile(AFileName);
    end;

    SetLength(LString, Length(LStrList.Text));
    LString := LStrList.Text;

    LoadFromJson(LString);
//    JSONToObject(Self, PUTF8Char(LString), LValid, TMacroManagement, [j2oIgnoreUnknownProperty]);
  finally
    LStrList.Free;
  end;
end;

function TMacroManagements.LoadFromStream(AStream: TStream; APassPhrase: string;
  AIsEncrypt: Boolean): integer;
begin

end;

function TMacroManagements.SaveToJSONFile(AFileName, APassPhrase: string;
  AIsEncrypt: Boolean): integer;
var
  LStrList: TStringList;
  LMemStream: TMemoryStream;
  Fs: TFileStream;
  LStr: RawUTF8;
  LDocList: IDocList;
  LDocDict: IDocDict;
//  LVar: variant;
  LMacroManagement: TMacroManagement;
  i: integer;
begin
  LStrList := TStringList.Create;
  try
//    LStr := ObjectToJSON(Self,[woHumanReadable,woStoreClassName]);
    for LMacroManagement in FMacroManageList do
    begin
      LStr := LMacroManagement.FActionItemList.Data.SaveToJson();
      LStrList.Add(Utf8ToString(LStr));
    end;

    LStr := FMacroManageList.Data.SaveToJson();
    LDocList := DocList(LStr);

    i := 0;

    for LDocDict in LDocList.Objects do
    begin
      LDocDict['ActItemListJson'] := StringToUtf8(LStrList.Strings[i]);
      Inc(i);
    end;

    LStrList.Clear;
    LStr := LDocList.Json;
    LStrList.Add(UTF8ToString(LStr));

    if AIsEncrypt then
    begin
      LMemStream := TMemoryStream.Create;
      Fs := TFileStream.Create(AFileName, fmCreate);
      try
        LStrList.SaveToStream(LMemStream);
        LMemStream.Position := 0;
        EncryptStream(LMemStream, Fs, APassphrase);
      finally
        Fs.Free;
        LMemStream.Free;
      end;
   end
    else
      LStrList.SaveToFile(AFileName);
  finally
    LStrList.Free;
  end;
end;

function TMacroManagements.SaveToStream(AStream: TStream; APassPhrase: string;
  AIsEncrypt: Boolean): integer;
begin

end;

{ TActionItem }

class function TActionItem.AddActionItem2List(var AList: IList<IAction>;
  AItem: TActionItem; ADesc: string): IAction;
var
  action: IAction;
//  actionType: TActionType;
  x, y: Integer;
begin
//  actionType := TActionTypeHelper.GetActionTypeFromDesc(AItem.ActionCode);

  case AItem.ActionType of
    atNull: exit;
    atMousePos:
      begin
        if (AItem.XPos < 0) or (AItem.YPos < 0) then
          raise Exception.Create('Fields must contain valid coordinates');

        action := TAction<Integer>.Create(AItem.ActionType, TParameters<Integer>.Create(AItem.XPos, AItem.YPos), AItem.CustomDesc);
      end;
    atMouseLClick, atMouseLDClick, atMouseRClick, atMouseRDClick,
    atMouseLDown, atMouseLUp, atMouseRDown, atMouseRUp, atMouseMDown, atMouseMUp:
      action := TAction<String>.Create(AItem.ActionType, TParameters<String>.Create('', ''), AItem.CustomDesc);
    atKey:
      begin
        if (AItem.InputText = '') then
          raise Exception.Create('Fields must contain valid key');
        action := TAction<String>.Create(AItem.ActionType, TParameters<String>.Create(AItem.InputText, ''), AItem.CustomDesc);
      end;
    atMessage:
      begin
        if (AItem.InputText = '') then
          raise Exception.Create('Fields must contain valid message');
        action := TAction<String>.Create(AItem.ActionType, TParameters<String>.Create(AItem.InputText, ''), AItem.CustomDesc);
      end;
    atWait:
      begin
        if (AItem.WaitSec = 0) then
          raise Exception.Create('Field must contain time greater than zero');
        action := TAction<Integer>.Create(AItem.ActionType, TParameters<Integer>.Create(AItem.WaitSec, 0), AItem.CustomDesc);
      end;
    atMessage_Dyn:
      begin
        if (AItem.GridIndex <= 0) then
          raise Exception.Create('Fields must contain valid Grid Index');
        action := TAction<Integer>.Create(AItem.ActionType, TParameters<Integer>.Create(AItem.GridIndex, 0), AItem.CustomDesc);
      end;
    atExecuteFunc:
      begin
        if (AItem.InputText = '') then
          raise Exception.Create('Fields must contain valid function name');
        action := TAction<String>.Create(AItem.ActionType, TParameters<String>.Create(AItem.InputText, ''), AItem.CustomDesc);
      end;
    atMouseDrag:
      begin
        if (AItem.XPos < 0) or (AItem.YPos < 0) then
          raise Exception.Create('Fields must contain valid coordinates');
        action := TAction<Integer>.Create(AItem.ActionType, TParameters<Integer>.Create(AItem.XPos, AItem.YPos), IntToStr(AItem.VKExtendKey));
      end;
    atDragBegin:
      begin
        action := TAction<Integer>.Create(AItem.ActionType, TParameters<Integer>.Create(AItem.XPos, AItem.YPos), IntToStr(AItem.VKExtendKey));
      end;
    atDragEnd:
      begin
        action := TAction<Integer>.Create(AItem.ActionType, TParameters<Integer>.Create(AItem.XPos, AItem.YPos), IntToStr(AItem.VKExtendKey));
      end;
  end;

  action.SetExecMode(AItem.ExecMode);

//  if not Assigned(AList) then
//    AList := TActionList.Create;

  AList.Add(action);

  Result := action;
end;

//procedure TActionItem.Assign(Source: TSynPersistent);
//begin
//  inherited;

//  ActionCode := TActionItem(Source).FActionCode;
//  ActionDesc := TActionItem(Source).FActionDesc;
//  ActionType := TActionItem(Source).FActionType;
//  ExecMode := TActionItem(Source).FExecMode;
//  XPos := TActionItem(Source).FXPos;
//  YPos := TActionItem(Source).FYPos;
//  WaitSec := TActionItem(Source).FWaitSec;
//  InputText := TActionItem(Source).FInputText;
//  GridIndex := TActionItem(Source).FGridIndex;
//  VKExtendKey := TActionItem(Source).FVKExtendKey;
//end;

procedure TActionItem.AssignTo(ADest: TActionItem);
begin
  ADest.ActionCode := ActionCode;
  ADest.ActionDesc := ActionDesc;
  ADest.CustomDesc := CustomDesc;
  ADest.ActionType := ActionType;
  ADest.ExecMode := ExecMode;
  ADest.XPos := XPos;
  ADest.YPos := YPos;
  ADest.WaitSec := WaitSec;
  ADest.GridIndex := GridIndex;
  ADest.VKExtendKey := VKExtendKey;
  ADest.InputText := InputText;
end;

procedure TActionItem.CopyActionList(ADest: TActionList);
begin

end;

function TActionItem.ToString: string;
begin

end;

{ TMacroManagement }

procedure TMacroManagement.Action2HW(Action: IAction);
begin
end;

procedure TMacroManagement.AddTypeMsgMacro2ActItemList(AMsg: string);
var
  LActItem: TActionItem;
begin
  LActItem := TActionItem.Create;
  LActItem.FActionCode := g_ActionType.ToString(atMessage);//'Type message';
  LActItem.FActionType := atMessage;
  LActItem.InputText := AMsg;
  LActItem.FActionDesc := 'Message: ' + AMsg;

  FActionItemList.Add(LActItem)
end;

procedure TMacroManagement.ChangeMacroName(AMacroName: string);
begin
  FMacroName := AMacroName;
end;

procedure TMacroManagement.Clear;
var
  i: integer;
begin
//  if Assigned(FActionList) then
//    FActionList.Free;

//  if Assigned(FActionCollection) then
//    FActionCollection.Free;

  for i := Low(MacroArray) to High(MacroArray) do
    MacroArray[i].Free;

  SetLength(FMacroArray, 0);
end;

procedure TMacroManagement.CopyActionItemList(ASrc: IList<TActionItem>;
  var ADest: IList<TActionItem>);
var
  LSrcItem, LDestItem: TActionItem;
  i: integer;
begin
  ADest.Clear;

//  ASrc.Data.CopyTo(ADest.Data^, True);

  for i := 0 to ASrc.Count - 1 do
  begin
    LSrcItem := ASrc.Items[i];
    LDestItem := TActionItem.Create;
    LSrcItem.AssignTo(LDestItem);
//    ASrc.Data.ItemCopyAt(i, @LDestItem);
//    ASrc.Pop(LActionItem, [popFromHead]);
    ADest.Add(LDestItem);
  end;
end;

destructor TMacroManagement.Destroy;
begin
  Clear;
//  inherited;
end;

function TMacroManagement.MacroArrayAdd: TMacros;
var
  i: integer;
begin
  Result := nil;
  i := High(FMacroArray) + 1;

  if i = 0 then
    i := 1;

  SetLength(FMacroArray, i);

  FMacroArray[i-1] := TMacroCollection.Create;
  Result := TMacroCollection(FMacroArray[i-1]).Add;
end;

procedure TMacroManagement.ExecuteActionList;
var
  i, j: Integer;
  action: IAction;
  LIsVKDownStatus: Boolean;
  LPrevVKExtendKey: integer;
  LRec    : TMacroSignalEventRec;
  LExecuteMode: TExecuteMode;
begin
  if RepeatCount > 0 then
  begin
    for j := 0 to RepeatCount - 1 do
    begin
      LIsVKDownStatus := False;
      LPrevVKExtendKey := -1;

      for i := 0 to FActionList.Count - 1 do
      begin
        FRepeatPos := i;
        action := FActionList.Items[i];

        //이전 Action의 VKExtendKey(LPrevVKExtendKey)와 현재 Action의 VKExtendKey가 다르면
        //VKExtendKey 키 누름(Mouse Event 중 Extend Key가 눌려진 경우에만 사용됨)
        if action.GetVKExtendKey <> LPrevVKExtendKey then
        begin
          if (action.GetActionType = atDragEnd) then
          begin
            if (LIsVKDownStatus) then
            begin
              action.SetVKExtendKey(LPrevVKExtendKey);
              action.SetVKAction(2); //VKKey Key_Up
              LIsVKDownStatus := False;
            end;
          end
          else
          begin
            LPrevVKExtendKey := action.GetVKExtendKey;

            //VKKey가 Key_Down되었다가 Key_Up 된 경우
            if LPrevVKExtendKey = -1 then
            begin
              action.SetVKAction(2); //VKKey Key_Up
              LIsVKDownStatus := False;
            end
            else if LPrevVKExtendKey <> 0 then
            begin//VKKey가 Key_Down된 경우
              action.SetVKAction(1); //VKKey Key_Down
              LIsVKDownStatus := True;
            end;
          end;
        end
        else
        begin
          //Extend Key_Down상태에서 Mouse Drag가 끝난 경우 이전에 Key_Down을 Key_Up 해 줘야함
          if (action.GetActionType = atDragEnd) then
          begin
            if (LIsVKDownStatus) then
            begin
              action.SetVKExtendKey(LPrevVKExtendKey);
              action.SetVKAction(2); //VKKey Key_Up
            end;
          end
          else
          if LPrevVKExtendKey <> -1 then//이전 Action과 VKExtendKey값이 같으므로 VKExtendKey키를 누르지 않기 위해 -1을 할당함
            action.SetVKExtendKey(-1);
        end;

        LExecuteMode := action.GetExecMode();

        case LExecuteMode of
          emSWEvent,
          emDriver : action.Execute(LExecuteMode,True);
          emHardware: Action2HW(action);
        end;

        Sleep(10);//200

        if FBreakExecute then
          break;
      end;

      if FBreakExecute then
        break;
    end;//for
  end;
end;

procedure TMacroManagement.ExecuteActItemList;
begin
  SetActionItemList2ActionList();
  ExecuteActionList();
end;

procedure TMacroManagement.SetActionColl2ActionList;
var
  i: integer;
begin
//  for i := 0 to FActionCollection.Count - 1 do
//  begin
//    TActionItem.AddActionItem2List(FActionList, FActionCollection.Item[i].ActionItem);
//  end;
end;

procedure TMacroManagement.SetActionItemList2ActionList;
var
  LActionItem: TActionItem;
begin
  FActionList.Clear;

  for LActionItem in FActionItemList do
  begin
    TActionItem.AddActionItem2List(FActionList, LActionItem);
  end;
end;

{ TActions }

procedure TActions.AssignActionItem(ASource: TActions);
begin
  if ASource is TActions then
  begin
    PersistentCopy(TPersistent(ASource.ActionItem), TPersistent(Self.FActionItem));
  end
  else
    inherited;
end;

procedure TActions.AssignActionItem2(ASource: TActionItem);
begin
  PersistentCopy(TPersistent(ASource), TPersistent(Self.FActionItem));
end;

initialization
  TJSONSerializer.RegisterObjArrayForJSON([TypeInfo(TMacroArray), TMacroCollection]);

end.
