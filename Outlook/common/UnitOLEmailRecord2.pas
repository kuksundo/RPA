unit UnitOLEmailRecord2;

interface

uses System.SysUtils, Classes,
  OtlCommon, OtlComm,
  mormot.orm.core, mormot.core.base, mormot.rest.sqlite3, mormot.core.os,
  mormot.core.data, mormot.orm.base, mormot.core.variants, mormot.core.datetime;

type
  THiconisASTaskEditConfig = record
    // False = FrameOLEmaiList4Ole에서 TOLControlWorker 사용 안 함(HiconisASManager에서 Start함)
    IsUseOLControlWorkerFromEmailList: Boolean;
    //True = FrameOLEmailList->grid_Mail에 HullNo을 채워 줌
    //False = Selected Email List를 Json Ary로 가져오기 위해서는 HullNo를 채우지 않음
    //        이후 AI로 부터 Hull No/Claim No를 가져와서 채움
    IsAllowUpdateHullNo2Grid,
    IsDocFromInvoiceManage
    : Boolean;

    IPCMQCommandOLEmail, //FrameOLEmaiList4Ole에서 HiconisASManager로 OL Command 요청
    IPCMQ2RespondOLEmail,    //HiconisASManager에서 FrameOLEmaiList4Ole로 OL Respond 전송
    IPCMQCommandOLCalendar,//FrmTodo_Detail.TToDoDetailF에서 HiconisASManager로 OL Command 요청
    IPCMQ2RespondOLCalendar//HiconisASManager에서 FrmTodo_Detail.TToDoDetailF로 OL Respond 전송
    : TOmniMessageQueue;
  end;

  TOLEmailSrchRec = packed record
    //OLEmailListF를 생성한 Form Handle, Close시에 Owner에 Notify하기 위함
    fOwnerFormHandle: THandle;
    FTaskID: TID;
    FHullNo,
    FClaimNo,
    FProjectNo,
    fOrderBy
    : RawUTF8;
    AutoMoveCBCheck,
    SaveToDBButtonEnable,
    CloseButtonEnable,
    //HiconisASManageF->Tool->ShowEmailListForm1 메뉴 선택 시
    //OLEmailListF의 grid_Mail Double Click 시 Frame의 DblClick Event를 변경하기 위함(True 일때 변경됨)
    FHiconisASManageMode
    :Boolean;
    //FrmHiconisASTaskEdit Form을부터 전달 받은 Config Data
    FTaskEditConfig: THiconisASTaskEditConfig;
  end;

  TSQLOLEmailMsg = class(TSQLRecord)
  private
    fTaskID: TID;
//    fDBKey,//Email EntryId를 Key로 사용
    fHullNo,
    fProjectNo,
    fClaimNo: RawUTF8;
    fPrevFolderPath,
    fSavedOLFolderPath,
    fLocalEntryId,
    fLocalStoreId,
    fRemoteEntryId, //원격지의 pst파일에 저장할 때 Id
    fRemoteStoreId,
    fFolderEntryId,
    fFolderStoreId,
    fSenderName,
    fSenderEmail,
    fRecipients, //수신자 Email List
    fCarbonCopy,
    fBlindCC,
    fSubject,
    fSavedMsgFilePath,
    fSavedMsgFileName,
    fFlagRequest,
    fDescription //메일 보충 설명
    : RawUTF8;
    fAttachCount: integer;
    FContainData: integer;//TContainData4Mail;
    //해당 메일이 누구한테 보내는 건지 구분하기 위함
    FProcDirection: integer;//TProcessDirection;
    fRecvDate: TTimeLog;
  public
    FIsUpdate: Boolean;
    property IsUpdate: Boolean read FIsUpdate write FIsUpdate;
  published
    property TaskID: TID read fTaskID write fTaskID;
//    property DBKey: RawUTF8 read fDBKey write fDBKey;// stored AS_UNIQUE;
    property HullNo: RawUTF8 read fHullNo write fHullNo;
    property ProjectNo: RawUTF8 read fProjectNo write fProjectNo;
    property ClaimNo: RawUTF8 read fClaimNo write fClaimNo;
    property PrevFolderPath: RawUTF8 read fPrevFolderPath write fPrevFolderPath;
    property SavedOLFolderPath: RawUTF8 read fSavedOLFolderPath write fSavedOLFolderPath;
    property LocalEntryId: RawUTF8 read fLocalEntryId write fLocalEntryId;
    property LocalStoreId: RawUTF8 read fLocalStoreId write fLocalStoreId;
    property RemoteEntryId: RawUTF8 read fRemoteEntryId write fRemoteEntryId;
    property RemoteStoreId: RawUTF8 read fRemoteStoreId write fRemoteStoreId;
    property FolderEntryId: RawUTF8 read fFolderEntryId write fFolderEntryId;
    property FolderStoreId: RawUTF8 read fFolderStoreId write fFolderStoreId;
    property SenderName: RawUTF8 read fSenderName write fSenderName;
    property SenderEmail: RawUTF8 read fSenderEmail write fSenderEmail;
    property Recipients: RawUTF8 read fRecipients write fRecipients;
    property CC: RawUTF8 read fCarbonCopy write fCarbonCopy;
    property BCC: RawUTF8 read fBlindCC write fBlindCC;
    property Subject: RawUTF8 read fSubject write fSubject;
    property FlagRequest: RawUTF8 read fFlagRequest write fFlagRequest;
    property Description: RawUTF8 read fDescription write fDescription;
    property SavedMsgFilePath: RawUTF8 read fSavedMsgFilePath write fSavedMsgFilePath;
    property SavedMsgFileName: RawUTF8 read fSavedMsgFileName write fSavedMsgFileName;
    property AttachCount: integer read fAttachCount write fAttachCount;
    property ContainData: integer read FContainData write FContainData;
    property ProcDirection: integer read FProcDirection write FProcDirection;
    property RecvDate: TTimeLog read fRecvDate write fRecvDate;
  end;

var
  g_OLEmailMsgDB: TRestClientDB;
  OLEmailMsgModel: TSQLModel;
  g_OLEmailMsgDBFileName: string;

procedure InitOLEmailMsgClient(AExeName: string = ''; ADBFileName: string = '');
function CreateOLEmailMsgModel: TSQLModel;
procedure DestroyOLEmailMsg;

function GetEMailDBName(AExeName, AProdType: string): String;
//function GetSQLOLEmailMsgFromDBKey(ADBKey: string): TSQLOLEmailMsg;
function GetSQLOLEmailMsgFromTaskID(ATaskID: TID): TSQLOLEmailMsg;
function GetFirstStoreIdFromDBKey(ADBKey: TID): string;
function GetOLEmailList2JSONArrayFromDBKey(ADBKey: TID): RawUTF8;
function GetEmailList2JSONArrayFromSearchRec(ASearchRec: TOLEmailSrchRec): RawUTF8;
function GetSQLOLEmailMsgFromSearchRec(ASearchRec: TOLEmailSrchRec): TSQLOLEmailMsg;
procedure GetContainDataNDirFromID(AEntryID: string; out AConData, AProcDir: integer);
//function GetEmailCountFromDBKey(ADBKey: string): integer;
function GetEmailCountFromTaskID(ATaskID: TID): integer;

function AddOLMail2DBFromDroppedMail(AJson: string;
  AAddedMailList: TStringList; AFromRemote: Boolean=False): Boolean;
function UpdateOLMail2DBFromMovedMail(AMovedMailList: TStringList; AFromRemote: Boolean=False): Boolean;
function DeleteOLMail2DBFromTaskID(AID: integer): Boolean;
function DeleteOLMail2DBFromEntryID(AEntryID: string): Boolean;

implementation

uses UnitFolderUtil2, VarRecUtils, Forms, UnitVariantUtil;

procedure InitOLEmailMsgClient(AExeName: string; ADBFileName: string);
var
  LStr, LFileName, LFilePath: string;
begin
  if Assigned(g_OLEmailMsgDB) then
    exit;
//    DestroyOLEmailMsg;

  if AExeName = '' then
    AExeName := Application.ExeName;

  LStr := ExtractFileExt(AExeName);
  LFileName := ExtractFileName(AExeName);
  LFilePath := ExtractFilePath(AExeName);

  if LStr = '.exe' then
  begin
    LFileName := ChangeFileExt(ExtractFileName(AExeName),'.sqlite');
    LFileName := LFileName.Replace('.sqlite', '_Email.sqlite');
    LFilePath := GetSubFolderPath(LFilePath, 'db');
  end;

  LFilePath := EnsureDirectoryExists(LFilePath);

  if ADBFileName = '' then
    g_OLEmailMsgDBFileName := LFilePath + LFileName
  else
    g_OLEmailMsgDBFileName := ADBFileName;

  OLEmailMsgModel := CreateOLEmailMsgModel;
  g_OLEmailMsgDB:= TSQLRestClientDB.Create(OLEmailMsgModel, CreateOLEmailMsgModel,
    g_OLEmailMsgDBFileName, TSQLRestServerDB);
  TSQLRestClientDB(g_OLEmailMsgDB).Server.CreateMissingTables;
end;

function CreateOLEmailMsgModel: TSQLModel;
begin
  result := TSQLModel.Create([TSQLOLEmailMsg]);
end;

procedure DestroyOLEmailMsg;
begin
  if Assigned(g_OLEmailMsgDB) then
    FreeAndNil(g_OLEmailMsgDB);

  if Assigned(OLEmailMsgModel) then
    FreeAndNil(OLEmailMsgModel);
end;

function GetEMailDBName(AExeName, AProdType: string): String;
begin
  Result := AExeName;
  Result := Result.Replace('.exe', '_' + AProdType + '.exe');
end;

//function GetSQLOLEmailMsgFromDBKey(ADBKey: string): TSQLOLEmailMsg;
//begin
//  Result := TSQLOLEmailMsg.CreateAndFillPrepare(g_OLEmailMsgDB.Orm,
//    'DBKey = ?', [ADBKey]);
//
//  if Result.FillOne then
//    Result.IsUpdate := True
//  else
//    Result.IsUpdate := False;
//end;

function GetSQLOLEmailMsgFromTaskID(ATaskID: TID): TSQLOLEmailMsg;
begin
  Result := TSQLOLEmailMsg.CreateAndFillPrepare(g_OLEmailMsgDB.Orm,
    'TaskID = ?', [ATaskID]);

  if Result.FillOne then
    Result.IsUpdate := True
  else
    Result.IsUpdate := False;
end;

function GetFirstStoreIdFromDBKey(ADBKey: TID): string;
var
  LIds: TIDDynArray;
  LSQLEmailMsg: TSQLOLEmailMsg;
begin
  LSQLEmailMsg:= TSQLOLEmailMsg.CreateAndFillPrepare(g_OLEmailMsgDB.Orm, 'TaskID = ?', [ADBKey]);

  try
    if LSQLEmailMsg.FillOne then
    begin
      Result := LSQLEmailMsg.LocalStoreId;
    end;
  finally
    FreeAndNil(LSQLEmailMsg);
  end;
end;

function GetOLEmailList2JSONArrayFromDBKey(ADBKey: TID): RawUTF8;
var
  LSQLEmailMsg: TSQLOLEmailMsg;
  LUtf8: RawUTF8;
  LDynUtf8: TRawUTF8DynArray;
  LDynArr: TDynArray;
begin
  LDynArr.Init(TypeInfo(TRawUTF8DynArray), LDynUtf8);
  LSQLEmailMsg := GetSQLOLEmailMsgFromTaskID(ADBKey);

  try
    LSQLEmailMsg.FillRewind;

    while LSQLEmailMsg.FillOne do
    begin
      LUtf8 := LSQLEmailMsg.GetJSONValues(true, true, soSelect);
      LDynArr.Add(LUtf8);
    end;

    LUtf8 := LDynArr.SaveToJSON;
    Result := LUtf8;
  finally
    FreeAndNil(LSQLEmailMsg);
  end;
end;

function GetEmailList2JSONArrayFromSearchRec(ASearchRec: TOLEmailSrchRec): RawUTF8;
var
  LSQLEmailMsg: TSQLOLEmailMsg;
  LUtf8: RawUTF8;
  LDynUtf8: TRawUTF8DynArray;
  LDynArr: TDynArray;
  LStr: string;
begin
  LDynArr.Init(TypeInfo(TRawUTF8DynArray), LDynUtf8);
  LSQLEmailMsg := GetSQLOLEmailMsgFromSearchRec(ASearchRec);

  try
    LSQLEmailMsg.FillRewind;

    while LSQLEmailMsg.FillOne do
    begin
      LUtf8 := LSQLEmailMsg.GetJSONValues(true, True, soSelect);
      LStr := Utf8ToString(LUtf8);//메일 제목 한글이 깨지는 문제 때문에 추가함 - 2025-03-20
      LDynArr.Add(LStr);
    end;

    LUtf8 := LDynArr.SaveToJSON;
    Result := LUtf8;
  finally
    FreeAndNil(LSQLEmailMsg);
  end;
end;

function GetSQLOLEmailMsgFromSearchRec(ASearchRec: TOLEmailSrchRec): TSQLOLEmailMsg;
var
  ConstArray: TConstArray;
  LWhere, LStr: string;
begin
  LWhere := '';
  ConstArray := CreateConstArray([]);
  try
    if ASearchRec.FTaskID <> -1 then
    begin
      AddConstArray(ConstArray, [ASearchRec.FTaskID]);

      if LWhere <> '' then
        LWhere := LWhere + ' and ';

      LWhere := LWhere + 'TaskID = ? ';
    end;

    if ASearchRec.FHullNo <> '' then
    begin
      AddConstArray(ConstArray, ['%'+ASearchRec.FHullNo+'%']);

      if LWhere <> '' then
        LWhere := LWhere + ' and ';

      LWhere := LWhere + 'HullNo LIKE ? ';
    end;

    if ASearchRec.FClaimNo <> '' then
    begin
      AddConstArray(ConstArray, ['%'+ASearchRec.FClaimNo+'%']);

      if LWhere <> '' then
        LWhere := LWhere + ' and ';

      LWhere := LWhere + 'ClaimNo LIKE ? ';
    end;

    if ASearchRec.FProjectNo <> '' then
    begin
      AddConstArray(ConstArray, ['%'+ASearchRec.FProjectNo+'%']);

      if LWhere <> '' then
        LWhere := LWhere + ' and ';

      LWhere := LWhere + 'ProjectNo LIKE ? ';
    end;

    if LWhere = '' then
    begin
      AddConstArray(ConstArray, [-1]);
      LWhere := 'ID <> ? ';
    end;

    if ASearchRec.fOrderBy <> '' then
      LWhere := LWhere + ' ' + ASearchRec.fOrderBy;

    Result := TSQLOLEmailMsg.CreateAndFillPrepare(g_OLEmailMsgDB.Orm, Lwhere, ConstArray);

    if Result.FillOne then
    begin
      Result.IsUpdate := True;
    end
    else
    begin
      Result.IsUpdate := False;
    end
  finally
    FinalizeConstArray(ConstArray);
  end;
end;

procedure GetContainDataNDirFromID(AEntryID: string; out AConData, AProcDir: integer);
var
  i: integer;
  LEmailMsg: TSQLOLEmailMsg;
begin
  AConData := -1;
  AProcDir := -1;

  LEmailMsg := TSQLOLEmailMsg.Create(g_OLEmailMsgDB.Orm, 'LocalEntryId = ?', [AEntryID]);

  try
    if LEmailMsg.FillOne then
    begin
      AConData := Ord(LEmailMsg.ContainData);
      AProcDir := Ord(LEmailMsg.ProcDirection);
    end;
  finally
    FreeAndNil(LEmailMsg);
  end;
end;

//function GetEmailCountFromDBKey(ADBKey: TID): integer;
//var
//  LSQLEmailMsg: TSQLOLEmailMsg;
//begin
//  Result := 0;
//  LSQLEmailMsg := GetSQLOLEmailMsgFromDBKey(ADBKey);
//  try
//    if LSQLEmailMsg.IsUpdate then
//    begin
//      Result := LSQLEmailMsg.fFill.Table.RowCount;
//    end;
//  finally
//    FreeAndNil(LSQLEmailMsg);
//  end;
//end;

function GetEmailCountFromTaskID(ATaskID: TID): integer;
var
  LSQLEmailMsg: TSQLOLEmailMsg;
begin
  Result := 0;
  LSQLEmailMsg := GetSQLOLEmailMsgFromTaskID(ATaskID);
  try
    if LSQLEmailMsg.IsUpdate then
    begin
      Result := LSQLEmailMsg.fFill.Table.RowCount;
    end;
  finally
    FreeAndNil(LSQLEmailMsg);
  end;
end;

function AddOLMail2DBFromDroppedMail(AJson: string;
  AAddedMailList: TStringList; AFromRemote: Boolean): Boolean;
var
  LVarArr: TVariantDynArray;
  LVar: Variant;
  i, LID: integer;
  LEmailMsg: TSQLOLEmailMsg;
  LUtf8: RawUTF8;
  LEntryId, LStoreId, LWhere, LStr: string;
begin
  Result := False;
  LEntryId := '';
  LStoreId := '';

  if AFromRemote then
    LWhere := 'RemoteEntryID = ? AND RemoteStoreID = ?'
  else
    LWhere := 'LocalEntryID = ? AND LocalStoreID = ?';

  LVarArr := JSONToVariantDynArray(AJson);

  for i := 0 to High(LVarArr) do
  begin
    LVar := _JSON(LVarArr[i]);

    if LVar.EntryId <> Null then
      LEntryId := LVar.EntryId
    else
    if LVar.LocalEntryId <> Null then
    begin
      LEntryId := LVar.LocalEntryId;
    end
    else
    if LVar.RemoteEntryId <> Null then
      LEntryId := LVar.RemoteEntryId;

    if LVar.StoreId  <> Null then
      LStoreId := LVar.StoreId
    else
    if LVar.LocalStoreId  <> Null then
      LStoreId := LVar.LocalStoreId
    else
    if LVar.RemoteStoreId  <> Null then
      LStoreId := LVar.RemoteStoreId;

    LID := StrToIntDef(LVar.TaskID, 0);

    if (LEntryId <> '') and (LStoreId <> '') then
    begin
      LEmailMsg := TSQLOLEmailMsg.CreateAndFillPrepare(g_OLEmailMsgDB.Orm,
        'TaskID = ? AND ' + LWhere, [LID, LEntryId, LStoreId]);
      try
        if LEmailMsg.FillOne then
          LEmailMsg.IsUpdate := True
        else
          LEmailMsg.IsUpdate := False;

        LEmailMsg.TaskID := LID;
        LEmailMsg.HullNo := LVar.HullNo;

        if AFromRemote then
        begin
          LEmailMsg.RemoteEntryId := LEntryId;
          LEmailMsg.RemoteStoreId := LStoreId;
        end
        else
        begin
          LEmailMsg.LocalEntryID := LEntryId;
          LEmailMsg.LocalStoreId := LStoreId;
        end;

        LEmailMsg.FolderEntryId := LVar.FolderEntryId;

        LEmailMsg.SenderName := LVar.SenderName;
        LEmailMsg.SenderEmail := LVar.SenderEmail;
        LEmailMsg.Recipients := LVar.Recipients;
        LEmailMsg.CC := LVar.CC;
        LEmailMsg.BCC := LVar.BCC;
        LEmailMsg.Subject := LVar.Subject;
        LUtf8 := LVar.SavedOLFolderPath;
        LEmailMsg.SavedOLFolderPath := LUtf8;
        //"jhpark@hyundai-gs.com\VDR\" 형식으로 저장 됨
        LEmailMsg.SavedMsgFilePath := GetFolderPathFromEmailPath(LUtf8);
        //GUID.msg 형식으로 저장됨
        LEmailMsg.SavedMsgFileName := LVar.SavedMsgFileName;
        LEmailMsg.AttachCount := StrToIntDef(LVar.AttachCount, 0);
        LEmailMsg.RecvDate := TimeLogFromDateTime(VarToDateWithTimeLog(LVar.RecvDate));//TimeLogFromDateTime(StrToDateTime(LVar.RecvDate));
//        LStr := LVar.ContainData;
        LEmailMsg.ContainData := StrToIntDef(LVar.ContainData, 0);//g_ContainData4Mail.ToType(LStr);
//        LStr := LVar.ProcDirection;
        LEmailMsg.ProcDirection := StrToIntDef(LVar.ProcDirection, 0);//g_ProcessDirection.ToType(LStr);

        LEmailMsg.HullNo := LVar.HullNo;
        LEmailMsg.ClaimNo := LVar.ClaimNo;
        LEmailMsg.ProjectNo := LVar.ProjectNo;
        LEmailMsg.Description := LVar.Description;
        LEmailMsg.FlagRequest := LVar.FlagRequest;

        if LEmailMsg.IsUpdate then
          g_OLEmailMsgDB.Update(LEmailMsg)
        else
        //DB에 동일한 데이터가 없으면 email을 DB에 추가
        begin
          LID := g_OLEmailMsgDB.Add(LEmailMsg, true);

          if Assigned(AAddedMailList) then //신규 추가인 경우 Grid의 RawID를 갱신하기 위해 반환함
            AAddedMailList.Add(LEmailMsg.LocalEntryId + '=' + IntToStr(LID));
        end;

        Result := True;
      finally
        FreeAndNil(LEmailMsg);
      end;
    end;
  end;//for
end;

function UpdateOlMail2DBFromMovedMail(AMovedMailList: TStringList; AFromRemote: Boolean): Boolean;
var
  LEmailMsg: TSQLOLEmailMsg;
  LUtf8, LOldPath: RawUTf8;
  LSrcFile, LDestFile, LWhere: string;
begin
  Result := False;

  if AFromRemote then
    LWhere := 'RemoteEntryID = ? AND RemoteStoreID = ?'
  else
    LWhere := 'LocalEntryID = ? AND LocalStoreID = ?';

  LEmailMsg := TSQLOLEmailMsg.CreateAndFillPrepare(g_OLEmailMsgDB.Orm,
    LWhere, [AMovedMailList.Values['OriginalEntryId'],
                                    AMovedMailList.Values['OriginalStoreId']]);
  try
    if LEmailMsg.FillOne then
    begin
      LEmailMsg.LocalEntryId := AMovedMailList.Values['NewEntryId'];
      LEmailMsg.LocalStoreID := AMovedMailList.Values['MovedStoreId'];
      LUtf8 := AMovedMailList.Values['MovedFolderPath'];
      LOldPath := LEmailMsg.SavedOLFolderPath;

      if LUtf8 <> LOldPath then
      begin
        LEmailMsg.SavedOLFolderPath := LUtf8;
        LEmailMsg.SavedMsgFilePath := GetFolderPathFromEmailPath(LUtf8);
        LSrcFile := ExtractFilePath(g_OLEmailMsgDBFileName) + LOldPath + LEmailMsg.SavedMsgFileName;
        LDestFile := ExtractFilePath(g_OLEmailMsgDBFileName) + LEmailMsg.SavedMsgFilePath + LEmailMsg.SavedMsgFileName;

        if FileExists(LSrcFile) then
        begin
          if CopyFile(LSrcFile, LDestFile, True) then
            DeleteFile(LSrcFile);
        end;
      end;

      Result := g_OLEmailMsgDB.Update(LEmailMsg);
    end;
  finally
    FreeAndNil(LEmailMsg);
  end;
end;

function UpdateOLMail2DBFromContainDataNProcdir(AID, AConData, AProcDir: integer): Boolean;
var
  LEmailMsg: TSQLOLEmailMsg;
begin
  Result := False;

  LEmailMsg := TSQLOLEmailMsg.CreateAndFillPrepare(g_OLEmailMsgDB.Orm, 'ID = ?', [AID]);

  try
    if LEmailMsg.FillOne then
    begin
      LEmailMsg.ContainData := AConData;
      LEmailMsg.ProcDirection := AProcDir;
      Result := g_OLEmailMsgDB.Update(LEmailMsg);
    end;
  finally
    FreeAndNil(LEmailMsg);
  end;
end;

function DeleteOLMail2DBFromTaskID(AID: integer): Boolean;
begin
  Result := g_OLEmailMsgDB.Delete(TSQLOLEmailMsg, 'TaskID = ?', [AID]);
end;

function DeleteOLMail2DBFromEntryID(AEntryID: string): Boolean;
begin
  Result := g_OLEmailMsgDB.Delete(TSQLOLEmailMsg, 'LocalEntryID = ?', [AEntryID]);
end;

initialization
  g_OLEmailMsgDB := nil;

finalization
  DestroyOLEmailMsg;

end.
