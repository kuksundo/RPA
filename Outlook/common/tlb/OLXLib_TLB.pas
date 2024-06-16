unit OLXLib_TLB;

// ************************************************************************ //
// WARNING                                                                    
// -------                                                                    
// The types declared in this file were generated from data read from a       
// Type Library. If this type library is explicitly or indirectly (via        
// another type library referring to this type library) re-imported, or the   
// 'Refresh' command of the Type Library Editor activated while editing the   
// Type Library, the contents of this file will be regenerated and all        
// manual modifications will be lost.                                         
// ************************************************************************ //

// $Rev: 52393 $
// File generated on 2024-06-15 오전 9:59:06 from Type Library described below.

// ************************************************************************  //
// Type Lib: C:\Program Files (x86)\Microsoft Office\Root\Office16\OUTLCTL.DLL (1)
// LIBID: {0006F062-0000-0000-C000-000000000046}
// LCID: 0
// Helpfile: 
// HelpString: Microsoft Outlook View Control
// DepndLst: 
//   (1) v2.0 stdole, (C:\Windows\SysWOW64\stdole2.tlb)
// SYS_KIND: SYS_WIN32
// ************************************************************************ //
{$TYPEDADDRESS OFF} // Unit must be compiled without type-checked pointers. 
{$WARN SYMBOL_PLATFORM OFF}
{$WRITEABLECONST ON}
{$VARPROPSETTER ON}
{$ALIGN 4}

interface

uses Winapi.Windows, System.Classes, System.Variants, System.Win.StdVCL, Vcl.Graphics, Vcl.OleCtrls, Vcl.OleServer, Winapi.ActiveX;
  


// *********************************************************************//
// GUIDS declared in the TypeLibrary. Following prefixes are used:        
//   Type Libraries     : LIBID_xxxx                                      
//   CoClasses          : CLASS_xxxx                                      
//   DISPInterfaces     : DIID_xxxx                                       
//   Non-DISP interfaces: IID_xxxx                                        
// *********************************************************************//
const
  // TypeLibrary Major and minor versions
  OLXLibMajorVersion = 1;
  OLXLibMinorVersion = 2;

  LIBID_OLXLib: TGUID = '{0006F062-0000-0000-C000-000000000046}';

  DIID_ViewCtlEvents: TGUID = '{BA4CF450-EE13-11D3-8C45-00C04F4C517C}';
  IID_IObjectModelAccess: TGUID = '{00067276-0000-0000-C000-000000000046}';
  CLASS_ObjectModelCtl: TGUID = '{0006F069-0000-0000-C000-000000000046}';
  IID_IViewCtl: TGUID = '{00067274-0000-0000-C000-000000000046}';
  CLASS_ViewCtl: TGUID = '{261B8CA9-3BAF-4BD0-B0C2-BF04286785C6}';
  IID_IDataCtl: TGUID = '{0468C084-CA5B-11D0-AF08-00609797F0E0}';
  CLASS_DataCtl: TGUID = '{0468C085-CA5B-11D0-AF08-00609797F0E0}';

// *********************************************************************//
// Declaration of Enumerations defined in Type Library                    
// *********************************************************************//
// Constants for enum OlxDefaultFolders
type
  OlxDefaultFolders = TOleEnum;
const
  olxFolderDeletedItems = $00000003;
  olxFolderOutbox = $00000004;
  olxFolderSentMail = $00000005;
  olxFolderInbox = $00000006;
  olxFolderCalendar = $00000009;
  olxFolderContacts = $0000000A;
  olxFolderJournal = $0000000B;
  olxFolderNotes = $0000000C;
  olxFolderTasks = $0000000D;
  olxFolderDrafts = $00000010;

// Constants for enum FIELDREGISTRY_REFRESHOPTIONS
type
  FIELDREGISTRY_REFRESHOPTIONS = TOleEnum;
const
  fro_Forms = $00000000;
  fro_Fields = $00000001;
  fro_View = $00000002;

type

// *********************************************************************//
// Forward declaration of types defined in TypeLibrary                    
// *********************************************************************//
  ViewCtlEvents = dispinterface;
  IObjectModelAccess = interface;
  IObjectModelAccessDisp = dispinterface;
  IViewCtl = interface;
  IViewCtlDisp = dispinterface;
  IDataCtl = interface;

// *********************************************************************//
// Declaration of CoClasses defined in Type Library                       
// (NOTE: Here we map each CoClass to its Default Interface)              
// *********************************************************************//
  ObjectModelCtl = IObjectModelAccess;
  ViewCtl = IViewCtl;
  DataCtl = IDataCtl;


// *********************************************************************//
// Declaration of structures, unions and aliases.                         
// *********************************************************************//
  PWordBool1 = ^WordBool; {*}

  FR_REFRESHOPTIONS = FIELDREGISTRY_REFRESHOPTIONS; 

// *********************************************************************//
// DispIntf:  ViewCtlEvents
// Flags:     (4096) Dispatchable
// GUID:      {BA4CF450-EE13-11D3-8C45-00C04F4C517C}
// *********************************************************************//
  ViewCtlEvents = dispinterface
    ['{BA4CF450-EE13-11D3-8C45-00C04F4C517C}']
    procedure BeforeViewSwitch(const newView: WideString; var Cancel: WordBool); dispid 4;
    procedure ViewSwitch; dispid 5;
    procedure Activate; dispid 1044;
    procedure SelectionChange; dispid 1037;
  end;

// *********************************************************************//
// Interface: IObjectModelAccess
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {00067276-0000-0000-C000-000000000046}
// *********************************************************************//
  IObjectModelAccess = interface(IDispatch)
    ['{00067276-0000-0000-C000-000000000046}']
    function Get_OutlookApplication: IDispatch; safecall;
    procedure Set_OutlookApplication(const pVal: IDispatch); safecall;
    procedure SetPref(const szname: WideString; const szvalue: WideString); safecall;
    function GetPref(const szname: WideString): WideString; safecall;
    procedure DeletePrefs; safecall;
    function Get_ActiveDesktop: Integer; safecall;
    procedure Set_ActiveDesktop(pfActiveDesktop: Integer); safecall;
    procedure EnableInProcOptimizations; safecall;
    procedure FindPerson(const bstrName: WideString); safecall;
    function GetDate: WideString; safecall;
    procedure PickEmailFolders; safecall;
    property OutlookApplication: IDispatch read Get_OutlookApplication write Set_OutlookApplication;
    property ActiveDesktop: Integer read Get_ActiveDesktop write Set_ActiveDesktop;
  end;

// *********************************************************************//
// DispIntf:  IObjectModelAccessDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {00067276-0000-0000-C000-000000000046}
// *********************************************************************//
  IObjectModelAccessDisp = dispinterface
    ['{00067276-0000-0000-C000-000000000046}']
    property OutlookApplication: IDispatch dispid 1;
    procedure SetPref(const szname: WideString; const szvalue: WideString); dispid 2;
    function GetPref(const szname: WideString): WideString; dispid 3;
    procedure DeletePrefs; dispid 4;
    property ActiveDesktop: Integer dispid 5;
    procedure EnableInProcOptimizations; dispid 6;
    procedure FindPerson(const bstrName: WideString); dispid 7;
    function GetDate: WideString; dispid 8;
    procedure PickEmailFolders; dispid 9;
  end;

// *********************************************************************//
// Interface: IViewCtl
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {00067274-0000-0000-C000-000000000046}
// *********************************************************************//
  IViewCtl = interface(IDispatch)
    ['{00067274-0000-0000-C000-000000000046}']
    function Get_View: WideString; safecall;
    procedure Set_View(const pVal: WideString); safecall;
    function Get_Folder: WideString; safecall;
    procedure Set_Folder(const pVal: WideString); safecall;
    function Get_Namespace: WideString; safecall;
    procedure Set_Namespace(const pVal: WideString); safecall;
    function Get_ActiveFolder: IDispatch; safecall;
    function Get_OutlookApplication: IDispatch; safecall;
    function Get_Restriction: WideString; safecall;
    procedure Set_Restriction(const pVal: WideString); safecall;
    function Get_DeferUpdate: WordBool; safecall;
    procedure Set_DeferUpdate(pVal: WordBool); safecall;
    procedure Open; safecall;
    procedure Reply; safecall;
    procedure ReplyAll; safecall;
    procedure Forward; safecall;
    procedure ReplyInFolder; safecall;
    procedure NewDefaultItem; safecall;
    procedure NewOfficeDocument; safecall;
    procedure SaveAs; safecall;
    procedure PrintItem; safecall;
    procedure FlagItem; safecall;
    procedure ForceUpdate; safecall;
    procedure Categories; safecall;
    procedure MarkAllAsRead; safecall;
    procedure GoToToday; safecall;
    procedure Delete; safecall;
    procedure AdvancedFind; safecall;
    procedure AddressBook; safecall;
    procedure MoveItem; safecall;
    procedure NewForm; safecall;
    procedure NewMessage; safecall;
    procedure NewPost; safecall;
    procedure NewAppointment; safecall;
    procedure NewMeetingRequest; safecall;
    procedure NewTask; safecall;
    procedure NewTaskRequest; safecall;
    procedure NewContact; safecall;
    procedure NewJournalEntry; safecall;
    procedure CustomizeView; safecall;
    procedure Sort; safecall;
    procedure GroupBy; safecall;
    procedure ShowFields; safecall;
    procedure CollapseAllGroups; safecall;
    procedure ExpandAllGroups; safecall;
    procedure CollapseGroup; safecall;
    procedure ExpandGroup; safecall;
    procedure AddToPFFavorites; safecall;
    procedure SynchFolder; safecall;
    procedure SendAndReceive; safecall;
    procedure MarkAsRead; safecall;
    procedure MarkAsUnread; safecall;
    procedure OpenSharedDefaultFolder(const bstrRecipient: WideString; FolderType: OlxDefaultFolders); safecall;
    procedure NewNote; safecall;
    procedure NewDistributionList; safecall;
    procedure AddressMessage(const pdispContainer: IDispatch); safecall;
    procedure AddressMeeting(const pdispContainer: IDispatch); safecall;
    function Get_Dirty: WordBool; safecall;
    procedure Set_Dirty(pVal: WordBool); safecall;
    procedure SaveView(const ViewName: WideString); safecall;
    procedure SetDesignMode; safecall;
    procedure GoToDate(const newDate: WideString); safecall;
    function Get_Filter: WideString; safecall;
    procedure Set_Filter(const pVal: WideString); safecall;
    function Get_FilterAppend: WideString; safecall;
    procedure Set_FilterAppend(const pVal: WideString); safecall;
    function Get_ItemCount: Integer; safecall;
    procedure RefreshFieldRegistry(fro: FR_REFRESHOPTIONS); safecall;
    function Get_EnableRowPersistance: WordBool; safecall;
    procedure Set_EnableRowPersistance(pVal: WordBool); safecall;
    function Get_Selection: IDispatch; safecall;
    function Get_ViewXML: WideString; safecall;
    procedure Set_ViewXML(const pVal: WideString); safecall;
    function Get_SelectedDate: TDateTime; safecall;
    procedure ExplorerActivate; safecall;
    procedure ExplorerSelectionChange; safecall;
    procedure ExplorerViewSwitch; safecall;
    procedure ExplorerBeforeViewSwitch(const bStrNewView: WideString; var pVarCancel: WordBool); safecall;
    property View: WideString read Get_View write Set_View;
    property Folder: WideString read Get_Folder write Set_Folder;
    property Namespace: WideString read Get_Namespace write Set_Namespace;
    property ActiveFolder: IDispatch read Get_ActiveFolder;
    property OutlookApplication: IDispatch read Get_OutlookApplication;
    property Restriction: WideString read Get_Restriction write Set_Restriction;
    property DeferUpdate: WordBool read Get_DeferUpdate write Set_DeferUpdate;
    property Dirty: WordBool read Get_Dirty write Set_Dirty;
    property Filter: WideString read Get_Filter write Set_Filter;
    property FilterAppend: WideString read Get_FilterAppend write Set_FilterAppend;
    property ItemCount: Integer read Get_ItemCount;
    property EnableRowPersistance: WordBool read Get_EnableRowPersistance write Set_EnableRowPersistance;
    property Selection: IDispatch read Get_Selection;
    property ViewXML: WideString read Get_ViewXML write Set_ViewXML;
    property SelectedDate: TDateTime read Get_SelectedDate;
  end;

// *********************************************************************//
// DispIntf:  IViewCtlDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {00067274-0000-0000-C000-000000000046}
// *********************************************************************//
  IViewCtlDisp = dispinterface
    ['{00067274-0000-0000-C000-000000000046}']
    property View: WideString dispid 1;
    property Folder: WideString dispid 2;
    property Namespace: WideString dispid 3;
    property ActiveFolder: IDispatch readonly dispid 4;
    property OutlookApplication: IDispatch readonly dispid 5;
    property Restriction: WideString dispid 6;
    property DeferUpdate: WordBool dispid 7;
    procedure Open; dispid 8;
    procedure Reply; dispid 9;
    procedure ReplyAll; dispid 10;
    procedure Forward; dispid 11;
    procedure ReplyInFolder; dispid 12;
    procedure NewDefaultItem; dispid 13;
    procedure NewOfficeDocument; dispid 14;
    procedure SaveAs; dispid 15;
    procedure PrintItem; dispid 16;
    procedure FlagItem; dispid 17;
    procedure ForceUpdate; dispid 18;
    procedure Categories; dispid 19;
    procedure MarkAllAsRead; dispid 20;
    procedure GoToToday; dispid 21;
    procedure Delete; dispid 22;
    procedure AdvancedFind; dispid 23;
    procedure AddressBook; dispid 24;
    procedure MoveItem; dispid 25;
    procedure NewForm; dispid 26;
    procedure NewMessage; dispid 27;
    procedure NewPost; dispid 28;
    procedure NewAppointment; dispid 29;
    procedure NewMeetingRequest; dispid 30;
    procedure NewTask; dispid 31;
    procedure NewTaskRequest; dispid 32;
    procedure NewContact; dispid 33;
    procedure NewJournalEntry; dispid 34;
    procedure CustomizeView; dispid 35;
    procedure Sort; dispid 36;
    procedure GroupBy; dispid 37;
    procedure ShowFields; dispid 38;
    procedure CollapseAllGroups; dispid 39;
    procedure ExpandAllGroups; dispid 40;
    procedure CollapseGroup; dispid 41;
    procedure ExpandGroup; dispid 42;
    procedure AddToPFFavorites; dispid 43;
    procedure SynchFolder; dispid 44;
    procedure SendAndReceive; dispid 45;
    procedure MarkAsRead; dispid 46;
    procedure MarkAsUnread; dispid 47;
    procedure OpenSharedDefaultFolder(const bstrRecipient: WideString; FolderType: OlxDefaultFolders); dispid 48;
    procedure NewNote; dispid 49;
    procedure NewDistributionList; dispid 50;
    procedure AddressMessage(const pdispContainer: IDispatch); dispid 51;
    procedure AddressMeeting(const pdispContainer: IDispatch); dispid 52;
    property Dirty: WordBool dispid 53;
    procedure SaveView(const ViewName: WideString); dispid 54;
    procedure SetDesignMode; dispid 55;
    procedure GoToDate(const newDate: WideString); dispid 60;
    property Filter: WideString dispid 63;
    property FilterAppend: WideString dispid 64;
    property ItemCount: Integer readonly dispid 65;
    procedure RefreshFieldRegistry(fro: FR_REFRESHOPTIONS); dispid 66;
    property EnableRowPersistance: WordBool dispid 67;
    property Selection: IDispatch readonly dispid 68;
    property ViewXML: WideString dispid 70;
    property SelectedDate: TDateTime readonly dispid 71;
    procedure ExplorerActivate; dispid 61441;
    procedure ExplorerSelectionChange; dispid 61447;
    procedure ExplorerViewSwitch; dispid 61444;
    procedure ExplorerBeforeViewSwitch(const bStrNewView: WideString; var pVarCancel: WordBool); dispid 61445;
  end;

// *********************************************************************//
// Interface: IDataCtl
// Flags:     (0)
// GUID:      {0468C084-CA5B-11D0-AF08-00609797F0E0}
// *********************************************************************//
  IDataCtl = interface(IUnknown)
    ['{0468C084-CA5B-11D0-AF08-00609797F0E0}']
  end;

// *********************************************************************//
// The Class CoObjectModelCtl provides a Create and CreateRemote method to          
// create instances of the default interface IObjectModelAccess exposed by              
// the CoClass ObjectModelCtl. The functions are intended to be used by             
// clients wishing to automate the CoClass objects exposed by the         
// server of this typelibrary.                                            
// *********************************************************************//
  CoObjectModelCtl = class
    class function Create: IObjectModelAccess;
    class function CreateRemote(const MachineName: string): IObjectModelAccess;
  end;

// *********************************************************************//
// The Class CoDataCtl provides a Create and CreateRemote method to          
// create instances of the default interface IDataCtl exposed by              
// the CoClass DataCtl. The functions are intended to be used by             
// clients wishing to automate the CoClass objects exposed by the         
// server of this typelibrary.                                            
// *********************************************************************//
  CoDataCtl = class
    class function Create: IDataCtl;
    class function CreateRemote(const MachineName: string): IDataCtl;
  end;

implementation

uses System.Win.ComObj;

class function CoObjectModelCtl.Create: IObjectModelAccess;
begin
  Result := CreateComObject(CLASS_ObjectModelCtl) as IObjectModelAccess;
end;

class function CoObjectModelCtl.CreateRemote(const MachineName: string): IObjectModelAccess;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_ObjectModelCtl) as IObjectModelAccess;
end;

class function CoDataCtl.Create: IDataCtl;
begin
  Result := CreateComObject(CLASS_DataCtl) as IDataCtl;
end;

class function CoDataCtl.CreateRemote(const MachineName: string): IDataCtl;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_DataCtl) as IDataCtl;
end;

end.
