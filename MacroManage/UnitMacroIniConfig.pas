unit UnitMacroIniConfig;

interface

uses TypInfo, UnitIniAttriPersist, UnitIniConfigBase;

type
  TMacroIniConfig = class(TJHPIniConfigBase)
    FRepeatCount : integer;
    FIsExecute,
    FIsDisplayCustomDesc //True: ActionDesc 대신 CustomDesc를 표시함
    : Boolean;
    FActionDesc,
    FCommIniFileName,
    FMacroName,
    FMacroDesc
    : string;
  published
    //Section Name, Key Name, Default Key Value, Tag Value, TypeKind
    [JHPIni('Macro','RepeatCount','1',1, tkInteger)]
    property RepeatCount : integer read FRepeatCount write FRepeatCount;
    [JHPIni('Macro','IsExecute','True',2, tkEnumeration)]
    property IsExecute : Boolean read FIsExecute write FIsExecute;
    [JHPIni('Macro','IsDisplayCustomDesc','False',3, tkEnumeration)]
    property IsDisplayCustomDesc : Boolean read FIsDisplayCustomDesc write FIsDisplayCustomDesc;
    [JHPIni('Macro','ActionDesc','',4, tkString)]
    property ActionDesc : string read FActionDesc write FActionDesc;
    [JHPIni('Macro','CommIniFileName','',5, tkString)]
    property CommIniFileName : string read FCommIniFileName write FCommIniFileName;
    [JHPIni('Macro','FMacroName','',6, tkString)]
    property MacroName : string read FMacroName write FMacroName;
    [JHPIni('Macro','MacroDesc','',7, tkString)]
    property MacroDesc : string read FMacroDesc write FMacroDesc;
  end;

implementation
end.
