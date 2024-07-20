unit UnitMacroIniConfig;

interface

uses TypInfo, UnitIniAttriPersist, UnitIniConfigBase;

type
  TMacroIniConfig = class(TJHPIniConfigBase)
    FRepeatCount : integer;
    FIsExecute,
    FIsDisplayCustomDesc //True: ActionDesc ��� CustomDesc�� ǥ����
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

    [JHPIni('Macro','IsDisplayCustomDesc','False',3, tkEnumeration)]

    [JHPIni('Macro','ActionDesc','',4, tkString)]

    [JHPIni('Macro','CommIniFileName','',5, tkString)]

    [JHPIni('Macro','FMacroName','',6, tkString)]

    [JHPIni('Macro','MacroDesc','',7, tkString)]

  end;


