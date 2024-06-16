unit UnitMacroConfigClass2;

interface

uses classes, GpCommandLineParser;

type
  TMacroCommandLineOption = class
    FMacroFileName: string;
    FAutoExecute,//���α׷� ���۽� Macro �ڵ� ����
    FNoScreenSaver,//ȭ�� ��ȣ�� ���� Check
    FCheckExecuteTime //����ð� Check
    : Boolean;
    FSetExecuteTime: string;
  public
    [CLPName('m'), CLPLongName('macro'), CLPDescription('macro file name'), CLPDefault('')]
    property MacroFileName: string read FMacroFileName write FMacroFileName;
    [CLPName('a'), CLPLongName('AutoExcute', 'Auto'), CLPDescription('Enable autoExecute mode.')]
    property AutoExecute: boolean read FAutoExecute write FAutoExecute;
    [CLPName('n'), CLPLongName('NoScrSaver', 'No Screen Saver'), CLPDescription('Inhibit Screen Saver')]
    property NoScreenSaver: boolean read FNoScreenSaver write FNoScreenSaver;
    [CLPName('c'), CLPLongName('CheckExeTime', 'Check Exe Time'), CLPDescription('Check Execute Time')]
    property CheckExecuteTime: boolean read FCheckExecuteTime write FCheckExecuteTime;
    [CLPName('t'), CLPLongName('SetExtTime', 'Set Execute Time'), CLPDescription('Set Execute Time')]
    property SetExecuteTime: string read FSetExecuteTime write FSetExecuteTime;
  end;

implementation

end.
