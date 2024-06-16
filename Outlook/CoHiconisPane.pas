unit CoHiconisPane;

{$WARN SYMBOL_PLATFORM OFF}

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ActiveX, AxCtrls, OLAddin4HiconisAS_TLB, OLAddin4HiconisAS_IMPL, StdVcl,
  StdCtrls;

type
  TCoHiconisPane = class(TActiveForm, ICoHiconisPane)
    Label1: TLabel;
  private
    FEvents: ICoHiconisPaneEvents;
    procedure ActivateEvent(Sender: TObject);
    procedure ClickEvent(Sender: TObject);
    procedure CreateEvent(Sender: TObject);
    procedure DblClickEvent(Sender: TObject);
    procedure DeactivateEvent(Sender: TObject);
    procedure DestroyEvent(Sender: TObject);
    procedure KeyPressEvent(Sender: TObject; var Key: Char);
    procedure PaintEvent(Sender: TObject);
    procedure WMMouseActivate(var Message: TWMMouseActivate); message WM_MOUSEACTIVATE;
  protected
    procedure DefinePropertyPages(DefinePropertyPage: TDefinePropertyPage); override;
    procedure EventSinkChanged(const EventSink: IUnknown); override;
    function Get_Active: WordBool; safecall;
    function Get_AutoScroll: WordBool; safecall;
    function Get_AutoSize: WordBool; safecall;
    function Get_AxBorderStyle: TxActiveFormBorderStyle; safecall;
    function Get_Caption: WideString; safecall;
    function Get_Color: Integer; safecall;
    function Get_Cursor: Smallint; safecall;
    function Get_DoubleBuffered: WordBool; safecall;
    function Get_DropTarget: WordBool; safecall;
    function Get_Enabled: WordBool; safecall;
    function Get_HelpFile: WideString; safecall;
    function Get_KeyPreview: WordBool; safecall;
    function Get_PixelsPerInch: Integer; safecall;
    function Get_PrintScale: TxPrintScale; safecall;
    function Get_Scaled: WordBool; safecall;
    function Get_Visible: WordBool; safecall;
    function Get_VisibleDockClientCount: Integer; safecall;
    procedure Set_AutoScroll(Value: WordBool); safecall;
    procedure Set_AutoSize(Value: WordBool); safecall;
    procedure Set_AxBorderStyle(Value: TxActiveFormBorderStyle); safecall;
    procedure Set_Caption(const Value: WideString); safecall;
    procedure Set_Color(Value: Integer); safecall;
    procedure Set_Cursor(Value: Smallint); safecall;
    procedure Set_DoubleBuffered(Value: WordBool); safecall;
    procedure Set_DropTarget(Value: WordBool); safecall;
    procedure Set_Enabled(Value: WordBool); safecall;
    procedure Set_HelpFile(const Value: WideString); safecall;
    procedure Set_KeyPreview(Value: WordBool); safecall;
    procedure Set_PixelsPerInch(Value: Integer); safecall;
    procedure Set_PrintScale(Value: TxPrintScale); safecall;
    procedure Set_Scaled(Value: WordBool); safecall;
    procedure Set_Visible(Value: WordBool); safecall;
  public
    procedure Initialize; override;
    destructor Destroy; override;
  end;

implementation

uses ComObj, ComServ;

{$R *.DFM}

{ TCoHiconisPane }

procedure TCoHiconisPane.DefinePropertyPages(DefinePropertyPage: TDefinePropertyPage);
begin
  { Define property pages here.  Property pages are defined by calling
    DefinePropertyPage with the class id of the page.  For example,
      DefinePropertyPage(Class_CoHiconisPane); }
end;

procedure TCoHiconisPane.EventSinkChanged(const EventSink: IUnknown);
begin
  FEvents := EventSink as ICoHiconisPaneEvents;
end;

procedure TCoHiconisPane.Initialize;
begin
  inherited Initialize;
  OnActivate := ActivateEvent;
  OnClick := ClickEvent;
  OnCreate := CreateEvent;
  OnDblClick := DblClickEvent;
  OnDeactivate := DeactivateEvent;
  OnDestroy := DestroyEvent;
  OnKeyPress := KeyPressEvent;
  OnPaint := PaintEvent;
end;

function TCoHiconisPane.Get_Active: WordBool;
begin
  Result := Active;
end;

function TCoHiconisPane.Get_AutoScroll: WordBool;
begin
  Result := AutoScroll;
end;

function TCoHiconisPane.Get_AutoSize: WordBool;
begin
  Result := AutoSize;
end;

function TCoHiconisPane.Get_AxBorderStyle: TxActiveFormBorderStyle;
begin
  Result := Ord(AxBorderStyle);
end;

function TCoHiconisPane.Get_Caption: WideString;
begin
  Result := WideString(Caption);
end;

function TCoHiconisPane.Get_Color: Integer;
begin
  Result := Integer(Color);
end;

function TCoHiconisPane.Get_Cursor: Smallint;
begin
  Result := Smallint(Cursor);
end;

function TCoHiconisPane.Get_DoubleBuffered: WordBool;
begin
  Result := DoubleBuffered;
end;

function TCoHiconisPane.Get_DropTarget: WordBool;
begin
  Result := DropTarget;
end;

function TCoHiconisPane.Get_Enabled: WordBool;
begin
  Result := Enabled;
end;

function TCoHiconisPane.Get_HelpFile: WideString;
begin
  Result := WideString(HelpFile);
end;

function TCoHiconisPane.Get_KeyPreview: WordBool;
begin
  Result := KeyPreview;
end;

function TCoHiconisPane.Get_PixelsPerInch: Integer;
begin
  Result := PixelsPerInch;
end;

function TCoHiconisPane.Get_PrintScale: TxPrintScale;
begin
  Result := Ord(PrintScale);
end;

function TCoHiconisPane.Get_Scaled: WordBool;
begin
  Result := Scaled;
end;

function TCoHiconisPane.Get_Visible: WordBool;
begin
  Result := Visible;
end;

function TCoHiconisPane.Get_VisibleDockClientCount: Integer;
begin
  Result := VisibleDockClientCount;
end;

procedure TCoHiconisPane.ActivateEvent(Sender: TObject);
begin
  if FEvents <> nil then FEvents.OnActivate;
end;

procedure TCoHiconisPane.ClickEvent(Sender: TObject);
begin
  if FEvents <> nil then FEvents.OnClick;
end;

procedure TCoHiconisPane.CreateEvent(Sender: TObject);
begin
  if FEvents <> nil then FEvents.OnCreate;
end;

procedure TCoHiconisPane.DblClickEvent(Sender: TObject);
begin
  if FEvents <> nil then FEvents.OnDblClick;
end;

procedure TCoHiconisPane.DeactivateEvent(Sender: TObject);
begin
  if FEvents <> nil then FEvents.OnDeactivate;
end;

procedure TCoHiconisPane.DestroyEvent(Sender: TObject);
begin
  if FEvents <> nil then FEvents.OnDestroy;
end;

procedure TCoHiconisPane.KeyPressEvent(Sender: TObject; var Key: Char);
var
  TempKey: Smallint;
begin
  TempKey := Smallint(Key);
  if FEvents <> nil then FEvents.OnKeyPress(TempKey);
  Key := Char(TempKey);
end;

procedure TCoHiconisPane.PaintEvent(Sender: TObject);
begin
  if FEvents <> nil then FEvents.OnPaint;
end;

procedure TCoHiconisPane.Set_AutoScroll(Value: WordBool);
begin
  AutoScroll := Value;
end;

procedure TCoHiconisPane.Set_AutoSize(Value: WordBool);
begin
  AutoSize := Value;
end;

procedure TCoHiconisPane.Set_AxBorderStyle(Value: TxActiveFormBorderStyle);
begin
  AxBorderStyle := TActiveFormBorderStyle(Value);
end;

procedure TCoHiconisPane.Set_Caption(const Value: WideString);
begin
  Caption := TCaption(Value);
end;

procedure TCoHiconisPane.Set_Color(Value: Integer);
begin
  Color := TColor(Value);
end;

procedure TCoHiconisPane.Set_Cursor(Value: Smallint);
begin
  Cursor := TCursor(Value);
end;

procedure TCoHiconisPane.Set_DoubleBuffered(Value: WordBool);
begin
  DoubleBuffered := Value;
end;

procedure TCoHiconisPane.Set_DropTarget(Value: WordBool);
begin
  DropTarget := Value;
end;

procedure TCoHiconisPane.Set_Enabled(Value: WordBool);
begin
  Enabled := Value;
end;

procedure TCoHiconisPane.Set_HelpFile(const Value: WideString);
begin
  HelpFile := String(Value);
end;

procedure TCoHiconisPane.Set_KeyPreview(Value: WordBool);
begin
  KeyPreview := Value;
end;

procedure TCoHiconisPane.Set_PixelsPerInch(Value: Integer);
begin
  PixelsPerInch := Value;
end;

procedure TCoHiconisPane.Set_PrintScale(Value: TxPrintScale);
begin
  PrintScale := TPrintScale(Value);
end;

procedure TCoHiconisPane.Set_Scaled(Value: WordBool);
begin
  Scaled := Value;
end;

procedure TCoHiconisPane.Set_Visible(Value: WordBool);
begin
  Visible := Value;
end;

destructor TCoHiconisPane.Destroy;
var
  ParkingHandle: HWND;
begin
  ParkingHandle := FindWindowEx(0, 0, 'DAXParkingWindow', nil);
  if ParkingHandle <> 0 then
    SendMessage(ParkingHandle, WM_CLOSE, 0, 0);
  inherited Destroy;
end;

function SearchForHWND(const AControl: TWinControl; Focused: HWND): boolean;
var
  i: Integer;
begin
  Result := (AControl.Handle = Focused);
  if not Result then
    for i := 0 to AControl.ControlCount - 1 do
      if AControl.Controls[i] is TWinControl then begin
        if TWinControl(AControl.Controls[i]).Handle = Focused then begin
          Result := True;
          Break;
        end
        else
          if TWinControl(AControl.Controls[i]).ControlCount > 0 then begin
            Result := SearchForHWND(TWinControl(AControl.Controls[i]), Focused);
            if Result then Break;
          end;
      end;
end;

procedure TCoHiconisPane.WMMouseActivate(var Message: TWMMouseActivate);
var
  FocusedWindow: HWND;
  CursorPos: TPoint;
begin
  inherited;
  FocusedWindow := Windows.GetFocus;
  if not SearchForHWND(Self, FocusedWindow) then begin
    Windows.GetCursorPos(CursorPos);
    FocusedWindow := WindowFromPoint(CursorPos);
    Windows.SetFocus(FocusedWindow);
    Message.Result := MA_ACTIVATE;
  end;
end;

initialization
  TActiveFormFactory.Create(
    ComServer,
    TActiveFormControl,
    TCoHiconisPane,
    Class_CoHiconisPane,
    1,
    '',
    OLEMISC_SIMPLEFRAME or OLEMISC_ACTSLIKELABEL,
    tmApartment);
end.
