/// <summary>
/// 測試助手功能
/// </summary>
unit Unit_TestHelper;

interface

uses
  Windows, Classes, Clipbrd, SysUtils, Forms, Messages, DB, ADODB, Dialogs,
  ExtCtrls;

type
  TCloseMessageBoxBackgroundWorker = class(TThread)
  protected
    procedure Execute; override;
    procedure CloseMessageBox();
  end;

  TGetMessageBoxContentBackgroundWorker = class(TThread)
  private
    CloseMessageBoxBackgroundWorker: TCloseMessageBoxBackgroundWorker;
  protected
    procedure Execute; override;
    procedure GetMessageBoxContent();
  end;

  /// <summary>
  /// 測試助手類別
  /// </summary>
  TTestHelper = class
  private
    GetMessageBoxContentBackgroundWorker: TGetMessageBoxContentBackgroundWorker;
  public
    /// <summary>
    /// 取得MessageBox內容
    /// </summary>
    /// <param name="Title">要取得MessageBox的標題</param>
    procedure WaitForMessageBox(Title: string);
    function GetMessageBoxContent(): string;
  end;

implementation

var
  FindingMessageBoxTitle: string;
  MessageBoxContent: string;
  IsMessageBoxContentCopied: Boolean;
  ShowingMessageBoxHwnd: hwnd;

procedure TCloseMessageBoxBackgroundWorker.CloseMessageBox();
begin
  if IsMessageBoxContentCopied = True then
  begin
    keybd_event(VK_ESCAPE, 1, 0, 0); //關閉MessageBox
    keybd_event(VK_ESCAPE, 1, KEYEVENTF_KEYUP, 0);
  end
  else
  begin
    CloseMessageBox();
  end;
end;

procedure TGetMessageBoxContentBackgroundWorker.GetMessageBoxContent();
var
  MessageBoxCopyText: TStringList;
  H, Len: Integer;
  Clips: string;
  ClipboardOwner: HWND;
begin
  MessageBoxContent := '';
  IsMessageBoxContentCopied := False;

  ShowingMessageBoxHwnd := FindWindow('TMessageForm', PChar(FindingMessageBoxTitle));
  if ShowingMessageBoxHwnd > 0 then
  begin
    SetForegroundWindow(ShowingMessageBoxHwnd);
    Clipboard.Open();
    keybd_event(VK_CONTROL, 1, 0, 0);
    keybd_event(VkKeyScan('C'), 1, 0, 0);
    keybd_event(VkKeyScan('C'), 1, KEYEVENTF_KEYUP, 0);
    keybd_event(VK_CONTROL, 1, KEYEVENTF_KEYUP, 0);
    try
      H := Clipboard.GetAsHandle(CF_TEXT);
      Len := GlobalSize(H);
      SetLength(Clips, Len);
      SetLength(Clips, Clipboard.GetTextBuf(PChar(Clips), Len));
      MessageBoxCopyText := TStringList.Create;
      MessageBoxCopyText.Text := Clips;

      if (MessageBoxCopyText.Count = 7) and (MessageBoxCopyText[3] <> 'Cannot open clipboard.') and (MessageBoxCopyText[0] = '---------------------------') and (MessageBoxCopyText[2] = '---------------------------') and (MessageBoxCopyText[4] = '---------------------------') and (MessageBoxCopyText[6] = '---------------------------') then
      begin
        MessageBoxContent := MessageBoxCopyText[3];
        IsMessageBoxContentCopied := True;
        CloseMessageBoxBackgroundWorker := TCloseMessageBoxBackgroundWorker.Create(False); //立即執行TCloseMessageBoxContentBackgroundWorker.Execute();
      end
      else
      begin
        GetMessageBoxContent();
      end;
    except
      on Exception do
      begin
        GetMessageBoxContent();
      end
    end;
  end
  else
  begin
    GetMessageBoxContent();
  end;
end;

procedure TGetMessageBoxContentBackgroundWorker.Execute;
begin
  FreeOnTerminate := True;
  GetMessageBoxContent();
end;

procedure TCloseMessageBoxBackgroundWorker.Execute;
begin
  FreeOnTerminate := True;
  CloseMessageBox();
end;

procedure TTestHelper.WaitForMessageBox(Title: string);
begin
  FindingMessageBoxTitle := Title;
  ShowingMessageBoxHwnd := 0;

  GetMessageBoxContentBackgroundWorker := TGetMessageBoxContentBackgroundWorker.Create(False); //立即執行TGetMessageBoxContentBackgroundWorker.Execute();
end;

function TTestHelper.GetMessageBoxContent(): string;
begin
  Result := MessageBoxContent;
end;

end.

