//*****************************************//
//project name:convert.dpr                 //
//purpose:file convert                     //
//writ code:xiao bin                       //
//date:2005.12.21                          //
//*****************************************//

program convert;

uses
  Forms,windows,SysUtils,
  Uconvert in 'Uconvert.pas' {frm_convert};

{$R *.res}

Var

hMutex:HWND;

Ret:Integer;

begin
  hMutex:=CreateMutex(nil,False,'�ļ���ʽת��');
  Ret:=GetLastError;
  If Ret<>ERROR_ALREADY_EXISTS Then
  begin
    Application.Initialize;
    Application.Title := '�ļ���ʽת��';
    Application.CreateForm(Tfrm_convert, frm_convert);
    Application.Run;
  end
  else
  begin
    Application.MessageBox('�����Ѿ������У�','��ʾ',MB_OK+MB_ICONHAND);
    Exit;
  end;

end.
