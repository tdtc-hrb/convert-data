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
  hMutex:=CreateMutex(nil,False,'文件格式转换');
  Ret:=GetLastError;
  If Ret<>ERROR_ALREADY_EXISTS Then
  begin
    Application.Initialize;
    Application.Title := '文件格式转换';
    Application.CreateForm(Tfrm_convert, frm_convert);
    Application.Run;
  end
  else
  begin
    Application.MessageBox('程序已经在运行！','提示',MB_OK+MB_ICONHAND);
    Exit;
  end;

end.
