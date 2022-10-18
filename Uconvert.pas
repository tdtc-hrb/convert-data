//***************************************//
//unit name:Uconvert.pas                 //
//purpose:read file                      //
//write coder:xiao bin                   //
//date:2004.12.24                        //
//***************************************//

unit Uconvert;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, ComCtrls, StdCtrls, Buttons, Gauges, Menus,IniFiles ,
  CoolTrayIcon, DB, ADODB, ShellApi,Grids, DBGrids;

type
  Tfrm_convert = class(TForm)
    Panel1: TPanel;
    StatusBar1: TStatusBar;
    Panel2: TPanel;
    Panel3: TPanel;
    Splitter1: TSplitter;
    Panel4: TPanel;
    mdb_path: TEdit;
    Label1: TLabel;
    Label2: TLabel;
    rep_path: TEdit;
    Label3: TLabel;
    compound_path: TEdit;
    BitBtn_set: TBitBtn;
    BitBtn_run: TBitBtn;
    Memo_recond: TMemo;
    Panel5: TPanel;
    PopupM1: TPopupMenu;
    S1: TMenuItem;
    N1: TMenuItem;
    Q1: TMenuItem;
    Timer_loop: TTimer;
    S2: TMenuItem;
    btn_timeloop: TButton;
    Timer_continue: TTimer;
    Timer_status: TTimer;
    adoconn_hz: TADOConnection;
    adoquery_hz: TADOQuery;
    DataSource1: TDataSource;
    DBGrid1: TDBGrid;
    Memo1: TMemo;
    Timer_flag: TTimer;
    adoquery_continue: TADOQuery;
    adoquery_max: TADOQuery;
    creat_delfile: TMemo;
    Memo_view: TMemo;
    CoolTrayIcon1: TCoolTrayIcon;
    procedure BitBtn_runClick(Sender: TObject);
    procedure Q1Click(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure BitBtn_setClick(Sender: TObject);
    procedure S1Click(Sender: TObject);
    procedure S2Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure btn_timeloopClick(Sender: TObject);
    procedure Timer_continueTimer(Sender: TObject);
    procedure Timer_statusTimer(Sender: TObject);
    procedure adoconn_hzAfterConnect(Sender: TObject);
    procedure Timer_flagTimer(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
    lastfile_xls:string;
    inifile_total:TIniFile;
    number1:integer;
    del_num:integer;
    flag1:Boolean;
    t2:real;//timer_flag interval
    //dynamic link library using HWND
    h1:THandle;

    { Private declarations }
  public
    recondfile:TStringList;
    subsection1,subsection2,subsection3,bakSubsection03:string;
    CRCsection1:string;//CRC32path
    filepath:string;
    xiaobin_path:string;
    flag_number:integer;
    //remote hard disk---2006.11.26
    localName2,remoteName2,userName2,passWord2:PChar;
    function WNetConnect(localName1,remoteName1,user1,passWord1:PChar):DWORD;
    function WNetCancel(localName3:PChar):DWORD;
    //
    procedure WMSysCommand(var Message: TWMSysCommand);
    message WM_SYSCOMMAND;
    { Public declarations }
  end;

type
  Tpro_saveFCN=procedure(saveFile1,CheckFilePath:WideString);stdcall;  

var
  frm_convert: Tfrm_convert;

  saveFCNA:Tpro_saveFCN;
  //
  function UpTime: string;
  function GetTextFromFile(AFile : String; var ReturnString : string) : boolean;

implementation

{$R *.dfm}


procedure Tfrm_convert.WMSysCommand(var Message:TWMSysCommand);
begin
  Inherited;
  if (Message.CmdType and $FFF0 = SC_MINIMIZE) then
  begin
    if not BitBtn_run.Enabled then
    begin
      CoolTrayIcon1.IconVisible:=True;
      CoolTrayIcon1.MinimizeToTray:=True;
      frm_convert.Visible:=False;
    end;
  end;
end;

function Tfrm_convert.WNetConnect(localName1,remoteName1,user1,passWord1:PChar):DWORD;
var
  NetR:NETRESOURCE;
  ErrorInfo:Longint;
begin
  with NetR do
  begin
    dwType:=RESOURCETYPE_ANY;
    lpLocalName:=localName1;
    lpRemoteName:=remoteName1;
    lpProvider:='';
  end;
  ErrorInfo:=WNetAddConnection2(NetR,passWord1,user1,CONNECT_UPDATE_PROFILE);
  Result:=ErrorInfo;
end;

function Tfrm_convert.WNetCancel(localName3:PChar):DWORD;
var
  ErrInfo:Longint;
begin
  ErrInfo:=WNetCancelConnection2(localName3,CONNECT_UPDATE_PROFILE,False);
  Result:=ErrInfo;
end;

procedure Tfrm_convert.FormCreate(Sender: TObject);
//var
  //c1:Cardinal;
  //adoconnectstr2:string;
begin
  flag1:=True;
  xiaobin_path:=ExtractFilePath(ParamStr(0))+'xiaobin1224.xbf';
  t2:=Timer_flag.Interval;
end;

procedure Tfrm_convert.FormShow(Sender: TObject);
var
  j1:integer;
begin
  //读取INI文件
  filepath:=ExtractFilePath(ParamStr(0))+'filepath.ini';
  inifile_total:=TIniFile.Create(filepath);
  subsection1:=inifile_total.ReadString('access_path','mdb1',mdb_path.Text);
  subsection2:=inifile_total.ReadString('report_path','rep1',rep_path.Text);
  subsection3:=inifile_total.ReadString('convert_path','cnt1',compound_path.Text);
  bakSubsection03:=inifile_total.ReadString('convert_path','cnt2','c:\f\13754037907.log');
  CRCsection1:=inifile_total.ReadString('convert_path','cnt3','x:\datasb.fcn');
  mdb_path.Text:=subsection1;
  rep_path.Text:=subsection2;
  compound_path.Text:=subsection3;
  recondfile:=TStringList.Create;
  //使用网络映射硬盘
  localName2:=pchar(inifile_total.ReadString('net_harddisk','nhd1','x:'));
  remoteName2:=pchar(inifile_total.ReadString('remote_path','rmt1','\\brxl_server\s'));
  userName2:=pchar(inifile_total.ReadString('remote_login','rlg1','administrator'));
  passWord2:=pchar(inifile_total.ReadString('remote_login','rlg2','123'));
  if WNetConnect(localName2,remoteName2,userName2,passWord2)<>NO_ERROR then
  begin
    for j1:=0 to 3 do//重试3次
    begin
      WNetCancel(localName2);
      WNetConnect(localName2,remoteName2,userName2,passWord2);
    end;
  end
  else
  begin
    Application.MessageBox('使用网络映射硬盘，失败！','Hint',MB_OK);
  end;

  //连接数据库
  if not adoconn_hz.Connected then
  begin
    adoconn_hz.Close;
    adoconn_hz.ConnectionString:='Provider=Microsoft.Jet.OLEDB.4.0;Data Source='+subsection1+';Persist Security Info=False';
    adoconn_hz.Open;
  end;

  //程序启动

  BitBtn_runClick(nil);
  BitBtn_run.Enabled:=False;
end;

procedure Tfrm_convert.BitBtn_setClick(Sender: TObject);
var
  connectstr:string;
  connectfile:string;
begin
  //写INI文件
  inifile_total.WriteString('access_path','mdb1',mdb_path.Text);
  inifile_total.WriteString('report_path','rep1',rep_path.Text);
  inifile_total.WriteString('convert_path','cnt1',compound_path.Text);
  //连接华章数据库
  if GetTextFromFile(xiaobin_path,connectfile)then
  begin
    connectstr:=connectfile;
  end;
  try
    if not adoconn_hz.Connected then
    begin
      adoconn_hz.ConnectionString:=connectstr;
      adoconn_hz.Connected:=True;
    end;
    //读取重量数据
    adoquery_hz.Close;
    adoquery_hz.SQL.Clear;
    adoquery_hz.SQL.Add('select Val,ord from FloatTable');
    adoquery_hz.Open;
    
  except
    Application.MessageBox('连接MicroSoft ACCESS数据库失败！','提示',MB_OK+MB_ICONINFORMATION);
    Exit;
  end;
  //检查发送文件夹是否存在
  if not FileExists(compound_path.Text) then
  begin
    Application.MessageBox('合成文件夹和文件是否存在？'+#13+'请建立合成文件夹!','提示',MB_OK+MB_ICONINFORMATION);
    BitBtn_run.Enabled:=False;
    Exit;
  end;
  Application.MessageBox('所有设置选项完成！','提示',MB_OK+MB_ICONINFORMATION);
  BitBtn_run.Enabled:=True;
  BitBtn_set.Enabled:=False;
  mdb_path.Enabled:=False;
  rep_path.Enabled:=False;
  compound_path.Enabled:=False;
end;

procedure Tfrm_convert.BitBtn_runClick(Sender: TObject);
begin
  number1:=inifile_total.ReadInteger('last_rep','ltp2',1);
  
  lastfile_xls:=inifile_total.ReadString('excel_last','elt1','book1');

  BitBtn_run.Enabled:=False;
  BitBtn_set.Enabled:=False;
  mdb_path.Enabled:=False;
  rep_path.Enabled:=False;
  compound_path.Enabled:=False;
  Timer_loop.OnTimer:=btn_timeloop.OnClick;
end;

function UpTime: string;
const
  ticksperday : integer = 1000 * 60 * 60 * 24;
  ticksperhour : integer = 1000 * 60 * 60;
  ticksperminute : integer = 1000 * 60;
  tickspersecond : integer = 1000;
var
  t : longword;
  d,h,m,s : integer;
begin
  t := GetTickCount;  //驻留内存
  d := t div ticksperday;
  dec(t,d * ticksperday);
  h := t div ticksperhour;
  dec(t,h * ticksperhour);
  m := t div ticksperminute;
  dec(t,m * ticksperminute);
  s := t div tickspersecond;
  Result := '运行时间: '+IntToStr(d)+ ' 天 '+IntToStr(h)+' 小时 '+IntToStr(m)+' 分钟 '+IntToStr(s)+' 秒钟 ';
end;

function GetTextFromFile(AFile : String; var ReturnString : string) : boolean;
var
FileStream : TFileStream;
begin
  result := false;
  if not fileexists(AFile) then exit;
  FileStream := TFileStream.Create(AFile,fmopenreadwrite);
  try
    if FileStream.Size > 0 then
    begin
      SetLength(ReturnString,FileStream.size);
      FileStream.Read(ReturnString[1],FileStream.Size);
      result := true;
    end
  finally
    FileStream.Free;
  end;

end;

procedure Tfrm_convert.btn_timeloopClick(Sender: TObject);
var
  reppath:string;
  repfile:string;
  number2:integer;
  //head1:integer;
  hzdata:string;
  hz1number:integer;
  delfile2:string;//生成删除文件名
begin
  if flag1 then
  begin
    //Memo_recond.Lines.Append('序号,车速,载重,车型,车号,自重,皮重,盈亏,日期,时间');
    StatusBar1.Panels[0].Text:='开始纪录时间：'+DateTimeToStr(now);
    StatusBar1.Panels[1].Text:=uptime;
    {
     静态衡车号识别系统是按照累加的方式来进行报文的生成！
     如果已经生成了hao138则下一次过车就是hao139!
     还有就是遇到hao999则下一次过车就是hao001！
     但是在现场会出现拉空车进入装车口，从而多出一列空车记录！

    }
    //取最后过车的报文记录数
    number1:=inifile_total.ReadInteger('last_rep','ltp1',flag_number);
    del_num:=inifile_total.ReadInteger('last_rep','ltp2',2);//删除报文计数

    //取华章数据####################################################需要修改
    adoquery_max.Close;
    adoquery_max.SQL.Clear;
    adoquery_max.SQL.Add('select Max(ord) from FloatTable');
    adoquery_max.Open;
    hz1number:=adoquery_max.Fields[0].Value;

    //
    adoquery_hz.Close;
    adoquery_hz.SQL.Clear;
    adoquery_hz.SQL.Add('select Val,ord from FloatTable where Val<>0 and ord>'+IntToStr(hz1number-100));
    adoquery_hz.Open;


    for number2:=number1 to 999 do
    begin
      if number2<10 then
      begin
        reppath:=subsection2+'\hao00'+IntToStr(number2)+'.rep';
        delfile2:='del c:\tran\chehao'+'\hao00'+IntToStr(number2)+'.rep';
      end
      else
      begin
        if number2<100 then
        begin
          reppath:=subsection2+'\hao0'+IntToStr(number2)+'.rep';
          delfile2:='del c:\tran\chehao'+'\hao0'+IntToStr(number2)+'.rep';
        end
        else
        begin
          //报文数尾
          if number2=999 then
          begin
            //把结尾数写进INI文件
            inifile_total.WriteInteger('last_rep','ltp1',1);
            reppath:=subsection2+'\hao'+IntToStr(number2)+'.rep';
            delfile2:='del c:\tran\chehao'+'\hao'+IntToStr(number2)+'.rep';
          end
          else
          begin
            reppath:=subsection2+'\hao'+IntToStr(number2)+'.rep';
            delfile2:='del c:\tran\chehao'+'\hao'+IntToStr(number2)+'.rep';
          end
        end;
      end;//生成删除报文列表

      if not FileExists(reppath)then//找不到文件
      begin
        //Application.MessageBox('没有找到HTK-196系统生成的文件！','提示',MB_OK+MB_ICONERROR);

        flag_number:=number2;
        StatusBar1.Panels[2].Text:='本次共有'+IntToStr(flag_number-number1)+'节过衡';
        //head1:=1;
        flag1:=False;
        inifile_total.WriteInteger('last_rep','ltp1',flag_number);
        //Break;

        
        //写数据
        Memo_recond.Lines.SaveToFile(compound_path.Text);
        Memo_recond.Lines.SaveToFile(bakSubsection03);
        //写CRC32值

        h1:=0;
        try
        h1:=LoadLibrary('FCN.dll');

        if h1<>0 then
          @saveFCNA:=GetprocAddress(h1,'saveFCN');
        if (@saveFCNA<>nil)then
          saveFCNA(CRCsection1,compound_path.Text);
        finally
          FreeLibrary(h1);
        end;
        inifile_total.WriteInteger('last_rep','ltp2',1);

        //写批处理文件
        creat_delfile.Lines.SaveToFile(ExtractFilePath(ParamStr(0))+'del_all.bat');

        
        //写完数据进行发送（暂时不用）
        //ShellExecute(frm_convert.Handle,'open','Send.exe',nil,'c:\Receive',SW_SHOWNORMAL);
        //删除现存所有文件
        if del_num=1 then
        begin
          ShellExecute(Application.Handle,'open','del_all.bat',nil,'c:\convert',SW_SHOWNORMAL);
          //清空数据然后再写数据
          Memo_recond.Lines.Clear;

          Q1Click(nil);
        end;
        Exit;
      end;

      //循环读取rep文件
      recondfile.LoadFromFile(reppath);
      repfile:=recondfile.Strings[0];
      //删除文件列表生成
      creat_delfile.Lines.Append(delfile2);

      
      //取华章数据写进MEMO中
      try
        hzdata:=copy((adoquery_hz.Fields[0].Text),1,6);
        if adoquery_hz.RecordCount-1<number2-number1 then
        begin
          Memo_recond.Lines.Append(IntToStr(number2)+',0'+','+repfile);
         //显示数据，不写文件！
          Memo_view.Lines.Append(IntToStr(number2)+',0'+','+repfile);
        end
        else
        begin
          Memo_recond.Lines.Append(IntToStr(number2)+','+hzdata+','+repfile);
          //显示数据，不写文件！
          Memo_view.Lines.Append(IntToStr(number2)+','+hzdata+','+repfile);
        end;
        adoquery_hz.Next;
      except
        Application.MessageBox('取华章数据出错！','提示',MB_OK+MB_ICONINFORMATION);
        Break;
      end;

      repfile:='';
      //写删除文件计数
      inifile_total.WriteInteger('last_rep','ltp2',1);
    end;

  end; //if end
end;

procedure Tfrm_convert.Timer_continueTimer(Sender: TObject);
var
  continuefile:string;
begin
  if flag_number<10 then
  begin
    continuefile:=subsection2+'\hao00'+IntToStr(flag_number)+'.rep';
  end
  else
  begin
    if flag_number<100 then
    begin
      continuefile:=subsection2+'\hao0'+IntToStr(flag_number)+'.rep';
    end
    else
    begin
      continuefile:=subsection2+'\hao'+IntToStr(flag_number)+'.rep';
    end;
  end;
  StatusBar1.Panels[0].Text:='正在等待下一次过车... ...';
  if FileExists(continuefile)then
  begin
    Memo_recond.Lines.Clear;
    flag1:=True;
    StatusBar1.Panels[0].Text:='正在过车... ...';
    btn_timeloopClick(nil);
  end;
end;

procedure Tfrm_convert.Timer_statusTimer(Sender: TObject);
var
  t1:Real;
begin
  StatusBar1.Panels[3].Text:=DateTimeToStr(now);
  StatusBar1.Panels[1].Text:=uptime;
  //new add 2006.1.22

  t1:=(t2-1000)/1000;
  StatusBar1.Panels[2].Text:='检测数据：'+FloatToStr(t1)+'秒';
  if t1=0 then
  begin
    t2:=Timer_flag.Interval;
  end
  else
  begin
    t2:=t1*1000;
  end;

  CoolTrayIcon1.IconVisible:=True;
  CoolTrayIcon1.MinimizeToTray:=True;
  frm_convert.Visible:=False;

end;

procedure Tfrm_convert.adoconn_hzAfterConnect(Sender: TObject);
begin
  //写连接串
  try
    Memo1.Lines.Clear;
    Memo1.Lines.Add(adoconn_hz.ConnectionString);
    Memo1.Lines.SaveToFile(xiaobin_path);
  finally

  end;


end;

procedure Tfrm_convert.Timer_flagTimer(Sender: TObject);
var
  hz2number:integer;
  sqlstr1,sqlstr3:string;
begin
  //限制数据集大小
   adoquery_max.Close;
   adoquery_max.SQL.Clear;
   adoquery_max.SQL.Add('select Max(ord) from FloatTable');
   adoquery_max.Open;
   hz2number:=adoquery_max.Fields[0].Value;

    //
   sqlstr1:='select Val,ord,TagIndex from FloatTable where ord>'+IntToStr(hz2number-100);
   //sqlstr2是adoquery_continue的sql语句
   //sqlstr2:='select Val,TagIndex from FloatTable where Val=1 and TagIndex<98';
   sqlstr3:='select val,ord,TagIndex from('+sqlstr1+')where Val=1 and TagIndex<100';
   adoquery_hz.Close;
   adoquery_hz.SQL.Clear;
   adoquery_hz.SQL.Add(sqlstr1);
   adoquery_hz.Open;


  //启动循环条件
  adoquery_continue.Close;
  adoquery_continue.SQL.Clear;
  adoquery_continue.SQL.Add(sqlstr3);
  adoquery_continue.Open;

  if adoquery_continue.Fields[0].Text='1' then
  begin
    //flag1:=True;
    btn_timeloopClick(nil);
    //清空显示的条件
    if Memo_view.Lines.Count>999 then
    begin
      Memo_view.Lines.SaveToFile('c:\f\datasb1.xbf'); 
      Memo_view.Lines.Clear;
    end;
  end;
  
end;

procedure Tfrm_convert.S1Click(Sender: TObject);
begin
  CoolTrayIcon1.IconVisible:=False;
  CoolTrayIcon1.MinimizeToTray:=False;
  frm_convert.Visible:=True;
end;

procedure Tfrm_convert.S2Click(Sender: TObject);
begin
  BitBtn_run.Enabled:=True;
  BitBtn_set.Enabled:=True;
  mdb_path.Enabled:=True;
  rep_path.Enabled:=True;
  compound_path.Enabled:=True;
  Timer_loop.OnTimer:=nil;
end;

procedure Tfrm_convert.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  inifile_total.WriteString('excel_last','elt1',lastfile_xls);
  inifile_total.WriteInteger('last_rep','ltp2',1);//删除报文计数
  //inifile_total.WriteInteger('last_rep','ltp1',flag_number);
  inifile_total.Destroy;
  recondfile.Destroy;
  //2006.6.24
  ShellExecute(Application.Handle,'open','del_all2.bat',nil,'c:\convert',SW_SHOWNORMAL);
end;

procedure Tfrm_convert.Q1Click(Sender: TObject);
begin
  Close;
end;

end.
