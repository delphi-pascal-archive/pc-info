unit PC_Info;

interface

uses
  SysUtils, Classes, Controls, Grids, ValEdit,
  WinSock,                                      {HOSTName}
  Menus,                                        {PouUpMenu}
  WbemScripting_TLB,smartapi,magwmi,magsubs1;   {WMI ...}

type
  TPC_Info = class(TValueListEditor)

  private

    FPopUpMenu:TPopUpMenu;       // Для отобажения меню в компоненте

          FName          : String;     // Name
          FLocation      : String;     // Location
{Auto}    GetCaption       : TStrings;   // HostName, UserName, Model, Manufacturer
{Auto}    FHostname        : String;     // procedure SetCaption (Caption)
{Auto}    FUserName_l      : String;     // procedure SetCaption (UserName)
{    }    FUserName_nw   : String;     // Other User Name (NW)
{Auto}    GetBIOS          : TStrings;   // pcBIOSVer,pcSMBIOSBIOSVer,pcVer,pcReleaseDate,pcManufacturer,pcSerialNumber
{Auto}    FBIOS_ver        : String;     // procedure SetBIOS (Version)
{Auto}    FBIOS_date       : String;     // pcBIOSVer
{Auto}    FModel           : String;     // procedure SetCaption (Model)
{Auto}    FSerialNumber    : String;     // procedure SetBIOS (pcSerialNumber)
          FPartNumber      : String;     // Patr Number
{Auto}    GetMatherBoard   : TStrings;   // pcManufacturer,pcSerialNumber,pcPartNumber,pcProduct
{Auto}    GetCPU           : TStrings;   // pcCaption,pcDeviceID,pcPN,pcSocketDesignation,pcManufacturer,pcName,
                                         // pcMaxClockSpeed,pcCurrentClockSpeed,pcNumberOfCores,pcNumberOfLogicalProcessors,
                                         // pcL2CacheSize,pcProcessorId,pcAllCPU
{Auto}    FCPU_All         : String;     // procedure SetCPU
{Auto}    GetRAM           : TStrings;   // pcTotalVirtualMemorySize,pcTotalVisibleMemorySize,pcTotalSwapSpaceSize,
                                         // pcFreePhysicalMemory,pcFreeVirtualMemory,pcFreeSpaceInPagingFiles
{Auto}    FRAM_Size        : String;     // procedure SetRAM (TotalVisibleMemorySize)
{Auto}    FRAM_Swap_Size   : String;     // procedure SetRAM (pcTotalSwapSpaceSize)
{Auto}    GetRAM_Phisic    : TStrings;   // pcCaption,pcDeviceLocator,pcSpeed
{Auto}    FRAM_Speed       : String;     // procedure SetRAM_Phisic (Speed)
{Auto}    FSwap            : String;     // Disk Swap
{Auto}    GetVideo_Adapter : TStrings;   // pcAdapterCompatibility,pcAdapterRAM,pcCaption,pcVideoProcessor,
                                         // pcCurrentBitsPerPixel,pcVideoModeDescription,pcMaxRefreshRate,
                                         // pcCurrentRefreshRate,pcDriverDate,pcInfFilename,pcInstalledDisplayDrivers
{Auto}    FVideo           : String;     // procedure Video_adapter (pcCaption,pcAdapterRAM)
{Auto}    GetAudio_Adapter : TStrings;   // pcManufacturer,pcName,pcDeviceID
{Auto}    FAudio           : String;     // procedure Audio_adapter (pcName)
{Auto}    GetHDD_Phisic    : TStrings;   // pcModel,pcSize,pcInterfaceType,pcManufacturer,pcName,pcTotalCylinders,
                                         // pcTotalHeads,pcTotalSectors,pcTotalTracks,pcTracksPerCylinder,pcBytesPerSector,
                                         // pcSectorsPerTrack,pcSignature
{Auto}    FHDD_p           : String;     // procedure SetHDD_Phisic (Model)
{Auto}    GetHDD_Map       : TStrings;   // pcName,pcProviderName,pcVolumeName,pcFileSystem,pcVolumeSerialNumber,
                                         // pcSize,pcFreeSpace,pcSessionID
{Auto}    FHDD_m           : String;     // procedure SetHDD_Map (pcProviderName)
{Auto}    GetHDD_Vol       : TStrings;   // Volum(C:..), Size, ..
{Auto}    FHDD_v           : String;     // procedure SetHDD_Vol ()
{Auto}    FDVD             : String;     // Volum(C:..), Name(Firm)
{Auto}    GetLAN_Adapter   : TStrings;   // pcName,pcManufacturer,NetConnectionID,MACAddress
{Auto}    GetNetwork       : TStrings;   // IPAddress,MACAddress,cpIPSubnet,
                                         // pcDefaultIPGateway,pcDHCPServer,pcDNSHostName,DNSDomain,pcDNSServerSearchOrder,
                                         // pcServiceName
{Auto}    FIP              : String;     // procedure SetNetwork (IPaddress)
{Auto}    FMAC             : String;     // procedure SetNetwork (MACaddress)
{Auto}    FGroup           : String;     // procedure SetNetConnect (Domain)
{Auto}    FDNSServer       : String;     // procedure SetNetwork (pcDNSServerSearchOrder)
{Auto}    FDNSDomain       : String;     // procedure SetNetwork (DNSDomain)
{Auto}    FDNSHostName     : String;     // procedure SetNetwork (pcDNSHostName)
{Auto}    FIPSubnet        : String;     // procedure SetNetwork (cpIPSubnet)
{Auto}    FGateway         : String;     // procedure SetNetwork (pcDefaultIPGateway)
{Auto}    FDHCPServer      : String;     // procedure SetNetwork (pcDHCPServer)
{Auto}    GetNetConnect    : TStrings;   // pcName,pcConnectionType,pclocalName,pcUserName,pcResourceType,pcDisplayType
{Auto}    FConnectNeme     : String;     // procedure SetNetConnect(Name)
{Auto}    GetIDE_USBController: TStrings;// IDE_USB
{Auto}    FOS              : String;     // procedure SetOS()
{Auto}    GetOS            : TStrings;   // pcManufacturer,pcCaption,pcCSDVersion,pcBuildNumber,pcBuildType,pcVersion,
                                         // pcBootDevice,pcName,pcSystemStartupOptions,pcSystemDrive,pcWindowsDirectory,
                                         // pcSystemDevice,pcSystemDirectory,pcTempDirectory,pcOSLanguage,pcSystemType,
                                         // pcLocale,pcCSName,pcOrganization,pcLastBootUpTime,pcLocalDateTime
{Auto}    GetOS_Variable   : TStrings;   // OS Path
{Auto}    GetUserAccount   : TStrings;   // List_UserAccount (pcName,pcSID,pcDisabled)
{Auto}    GetSystemAccount : TStrings;   // List_UserAccount (pcName,pcSID)
{Auto}    GetGroupName     : TStrings;   // List_GroupName   (pcName,pcSID)
{Auto}    GetProcess       : TStrings;   // List_Process     (pcName,pcSID)
{Default} FSaveToPath      : String;     // Path to Save File (Default = C:\)

    procedure SetCaption(Value: TStrings);
    procedure Caption(Value: String);
    procedure SetBIOS(Value: TStrings);
    procedure BIOS(Value: String);
    procedure SetMatherBoard(Value: TStrings);
    procedure MatherBoard(Value: String);
    procedure SetCPU(Value: TStrings);
    procedure CPU(Value: String);
    procedure SetRAM(Value: TStrings);
    procedure RAM(Value: String);
    procedure SetRAM_Phisic(Value: TStrings);
    procedure RAM_Phisic(Value: String);
    procedure SetVideo_Adapter(Value: TStrings);
    procedure Video_Adapter(Value: String);
    procedure SetAudio_Adapter(Value: TStrings);
    procedure Audio_Adapter(Value: String);
    procedure SetHDD_Phisic(Value: tStrings);
    procedure HDD_Phisic(Value: String);
    procedure SetHDD_Map(Value: TStrings);
    procedure HDD_Map(Value: String);
    procedure SetHDD_Vol(Value: TStrings);
    procedure HDD_Vol(Value: String);
//    procedure SetDVD(Value: String);
    procedure SetLAN_Adapter(Value: TStrings);
    procedure LAN_Adapter(Value: String);
    procedure SetNetwork(Value: TStrings);
    procedure Network(Value: String);
    procedure SetNetConnect(Value: TStrings);
    procedure NetConnect(Value: String);
    procedure SetIDE_USBController(Value: TStrings);
    procedure IDE_USBController(Value: String);
    procedure SetOS(Value: TStrings);
    procedure OS(Value: String);
    procedure SetOS_Variable(Value: TStrings);
    procedure OS_Variable (Value: String);
    procedure SetUserAccount(Value: TStrings);
    procedure UserAccount(Value: String);
    procedure SetSystemAccount(Value: TStrings);
    procedure SystemAccount(Value: String);
    procedure SetGroupName(Value: TStrings);
    procedure GroupName(Value: String);
    procedure SetProcess(Value: TStrings);
    procedure Process(Value: String);

  protected
    //PopUpMenu1: TPopUpMenu;//------------------------------------
    procedure PCreatePopUpMenu;                          // Создаю Меню для объекта
    procedure MenuItem_BIOSClick(Sender:TObject);        // реакция на пунты меню
    procedure MenuItem_MatherBoardClick(Sender:TObject); // реакция на пунты меню
    procedure MenuItem_CPUClick(Sender:TObject);         // реакция на пунты меню
    procedure MenuItem_RAMClick(Sender:TObject);         // реакция на пунты меню
    procedure MenuItem_VideoClick(Sender:TObject);       // реакция на пунты меню
    procedure MenuItem_AudioClick(Sender:TObject);       // реакция на пунты меню
    procedure MenuItem_HDDClick(Sender:TObject);         // реакция на пунты меню
    procedure MenuItem_NetworkClick(Sender:TObject);     // реакция на пунты меню
    procedure MenuItem_IDE_USBControllerClick(Sender:TObject);// реакция на пунты меню
    procedure MenuItem_OSClick(Sender:TObject);          // реакция на пунты меню
    procedure MenuItem_ProcessClick(Sender:TObject);     // реакция на пунты меню
    procedure MenuItem_CleanClick(Sender:TObject);       // реакция на пунты меню
    procedure MenuItem_SaveClick(Sender:TObject);        // реакция на пунты меню
    procedure MenuItem_MailClick(Sender:TObject);        // реакция на пунты меню


      //PopUpMenu1: TPopUpMenu;//------------------- END ------------


  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;

  published
    property PC_Name           :String read FName write FName;
    property PC_Location       :String read FLocation write FLocation;
    //property PC_Caption_        :String read GetCaption write SetCaption;
    property PC__Caption       :TStrings read GetCaption write SetCaption;
    property PC_HostName       :String read FHostname write FHostname;
    property PC_UserName       :String read FUserName_l write FUserName_l;
    property PC_UserName_nw    :String read FUserName_nw write FUserName_nw;
//    property PC_BIOS           :String read FBIOS  write FBIOS;
    property PC__BIOS          :TStrings read GetBIOS  write SetBIOS;
    property PC_BIOS_ver       :String read FBIOS_ver write FBIOS_ver;
    property PC_BIOS_date      :String read FBIOS_date write FBIOS_date;
    property PC_Model          :String read FModel write FModel;
    property PC_SeialNumber    :String read FSerialNumber write FSerialNumber;
    property PC_PartNumber     :String read FPartNumber write FPartNumber;
    property PC__MatherBoard   :TStrings read GetMatherBoard write SetMatherBoard;
    property PC__CPU           :TStrings read GetCPU write SetCPU;
    property PC_inf_CPU        :String read FCPU_All write FCPU_All;
    property PC__RAM           :TStrings read GetRAM write SetRAM;
    property PC_RAM_Size       :String read FRAM_Size write FRAM_Size;
    property PC_RAM_Swap_Size  :String read FRAM_Swap_Size write FRAM_Swap_Size;
    property PC__RAM_Phisic    :TStrings read GetRAM_Phisic write SetRAM_Phisic;
    property PC_RAM_Speed      :String read FRAM_Speed write FRAM_Speed;
    property PC_Swap           :String read FSwap  write FSwap;
    property PC__Video_Adapter :TStrings read GetVideo_Adapter write SetVideo_Adapter;
    property PC_Video          :String read FVideo write FVideo;
    property PC__Audio_Adapter :TStrings read GetAudio_Adapter write SetAudio_Adapter;
    property PC_Audio          :String read FAudio write FAudio;
    property PC__HDD_Phisic    :TStrings read GetHDD_Phisic write SetHDD_Phisic;
    property PC_HDD_p          :String read FHDD_p write FHDD_p;
    property PC__HDD_Map       :TStrings read GetHDD_Map write SetHDD_Map;
    property PC_HDD_m          :String read FHDD_m write FHDD_m;
    property PC__HDD_Vol       :TStrings read GetHDD_Vol write SetHDD_Vol;
    property PC_HDD_v          :String read FHDD_v write FHDD_v;
    property NET__LAN_Adapter  :TStrings read GetLAN_Adapter write SetLAN_Adapter;
    property NET__Network      :TStrings read GetNetwork write SetNetwork;
    property NET_IP            :String read FIP write FIP;
    property NET_MAC           :String read FMAC write FMAC;
    property NET_Group         :String read FGroup write FGroup;
    property NET_DNSServer     :String read FDNSServer write FDNSServer;
    property NET_DNSDomain     :String read FDNSDomain write FDNSDomain;
    property NET_DNSHostName   :String read FDNSHostName write FDNSHostName;
    property NET_IPSubnet      :String read FIPSubnet write FIPSubnet;
    property NET_Gateway       :String read FGateway write FGateway;
    property NET_DHCPServer    :String read FDHCPServer write FDHCPServer;
    property NET__NetConnect   :TStrings read GetNetConnect write SetNetConnect;
    property NET_FConnectNeme  :String read FConnectNeme write FConnectNeme;
    property PC__IDE_USBController:TStrings read GetIDE_USBController write SetIDE_USBController;
    property OS__OS_About      :TStrings read GetOS write SetOS;
    property OS_OS             :String read FOS write FOS;
    property OS__OS_Variable   :TStrings read GetOS_Variable write SetOS_Variable;
    property OS__UserAccount   :TStrings read GetUserAccount write SetUserAccount;
    property OS__SystemAccount :TStrings read GetSystemAccount write SetSystemAccount;
    property OS__GroupName     :TStrings read GetGroupName write SetGroupName;
    property OS__Process       :TStrings read GetProcess write SetProcess;
    property PC_SaveToPath     :String read FSaveToPath write FSaveToPath;
//    property DVD            :String read FDVD write SetDVD;
  end;

procedure Register;

implementation
var
    MItem:TMenuItem; // описываю переменные для работы с подключенным TMenuItem
    vCaption:array [0..4] of string;
    vCaption_v:array [0..4] of string;
    vBIOS: array [0..7] of string;
    vBIOS_v: array [0..7] of string;
    vMatherBoard: array [0..8] of string;
    vMatherBoard_v: array [0..8] of string;
    vCPU: array [0..11] of string;
    vCPU_v: array [0..11] of string;
    vRAM: array [0..5] of string;
    vRAM_v: array [0..5] of string;
    vRAM_Phisic: array [0..2] of string;
    vRAM_Phisic_v: array [0..2] of string;
    vVideo_Adapter: array [0..14] of string;
    vVideo_Adapter_v: array [0..14] of string;
    vAudio_Adapter: array [0..2] of string;
    vAudio_Adapter_v: array [0..2] of string;
    vHDD_Phisic: array [0..12] of string;
    vHDD_Phisic_v: array [0..12] of string;
    vHDD_Map: array [0..7] of string;
    vHDD_Map_v: array [0..7] of string;
    vHDD_Vol: array [0..11] of string;
    vHDD_Vol_v: array [0..11] of string;
    vLAN_Adapter: array [0..5] of string;
    vLAN_Adapter_v: array [0..5] of string;
    vNetwork: array [0..8] of string;
    vNetwork_v: array [0..8] of string;
    vNetConnect: array [0..6] of string;
    vNetConnect_v: array [0..6] of string;
    vIDE_USBController: array [0..1] of string;
    vIDE_USBController_v: array [0..1] of string;
    vOS: array [0..20] of string;
    vOS_v: array [0..20] of string;
    vOS_Variable: array [0..2] of string;
    vOS_Variable_v: array [0..2] of string;
    vUserAccount: array [0..2] of string;
    vUserAccount_v: array [0..2] of string;
    vSystemAccount: array [0..1] of string;
    vSystemAccount_v: array [0..1] of string;
    vGroupName: array [0..1] of string;
    vGroupName_v: array [0..1] of string;
    vProcess: array [0..3] of string;
    vProcess_v: array [0..3] of string;





/////////////////////////////////////////////////////////

constructor TPC_Info.Create(AOwner: TComponent);
Begin
  inherited Create(AOwner);

  Self.FPopUpMenu:=TpopUpMenu.Create(Self);
  PCreatePopUpMenu;

  self.GetCaption         := TStringList.Create;
  self.GetBIOS            := TStringList.Create;
  self.GetMatherBoard     := TStringList.Create;
  self.GetCPU             := TStringList.Create;
  Self.GetRAM             := TStringList.Create;
  self.GetRAM_Phisic      := TStringList.Create;
  self.GetVideo_Adapter   := TStringList.Create;
  self.GetAudio_Adapter   := TStringList.Create;
  self.GetHDD_Phisic      := TStringList.Create;
  self.GetHDD_Map         := TStringList.Create;
  self.GetHDD_Vol         := TStringList.Create;
  self.GetLAN_Adapter     := TStringList.Create;
  self.GetNetwork         := TStringList.Create;
  self.GetNetConnect      := TStringList.Create;
  self.GetIDE_USBController := TStringList.Create;
  self.GetOS              := TStringList.Create;
  self.GetOS_Variable     := TStringList.Create;
  self.GetUserAccount     := TStringList.Create;
  self.GetSystemAccount   := TStringList.Create;
  self.GetGroupName       := TStringList.Create;
  self.GetProcess         := TStringList.Create;

  //Interface
  Self.ShowHint:=True;
  Self.TitleCaptions.Clear;
  Self.InsertRow('','',True);
  Self.Hint:=' PC_Info ';
  Self.FSaveToPath := 'C:\';
end;


//////////////////////////////////////////////////
// описываю поля TStings

procedure TPC_Info.SetCaption(Value: TStrings);
begin
GetCaption.Assign(Value);
end;

procedure TPC_Info.SetBIOS(Value: TStrings);
begin
GetBIOS.Assign(Value);
end;

procedure TPC_Info.SetMatherBoard(Value: TStrings);
begin
GetMatherBoard.Assign(Value);
end;

procedure TPC_Info.SetCPU(Value: TStrings);
begin
GetCPU.Assign(Value);
end;

procedure TPC_Info.SetRAM(Value: TStrings);
begin
GetRAM.Assign(Value);
end;

procedure TPC_Info.SetRAM_Phisic(Value: TStrings);
begin
GetRAM_Phisic.Assign(Value);
end;

procedure TPC_Info.SetVideo_Adapter(Value: TStrings);
begin
GetVideo_Adapter.Assign(Value);
end;

procedure TPC_Info.SetAudio_Adapter(Value: TStrings);
begin
GetAudio_Adapter.Assign(Value);
end;

procedure TPC_Info.SetHDD_Phisic(Value: TStrings);
begin
GetHDD_Phisic.Assign(Value);
end;

procedure TPC_Info.SetHDD_Map(Value: TStrings);
begin
GetHDD_Map.Assign(Value);
end;

procedure TPC_Info.SetHDD_Vol(Value: TStrings);
begin
GetHDD_Vol.Assign(Value);
end;

procedure TPC_Info.SetLAN_Adapter(Value: TStrings);
begin
GetLAN_Adapter.Assign(Value);
end;

procedure TPC_Info.SetNetwork(Value: TStrings);
begin
GetNetwork.Assign(Value);
end;

procedure TPC_Info.SetNetConnect(Value: TStrings);
begin
GetNetConnect.Assign(Value);
end;

procedure TPC_Info.SetIDE_USBController(Value: TStrings);
begin
GetIDE_USBController.Assign(Value);
end;

procedure TPC_Info.SetOS(Value: TStrings);
begin
GetOS.Assign(Value);
end;

procedure TPC_Info.SetOS_Variable(Value: TStrings);
begin
GetOS_Variable.Assign(Value);
end;

procedure TPC_Info.SetUserAccount(Value: TStrings);
begin
GetUserAccount.Assign(Value);
end;

procedure TPC_Info.SetSystemAccount(Value: TStrings);
begin
GetSystemAccount.Assign(Value);
end;

procedure TPC_Info.SetGroupName(Value: TStrings);
begin
GetGroupName.Assign(Value);
end;

procedure TPC_Info.SetProcess(Value: TStrings);
begin
GetProcess.Assign(Value);
end;



///////////////////////////////
//                           //
//            START          //
//                           //
//        PC_INFO WMI        //
//                           //
///////////////////////////////
///////////////////////////////
  /////////////////////////
       ///////////////

//--- PC_HostName,Model,Manufacturer ---
procedure TPC_Info.Caption(Value: String);
var
    rows, instances, I, J, a: integer ;
    WmiResults: T2DimStrArray ;
//    pcHostName,pcUserName,pcModel,pcManufacturer,pcDomain:string;
begin
if GetCaption.Count < 1 then
 begin //Count
  rows := MagWmiGetInfo ('.', 'root\CIMV2', '', '',
    'Win32_ComputerSystem',
    WmiResults, instances) ;
  if rows > 0 then
    begin
      for J := 1 to instances do
        begin
          for I := 1 to rows do
            begin ////////////////////////////////////////////////////////
              for a := 0 to 4 do
                begin
                  if WmiResults [0, I] = vCaption[a] then vCaption_v[a]:= WmiResults [J, I];
                end;
            end; // I ////////////////////////////////////////////////////
        end ; // J
    end; //rows
  end; // Count
end; //--- PC_HostName,Model,Manufacturer,pcDomain ---END


//--- BIOS (pcBIOSVer,pcSMBIOSBIOSVer,pcVer,pcReleaseDate,pcManufacturer,pcSerialNumber)---


//procedure TPC_Info.SetBIOS(Value: String);
procedure TPC_Info.BIOS(Value: String);
var
    rows, instances, I, J,a: integer ;
    WmiResults: T2DimStrArray ;
//    pcBIOSVer,pcSMBIOSBIOSVer,pcVer,pcReleaseDate,pcManufacturer,pcSerialNumber:string;
begin
if GetBIOS.Count < 1 then
 begin //Count
    rows := MagWmiGetInfo ('.', 'root\CIMV2', '', '',
    'Win32_BIOS',
    WmiResults, instances) ;
  if rows > 0 then
    begin
      for J := 1 to instances do
        begin
          for I := 1 to rows do
            begin ////////////////////////////////////////////////////////
              for a := 0 to 7 do
                begin
                  if WmiResults [0, I] = vBIOS[a] then vBIOS_v[a]:= WmiResults [J, I];
                end;
            end; // I ////////////////////////////////////////////////////
        end ; // J
    end; //rows
  end; //Count
end; //--- BIOS (pcBIOSVer,pcSMBIOSBIOSVer,pcVer,pcReleaseDate,pcManufacturer,pcSerialNumber)---END

//--- MatherBoard  (pcManufacturer,pcSerialNumber,pcPartNumber,pcProduct)---
procedure TPC_Info.MatherBoard(Value: String);
var
    rows, instances, I, J, a: integer ;
    WmiResults: T2DimStrArray ;
//    pcManufacturer,pcSerialNumber,pcPN,pcProduct:string;
begin
if GetMatherBoard.Count < 1 then
 begin //Count
  rows := MagWmiGetInfo ('.', 'root\CIMV2', '', '',
    'win32_baseboard',
    WmiResults, instances) ;
  if rows > 0 then
    begin
      for J := 1 to instances do
        begin
          for I := 1 to rows do
            begin ////////////////////////////////////////////////////////
              for a := 0 to 3 do
                begin
                  if WmiResults [0, I] = vMatherBoard[a] then vMatherBoard_v[a]:= WmiResults [J, I];
                end;
            end; // I ////////////////////////////////////////////////////
        end ; // J
    end; //rows
    rows := MagWmiGetInfo ('.', 'root\CIMV2', '', '',
    'win32_motherboarddevice',
    WmiResults, instances) ;
  if rows > 0 then
    begin
      for J := 1 to instances do
        begin
          for I := 1 to rows do
            begin ////////////////////////////////////////////////////////
              for a := 4 to 8 do
                begin
                  if WmiResults [0, I] = vMatherBoard[a] then vMatherBoard_v[a]:= WmiResults [J, I];
                end;
            end; // I ////////////////////////////////////////////////////
        end ; // J
    end; //rows
  end // Count
end; //--- MatherBoard (pcManufacturer,pcSerialNumber,pcPartNumber,pcProduct)---END


//--- CPU (pcCaption,pcDeviceID,pcPN,pcSocketDesignation,pcManufacturer,pcName,pcMaxClockSpeed,pcCurrentClockSpeed,pcNumberOfCores,pcNumberOfLogicalProcessors,pcL2CacheSize,pcProcessorId,pcAllCPU)---
procedure TPC_Info.CPU(Value: String);
var
    rows, instances, I, J, a: integer ;
    WmiResults: T2DimStrArray ;
 //   pcManufacturer,pcName,pcMaxClockSpeed,pcAllCPU :String;
//    pcCaption,pcDeviceID,pcSocketDesignation,pcManufacturer,pcName,pcMaxClockSpeed,pcCurrentClockSpeed,
//    pcNumberOfCores,pcNumberOfLogicalProcessors,pcL2CacheSize,pcRevision,pcProcessorId, pcAllCPU   :string;
begin
if GetCPU.Count < 1 then
 begin //Count
  rows := MagWmiGetInfo ('.', 'root\CIMV2', '', '',
    'Win32_Processor',
    WmiResults, instances) ;
  if rows > 0 then
    begin
      for J := 1 to instances do
        begin
          for I := 1 to rows do
            begin ////////////////////////////////////////////////////////
              for a := 0 to 11 do
                begin
                  if WmiResults [0, I] = vCPU[a] then vCPU_v[a]:=vCPU_v[a] + WmiResults [J, I] + ' | ';
                end; //IntToStr(StrToInt(pcNumberOfCores) * StrToInt(pcNumberOfLogicalProcessors)) ;
            end; // I ////////////////////////////////////////////////////
         FCPU_All :=vCPU_v[3] + ' ' + vCPU_v[4] + '    ' + vCPU_v[5] + 'MHz  ';
         end ; // J
    end; //rows
  end; // Count
end; //--- CPU (pcCaption,pcDeviceID,pcPN,pcSocketDesignation,pcManufacturer,pcName,pcMaxClockSpeed,pcCurrentClockSpeed,pcNumberOfCores,pcNumberOfLogicalProcessors,pcL2CacheSize,pcProcessorId,pcAllCPU)---END

// --- RAM(pcTotalVirtualMemorySize,pcTotalVisibleMemorySize,pcTotalSwapSpaceSize,pcFreePhysicalMemory,pcFreeVirtualMemory,pcFreeSpaceInPagingFiles)---
procedure TPC_Info.RAM(Value: String);
var
    rows, instances, I, J, a: integer ;
    WmiResults: T2DimStrArray ;
//    pcTotalVisibleMemorySize:String;
begin
if GetRAM.Count < 1 then
 begin //Count
 rows := MagWmiGetInfo ('.', 'root\CIMV2', '', '',
    'Win32_OperatingSystem',
    WmiResults, instances) ;
  if rows > 0 then
    begin
      for J := 1 to instances do
        begin
          for I := 1 to rows do
            begin ////////////////////////////////////////////////////////
              for a := 0 to 5 do
                begin
                  if WmiResults [0, I] = vRAM[a] then vRAM_v[a]:=WmiResults [J, I];
                end;
            end; // I ////////////////////////////////////////////////////
        end ; // J
    end; //rows
  end; //Count
end; //--- RAM (pcTotalVirtualMemorySize,pcTotalVisibleMemorySize,pcTotalSwapSpaceSize,pcFreePhysicalMemory,pcFreeVirtualMemory,pcFreeSpaceInPagingFiles)---END

//--- RAM_Phisic (pcCaption,pcDeviceLocator,pcSpeed) ---
procedure TPC_Info.RAM_Phisic(Value: String);
var
    rows, instances, I, J, a: integer ;
    WmiResults: T2DimStrArray ;
//    pcCaption,pcDeviceLocator,pcSpeed:string;
begin
if GetRAM_Phisic.Count < 1 then
 begin //Count
   rows := MagWmiGetInfo ('.', 'root\CIMV2', '', '',
    'Win32_PhysicalMemory',
    WmiResults, instances) ;
  if rows > 0 then
    begin
      for J := 1 to instances do
        begin
          for I := 1 to rows do
            begin ////////////////////////////////////////////////////////
              for a := 0 to 5 do
                begin
                  if WmiResults [0, I] = vRAM_Phisic[a] then vRAM_Phisic_v[a]:=vRAM_Phisic_v[a] + WmiResults [J, I] + ' | ';
                end;
            end; // I ////////////////////////////////////////////////////
        end ; // J
    end; //rows
  end; //Count
end; //--- RAM_Phisic (pcCaption,pcDeviceLocator,pcSpeed) ---END


// ---Video (pcAdapterCompatibility,pcAdapterRAM,pcCaption,pcVideoProcessor,pcCurrentBitsPerPixel,pcVideoModeDescription,pcMaxRefreshRate,pcCurrentRefreshRate,pcDriverDate,pcInfFilename,pcInstalledDisplayDrivers)---
procedure TPC_Info.Video_Adapter(Value: String);
var
    rows, instances, I, J, a: integer ;
    WmiResults: T2DimStrArray ;
//    pcAdapterRAM,pcCaption :string;
//    pcAdapterCompatibility,pcAdapterRAM,pcCaption,pcVideoProcessor,pcCurrentBitsPerPixel,pcVideoModeDescription,
//    pcMaxRefreshRate,pcCurrentRefreshRate,pcDriverDate,pcInfFilename,pcInstalledDisplayDrivers:string;
begin
if GetVideo_Adapter.Count < 1 then
 begin //Count
  rows := MagWmiGetInfo ('.', 'root\CIMV2', '', '',
    'win32_videocontroller',
    WmiResults, instances) ;
  if rows > 0 then
    begin
      for J := 1 to instances do
        begin
          for I := 1 to rows do
            begin ////////////////////////////////////////////////////////
              for a := 0 to 14 do
                begin
                  if WmiResults [0, I] = vVideo_Adapter[a] then vVideo_Adapter_v[a]:=vVideo_Adapter_v[a] + WmiResults [J, I] + ' | ';
                end;
            end; // I ////////////////////////////////////////////////////
        end ; // J  AdapterRAM
    end; //rows
  end; // Count
end; // ---Video (pcAdapterCompatibility,pcAdapterRAM,pcCaption,pcVideoProcessor,pcCurrentBitsPerPixel,pcVideoModeDescription,pcMaxRefreshRate,pcCurrentRefreshRate,pcDriverDate,pcInfFilename,pcInstalledDisplayDrivers)---END


// ---Audio (pcManufacturer,pcName,pcDeviceID)---
procedure TPC_Info.Audio_Adapter(Value: String);
var
    rows, instances, I, J, a: integer ;
    WmiResults: T2DimStrArray ;
//    pcManufacturer,pcName,pcDeviceID:string;
begin
if GetAudio_Adapter.Count < 1 then
 begin //Count
  rows := MagWmiGetInfo ('.', 'root\CIMV2', '', '',
    'Win32_SoundDevice',
    WmiResults, instances) ;
  if rows > 0 then
    begin
      for J := 1 to instances do
        begin
          for I := 1 to rows do
            begin ////////////////////////////////////////////////////////
              for a := 0 to 2 do
                begin
                  if WmiResults [0, I] = vAudio_Adapter[a] then vAudio_Adapter_v[a]:=WmiResults [J, I];
                end;
            end; // I ////////////////////////////////////////////////////
        end ; // J
    end; //rows
  end; //Cout
end; // ---Audio (pcManufacturer,pcName,pcDeviceID)---END

// ---HDD PHISIC (pcModel,pcSize,pcInterfaceType,pcManufacturer,pcName,pcTotalCylinders,pcTotalHeads,pcTotalSectors,pcTotalTracks,pcTracksPerCylinder,pcBytesPerSector,pcSectorsPerTrack,pcSignature)---
procedure TPC_Info.HDD_Phisic(Value: String);
var
    rows, instances, I, J, a: integer ;
    WmiResults: T2DimStrArray ;
//    pcModel,pcSize,pcInterfaceType,pcManufacturer,pcName,pcTotalCylinders,pcTotalHeads,pcTotalSectors,
//    pcTotalTracks,pcTracksPerCylinder,pcBytesPerSector,pcSectorsPerTrack,pcSignature:string;
begin
if GetHDD_Phisic.Count < 1 then
 begin //Count
   rows := MagWmiGetInfo ('.', 'root\CIMV2', '', '',
    'win32_diskdrive',
    WmiResults, instances) ;
  if rows > 0 then
    begin
      for J := 1 to instances do
        begin
          for I := 1 to rows do
            begin ////////////////////////////////////////////////////////
              for a := 0 to 12 do
                begin
                  if WmiResults [0, I] = vHDD_Phisic[a] then vHDD_Phisic_v[a]:=vHDD_Phisic_v[a] + WmiResults [J, I] + ' | ';
                end;
end; // I ////////////////////////////////////////////////////
//Self.GetVideo_Adapter.Add(#10 + '--- HDD ' + IntTostr(J) + ' ---');
        end ; // J
    end; //rows
  end // Count
end; // ---HDD PHISIC (pcModel,pcSize,pcInterfaceType,pcManufacturer,pcName,pcTotalCylinders,pcTotalHeads,pcTotalSectors,pcTotalTracks,pcTracksPerCylinder,pcBytesPerSector,pcSectorsPerTrack,pcSignature)---END

// ---HDD MAP (pcName,pcProviderName,pcVolumeName,pcFileSystem,pcVolumeSerialNumber,pcSize,pcFreeSpace,pcSessionID)---
procedure TPC_Info.HDD_Map(Value: String);
var
    rows, instances, I, J, a: integer ;
    WmiResults: T2DimStrArray ;
//    HDD_Map_pcName,HDD_Map_pcProviderName,HDD_Map_pcVolumeName,HDD_Map_pcFileSystem,HDD_Map_pcVolumeSerialNumber,HDD_Map_pcSize,HDD_Map_pcFreeSpace,HDD_Map_pcSessionID:string;
begin
if GetHDD_Map.Count < 1 then
 begin //Count
  rows := MagWmiGetInfo ('.', 'root\CIMV2', '', '',
    'Win32_MappedLogicalDisk',
    WmiResults, instances) ;
  if rows > 0 then
    begin
      for J := 1 to instances do
        begin
          for I := 1 to rows do
            begin ////////////////////////////////////////////////////////
            for a := 0 to 7 do
                begin
                  if WmiResults [0, I] = vHDD_Map[a] then vHDD_Map_v[a]:=vHDD_Map_v[a] + WmiResults [J, I] + ' | ';
                end;
            end; // I ////////////////////////////////////////////////////
        end ; // J
    end; //rows
  end // Count
end; // ---HDD Map (pcName,pcProviderName,pcVolumeName,pcFileSystem,pcVolumeSerialNumber,pcSize,pcFreeSpace,pcSessionID)---END

//    --HDD Vol (pcCaption,pcDescription,pcName,pcVolumeName,pcProviderName,pcMediaType,pcDriveType,pcFileSystem,pcFreeSpace,pcSize,pcMaximumComponentLength,pcVolumeSerialNumber)---
procedure TPC_Info.HDD_Vol(Value: String);
var
    rows, instances, I, J, a: integer ;
    WmiResults: T2DimStrArray ;
//    pcCaption,pcDescription,pcName,pcVolumeName,pcProviderName,pcMediaType,pcDriveType,
//    pcFileSystem,pcFreeSpace,pcSize,pcMaximumComponentLength,pcVolumeSerialNumber:string;
begin
if GetHDD_Vol.Count < 1 then
 begin //Count
   rows := MagWmiGetInfo ('.', 'root\CIMV2', '', '',
    'win32_logicaldisk',
    WmiResults, instances) ;
  if rows > 0 then
    begin
      for J := 1 to instances do
        begin
          for I := 1 to rows do
            begin ////////////////////////////////////////////////////////
            for a := 0 to 11 do
                begin
                  if WmiResults [0, I] = vHDD_Vol[a] then vHDD_Vol_v[a]:=vHDD_Vol_v[a] + WmiResults [J, I] + ' | ';
                end;
            end; // I ////////////////////////////////////////////////////
        end ; // J
    end; //rows
  end; // Count
end; // ---HDD Vol (pcCaption,pcDescription,pcName,pcVolumeName,pcProviderName,pcMediaType,pcDriveType,pcFileSystem,pcFreeSpace,pcSize,pcMaximumComponentLength,pcVolumeSerialNumber)---END

// --- LAN (pcName,pcManufacturer,NetConnectionID,MACAddress)---
procedure TPC_Info.LAN_Adapter(Value: String);
var
    rows, instances, I, J, a: integer ;
    WmiResults: T2DimStrArray ;
    //pcName,pcManufacturer,NetConnectionID,MACAddress:string;
begin
if GetLAN_Adapter.Count < 1 then
 begin //Count
  rows := MagWmiGetInfo ('.', 'root\CIMV2', '', '',
    'Win32_NetworkAdapter',
    WmiResults, instances) ;
  if rows > 0 then
    begin
      for J := 1 to instances do
        begin
          for I := 1 to rows do
            begin ////////////////////////////////////////////////////////
            for a := 0 to 5 do
                begin
                  if WmiResults [0, I] = vLAN_Adapter[a] then vLAN_Adapter_v[a]:=vLAN_Adapter_v[a] + WmiResults [J, I] + ' | ';
                end;
            end; // I ////////////////////////////////////////////////////
        end ; // J
    end; //rows
  end; // Count
end;  //win32_networkconnection

// --- SetNetwork(pcIPAddress,pcMACAddress,pcIPSubnet,pcDefaultIPGateway,pcDHCPServer,pcDNSHostName,pcDNSDomain,pcDNSServerSearchOrder,pcServiceName)---
procedure TPC_Info.Network(Value: String);
var
    rows, instances, I, J, a: integer ;
    WmiResults: T2DimStrArray ;
    //pcIPAddress,pcMACAddress,pcIPSubnet,pcDefaultIPGateway,pcDHCPServer,pcDNSHostName,
    //pcDNSDomain,pcDNSServerSearchOrder,pcServiceName,pcDomain:string;
begin
if GetNetwork.Count < 1 then
 begin //Count
  rows := MagWmiGetInfo ('.', 'root\CIMV2', '', '',
    'Win32_NetworkAdapterConfiguration',
    WmiResults, instances) ;
  if rows > 0 then
    begin
      for J := 1 to instances do
        begin
          for I := 1 to rows do
            begin ////////////////////////////////////////////////////////
              for a := 0 to 8 do
                begin
                 if WmiResults [0, I] = vNetwork[a]{) and (WmiResults [J, I] <> 'NULL')} then vNetwork_v[a]:=vNetwork_v[a] + WmiResults [J, I] + ' | ';
                // if WmiResults [0, I] = vNetwork[a] then vNetwork_v[a]:=vNetwork_v[a] + WmiResults [J, I] + ' | ';
                end;
              end; // I ////////////////////////////////////////////////////
        end ; // J
    end; //rows
  end; // Count
end; // --- SetNetwork(pcIPAddress,pcMACAddress,pcIPSubnet,pcDefaultIPGateway,pcDHCPServer,pcDNSHostName,pcDNSDomain,pcDNSServerSearchOrder,pcServiceName)---

// --- SetNetConnect(pcName,pcConnectionType,pclocalName,pcUserName,pcResourceType,pcDisplayType)---
procedure TPC_Info.NetConnect(Value: String);
var
    rows, instances, I, J, a: integer ;
    WmiResults: T2DimStrArray ;
//    pcName,pcConnectionType,pclocalName,pcUserName,pcResourceType,pcDisplayType:string;
begin
if GetNetConnect.Count < 1 then
 begin //Count
  rows := MagWmiGetInfo ('.', 'root\CIMV2', '', '',
    'win32_networkconnection',
    WmiResults, instances) ;
  if rows > 0 then
    begin
      for J := 1 to instances do
        begin
          for I := 1 to rows do
            begin ////////////////////////////////////////////////////////
              for a := 0 to 6 do
                begin
                 if WmiResults [0, I] = vNetConnect[a] then vNetConnect_v[a]:=vNetConnect_v[a] + WmiResults [J, I] + ' | ';
                end;
            end; // I ////////////////////////////////////////////////////
        end ; // J
    end; //rows
  end; //Count
end; // --- SetNetConnect(pcName,pcConnectionType,pclocalName,pcUserName,pcResourceType,pcDisplayType)---END

// --- SetIDE_USBController(pcNameIDE,pcNameUSB) ---
procedure TPC_Info.IDE_USBController(Value: String);
var
    rows, instances, I, J: integer ;
    WmiResults: T2DimStrArray ;
//    pcNameIDE,pcNameUSB:string;
begin
  rows := MagWmiGetInfo ('.', 'root\CIMV2', '', '',
    'Win32_IDEController',
    WmiResults, instances) ;
  if rows > 0 then
    begin
      for J := 1 to instances do
        begin
          for I := 1 to rows do
            begin ////////////////////////////////////////////////////////
              if WmiResults [0, I] = vIDE_USBController[0] then vIDE_USBController_v[0]:=vIDE_USBController_v[0] + WmiResults [J, I] + ' | ';
            end; // I ////////////////////////////////////////////////////
        end ; // J
    end; //rows
  rows := MagWmiGetInfo ('.', 'root\CIMV2', '', '',
    'Win32_USBController',
    WmiResults, instances) ;
  if rows > 0 then
    begin
      for J := 1 to instances do
        begin
          for I := 1 to rows do
            begin ////////////////////////////////////////////////////////
              if WmiResults [0, I] = vIDE_USBController[0] then vIDE_USBController_v[0]:=vIDE_USBController_v[0] + WmiResults [J, I] + ' | ';
            end; // I ////////////////////////////////////////////////////
        end ; // J
    end; //rows
end; // --- SetIDE_USBController(pcNameIDE,pcNameUSB) ---END

// --- OS (pcManufacturer,pcCaption,pcCSDVersion,pcBuildNumber,pcBuildType,pcVersion,pcBootDevice,pcName,pcSystemStartupOptions,pcSystemDrive,pcWindowsDirectory,pcSystemDevice,pcSystemDirectory,pcTempDirectory,pcOSLanguage,pcSystemType,pcLocale,pcCSName,pcOrganization,pcLastBootUpTime,pcLocalDateTime) ---
procedure TPC_Info.OS(Value: String);
var
    rows, instances, I, J, a: integer ;
    WmiResults: T2DimStrArray ;
//    pcManufacturer,pcCaption,pcCSDVersion,pcBuildNumber,pcBuildType,pcVersion,
//    pcBootDevice,pcName,pcSystemStartupOptions,pcSystemDrive,pcWindowsDirectory,
//    pcSystemDevice,pcSystemDirectory,pcTempDirectory,pcOSLanguage,pcSystemType,
//    pcLocale,pcCSName,pcOrganization,pcLastBootUpTime,pcLocalDateTime:string;
begin
  rows := MagWmiGetInfo ('.', 'root\CIMV2', '', '',
    'Win32_OperatingSystem',
    WmiResults, instances) ;
  if rows > 0 then
    begin
      for J := 1 to instances do
        begin
          for I := 1 to rows do
            begin ////////////////////////////////////////////////////////
              for a := 0 to 20 do
                begin
                 if WmiResults [0, I] = vOS[a] then vOS_v[a]:=WmiResults [J, I];
                end;
              end; // I ////////////////////////////////////////////////////
        end ; // J
    end; //rows
end; // --- OS (pcManufacturer,pcCaption,pcCSDVersion,pcBuildNumber,pcBuildType,pcVersion,pcBootDevice,pcName,pcSystemStartupOptions,pcSystemDrive,pcWindowsDirectory,pcSystemDevice,pcSystemDirectory,pcTempDirectory,pcOSLanguage,pcSystemType,pcLocale,pcCSName,pcOrganization,pcLastBootUpTime,pcLocalDateTime) ---END

// --- OS_Variable (pcName,pcVariableValue,pcCaption) ---
procedure TPC_Info.OS_Variable(Value: String);
var
    rows, instances, I, J, a: integer ;
    WmiResults: T2DimStrArray ;
begin
  rows := MagWmiGetInfo ('.', 'root\CIMV2', '', '',
    'Win32_Environment',
    WmiResults, instances) ;
  if rows > 0 then
    begin
      for J := 1 to instances do
        begin
          for I := 1 to rows do
            begin ////////////////////////////////////////////////////////
              for a := 0 to 2 do
                begin
                 if WmiResults [0, I] = vOS_Variable[a] then vOS_Variable_v[a]:=vOS_Variable_v[a] + WmiResults [J, I] + ' | ';
                end;
            end; // I ////////////////////////////////////////////////////
        end ; // J
    end; //rows
end; // --- OS_Variable (pcName,pcVariableValue,pcCaption)---END

// ---SetUserAccount(Name,SID,Disabled) ---
procedure TPC_Info.UserAccount(Value: String);
var
    rows, instances, I, J, a: integer ;
    WmiResults: T2DimStrArray ;
//    pcName,pcSID,pcDisabled:string;
begin
  rows := MagWmiGetInfo ('.', 'root\CIMV2', '', '',
    'Win32_UserAccount',
    WmiResults, instances) ;
  if rows > 0 then
    begin
      for J := 1 to instances do
        begin
          for I := 1 to rows do
            begin ////////////////////////////////////////////////////////
              for a := 0 to 2 do
                begin
                 if WmiResults [0, I] = vUserAccount[a] then vUserAccount_v[a]:=vUserAccount_v[a] + WmiResults [J, I] + ' | ';
                end;
              end; // I ////////////////////////////////////////////////////
        end ; // J
    end; //rows
end; // --- SetUserName(Name,SID,Disabled) ---END

// ---SetSystemAccount(Name,SID) ---
procedure TPC_Info.SystemAccount(Value: String);
var
    rows, instances, I, J, a: integer ;
    WmiResults: T2DimStrArray ;
//    pcName,pcSID:string;
begin
  rows := MagWmiGetInfo ('.', 'root\CIMV2', '', '',
    'Win32_SystemAccount',
    WmiResults, instances) ;
  if rows > 0 then
    begin
      for J := 1 to instances do
        begin
          for I := 1 to rows do
            begin ////////////////////////////////////////////////////////
              for a := 0 to 1 do
                begin
                 if WmiResults [0, I] = vSystemAccount[a] then vSystemAccount_v[a]:=vSystemAccount_v[a] + WmiResults [J, I] + ' | ';
                end;
              end; // I ////////////////////////////////////////////////////
        end ; // J
    end; //rows
end; // --- SetSystemAccount(Name,SID) ---END


// --- SetGroupName (Name,SID) ---
procedure TPC_Info.GroupName(Value: String);
var
    rows, instances, I, J, a: integer ;
    WmiResults: T2DimStrArray ;
//    pcName,pcSID:string;
begin
  rows := MagWmiGetInfo ('.', 'root\CIMV2', '', '',
    'Win32_Group',
    WmiResults, instances) ;
  if rows > 0 then
    begin
      for J := 1 to instances do
        begin
          for I := 1 to rows do
            begin ////////////////////////////////////////////////////////
              for a := 0 to 1 do
                begin
                 if WmiResults [0, I] = vGroupName[a] then vGroupName_v[a]:=vGroupName_v[a] + WmiResults [J, I] + ' | ';
                end;
              end; // I ////////////////////////////////////////////////////
        end ; // J
    end; //rows
end; // --- SetGroupName(Name,SID) ---END

// --- SetProcess (Name,CommandLine,ProcessId) ---
procedure TPC_Info.Process(Value: String);
var
    rows, instances, I, J, a: integer ;
    WmiResults: T2DimStrArray ;
begin
  rows := MagWmiGetInfo ('.', 'root\CIMV2', '', '',
    'win32_process',
    WmiResults, instances) ;
  if rows > 0 then
    begin
      for J := 1 to instances do
        begin
          for I := 1 to rows do
            begin ////////////////////////////////////////////////////////
              for a := 0 to 1 do
                begin
                 if WmiResults [0, I] = vProcess[a] then vProcess_v[a]:=vProcess_v[a] + WmiResults [J, I] + ' | ';
                end;
              end; // I ////////////////////////////////////////////////////
        end ; // J
    end; //rows
end; // --- SetProcess(Name,CommandLine,ProcessId) ---END



               /////
         ///////////////
   /////////////////////////
///////////////////////////////
///////////////////////////////
//                           //
//             END           //
//                           //
//        PC_INFO WMI        //
//                           //
///////////////////////////////


// ----------------- START REC TO FIELDS



// ---- TPopupMenu Click нажатия на пункты меню
procedure TPC_Info.MenuItem_MatherBoardClick;
var a:integer;
begin
Self.Hint:='MatherBoard';
// procedure Caption
  begin
  vCaption[0]:='Caption';
  vCaption[1]:='UserName';
  vCaption[2]:='Model';
  vCaption[3]:='Manufacturer';
  vCaption[4]:='Domain';

  Self.Caption(vCaption[0]);
  Self.Strings.Add('--- MatherBoard ---');
  for a :=0 to 4 do
    begin
      Self.GetCaption.Add(vCaption[a] + ' | ' + vCaption_v[a]);
      Self.Strings.Add(vCaption[a] + '=' + vCaption_v[a]);
    end;
Self.Strings.Add('');
  FHostname  := vCaption_v[0];
  FUserName_l:= vCaption_v[1];
  FModel     := vCaption_v[2];
  FGroup     := vCaption_v[4];

// procedure MatherBoard Caption
  vMatherBoard[0]:='Manufacturer';
  vMatherBoard[1]:='SerialNumber';
  vMatherBoard[2]:='PartNumber';
  vMatherBoard[3]:='Product';
  vMatherBoard[4]:='DeviceID';
  vMatherBoard[5]:='PrimaryBusType';
  vMatherBoard[6]:='SecondaryBusType';
  vMatherBoard[7]:='Name';
  vMatherBoard[8]:='SystemName';

  Self.MatherBoard(vMatherBoard[0]);
  for a :=0 to 8 do
    begin
      Self.GetMatherBoard.Add(vMatherBoard[a] + ' | ' + vMatherBoard_v[a]);
      Self.Strings.Add(vMatherBoard[a] + '=' + vMatherBoard_v[a]);
    end;
    Self.Strings.Add('');
    FSerialNumber  :=vMatherBoard_v[1];
    FPartNumber    :=vMatherBoard_v[2];
  end;
end;

procedure TPC_Info.MenuItem_BIOSClick;
var a:integer;
begin
Self.Hint:='BIOS';
  begin  // procedure BIOS
  vBIOS[0]:='SMBIOSBIOSVersion';
  vBIOS[1]:='Version';
  vBIOS[2]:='ReleaseDate';
  vBIOS[3]:='Manufacturer';
  vBIOS[4]:='SerialNumber';
  vBIOS[5]:='Name';
  vBIOS[6]:='CurrentLanguage';
  vBIOS[7]:='TargetOperatingSystem';

  Self.BIOS(vBIOS[0]);
  Self.Strings.Add('--- BIOS ---');
  for a :=0 to 7 do
    begin
      Self.GetBIOS.Add(vBIOS[a] + ' | ' + vBIOS_v[a]);
      Self.Strings.Add(vBIOS[a] + '=' + vBIOS_v[a]);
    end;
  Self.Strings.Add('');
  FBIOS_ver  := vBIOS_v[0];
  FBIOS_date := vBIOS_v[1];
  end;
end;

procedure TPC_Info.MenuItem_CPUClick;
var a:integer;
begin
Self.Hint:='CPU';
  begin  // procedure CPU
  vCPU[0]:='DeviceID';
  vCPU[1]:='SocketDesignation';
  vCPU[2]:='Caption';
  vCPU[3]:='Manufacturer';
  vCPU[4]:='Name';
  vCPU[5]:='MaxClockSpeed';
  vCPU[6]:='CurrentClockSpeed';
  vCPU[7]:='NumberOfCores';
  vCPU[8]:='NumberOfLogicalProcessors';
  vCPU[9]:='L2CacheSize';
  vCPU[10]:='Revision';
  vCPU[11]:='ProcessorId';

  Self.CPU(vCPU[0]);
  Self.Strings.Add('--- CPU ---');
    for a :=0 to 11 do
    begin
      Self.GetCPU.Add(vCPU[a] + ' | ' + vCPU_v[a]);
      Self.Strings.Add(vCPU[a] + '=' + vCPU_v[a]);
    end;
    Self.Strings.Add('');
    //  FCPU_All :=vCPU_v[3] + ' ' + vCPU_v[4] + '    ' + vCPU_v[5] + 'MHz  ';
  end;
end;

procedure TPC_Info.MenuItem_RAMClick;
var a:integer;
begin
Self.Hint:='RAM';
  begin  // procedure RAM
  vRAM[0]:='TotalVirtualMemorySize';
  vRAM[1]:='TotalVisibleMemorySize';
  vRAM[2]:='TotalSwapSpaceSize';
  vRAM[3]:='FreePhysicalMemory';
  vRAM[4]:='FreeVirtualMemory';
  vRAM[5]:='FreeSpaceInPagingFiles';

  Self.RAM(vRAM[0]);
  Self.Strings.Add('--- RAM ---');
  for a :=0 to 5 do
    begin
      Self.GetRAM.Add(vRAM[a] + ' | ' + vRAM_v[a]);
      Self.Strings.Add(vRAM[a] + '=' + vRAM_v[a]);
    end;
  Self.Strings.Add('');
  FRAM_Size:=IntToStr(StrToInt(vRAM_v[1]) div 1024) + 'Mb ';
  FRAM_Swap_Size:=vRAM_v[2];
  end;
//
  begin  // procedure RAM_Phisic
  vRAM_Phisic[0]:='Caption';
  vRAM_Phisic[1]:='DeviceLocator';
  vRAM_Phisic[2]:='Speed';

  Self.RAM_Phisic(vRAM_Phisic[0]);
  Self.Strings.Add('--- RAM Phisic---');
  for a :=0 to 2 do
    begin
      Self.GetRAM_Phisic.Add(vRAM_Phisic[a] + ' | ' + vRAM_Phisic_v[a]);
      Self.Strings.Add(vRAM_Phisic[a] + '=' + vRAM_Phisic_v[a]);
    end;
  Self.Strings.Add('');
  FRAM_Speed:=vRAM_Phisic_v[2];
  end;
end;

procedure TPC_Info.MenuItem_VideoClick;
var a:integer;
begin
Self.Hint:='Video';
  begin  // procedure  Video_Adapter
  vVideo_Adapter[0]:='Name';
  vVideo_Adapter[1]:='AdapterRAM';
  vVideo_Adapter[2]:='DriverVersion';
  vVideo_Adapter[3]:='DriverDate';
  vVideo_Adapter[4]:='InstalledDisplayDrivers';
  vVideo_Adapter[5]:='InfFilename';
  vVideo_Adapter[6]:='MaxRefreshRate';
  vVideo_Adapter[7]:='VideoModeDescription';
  vVideo_Adapter[8]:='CurrentBitsPerPixel';
  vVideo_Adapter[9]:='CurrentRefreshRate';
  vVideo_Adapter[10]:='AdapterCompatibility';
  vVideo_Adapter[11]:='PNPDeviceID';
  vVideo_Adapter[12]:='Caption';
  vVideo_Adapter[13]:='AdapterDACType';
  vVideo_Adapter[14]:='VideoMemoryType';

  Self.Video_Adapter(vVideo_Adapter[0]);
  Self.Strings.Add('--- Video Adapter---');
  for a :=0 to 14 do
    begin
      Self.GetVideo_Adapter.Add(vVideo_Adapter[a] + ' | ' + vVideo_Adapter_v[a]);
      Self.Strings.Add(vVideo_Adapter[a] + '=' + vVideo_Adapter_v[a]);
    end;
  Self.Strings.Add('');
  FVideo:=vVideo_Adapter_v[0];
  end;
end;

procedure TPC_Info.MenuItem_AudioClick;
var a:integer;
begin
Self.Hint:='Audio';
  begin  // procedure Audio_Adapter
  vAudio_Adapter[0]:='Manufacturer';
  vAudio_Adapter[1]:='Name';
  vAudio_Adapter[2]:='DeviceID';

  Self.Audio_Adapter(vAudio_Adapter[0]);
  Self.Strings.Add('--- Audio Adapter ---');
  for a :=0 to 2 do
    begin
      Self.GetAudio_Adapter.Add(vAudio_Adapter[a] + ' | ' + vAudio_Adapter_v[a]);
      Self.Strings.Add(vAudio_Adapter[a] + '=' + vAudio_Adapter_v[a]);
    end;
  Self.Strings.Add('');
  FAudio:=vAudio_Adapter_v[1];
  end;
end;

procedure TPC_Info.MenuItem_HDDClick;
var a:integer;
begin
Self.Hint:='HDD';
  begin  // procedure  HDD_Phisic
  vHDD_Phisic[0]:='Model';
  vHDD_Phisic[1]:='Size';
  vHDD_Phisic[2]:='InterfaceType';
  vHDD_Phisic[3]:='Signature';
  vHDD_Phisic[4]:='Manufacturer';
  vHDD_Phisic[5]:='Name';
  vHDD_Phisic[6]:='TotalCylinders';
  vHDD_Phisic[7]:='TotalHeads';
  vHDD_Phisic[8]:='TotalSectors';
  vHDD_Phisic[9]:='TotalTracks';
  vHDD_Phisic[10]:='TracksPerCylinder';
  vHDD_Phisic[11]:='BytesPerSector';
  vHDD_Phisic[12]:='SectorsPerTrack';

  Self.HDD_Phisic(vHDD_Phisic[0]);
  Self.Strings.Add('--- HDD Phisic ---');
  for a :=0 to 12 do
    begin
      Self.GetHDD_Phisic.Add(vHDD_Phisic[a] + ' | ' + vHDD_Phisic_v[a]);
      Self.Strings.Add(vHDD_Phisic[a] + '=' + vHDD_Phisic_v[a]);
    end;
  Self.Strings.Add('');
  FHDD_p:=vHDD_Phisic_v[0];
  end;
//
  begin  // procedure  HDD_Map
  vHDD_Map[0]:='Name';
  vHDD_Map[1]:='ProviderName';
  vHDD_Map[2]:='VolumeName';
  vHDD_Map[3]:='FileSystem';
  vHDD_Map[4]:='VolumeSerialNumber';
  vHDD_Map[5]:='Size';
  vHDD_Map[6]:='FreeSpace';
  vHDD_Map[7]:='SessionID';

  Self.HDD_Map(vHDD_Map[0]);
  Self.Strings.Add('--- HDD Map ---');
  for a :=0 to 7 do
    begin
      Self.GetHDD_Map.Add(vHDD_Map[a] + ' | ' + vHDD_Map_v[a]);
      Self.Strings.Add(vHDD_Map[a] + '=' + vHDD_Map_v[a]);
    end;
  Self.Strings.Add('');
  FHDD_m:=vHDD_Map_v[1];
  end;
//
  begin  // procedure  HDD_Map
  vHDD_Vol[0]:='Caption';
  vHDD_Vol[1]:='Description';
  vHDD_Vol[2]:='Name';
  vHDD_Vol[3]:='VolumeName';
  vHDD_Vol[4]:='ProviderName';
  vHDD_Vol[5]:='MediaType';
  vHDD_Vol[6]:='DriveType';
  vHDD_Vol[7]:='FileSystem';
  vHDD_Vol[8]:='FreeSpace';
  vHDD_Vol[9]:='Size';
  vHDD_Vol[10]:='MaximumComponentLength';
  vHDD_Vol[11]:='VolumeSerialNumber';

  Self.HDD_Vol(vHDD_Vol[0]);
  Self.Strings.Add('--- HDD Vol ---');
  for a :=0 to 11 do
    begin
      Self.GetHDD_Vol.Add(vHDD_Vol[a] + ' | ' + vHDD_Vol_v[a]);
      Self.Strings.Add(vHDD_Vol[a] + '=' + vHDD_Vol_v[a]);
    end;
  Self.Strings.Add('');
  FHDD_v:=vHDD_Vol_v[0];
  end;
end;

procedure TPC_Info.MenuItem_NetworkClick;
var a:integer;
begin
Self.Hint:='Network';
  begin  // procedure  LAN_Adapter
  vLAN_Adapter[0]:='Name';
  vLAN_Adapter[1]:='Manufacturer';
  vLAN_Adapter[2]:='AdapterType';
  vLAN_Adapter[3]:='MACAddress';
  vLAN_Adapter[4]:='NetConnectionID';
  vLAN_Adapter[5]:='TimeOfLastReset';

  Self.LAN_Adapter(vLAN_Adapter[0]);
  Self.Strings.Add('--- LAN Adapter ---');
  for a :=0 to 5 do
    begin
      Self.GetLAN_Adapter.Add(vLAN_Adapter[a] + ' | ' + vLAN_Adapter_v[a]);
      Self.Strings.Add(vLAN_Adapter[a] + '=' + vLAN_Adapter_v[a]);
    end;
    Self.Strings.Add('');
  //  FMAC:=vLAN_Adapter_v[3];
  end;
//
  begin  // procedure  Network	
  vNetwork[0]:='IPAddress';
  vNetwork[1]:='MACAddress';
  vNetwork[2]:='IPSubnet';
  vNetwork[3]:='DefaultIPGateway';
  vNetwork[4]:='DHCPServer';
  vNetwork[5]:='DNSHostName';
  vNetwork[6]:='DNSDomain';
  vNetwork[7]:='DNSServerSearchOrder';
  vNetwork[8]:='ServiceName';

  Self.Network(vNetwork[0]);
  Self.Strings.Add('--- Network ---');
  for a :=0 to 8 do
    begin
      Self.GetNetwork.Add(vNetwork[a] + ' | ' + vNetwork_v[a]);
      Self.Strings.Add(vNetwork[a] + '=' + vNetwork_v[a]);
    end;
  Self.Strings.Add('');
  FIP:=vNetwork_v[0];
  FMAC:=vNetwork_v[1];
  FIPSubnet:=vNetwork_v[2];
  FGateway:=vNetwork_v[3];
  FDHCPServer:=vNetwork_v[4];
  FDNSHostName:=vNetwork_v[5];
  FDNSDomain:=vNetwork_v[6];
  FDNSServer:=vNetwork_v[7];
  end;
//
  begin  // procedure  NetConnect
  vNetConnect[0]:='ConnectionType';
  vNetConnect[1]:='LocalName';
  vNetConnect[2]:='RemoteName';
  vNetConnect[3]:='UserName';
  vNetConnect[4]:='Name';
  vNetConnect[5]:='ResourceType';
  vNetConnect[6]:='DisplayType';

  Self.NetConnect(vNetConnect[0]);
  Self.Strings.Add('--- NetConnect ---');
  for a :=0 to 6 do
    begin
      Self.GetNetConnect.Add(vNetConnect[a] + ' | ' + vNetConnect_v[a]);
      Self.Strings.Add(vNetConnect[a] + '=' + vNetConnect_v[a]);
    end;
  Self.Strings.Add('');
  FConnectNeme:= vNetConnect_v[4]
  end;
end;

procedure TPC_Info.MenuItem_IDE_USBControllerClick;
begin
Self.Hint:='IDE_USBController';
  begin  // procedure  IDE_USBController
  vIDE_USBController[0]:='Name';
  vIDE_USBController[1]:='Name';

  Self.IDE_USBController(vIDE_USBController[0]);
  Self.Strings.Add('--- IDE_USBController ---');
  //for a :=0 to 1 do
    begin
      Self.GetIDE_USBController.Add(vIDE_USBController[0] + ' | ' + vIDE_USBController_v[0]);
      Self.Strings.Add(vIDE_USBController[0] + '=' + vIDE_USBController_v[0]);
    end;
  Self.Strings.Add('');
  end;
end;

procedure TPC_Info.MenuItem_OSClick;
var a:integer;
begin
Self.Hint:='OS';
 begin  // procedure  OS
  vOS[0]:='Manufacturer';
  vOS[1]:='Caption';
  vOS[2]:='CSDVersion';
  vOS[3]:='BuildNumber';
  vOS[4]:='BuildType';
  vOS[5]:='Version';
    // PACH  Variable
  vOS[6]:='BootDevice';
  vOS[7]:='Name';
  vOS[8]:='SystemStartupOptions';
  vOS[9]:='SystemDrive';
  vOS[10]:='WindowsDirectory';
  vOS[11]:='SystemDevice';
  vOS[12]:='SystemDirectory';
  vOS[13]:='TempDirectory';
    //
  vOS[14]:='OSLanguage';
  vOS[15]:='SystemType';
  vOS[16]:='Locale';
  vOS[17]:='CSName';
  vOS[18]:='Organization';
  vOS[19]:='LastBootUpTime';
  vOS[20]:='LocalDateTime';

  Self.OS(vOS[0]);
  Self.Strings.Add('--- OS ---');
  for a :=0 to 20 do
    begin
      Self.GetOS.Add(vOS[a] + ' | ' + vOS_v[a]);
      Self.Strings.Add(vOS[a] + '=' + vOS_v[a]);
    end;
  Self.Strings.Add('');
  FOS:= vOS_v[1] + vOS_v[2] + ', Version:' + vOS_v[5] + ', BildType:' + vOS_v[4];
  end;
//
  begin  // procedure  OS_Variable
  vOS_Variable[0]:='Name';
  vOS_Variable[1]:='VariableValue';
  vOS_Variable[2]:='Caption';

  Self.OS_Variable(vOS_Variable[0]);
  Self.Strings.Add('--- OS_Variable ---');
  for a :=0 to 2 do
    begin
      Self.GetOS_Variable.Add(vOS_Variable[a] + ' | ' + vOS_Variable_v[a]);
      Self.Strings.Add(vOS_Variable[a] + '=' + vOS_Variable_v[a]);
   end;
  Self.Strings.Add('');
  end;
//
  begin  // procedure  UserAccount
  vUserAccount[0]:='Name';
  vUserAccount[1]:='SID';
  vUserAccount[2]:='Disabled';

  Self.UserAccount(vUserAccount[0]);
  Self.Strings.Add('--- UserAccount ---');
  for a :=0 to 2 do
    begin
      Self.GetUserAccount.Add(vUserAccount[a] + ' | ' + vUserAccount_v[a]);
      Self.Strings.Add(vUserAccount[a] + '=' + vUserAccount_v[a]);
   end;
  Self.Strings.Add('');
  end;
//
  begin  // procedure  SystemAccount
  vSystemAccount[0]:='Name';
  vSystemAccount[1]:='SID';

  Self.SystemAccount(vSystemAccount[0]);
  Self.Strings.Add('--- SystemAccount ---');
  for a :=0 to 1 do
    begin
      Self.GetSystemAccount.Add(vSystemAccount[a] + ' | ' + vSystemAccount_v[a]);
      Self.Strings.Add(vSystemAccount[a] + '=' + vSystemAccount_v[a]);
   end;
  Self.Strings.Add('');
  end;
//
  begin  // procedure  GroupName
  vGroupName[0]:='Name';
  vGroupName[1]:='SID';

  Self.GroupName(vGroupName[0]);
  Self.Strings.Add('--- GroupName ---');
  for a :=0 to 1 do
    begin
      Self.GetGroupName.Add(vGroupName[a] + ' | ' + vGroupName_v[a]);
      Self.Strings.Add(vGroupName[a] + '=' + vGroupName_v[a]);
   end;
  Self.Strings.Add('');
  end;
end;

procedure TPC_Info.MenuItem_ProcessClick;
var a:integer;
begin
Self.Hint:='Process';
  begin  // procedure
  vProcess[0]:='Name';
  vProcess[1]:='CommandLine';
  vProcess[2]:='ProcessId';
  vProcess[3]:='Priority';

  Self.Process(vProcess[0]);
  Self.Strings.Add('--- Process ---');
  for a :=0 to 1 do
    begin
      Self.GetProcess.Add(vProcess[a] + ' | ' + vProcess_v[a]);
      Self.Strings.Add(vProcess[a] + '=' + vProcess_v[a]);
    end;
  Self.Strings.Add('');
  end;
end;

procedure TPC_Info.MenuItem_CleanClick;
begin
Self.Hint:='Clear';
  Self.Strings.Clear;
  Self.InsertRow('','',True);
end;

procedure TPC_Info.MenuItem_SaveClick;
begin
  Self.Strings.SaveToFile('c:\PC_' + FHostname + FSerialNumber + '.txt');
end;

procedure TPC_Info.MenuItem_MailClick;
begin
Self.Hint:='E-Mail';
Self.TitleCaptions.Add('E-Mail');
Self.TitleCaptions.Add('IMy@ukr.net');
end;

procedure TPC_Info.PCreatePopUpMenu;           // Создаю Меню для объекта
begin
Self.PopupMenu:=FPopUpMenu;                 //подключаю к TSock.PopUpMenu встроенный компонент TPopUpMenu через FPopUpMenu
if Self.PopupMenu is TPopUpMenu then        // проверяю подключен к TSock.PopupMenu  TPopUpMenu
    begin
    If Self.PopupMenu.Items.Count = 0 Then  //проверяю добавленны или нет пункты
      begin
      MItem:=TMenuItem.Create(FPopUpMenu);
      MItem.Caption:='BIOS';
      MItem.OnClick:=MenuItem_BIOSClick;
      Self.PopupMenu.Items.Add(MItem);// end N1

      MItem:=TMenuItem.Create(FPopUpMenu);
      MItem.Caption:='CPU';
      MItem.OnClick:=MenuItem_CPUClick;
      Self.PopupMenu.Items.Add(MItem);// end N2

      MItem:=TMenuItem.Create(FPopUpMenu);
      MItem.Caption:='MatherBoard';
      MItem.OnClick:=MenuItem_MatherBoardClick;
      Self.PopupMenu.Items.Add(MItem);// end N3

      MItem:=TMenuItem.Create(FPopUpMenu);
      MItem.Caption:='RAM';
      MItem.OnClick:=MenuItem_RAMClick;
      Self.PopupMenu.Items.Add(MItem);// end N4

      MItem:=TMenuItem.Create(FPopUpMenu);
      MItem.Caption:='Video';
      MItem.OnClick:=MenuItem_VideoClick;
      Self.PopupMenu.Items.Add(MItem);// end N5

      MItem:=TMenuItem.Create(FPopUpMenu);
      MItem.Caption:='Audio';
      MItem.OnClick:=MenuItem_AudioClick;
      Self.PopupMenu.Items.Add(MItem);// end N6

      MItem:=TMenuItem.Create(FPopUpMenu);
      MItem.Caption:='HDD';
      MItem.OnClick:=MenuItem_HDDClick;
      Self.PopupMenu.Items.Add(MItem);// end N7

      MItem:=TMenuItem.Create(FPopUpMenu);
      MItem.Caption:='Network';
      MItem.OnClick:=MenuItem_NetworkClick;
      Self.PopupMenu.Items.Add(MItem);// end N8

      MItem:=TMenuItem.Create(FPopUpMenu);
      MItem.Caption:='IDE_USBController';
      MItem.OnClick:=MenuItem_IDE_USBControllerClick;
      Self.PopupMenu.Items.Add(MItem);// end N9

      MItem:=TMenuItem.Create(FPopUpMenu);
      MItem.Caption:='OS';
      MItem.OnClick:=MenuItem_OSClick;
      Self.PopupMenu.Items.Add(MItem);// end N10

      MItem:=TMenuItem.Create(FPopUpMenu);
      MItem.Caption:='Process';
      MItem.OnClick:=MenuItem_ProcessClick;
      Self.PopupMenu.Items.Add(MItem);// end N11

      MItem:=TMenuItem.Create(FPopUpMenu);
      MItem.Caption:='-';
      Self.PopupMenu.Items.Add(MItem);// end '-'

      MItem:=TMenuItem.Create(FPopUpMenu);
      MItem.Caption:='Clean';
      MItem.OnClick:=MenuItem_CleanClick;
      Self.PopupMenu.Items.Add(MItem);// end N12

      MItem:=TMenuItem.Create(FPopUpMenu);
      MItem.Caption:='-';
      Self.PopupMenu.Items.Add(MItem);// end '-'

      MItem:=TMenuItem.Create(FPopUpMenu);
      MItem.Caption:='Save';
      MItem.OnClick:=MenuItem_SaveClick;
      Self.PopupMenu.Items.Add(MItem);// end N13

      MItem:=TMenuItem.Create(FPopUpMenu);
      MItem.Caption:='-';
      Self.PopupMenu.Items.Add(MItem);// end '-'

      MItem:=TMenuItem.Create(FPopUpMenu);
      MItem.Caption:='IMy@ukr.net ';
      MItem.OnClick:=MenuItem_MailClick;
      Self.PopupMenu.Items.Add(MItem);// end N14
      end;
    end;
end;
// EVENTs  END --------




destructor TPC_Info.Destroy;
begin
  FPopUpMenu.Free;
  inherited Destroy;
end;


procedure Register;
begin
  RegisterComponents('Standard', [TPC_Info]);
end;


//{Компонент проверен на DELPHI 7
//Данный компонент позволяет определить конфигурацию Вашего ПК (BIOS, MatherBoard, CPU, RAM, Video_Adapter, Audio_Adapter, HDD, LAN_Adapter, SetIDE_USBController, также под управлением какого MS Windows работает Ваш ПК OS, OS_VariableХпеременные среды}, UserAccount, SystemAccount, GroupName, Process, Network {настройки сети IPAddress, MACAddress, IPSubnet, DefaultIPGateway, DHCPServer, DNSHostName, DNSDomain, DNSServer}, HDD_Vol{информация о разделах дисков}, HDD_Map{информация о примапленных ресурсах}...
//Компонент снабжен своим 'PopUpMenu', с возможностью сохранения данный на жесткий диск
//В общем всего так и не описать большее количество всякой необходимой информации смотрите сами
//Разработан на основе стандартного компонента 'TValueListEditor' и 'TPopUpMenu'. Также для работы компонента необходимы файлы WbemScripting_TLB.pas, smartapi.pas, magwmi.pas, magsubs1.pas { для работы с WMI} последние вложены в архив.}
//
//{WbemScripting_TLB.pas, smartapi.pas, magwmi.pas, magsubs1.pas используются для работы с WMI

//Компонент проверен на DELPHI 7
//Разработан на основе стандартного компонента 'TValueListEditor' и 'TPopUpMenu'.}
end.