unit DDU.DispatchEx;

//*****************************************************************************
//
// DDULIBRARY (DDU.RTTIDispatch)
// COPYRIGHT 2020 Clinton R. Johnson (xepol@xepol.com)
//
// Licensed under the Apache License, Version 2.0 (the "License");
// you may not use this file except in compliance with the License.
// You may obtain a copy of the License at
//
//     http://www.apache.org/licenses/LICENSE-2.0
//
// Unless required by applicable law or agreed to in writing, software
// distributed under the License is distributed on an "AS IS" BASIS,
// WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
// See the License for the specific language governing permissions and
// limitations under the License.
//
// Version : 1.0
//
// Purpose : Provide IDispatchEx interface an as extended alternative to
//           IDispatch and Variant, using RTTI custom properties to control
//           what is and is not published via the interface.
//
// History : First version is based on code from System.Win.Com TObjectDispatch
//           but is being modified to use System.RTTI more extensively.
//
//           Since System.RTTI uses TValue instead of Variant, performance
//           testing will be required as migration towards an exclusive
//           System.RTTI based solution.
//
//           Check IDispatch on other platforms, see if decoupling from
//           WinApi is called for.
//
//*****************************************************************************

interface

Uses
  System.Classes,
  System.SysUtils,
  System.RTTI,
  System.Variants,
  System.Generics.Defaults,
  System.Generics.Collections,
  System.Win.ComObj,
  Winapi.Windows,
  System.ObjAuto,  // PMethodInfoHeader, PParamInfo etc.
  Winapi.ActiveX,  // TDispID, TBSTR
  System.TypInfo  // RTTI PPropInfo etc.
  ;

Const
  grfdexPropCanAll        = fdexPropCanGet Or fdexPropCanPut Or fdexPropCanPutRef Or fdexPropCanCall Or fdexPropCanConstruct Or fdexPropCanSourceEvents;
  grfdexPropCannotAll     = fdexPropCannotGet Or fdexPropCannotPut Or fdexPropCannotPutRef Or fdexPropCannotCall Or fdexPropCannotConstruct Or fdexPropCannotSourceEvents;
  grfdexPropExtraAll      = fdexPropNoSideEffects Or fdexPropDynamicType;
  grfdexPropAll           = grfdexPropCanAll Or grfdexPropCannotAll Or grfdexPropExtraAll;

Type
  IDispatchEx = interface(IDispatch)
    ['{A6EF9860-C720-11D0-9337-00A0C90DCAA9}']
    function GetDispID(const bstrName: TBSTR; const grfdex: DWORD; out id: TDispID): HResult; stdcall;
    function InvokeEx(const id: TDispID; const lcid: LCID; const wflags: WORD; Const pdp: PDispParams; Const varRes: POleVariant; Const pei: PExcepInfo; const pspCaller: PServiceProvider): HResult; stdcall;
    function DeleteMemberByName(const bstr: TBSTR; const grfdex: DWORD): HResult; stdcall;
    function DeleteMemberByDispID(const id: TDispID): HResult; stdcall;
    function GetMemberProperties(const id: TDispID; const grfdexFetch: DWORD; out grfdex: DWORD): HResult; stdcall;
    function GetMemberName(const id: TDispID; out bstrName: TBSTR): HResult; stdcall;
    function GetNextDispID(const grfdex: DWORD; const id: TDispID; out nid: TDispID): HResult; stdcall;
    function GetNameSpaceParent(out unk: IUnknown): HResult; stdcall;
  end;

//************************************************************************************************************************************************************************************
//************************************************************************************************************************************************************************************
Type
  IParameters=Interface
  ['{AB382E13-ADEC-4EA0-8022-E62E997FCF84}']
    function GetParameter(aIndex: Integer): TVariantArg;
    Property Parameter[aIndex : Integer] : TVariantArg Read GetParameter; Default;
  End;

  TParameters=Class(TInterfacedObject,IParameters)
  Private
    fDispatchParameterIDs     : PDispIdList;
    fDispatchParameterIDsSize : Integer;
    fDispatchParameters       : PDispParams;

    fNull                     : Variant;
  Protected
    procedure BuildPositionalDispatchIDs;
    function GetParameter(aIndex: Integer): TVariantArg;
  Public
    Constructor Create(Const aDPS : PDispParams); Virtual;
    Destructor Destroy; Override;
    Property Parameter[aIndex : Integer] : TVariantArg Read GetParameter; Default;
  End;

  TGetIDsOfNamesEvent    = function (const IID: TGUID; Names: Pointer; NameCount, LocaleID: Integer; DispIDs: Pointer): HResult Of Object;
  TGetTypeInfoEvent      = function (Index, LocaleID: Integer; out TypeInfo): HResult Of Object;
  TGetTypeInfoCountEvent = function (out Count: Integer): HResult Of Object;
  TInvokeEvent           = function (DispID: Integer; const IID: TGUID; LocaleID: Integer; Flags: Word; Parameters : IParameters; VarResult, ExcepInfo, ArgErr: Pointer): HResult Of Object;

  TDispatchHandler=Class(TInterfacedObject,IInterface,IDispatch)
  Private
    fGetIDsOfNames    : TGetIDsOfNamesEvent;
    fGetTypeInfo      : TGetTypeInfoEvent;
    fGetTypeInfoCount : TGetTypeInfoCountEvent;
    fInvoke           : TInvokeEvent;
  Protected
    fCustomIID        : TGUID;
    fName             : String;
    fNameDispIDs      : TDictionary<String,Integer>;
  Protected
    { IInterface / IUnknown }
    function _AddRef: Integer; stdcall;
    function _Release: Integer; stdcall;
    function QueryInterface(const IID: TGUID; out Obj): HRESULT; stdcall;

    { IDispatch }
    function IDispatch_GetIDsOfNames(const IID: TGUID; Names: Pointer; NameCount, LocaleID: Integer; DispIDs: Pointer): HResult; stdcall;
    function IDispatch_GetTypeInfo(Index, LocaleID: Integer; out TypeInfo): HResult; stdcall;
    function IDispatch_GetTypeInfoCount(out Count: Integer): HResult; stdcall;
    function IDispatch_Invoke(DispID: Integer; const IID: TGUID; LocaleID: Integer; Flags: Word; var Params; VarResult, ExcepInfo, ArgErr: Pointer): HResult; virtual; stdcall;

    function IDispatch.GetTypeInfoCount = IDispatch_GetTypeInfoCount;
    function IDispatch.GetTypeInfo = IDispatch_GetTypeInfo;
    function IDispatch.GetIDsOfNames = IDispatch_GetIDsOfNames;
    function IDispatch.Invoke = IDispatch_Invoke;
  Protected
    function GetIDsOfNames(const IID: TGUID; Names: Pointer; NameCount, LocaleID: Integer; DispIDs: Pointer): HResult; Virtual;
    function GetTypeInfo(Index, LocaleID: Integer; out TypeInfo): HResult; Virtual;
    function GetTypeInfoCount(out Count: Integer): HResult; Virtual;
    function Invoke(DispID: Integer; const IID: TGUID; LocaleID: Integer; Flags: Word; Parameters : IParameters; VarResult, ExcepInfo, ArgErr: Pointer): HResult; Virtual;
  Protected
    Procedure Init; Virtual;
  Public
    Constructor Create; Overload;
    Constructor Create(Const aName : String; aCustomIID : TGUID); Overload;
    Destructor Destroy; Override;

    Procedure MapNameDispID(Const aName : String; aDispID : Integer);
  Public
    Property OnGetIDsOfNames    : TGetIDsOfNamesEvent    Read fGetIDsOfNames    Write fGetIDsOfNames;
    Property OnGetTypeInfo      : TGetTypeInfoEvent      Read fGetTypeInfo      Write fGetTypeInfo;
    Property OnGetTypeInfoCount : TGetTypeInfoCountEvent Read fGetTypeInfoCount Write fGetTypeInfoCount;
    Property OnInvoke           : TInvokeEvent           Read fInvoke           Write fInvoke;

    Property Name               : String                 Read fName             Write fName;
    Property CustomIID          : TGUID                  Read fCustomIID        Write fCustomIID;
  End;

  TDeleteMemberByDispIDEvent = function (const id: TDispID): HResult Of Object;
  TDeleteMemberByNameEvent   = function (const Name: String; const grfdex: DWORD): HResult Of Object;
  TGetDispIDEvent            = function (const Name: String; const grfdex: DWORD; out id: TDispID): HResult Of Object;
  TGetMemberNameEvent        = function (const id: TDispID; out Name: String): HResult Of Object;
  TGetMemberPropertiesEvent  = function (const id: TDispID; const grfdexFetch: DWORD; Out grfdex: DWORD): HResult Of Object;
  TGetNameSpaceParentEvent   = function (out unk: IUnknown): HResult Of Object;
  TGetNextDispIDEvent        = function (const grfdex: DWORD; const id: TDispID; out nid: TDispID): HResult Of Object;
  TInvokeExEvent             = function (const id: TDispID; const lcid: LCID; const wflags: WORD; const Parameters : IParameters; const varRes: POleVariant; const pei: PExcepInfo; const pspCaller: PServiceProvider): HResult Of Object;

  TDispatchExHandler=Class(TDispatchHandler,IDispatchEx)
  private
    fOnDeleteMemberByDispID : TDeleteMemberByDispIDEvent;
    fOnDeleteMemberByName   : TDeleteMemberByNameEvent;
    fOnGetDispID            : TGetDispIDEvent;
    fOnGetMemberName        : TGetMemberNameEvent;
    fOnGetMemberProperties  : TGetMemberPropertiesEvent;
    fOnGetNameSpaceParent   : TGetNameSpaceParentEvent;
    fOnGetNextDispID        : TGetNextDispIDEvent;
    fOnInvokeEx             : TInvokeExEvent;
  Protected
    function IDispatchEx_GetDispID(const bstrName: TBSTR; const grfdex: DWORD; out id: TDispID): HResult; stdcall;
    function IDispatchEx_InvokeEx(const id: TDispID; const lcid: LCID; const wflags: WORD; Const pdp: PDispParams; Const varRes: POleVariant; Const pei: PExcepInfo; const pspCaller: PServiceProvider): HResult; stdcall;
    function IDispatchEx_DeleteMemberByName(const bstr: TBSTR; const grfdex: DWORD): HResult; stdcall;
    function IDispatchEx_DeleteMemberByDispID(const id: TDispID): HResult; stdcall;
    function IDispatchEx_GetMemberProperties(const id: TDispID; const grfdexFetch: DWORD; out grfdex: DWORD): HResult; stdcall;
    function IDispatchEx_GetMemberName(const id: TDispID; out bstrName: TBSTR): HResult; stdcall;
    function IDispatchEx_GetNextDispID(const grfdex: DWORD; const id: TDispID; out nid: TDispID): HResult; stdcall;
    function IDispatchEx_GetNameSpaceParent(out unk: IUnknown): HResult; stdcall;

    function IDispatchEx.GetDispID = IDispatchEx_GetDispID;
    function IDispatchEx.InvokeEx = IDispatchEx_InvokeEx;
    function IDispatchEx.DeleteMemberByName = IDispatchEx_DeleteMemberByName;
    function IDispatchEx.DeleteMemberByDispID = IDispatchEx_DeleteMemberByDispID;
    function IDispatchEx.GetMemberProperties = IDispatchEx_GetMemberProperties;
    function IDispatchEx.GetMemberName = IDispatchEx_GetMemberName;
    function IDispatchEx.GetNextDispID = IDispatchEx_GetNextDispID;
    function IDispatchEx.GetNameSpaceParent = IDispatchEx_GetNameSpaceParent;
  Protected
    function DeleteMemberByDispID(const id: TDispID): HResult; Virtual;
    function DeleteMemberByName(const Name: String; const grfdex: DWORD): HResult; Virtual;
    function GetDispID(const Name: String; const grfdex: DWORD; out id: TDispID): HResult; Virtual;
    function GetMemberName(const id: TDispID; out Name: String): HResult; Virtual;
    function GetMemberProperties(const id: TDispID; const grfdexFetch: DWORD; out grfdex: DWORD): HResult; Virtual;
    function GetNameSpaceParent(out unk: IUnknown): HResult; Virtual;
    function GetNextDispID(const grfdex: DWORD; const id: TDispID; out nid: TDispID): HResult; Virtual;
    function InvokeEx(const id: TDispID; const lcid: LCID; const wflags: WORD; const Parameters : IParameters; Const varRes: POleVariant; Const pei: PExcepInfo; const pspCaller: PServiceProvider): HResult; Virtual;

    Property OnDeleteMemberByDispID : TDeleteMemberByDispIDEvent Read fOnDeleteMemberByDispID Write fOnDeleteMemberByDispID;
    Property OnDeleteMemberByName   : TDeleteMemberByNameEvent   Read fOnDeleteMemberByName   Write fOnDeleteMemberByName;
    Property OnGetDispID            : TGetDispIDEvent            Read fOnGetDispID            Write fOnGetDispID;
    Property OnGetMemberName        : TGetMemberNameEvent        Read fOnGetMemberName        Write fOnGetMemberName;
    Property OnGetMemberProperties  : TGetMemberPropertiesEvent  Read fOnGetMemberProperties  Write fOnGetMemberProperties;
    Property OnGetNameSpaceParent   : TGetNameSpaceParentEvent   Read fOnGetNameSpaceParent   Write fOnGetNameSpaceParent;
    Property OnGetNextDispID        : TGetNextDispIDEvent        Read fOnGetNextDispID        Write fOnGetNextDispID;
    Property OnInvokeEx             : TInvokeExEvent             Read fOnInvokeEx             Write fOnInvokeEx;
  End;


// This is a similar to TObjectDispatch, but instead of exporting everything with random DispIDs,
// RTTI attributes are used to provide a DispID and exported name (which can differ from the
// method name)
//
// Also contains the ability to export alternate IDispatch interface IDs via the IID provided
// in the constructor and CustomIID property.
//
// If you provide a class instance to the constructor, this will proxy IDispatch and IDispatchEx
// calls for the instance, or you can descend from this class and provide nil as the instance.
// to proxy directly from this class.
Type
  IDispatchInfo=Class(TCustomAttribute)
  Private
    fIID    : String;
    fDispID : TDispID;
    fName   : String;
  Protected
  Public
    Constructor Create(Const aIID : String; aDispID : TDispID; Const aName : String=''); Overload;
    Constructor Create(aDispID : TDispID; Const aName : String=''); Overload;

    Property IID    : String Read fIID;
    Property DispID : TDispID Read fDispID;
    Property Name   : String  Read fName;
  End;

  TDDUObjectDispatch=Class(TInterfacedObject,IInterface,IDispatch)
  Private
    fGetIDsOfNames      : TGetIDsOfNamesEvent;
    fGetTypeInfo        : TGetTypeInfoEvent;
    fGetTypeInfoCount   : TGetTypeInfoCountEvent;
    fInvoke             : TInvokeEvent;
  Private
    fInstance           : TObject;
    fOwned              : Boolean;
    fCustomIID          : TGUID;    // Experimental support
    fIgnoreNamed        : Boolean;
    function  GetInstance: TObject;
  Private type
    TDispatchKind = (dkUnknown, dkMethod, dkProperty, dkSubComponent);
    TDispatchInfo = record
      DelphiName : String;
      Instance   : TObject;
      Case Kind: TDispatchKind of
        dkUnknown     : (Value : Pointer);
        dkMethod      : (MethodInfo: PMethodInfoHeader);
        dkProperty    : (PropInfo: PPropInfo);
        dkSubComponent: (Index: Integer);
    end;
  Private
    fMapDispIDByDispatchName : TDictionary<String, TDispID>;
    fMapDispatchNameByDispID : TDictionary<TDispID, String>;
    fDispatchInfo            : TDictionary<TDispID, TDispatchInfo>;
    fDispatchIDs             : TList<TDispID>;
    fName                    : String;
    fNames                   : TDictionary<TDispID,String>;
  Protected
    procedure AddDispatchInfo(aDispID: TDispID; Const aDelphiName : String; AKind: TDispatchKind; Value: Pointer; AInstance: TObject);
    Function  GetDispatchInfo(aDispID : TDispID; Out DispatchInfo : TDispatchInfo) : Boolean;
    Function  GetDispatchNameByDispID(Const DispID : TDispID) : String;
    Function  GetDispIDByDispatchName(Const DispatchName : String) : TDispID;
    Function  GetDispIDName(Const DispID : TDispID; Out Name : String) : Boolean;
    Procedure RegisterDispID(DispID : TDispID; Const DispatchName : String; DelphiName : String = ''; SortImmediate : Boolean=True);

  Protected
    Function  Dump_PDP(Const PDP : PDispParams) : String;
    function  GetAsDispatch(Const aName : String; Obj: TObject): TDDUObjectDispatch; virtual;
    function  GetMethodInfo(const DelphiName: ShortString; var AInstance: TObject): PMethodInfoHeader; virtual;
    function  GetPropInfo(const DelphiName: string; var AInstance: TObject; var CompIndex: Integer): PPropInfo; virtual;

    property  Instance: TObject read GetInstance;
  Protected
    { IInterface / IUnknown }
    function  _AddRef: Integer; stdcall;
    function  _Release: Integer; stdcall;
    function  QueryInterface(const IID: TGUID; out Obj): HRESULT; stdcall;

    { IDispatch }
    function  IDispatch_GetIDsOfNames(const IID: TGUID; Names: Pointer; NameCount, LocaleID: Integer; DispIDs: Pointer): HResult; stdcall;
    function  IDispatch_GetTypeInfo(Index, LocaleID: Integer; out TypeInfo): HResult; stdcall;
    function  IDispatch_GetTypeInfoCount(out Count: Integer): HResult; stdcall;
    function  IDispatch_Invoke(DispID: Integer; const IID: TGUID; LocaleID: Integer; Flags: Word; var Params; VarResult, ExcepInfo, ArgErr: Pointer): HResult; virtual; stdcall;

    function  IDispatch.GetTypeInfoCount = IDispatch_GetTypeInfoCount;
    function  IDispatch.GetTypeInfo = IDispatch_GetTypeInfo;
    function  IDispatch.GetIDsOfNames = IDispatch_GetIDsOfNames;
    function  IDispatch.Invoke = IDispatch_Invoke;

    { IDispatch }
    function  GetIDsOfNames(const IID: TGUID; Names: Pointer; NameCount: Integer; LocaleID: Integer; DispIDs: Pointer): HRESULT; virtual;
    function  GetTypeInfo(Index: Integer; LocaleID: Integer; out TypeInfo): HRESULT; Virtual;
    function  GetTypeInfoCount(out Count: Integer): HRESULT; Virtual;
    function  Invoke(DispID: Integer; const IID: TGUID; LocaleID: Integer; Flags: Word; var Params; VarResult: Pointer; ExcepInfo: Pointer; ArgErr: Pointer): HRESULT; virtual;
  Protected
    Procedure Init; Virtual;
   Public
    Constructor Create(Const aName : String; aCustomIID : TGUID; Instance: TObject=Nil; Owned: Boolean = True; IgnoreNamed : Boolean=False); Overload;
    Constructor Create(Const aName : String; Instance: TObject=Nil; Owned: Boolean = True; IgnoreNamed : Boolean=False); Overload;
    Destructor  Destroy; Override;

    Procedure NameDispID(DispID : TDispID; Const DispatchName : String);
  Public
    Property OnGetIDsOfNames    : TGetIDsOfNamesEvent    Read fGetIDsOfNames    Write fGetIDsOfNames;
    Property OnGetTypeInfo      : TGetTypeInfoEvent      Read fGetTypeInfo      Write fGetTypeInfo;
    Property OnGetTypeInfoCount : TGetTypeInfoCountEvent Read fGetTypeInfoCount Write fGetTypeInfoCount;
    Property OnInvoke           : TInvokeEvent           Read fInvoke           Write fInvoke;
  Public
    Property CustomIID   : TGUID   Read fCustomIID   Write fCustomIID;
    Property IgnoreNamed : Boolean Read fIgnoreNamed Write fIgnoreNamed Default False;
    Property Name        : String  Read fName;
  End;

  TDDUObjectDispatchEx=Class(TDDUObjectDispatch,IDispatchEx)
  Private
  Protected
    { IDispatchEx }
    function IDispatchEx_DeleteMemberByDispID(const id: TDispID): HResult; stdcall;
    function IDispatchEx_DeleteMemberByName(const bstr: TBSTR; const grfdex: DWORD): HResult; stdcall;
    function IDispatchEx_GetDispID(const bstrName: TBSTR; const grfdex: DWORD; out id: TDispID): HResult; stdcall;
    function IDispatchEx_GetMemberName(const id: TDispID; out bstrName: TBSTR): HResult; stdcall;
    function IDispatchEx_GetMemberProperties(const id: TDispID; const grfdexFetch: DWORD; Out grfdex: DWORD): HResult; stdcall;
    function IDispatchEx_GetNameSpaceParent(out unk: IUnknown): HResult; stdcall;
    function IDispatchEx_GetNextDispID(const grfdex: DWORD; const id: TDispID; out nid: TDispID): HResult; stdcall;
    function IDispatchEx_InvokeEx(const DispID: TDispID; const lcid: LCID; const wflags: WORD; const Pdp: PDispParams; Const varResult: POleVariant; Const pei: PExcepInfo; const pspCaller: PServiceProvider): HResult; stdcall;

    function IDispatchEx.GetDispID = IDispatchEx_GetDispID;
    function IDispatchEx.InvokeEx = IDispatchEx_InvokeEx;
    function IDispatchEx.DeleteMemberByName = IDispatchEx_DeleteMemberByName;
    function IDispatchEx.DeleteMemberByDispID = IDispatchEx_DeleteMemberByDispID;
    function IDispatchEx.GetMemberProperties = IDispatchEx_GetMemberProperties;
    function IDispatchEx.GetMemberName = IDispatchEx_GetMemberName;
    function IDispatchEx.GetNextDispID = IDispatchEx_GetNextDispID;
    function IDispatchEx.GetNameSpaceParent = IDispatchEx_GetNameSpaceParent;
  Public

    Function jsHasOwnProperty(Const Name : String) : Boolean;

  End;

//************************************************************************************************************************************************************************************
//************************************************************************************************************************************************************************************

Type
  TDispatchExCaller=Class
  Private
    fInterface       : IDispatchEx;
    fLCID            : Integer;
    fServiceProvider : IServiceProvider;

    function  GetValue(Const Name: String): Variant;
    function  GetValueByDispID(DispID: TDispID): Variant;
    procedure SetValue(Const Name: String; const Value: Variant);
    procedure SetValueByDispID(DispID: TDispID; const Value: Variant);
  Protected
    Class Function GetDispID(intf : IDispatchEx; Const Name : String) : TDispID; Overload;
    Class Function InvokeEx(intf : IDispatchEx; lcid: Integer; Flags : Integer; DispID : TDispID; Const Args : Array Of Variant; Const Named : Array Of TDispID) : Variant; Overload;
    Function GetDispID(Const Name : String) : TDispID; Overload;
    Function InvokeEx(Flags : Integer; DispID : TDispID; Const Args : Array Of Variant; Const Named : Array Of TDispID) : Variant; Overload;
  Public
    Constructor Create(Const intf : IDispatchEx; alcid : Integer=0); Virtual;
    Destructor Destroy; Override;

    Function Execute(Const Name : String; Const Args : Array Of Variant) : Variant; Overload;
    Function Execute(ID : TDispID; Const Args : Array Of Variant) : Variant; Overload;

    Class Function Execute(intf : IDispatchEx; lcid: Integer; Const Name : String; Const Args : Array Of Variant) : Variant; Overload;
    Class Function Execute(intf : IDispatchEx; lcid: Integer; DispID : TDispID; Const Args : Array Of Variant) : Variant; Overload;

    Class function &Property(intf: IDispatchEx; lcid: Integer; const DispID : TDispID): Variant; Overload;
    Class function &Property(intf: IDispatchEx; lcid: Integer; const Name: String): Variant; Overload;

    Class procedure &Property(intf: IDispatchEx; lcid: Integer; const DispID : TDispID; const Value: Variant); Overload;
    Class procedure &Property(intf: IDispatchEx; lcid: Integer; const Name: String; const Value: Variant); Overload;

    Class Function DispID(intf: IDispatchEx; const Name : String; EnsureName : Boolean): TDispID; Overload;
  Public
    Property Value[Const Name : String]                                     : Variant          Read GetValue         Write SetValue;         Default;
    Property ValueByDispID[DispID : TDispID]                                : Variant          Read GetValueByDispID Write SetValueByDispID;
  Public
    Property lcid                                                           : Integer          Read fLCID            Write fLCID;
    Property intf                                                           : IDispatchEx      Read fInterface       Write fInterface;
    Property ServiceProvider                                                : IServiceProvider Read fServiceProvider Write fServiceProvider;
  End;

  TDispatchCaller=Class
  Private
    fInterface       : IDispatch;
    fLCID            : Integer;
    fServiceProvider : IServiceProvider;

    function  GetValue(Const Name: String): Variant;
    function  GetValueByDispID(DispID: TDispID): Variant;
    procedure SetValue(Const Name: String; const Value: Variant);
    procedure SetValueByDispID(DispID: TDispID; const Value: Variant);

  Protected
    Class Function GetDispID(intf : IDispatch; Const Name : String) : TDispID; Overload;
    Class Function Invoke(intf : IDispatch; lcid: Integer; Flags : Integer; DispID : TDispID; Const Args : Array Of Variant; Const Named : Array Of TDispID) : Variant; Overload;

    Function GetDispID(Const Name : String) : TDispID; Overload;
    Function Invoke(Flags : Integer; DispID : TDispID; Const Args : Array Of Variant; Const Named : Array Of TDispID) : Variant; Overload;
  Public
    Constructor Create(Const intf : IDispatch; alcid : Integer=0); Virtual;
    Destructor Destroy; Override;

    Function Execute(Const Name : String; Const Args : Array Of Variant) : Variant; Overload;
    Function Execute(ID : TDispID; Const Args : Array Of Variant) : Variant; Overload;

    Class Function Execute(intf : IDispatch; lcid: Integer; Const Name : String; Const Args : Array Of Variant) : Variant; Overload;
    Class Function Execute(intf : IDispatch; lcid: Integer; DispID : TDispID; Const Args : Array Of Variant) : Variant; Overload;

    Class function &Property(intf: IDispatch; lcid: Integer; const DispID: TDispID): Variant; Overload;
    Class function &Property(intf: IDispatch; lcid: Integer; const Name: String): Variant; Overload;
    Class procedure &Property(intf: IDispatch; lcid: Integer; const DispID: TDispID; const Value: Variant); Overload;
    Class procedure &Property(intf: IDispatch; lcid: Integer; const Name: String; const Value: Variant); Overload;

    Class function PropertyDef(intf: IDispatch; lcid: Integer; const Name: String; Const Default : Variant): Variant; Overload;
    Class function PropertyDef(intf: IDispatch; lcid: Integer; const DispID : TDispID; Const Default : Variant): Variant; Overload;
  Public
    Property Value[Const Name : String]      : Variant          Read GetValue         Write SetValue;         Default;
    Property ValueByDispID[DispID : TDispID] : Variant          Read GetValueByDispID Write SetValueByDispID;
  Public
    Property lcid                            : Integer          Read fLCID            Write fLCID;
    Property intf                            : IDispatch        Read fInterface       Write fInterface;
    Property ServiceProvider                 : IServiceProvider Read fServiceProvider Write fServiceProvider;
  End;

Function SafeVarToStrDef(V : Variant; Const Default : String) : String;

implementation

type
  PVariantArray = ^TVariantArray;
  TVariantArray = array[0..65535] of Variant;
  PIntegerArray = ^TIntegerArray;
  TIntegerArray = array[0..65535] of Integer;

Function SafeVarToStrDef(V : Variant; Const Default : String) : String;

Begin
  Case FindVarData(V)^.VType Of
    varEmpty     : Result := '(varEmpty)';
    varError     : Result := '(varError)';
    varDispatch  : Result := '(IDispatch)';
    varUnknown   : Result := '(varUnknown)';
    varNull      : Result := '(varNull)';
  ELse
    Try
      Result := VarToStrDef(V,Default);
    Except
      Result := Default
    End;
  End;
End;

Procedure Debug(Const What : String); Overload;

{$IF not defined(RELEASE)}
Var
  S                       : TStringList;
  Loop                    : Integer;
{$ENDIF}

Begin
{$IF not defined(RELEASE)}
  S := TStringList.Create;
  Try
    S.Text := What.TrimRight;
    For Loop := 0 To S.Count-1 Do
    Begin
      OutputDebugString(PChar('DispatchEx: '+S[Loop]));
    End;
  Finally
    S.Free;
  End;
{$ENDIF}
End;

Procedure Debug(Const FormatStr : String; Const Args : Array Of Const); Overload;

Begin
  DDU.DispatchEx.Debug(Format(FormatStr,Args));
End;

function StringToShortUTF8String(S: string): ShortString;
begin
  Result[0] := AnsiChar(UnicodeToUtf8(@Result[1], 255, PWideChar(S), Cardinal(-1)) - 1);
end;


{$REGION 'TDispatchExCaller'}

{ TDispatchEx }

class function TDispatchExCaller.&Property(intf: IDispatchEx; lcid: Integer; const DispID: TDispID): Variant;
begin
  Result := InvokeEx(intf,lcid,DISPATCH_PROPERTYGET,DispID,[],[]);
end;

class procedure TDispatchExCaller.&Property(intf: IDispatchEx; lcid: Integer; const DispID: TDispID; const Value: Variant);
begin
  InvokeEx(intf,lcid,DISPATCH_PROPERTYPUT,DispID,[Value],[DISPID_PROPERTYPUT]);
end;

constructor TDispatchExCaller.Create(Const intf : IDispatchEx; alcid : Integer=0);
begin
  Inherited Create;
  fLCID := alcid;
  fInterface := intf;
end;

destructor TDispatchExCaller.Destroy;
begin

  inherited;
end;

class function TDispatchExCaller.DispID(intf: IDispatchEx; const Name : String; EnsureName: Boolean): TDispID;

Var
  lResult                 : HResult;
  Flags                   : Cardinal;

Begin
  Flags := fdexNameCaseInsensitive;
  If EnsureName Then
  Begin
    Flags := fdexNameEnsure;
  End;

  lResult := intf.GetDispID(PWideChar(Name),  Flags , Result);
  If Not Succeeded(lResult) Then
  Begin
    Raise EOleSysError.Create(Name, lResult, 0);
  End;
end;

class function TDispatchExCaller.Execute(intf : IDispatchEx; lcid: Integer; const Name: String; const Args: array of Variant): Variant;
begin
  Result := InvokeEx(intf,lcid,DISPATCH_METHOD,GetDispID(intf,Name),Args,[]);
end;

class function TDispatchExCaller.Execute(intf : IDispatchEx; lcid: Integer;DispID: TDispID; const Args: array of Variant): Variant;
begin
  Result := InvokeEx(intf,lcid,DISPATCH_METHOD,DispID,Args,[]);
end;

function TDispatchExCaller.Execute(const Name: String; const Args: array of Variant): Variant;
begin
  Result := InvokeEx(DISPATCH_METHOD,GetDispID(Name),Args,[]);
end;

function TDispatchExCaller.Execute(ID: TDispID; const Args: array of Variant): Variant;
begin
  Result := InvokeEx(DISPATCH_METHOD,ID,Args,[]);
end;

function TDispatchExCaller.GetDispID(const Name: String): TDispID;

Begin
// We could cache the results here for better speed.
  Result := GetDispID(intf,Name);
end;

class function TDispatchExCaller.GetDispID(intf : IDispatchEx; Const Name : String) : TDispID;

Var
  lResult                 : HResult;

Begin
  lResult := intf.GetDispID(PWideChar(Name),  fdexNameCaseInsensitive, Result);
  If Not Succeeded(lResult) Then
  Begin
    Raise EOleSysError.Create(Name, lResult, 0);
  End;
end;

function TDispatchExCaller.GetValue(Const Name: String): Variant;
begin
  Result := InvokeEx(DISPATCH_PROPERTYGET,GetDispID(Name),[],[]);
end;

function TDispatchExCaller.GetValueByDispID(DispID: TDispID): Variant;
begin
  Result := InvokeEx(DISPATCH_PROPERTYGET,DispID,[],[]);
end;

class function TDispatchExCaller.InvokeEx(intf: IDispatchEx; lcid: Integer; Flags : Integer; DispID: TDispID; const Args: array of Variant; const Named: array of TDispID): Variant;

Var
  lDispParams             : DISPPARAMS;
  lExceptInfo             : TExcepInfo;
  lArgs                   : Array Of OleVariant;
  lNamed                  : Array Of TDispID;
  Loop                    : Integer;
  lResult                 : HRESULT;
  varRes                  : OleVariant;

Begin
  SetLength(lArgs,Length(Args));
  For Loop := Low(Args) To High(Args) Do
    lArgs[Loop] := Args[Loop];

  SetLength(lNamed,Length(Named));
  For Loop := Low(Named) To High(Named) Do
    lNamed[Loop] := Named[Loop];

  FillChar(lDispParams,SizeOf(lDispParams),0);
  FillChar(lExceptInfo,SizeOf(lExceptInfo),0);

  lDispParams.cArgs             := Length(lArgs);
  lDispParams.rgvarg            := PVariantArgList(lArgs);
  lDispParams.cNamedArgs        := Length(lNamed);
  lDispParams.rgDispIDNamedArgs := PDispIDList(lNamed);

  lResult := intf.InvokeEx(DispID,LCID,Flags,@lDispParams,@varRes,@lExceptInfo,Nil);

  If Succeeded(lResult) Then
  Begin
    Result := varRes;
  End
  Else
  Begin
    DispatchInvokeError(lResult,lExceptInfo);
  End;
End;

function TDispatchExCaller.InvokeEx(Flags: Integer; DispID: TDispID; const Args: array of Variant; const Named: array of TDispID): Variant;

Begin
  Result := InvokeEx(intf,lcid,Flags,DispID,Args,Named);
End;

procedure TDispatchExCaller.SetValue(Const Name: String; const Value: Variant);
begin
  InvokeEx(DISPATCH_PROPERTYPUT,GetDispID(Name),[Value],[DISPID_PROPERTYPUT]);
end;

procedure TDispatchExCaller.SetValueByDispID(DispID: TDispID; const Value: Variant);
begin
  InvokeEx(DISPATCH_PROPERTYPUT,DispID,[Value],[DISPID_PROPERTYPUT]);
end;

class function TDispatchExCaller.&Property(intf: IDispatchEx; lcid: Integer; const Name: String): Variant;
begin
  Result := InvokeEx(intf,lcid,DISPATCH_PROPERTYGET,GetDispID(intf,Name),[],[]);
end;

class procedure TDispatchExCaller.&Property(intf: IDispatchEx; lcid: Integer; const Name: String; const Value: Variant);
begin
  InvokeEx(intf,lcid,DISPATCH_PROPERTYPUT,GetDispID(intf,Name),[Value],[DISPID_PROPERTYPUT]);
end;

{$ENDREGION}

{$REGION 'TDispatchCaller'}

{ TDispatch }

class function TDispatchCaller.&Property(intf: IDispatch; lcid: Integer; const DispID: TDispID): Variant;
begin
  Result := Invoke(intf,lcid,DISPATCH_PROPERTYGET,DispID,[],[]);
end;

class procedure TDispatchCaller.&Property(intf: IDispatch; lcid: Integer; const DispID: TDispID; const Value: Variant);
begin
  Invoke(intf,lcid,DISPATCH_PROPERTYPUT,DispID,[Value],[DispID_PROPERTYPUT]);
end;

class function TDispatchCaller.PropertyDef(intf: IDispatch; lcid: Integer; const Name: String; const Default: Variant): Variant;

Var
  aDispID                 : TDispID;

begin
  aDispID := GetDispID(intf,Name);
  Try
    If aDispID=DISPID_UNKNOWN Then
    Begin
      Result := Default;
    End
    Else
    Begin
      Result := &Property(intf,lcid,aDispID);
    End;
  Except
    Result := Default;
  End;
end;

class function TDispatchCaller.PropertyDef(intf: IDispatch; lcid: Integer; const DispID: TDispID; const Default: Variant): Variant;

begin
  Try
    Result := &Property(intf,lcid,DispID);
  Except
    Result := Default;
  End;
end;

constructor TDispatchCaller.Create(Const intf : IDispatch; alcid : Integer=0);
begin
  Inherited Create;
  fLCID := alcid;
  fInterface := intf;
end;

destructor TDispatchCaller.Destroy;
begin

  inherited;
end;

class function TDispatchCaller.Execute(intf : IDispatch; lcid: Integer; const Name: String; const Args: array of Variant): Variant;
begin
  Result := Invoke(intf,lcid,DISPATCH_METHOD,GetDispID(intf,Name),Args,[]);
end;

class function TDispatchCaller.Execute(intf : IDispatch; lcid: Integer;DispID: TDispID; const Args: array of Variant): Variant;
begin
  Result := Invoke(intf,lcid,DISPATCH_METHOD,DispID,Args,[]);
end;

function TDispatchCaller.Execute(const Name: String; const Args: array of Variant): Variant;
begin
  Result := Invoke(DISPATCH_METHOD,GetDispID(Name),Args,[]);
end;

function TDispatchCaller.Execute(ID: TDispID; const Args: array of Variant): Variant;
begin
  Result := Invoke(DISPATCH_METHOD,ID,Args,[]);
end;

function TDispatchCaller.GetDispID(const Name: String): TDispID;

Begin
// We could cache the results here for better speed.
  Result := GetDispID(intf,Name);
end;

class function TDispatchCaller.GetDispID(intf : IDispatch; Const Name : String) : TDispID;

Var
  lResult                 : HResult;
  Names                   : Array Of POleStr;
  IDs                     : Array Of TDispID;

Begin
  SetLength(Names,1);
  SetLength(IDs,1);
  Names[0] := StringToOleStr(Name);
  IDs[0]   := DISPID_UNKNOWN;

  lResult := intf.GetIDsOfNames(GUID_NULL,Names,1,0,IDs);
  If Succeeded(lResult) Then
  Begin
    Result := IDs[0];
  End
  Else
  Begin
    Raise EOleSysError.Create(Name, lResult, 0);
  End;
end;

function TDispatchCaller.GetValue(Const Name: String): Variant;
begin
  Result := Invoke(DISPATCH_PROPERTYGET,GetDispID(Name),[],[]);
end;

function TDispatchCaller.GetValueByDispID(DispID: TDispID): Variant;
begin
  Result := Invoke(DISPATCH_PROPERTYGET,DispID,[],[]);
end;

class function TDispatchCaller.Invoke(intf: IDispatch; lcid: Integer; Flags : Integer; DispID: TDispID; const Args: array of Variant; const Named: array of TDispID): Variant;

Var
  lDispParams             : DISPPARAMS;
  lExceptInfo             : TExcepInfo;
  lArgs                   : Array Of OleVariant;
  lNamed                  : Array Of TDispID;
  Loop                    : Integer;
  lResult                 : HRESULT;
  varRes                  : OleVariant;

Begin
  SetLength(lArgs,Length(Args));
  For Loop := Low(Args) To High(Args) Do
    lArgs[Loop] := Args[Loop];

  SetLength(lNamed,Length(Named));
  For Loop := Low(Named) To High(Named) Do
    lNamed[Loop] := Named[Loop];

  FillChar(lDispParams,SizeOf(lDispParams),0);
  FillChar(lExceptInfo,SizeOf(lExceptInfo),0);

  lDispParams.cArgs             := Length(lArgs);
  lDispParams.rgvarg            := PVariantArgList(lArgs);
  lDispParams.cNamedArgs        := Length(lNamed);
  lDispParams.rgDispIDNamedArgs := PDispIDList(lNamed);

  lResult := intf.Invoke(DispID,GUID_NULL,LCID,Flags,lDispParams, @varRes,@lExceptInfo,Nil);

  If Succeeded(lResult) Then
  Begin
    Result := varRes;
  End
  Else
  Begin
    DispatchInvokeError(lResult,lExceptInfo);
  End;
End;

function TDispatchCaller.Invoke(Flags: Integer; DispID: TDispID; const Args: array of Variant; const Named: array of TDispID): Variant;

Begin
  Result := Invoke(intf,lcid,Flags,DispID,Args,Named);
End;

procedure TDispatchCaller.SetValue(Const Name: String; const Value: Variant);
begin
  Invoke(DISPATCH_PROPERTYPUT,GetDispID(Name),[Value],[DISPID_PROPERTYPUT]);
end;

procedure TDispatchCaller.SetValueByDispID(DispID: TDispID; const Value: Variant);
begin
  Invoke(DISPATCH_PROPERTYPUT,DispID,[Value],[DISPID_PROPERTYPUT]);
end;

class function TDispatchCaller.&Property(intf: IDispatch; lcid: Integer; const Name: String): Variant;
begin
  Result := Invoke(intf,lcid,DISPATCH_PROPERTYGET,GetDispID(intf,Name),[],[]);
end;

class procedure TDispatchCaller.&Property(intf: IDispatch; lcid: Integer; const Name: String; const Value: Variant);
begin
  Invoke(intf,lcid,DISPATCH_PROPERTYPUT,GetDispID(intf,Name),[Value],[DISPID_PROPERTYPUT]);
end;

{$ENDREGION}

{$REGION 'IDispatchInfo'}

{ IDispatchInfo }

constructor IDispatchInfo.Create(Const aIID : String; aDispID: Integer; Const aName: String='');
begin
  Inherited Create;
  fIID    := aIID;
  fDispID := aDispID;
  fName   := aName.ToLower;
end;

constructor IDispatchInfo.Create(aDispID: TDispID; const aName: String);
begin
  Inherited Create;
  fDispID := aDispID;
  fName   := aName.ToLower;
end;


{$ENDREGION}

{$REGION 'TDDUObjectDispatch/IDispatch'}

{ TDDUObjectDispatch }

procedure TDDUObjectDispatch.AddDispatchInfo(aDispID: TDispID; Const aDelphiName : String; AKind: TDispatchKind; Value: Pointer; AInstance: TObject);

Var
  DispatchInfo            : TDispatchInfo;

begin
  DispatchInfo.DelphiName := aDelphiName;
  DispatchInfo.Instance   := AInstance;
  DispatchInfo.Kind       := AKind;
  DispatchInfo.Value      := Value;
  fDispatchInfo.AddOrSetValue(aDispID,DispatchInfo);
end;

constructor TDDUObjectDispatch.Create(Const aName : String; Instance: TObject=Nil; Owned: Boolean = True; IgnoreNamed : Boolean=False);
begin
  Inherited Create;
  DDU.DispatchEx.Debug('CREATE: %s//%s',[Self.ClassName,aName]);
  fName         := aName;
  fCustomIID    := GUID_NULL;
  fInstance     := Instance;
  fOwned        := Owned And Assigned(fInstance);
  fIgnoreNamed  := IgnoreNamed;
  Init;
end;

constructor TDDUObjectDispatch.Create(Const aName : String; aCustomIID : TGUID; Instance: TObject=Nil; Owned: Boolean = True; IgnoreNamed : Boolean=False );
begin
  Inherited Create;
  DDU.DispatchEx.Debug('CREATE: %s//%s',[Self.ClassName,aName]);
  fName        := aName;
  fCustomIID   := aCustomIID;
  fInstance    := Instance;
  fOwned       := Owned;
  fIgnoreNamed := IgnoreNamed;
  Init;
end;

function TDDUObjectDispatchEx.IDispatchEx_DeleteMemberByDispID(const id: TDispID): HResult;

begin
  If fMapDispatchNameByDispID.ContainsKey(ID) Then
  Begin
    Result := S_FALSE;
  End
  Else
  Begin
    Result := S_OK;
  End;
end;

function TDDUObjectDispatchEx.IDispatchEx_DeleteMemberByName(const bstr: TBSTR; const grfdex: DWORD): HResult;
begin
  If GetDispIDByDispatchName(bStr)<>DispID_UNKNOWN Then
  Begin
    Result := S_FALSE;
  End
  Else
  Begin
    Result := S_OK;
  End;
end;

destructor TDDUObjectDispatch.Destroy;
begin
  DDU.DispatchEx.Debug('DESTROY : %s//%s',[Self.ClassName,fName]);
  fMapDispIDByDispatchName.Free;
  fMapDispatchNameByDispID.Free;
  fDispatchInfo.Free;
  fNames.Free;

  if fOwned And Assigned(fInstance) then
  Begin
    fInstance.Free;
  End;
  inherited;
end;

function TDDUObjectDispatch.Dump_PDP(const PDP: PDispParams): String;

Var
  Loop                    : Integer;
  V                       : Variant;
  D                       : TDISPID;
  S                       : TStringList;

  Value                   : String;

begin
  If Assigned(Pdp) Then
  Begin
    S := TStringList.Create;
    Try
      S.Add(Format('  cArgs(%d)',[Pdp.cArgs]));
      For Loop := 0 To Pdp.cArgs-1 Do
      Begin
        V := PVariantArray(Pdp.rgvarg)[Loop];

        Value := SafeVarToStrDef( V,'(other?)');

        S.Add(Format('    %d: %s (%s)',[Loop,Value,VarTypeAsText(VarType(V))]));
      End;

      S.Add(Format('  cNamedArgs(%d)',[pdp.cNamedArgs]));
      For Loop := 0 To Pdp.cNamedArgs-1 Do
      Begin
        D      := Pdp.rgdispidNamedArgs[Loop];
        S.Add(Format('    %d: DispID %d', [Loop, D]));
      End;
      S.Add('--');

      Result := S.Text;
    Finally
      S.Free;
    End;
  End
  Else
  Begin
    Result := '';
  End;
end;

function TDDUObjectDispatch.GetDispIDByDispatchName(const DispatchName: String): TDispID;
begin
  If fMapDispIDByDispatchName.ContainsKey(DispatchName.ToLower) Then
  Begin
    Result := fMapDispIDByDispatchName.Items[DispatchName.ToLower];
  End
  Else
  Begin
    Result := DISPID_UNKNOWN;
  End;
end;

function TDDUObjectDispatch.GetDispIDName(const DispID: TDispID; out Name: String): Boolean;
begin
  If fMapDispatchNameByDispID.ContainsKey(DispID) Then
  Begin
    Name   := fMapDispatchNameByDispID.Items[DispID];
    Result := True;
  End
  Else
  Begin
    Name   := '';
    Result := False;
  End;
end;

function TDDUObjectDispatch.GetAsDispatch(Const aName : String; Obj: TObject): TDDUObjectDispatch;
begin
  Result := TDDUObjectDispatch.Create(aName,Obj, False);
end;

function TDDUObjectDispatch.GetDispatchInfo(aDispID : TDispID; Out DispatchInfo : TDispatchInfo) : Boolean;
begin
  If fDispatchInfo.ContainsKey(aDispID) Then
  Begin
    DispatchInfo := fDispatchInfo.Items[aDispID];
    Result := True;
  End
  Else
  Begin
    Result := False;
  End;
end;

function TDDUObjectDispatch.GetDispatchNameByDispID(const DispID: TDispID): String;
begin
  If fMapDispatchNameByDispID.ContainsKey(DispID) Then
  Begin
    Result := fMapDispatchNameByDispID.Items[DispID];
  End
  Else
  Begin
    Result := '';
  End;
end;

function TDDUObjectDispatchEx.IDispatchEx_GetDispID(const bstrName: TBSTR; const grfdex: DWORD; out id: TDispID): HResult;
begin
  ID := GetDispIDByDispatchName(bstrName);
  If id=DispID_UNKNOWN Then
  Begin
    Result := DISP_E_UNKNOWNNAME;
  End
  Else
  Begin
    Result := S_OK;
  End;
end;

function TDDUObjectDispatch.GetIDsOfNames(const IID: TGUID; Names: Pointer; NameCount, LocaleID: Integer; DispIDs: Pointer): HRESULT;

Type
  PNames   = ^TNames;
  TNames   = Array[0..100] Of POleStr;
  PDispIDs = ^TDispIDs;
  TDispIDs = Array[0..100] Of Cardinal;

Var
  Name                    : String;
  Info                    : PMethodInfoHeader;
  PropInfo                : PPropInfo;
  InfoEnd                 : Pointer;
  Params                  : PParamInfo;
  Param                   : PParamInfo;
  I                       : Integer;
  ID                      : Cardinal;
  CompIndex               : Integer;
  Instance                : TObject;

  aDispID                 : Integer;
  DispatchInfo            : TDispatchInfo;
//  DispName                : String;

begin
  DDU.DispatchEx.Debug('%s.GetIDsOFNames',[Self.Name]);
  Result := S_OK;
  /// This assumes that the DispIDs are provide as RTTI properties
  ///
  Name := PNames(Names)^[0];

  FillChar(DispIDs^, SizeOf(PDispIDs(DispIDs^)[0]) * NameCount, $FF);
  aDispID := GetDispIDByDispatchName(Name);

  If aDispID<>DispID_UNKNOWN Then
  Begin
    GetDispatchInfo(aDispID,DispatchInfo);
    Info := GetMethodInfo(StringToShortUTF8String(DispatchInfo.DelphiName), Instance);
    if Info = nil then
    begin
      // Not a  method, try a property.
      PropInfo := GetPropInfo(DispatchInfo.DelphiName, Instance, CompIndex);

      if PropInfo <> nil then
        PDispIDs(DispIDs)^[0] := aDispID
      else if CompIndex > -1 then
        PDispIDs(DispIDs)^[0] := aDispID
      else
        Result := DISP_E_UNKNOWNNAME
    end
    else
    begin
      // Ensure the method information has enough type information
      if Info.Len <= SizeOf(Info^) - SizeOf(TSymbolName) + 1 + Info.NameFld.UTF8Length then
      Begin
        Result := DISP_E_UNKNOWNNAME;  //
      End
      else
      begin
        PDispIDs(DispIDs)^[0] := aDispID;
        Result := S_OK;
        if NameCount > 1 then
        begin
          // Now find the parameters. The DispID is assumed to be the parameter index.
          InfoEnd := Pointer(PByte(Info) + Info^.Len);
          Params  := PParamInfo(PByte(Info) + SizeOf(Info^) - SizeOf(TSymbolName) + 1 + SizeOf(TReturnInfo) + Info.NameFld.UTF8Length);
          for I := 1 to NameCount - 1 do
          begin
            Name  := PNames(Names) ^[I];
            Param := Params;
            ID    := 0;
            while IntPtr(Param) < IntPtr(InfoEnd) do
            begin
              // ignore Self
              if (Param^.ParamType^.Kind <> tkClass) or not SameText(Param^.NameFld.ToString, 'SELF') then
              Begin
                if SameText(Param.NameFld.ToString, Name) then
                begin
                  PDispIDs(DispIDs)^[I] := ID;
                  Break;
                end;
              End;
              Inc(ID);
              Param := PParamInfo(PByte(Param) + SizeOf(Param^) - SizeOf(TSymbolName) + 1 + Param^.NameFld.UTF8Length);
            end;
            if IntPtr(Param) >= IntPtr(InfoEnd) then
            Begin
              Result := DISP_E_UNKNOWNNAME
            End;
          end;
        end;
      end;
    end;
  End
  Else
  Begin
    Result := DISP_E_UNKNOWNNAME
  End;
end;

function TDDUObjectDispatch.GetInstance: TObject;
begin
  If Assigned(fInstance) Then
    Result := fInstance
  Else
    Result := Self;
end;

function TDDUObjectDispatchex.IDispatchEx_GetMemberName(const id: TDispID; out bstrName: TBSTR): HResult;

begin
  if fMapDispatchNameByDispID.ContainsKey(ID) Then
  Begin
    bStrName := StringToOleStr(fMapDispatchNameByDispID.Items[id]);
    Result := S_OK;
  End
  Else
  Begin
    Result := DISP_E_UNKNOWNNAME;
  End;
end;

function TDDUObjectDispatchEx.IDispatchEx_GetMemberProperties(const id: TDispID; const grfdexFetch: DWORD; Out grfdex: DWORD): HResult;

Var
  DispatchInfo : TDispatchInfo;

begin
  if (ID<>DispID_UNKNOWN) And GetDispatchInfo(ID,DispatchInfo) Then
  Begin
    grfdex := (fdexPropCannotAll or fdexPropCanSourceEvents) And (Not fdexPropNoSideEffects);

    case DispatchInfo.Kind of
      dkProperty:
        begin
          grfdex := (grfdex or fdexPropCanGet) And (Not fdexPropCannotGet);
          grfdex := (grfdex or fdexPropCanPut) And (Not fdexPropCannotPut);

          if DispatchInfo.PropInfo^.PropType^.Kind = tkClass then
          Begin
            grfdex := (grfdex or fdexPropCanPutRef) And (Not (fdexPropCannotPutRef));
          End;
        end;
      dkMethod:
        begin
          grfdex := (grfdex or fdexPropCanCall) And Not (fdexPropCannotCall);
        end;
      dkSubComponent:
        Begin
          grfdex := (grfdex or fdexPropCanGet) And (Not fdexPropCannotGet);
          grfdex := (grfdex or fdexPropCanPut) And (Not fdexPropCannotPut);
          grfdex := (grfdex or fdexPropCanPutRef) And (Not (fdexPropCannotPutRef or fdexPropNoSideEffects));
        end;
    end;
    grfdex := grfdex And grfdexFetch;

    Result := S_OK;
  End
  Else
  Begin
    Result := DISP_E_UNKNOWNNAME;
  End;

//  Result := DISP_E_UNKNOWNNAME;
end;

function TDDUObjectDispatch.GetMethodInfo(const DelphiName: ShortString; var AInstance: TObject): PMethodInfoHeader;
begin
  Result := System.ObjAuto.GetMethodInfo(Instance, UTF8ToString(DelphiName));
  if Result <> nil then
    AInstance := Instance;
end;

function TDDUObjectDispatchEx.IDispatchEx_GetNameSpaceParent(out unk: IInterface): HResult;
begin
  Result := E_NOTIMPL;
end;

function TDDUObjectDispatchEx.IDispatchEx_GetNextDispID(const grfdex: DWORD; const id: TDispID; out nid: TDispID): HResult;

Var
  Loop                    : Integer;
  At                      : Integer;

begin
  If (fDispatchIDs.Count=0) Or (grfdex=fdexEnumDefault) Then
  Begin
    nid := id;
    Exit(S_FALSE);
  End;

  If (id=DISPID_STARTENUM) Then
  Begin
    nid := fDispatchIDs[0];
    Exit(S_OK);
  End
  Else
  Begin
    At := fDispatchIDs.IndexOf(ID);
    If At=-1 Then
    Begin
      For Loop := 0 To fDispatchIDs.Count-1 Do
      Begin
        If fDispatchIDs[Loop]>ID Then
        Begin
          At := Loop;
          Break;
        End;
      End;
    End
    Else
    Begin
      At := At+1;
      If At>=fDispatchIDs.Count Then At := -1;
    End;

    If At=-1 Then
    Begin
      nid := id;
      Exit(S_FALSE);
    End
    Else
    Begin
      nid := fDispatchIDs[At];
      Exit(S_OK);
    End;
  End;

end;

function TDDUObjectDispatch.GetPropInfo(const DelphiName: string; var AInstance: TObject; var CompIndex: Integer): PPropInfo;

Var
  Component               : TComponent;

begin
  CompIndex := -1;
  Result := System.TypInfo.GetPropInfo(Instance, DelphiName);
  if (Result = nil) and (Instance is TComponent) then
  begin
    // Not a property, try a sub component
    Component := TComponent(Instance).FindComponent(DelphiName);
    if Component <> nil then
    begin
      AInstance := Instance;
      CompIndex := Component.ComponentIndex;
    end;
  end else if Result <> nil then
    AInstance := Instance
  else
    AInstance := nil;
end;

function TDDUObjectDispatch.GetTypeInfo(Index, LocaleID: Integer; out TypeInfo): HRESULT;

begin
  DDU.DispatchEx.Debug('%s.GetTypeInfo %d',[Self.Name,Index]);
  Result := E_NOTIMPL;
end;

function TDDUObjectDispatch.GetTypeInfoCount(out Count: Integer): HRESULT;
begin
  DDU.DispatchEx.Debug('%s.GetTypeInfoCount',[Self.Name]);
  Count := 0;
  Result := S_OK;
end;

function TDDUObjectDispatch.IDispatch_GetIDsOfNames(const IID: TGUID; Names: Pointer; NameCount, LocaleID: Integer; DispIDs: Pointer): HResult;
begin
  If Assigned(fGetIDsOfNames) Then
  Begin
    Result := fGetIDsOfNames(IID,Names,NameCount,LocaleID,DispIDs);
  End
  ELse
  Begin
    Result := GetIDsOfNames(IID,Names,NameCount,LocaleID,DispIDs);
  End;
end;

function TDDUObjectDispatch.IDispatch_GetTypeInfo(Index, LocaleID: Integer; out TypeInfo): HResult;
begin
  If Assigned(fGetTypeInfo) Then
  Begin
    Result := fGetTypeInfo(Index,LocaleID,TypeInfo);
  End
  Else
  Begin
    Result := GetTypeInfo(Index,LocaleID,TypeInfo);
  End;
end;

function TDDUObjectDispatch.IDispatch_GetTypeInfoCount(out Count: Integer): HResult;
begin
  If Assigned(fGetTypeInfoCount) Then
  Begin
    Result := fGetTypeInfoCount(Count);
  End
  Else
  Begin
    Result := GetTypeInfoCount(Count);
  End;
end;

function TDDUObjectDispatch.IDispatch_Invoke(DispID: Integer; const IID: TGUID; LocaleID: Integer; Flags: Word; var Params; VarResult, ExcepInfo, ArgErr: Pointer): HResult;

Var
  Parameters              : IParameters;

begin
  If Assigned(fInvoke) Then
  Begin
    If @Params<>Nil Then
    Begin
      Parameters := TParameters.Create(@TDispParams(Params));
    End
    Else
    Begin
      Parameters := TParameters.Create(Nil);
    end;
    Result := fInvoke(DispID,IID,LocaleID,Flags,Parameters,VarResult,ExcepInfo,ArgErr);
  End
  Else
  Begin
    Result := Invoke(DispID,IID,LocaleID,Flags,Params,VarResult,ExcepInfo,ArgErr);
  End;

end;

procedure TDDUObjectDispatch.Init;

var
  ctx                     : TRttiContext; // Static record, does not require initilization
  aType                   : TRttiType;
  aMethod                 : TRttiMethod;
  aProperty               : TRttiProperty;
  anAttribute             : TCustomAttribute;
  DispatchInfo            : IDispatchInfo;

begin
  DDU.DispatchEx.Debug('Init %s',[Self.Name]);

  fNames                      := TDictionary<TDispID , String>.Create;
  fMapDispIDByDispatchName    := TDictionary<String,TDispID>.Create;
  fMapDispatchNameByDispID    := TDictionary<TDispID,String>.Create;
  fDispatchInfo               := TDictionary<TDispID,TDispatchInfo>.Create;
  fDispatchIDs                := TList<TDispID>.Create;

  aType := ctx.GetType(Instance.ClassType);
  For aMethod In aType.GetMethods Do
  Begin
    For anAttribute In aMethod.GetAttributes Do
    Begin
      if (anAttribute is IDispatchInfo) Then
      begin
        DispatchInfo := IDispatchInfo(anAttribute);

        If (DispatchInfo.IID='') Or SameText(DispatchInfo.IID,GUIDToString(fCustomIID)) Then
        Begin
          RegisterDispID(DispatchInfo.DispID,DispatchInfo.Name, aMethod.Name,False);
        End;
      End;
    End;
  End;

  For aProperty In aType.GetProperties Do
  Begin
    For anAttribute In aProperty.GetAttributes Do
    Begin
      if (anAttribute is IDispatchInfo) Then
      begin
        DispatchInfo := IDispatchInfo(anAttribute);

        If (DispatchInfo.IID='') Or SameText(DispatchInfo.IID,GUIDToString(fCustomIID)) Then
        Begin
          RegisterDispID(DispatchInfo.DispID,DispatchInfo.Name, aProperty.Name,False);
        End;
      End;
    End;
  End;

  fDispatchIDs.Sort;
end;

procedure TDDUObjectDispatch.RegisterDispID(DispID : TDispID; Const DispatchName : String; DelphiName : String = ''; SortImmediate : Boolean=True);

Var
  Info                    : PMethodInfoHeader;
  PropInfo                : PPropInfo;
  CompIndex               : Integer;
  Instance                : TObject;

begin
  If (DispatchName<>'') Then
  Begin
    fMapDispIDByDispatchName.AddOrSetValue(DispatchName, DispID);
    fMapDispatchNameByDispID.AddOrSetValue(DispID,DispatchName);
  End;

  If fDispatchIDs.IndexOf(DispID)=-1 Then
  Begin
    fDispatchIDs.Add(DispID);
    If SortImmediate Then
    Begin
      fDispatchIDs.Sort;
    End;
  End;

  if (DelphiName='') Then
  Begin
    // Do nothing
  End
  Else
  Begin
    Info := GetMethodInfo(StringToShortUTF8String(DelphiName), Instance);
    if Info = nil then
    begin
      // Not a  method, try a property.
      PropInfo := GetPropInfo(DelphiName, Instance, CompIndex);

      if PropInfo <> nil then
      Begin
        AddDispatchInfo(DispID,DelphiName,dkProperty, PropInfo, Instance);
      End else if CompIndex > -1 then
      Begin
        AddDispatchInfo(DispID,DelphiName,dkSubComponent, Pointer(CompIndex), Instance)
      end else
      Begin
        Raise Exception.CreateFmt('Unknown name %d,%s,%s',[DispID,DispatchName,DelphiName]);
      End;
    end
    else
    begin
      // Ensure the method information has enough type information
      if Info.Len <= SizeOf(Info^) - SizeOf(TSymbolName) + 1 + Info.NameFld.UTF8Length then
      Begin
        Raise Exception.Create('Insufficent Type Info');
      End
      else
      begin
        AddDispatchInfo(DispID, DelphiName, dkMethod, Info, Instance);
      End;
    End;
  End;
end;

function TDDUObjectDispatch.Invoke(DispID: Integer; const IID: TGUID; LocaleID: Integer; Flags: Word; var Params; VarResult, ExcepInfo, ArgErr: Pointer): HRESULT;

Var
  Parms                   : PDispParams;
  TempRet                 : OleVariant;
  DispatchInfo            : TDispatchInfo;
  ReturnInfo              : PReturnInfo;
  DispIDName              : String;

begin
  Result := S_OK;
  try
    Parms := @Params;

    If Not GetDispIDName(DispID,DispIDName) Then
    Begin
      DispIDName := 'unknown';

      If fNames.ContainsKey(DispID) Then
      Begin
        DispIDName := fNames.Items[DispID];
      End;
    End;

    if GetDispatchInfo(DispID,DispatchInfo) Then
    begin
      DDU.DispatchEx.Debug('%s.Invoke %s(%d) [handled]'#13#10'%s',[Self.Name,DispIDName,DispID,dump_PDP(Parms)]);
      if VarResult = nil then
        VarResult := @TempRet;
      case DispatchInfo.Kind of
        dkProperty:
          begin
            // The high bit set means the DispID is a property not a method.
            // See GetIDsOfNames
            if Flags and (DISPATCH_PROPERTYPUTREF or DISPATCH_PROPERTYPUT) <> 0 then
            Begin
              if (Parms.cNamedArgs <> 1) or (PIntegerArray(Parms.rgDispIDNamedArgs)^[0] <> DispID_PROPERTYPUT) then
              Begin
                Result := DISP_E_MEMBERNOTFOUND;
              End
              else
              Begin
                SetPropValue(DispatchInfo.Instance, DispatchInfo.PropInfo, PVariantArray(Parms.rgvarg)^[0]);
              End;
            End
            else
            Begin
              if Parms.cArgs <> 0 then
              Begin
                Result := DISP_E_BADPARAMCOUNT;
              End Else if DispatchInfo.PropInfo^.PropType^.Kind = tkClass then
              Begin
                POleVariant(VarResult)^ := GetAsDispatch( Format('%s[property %s]',[Self.Name,DispatchInfo.DelphiName]),
                                                          TObject(GetOrdProp(DispatchInfo.Instance, DispatchInfo.PropInfo))
                                                        ) as IDispatch
              End else
              Begin
                POleVariant(VarResult)^ := GetPropValue(DispatchInfo.Instance,DispatchInfo.PropInfo, False);
              End;
            End;
          end;
        dkMethod:
          begin
            ReturnInfo := PReturnInfo(DispatchInfo.MethodInfo.NameFld.Tail);
            if (ReturnInfo.ReturnType <> nil) and (ReturnInfo.ReturnType^.Kind = tkClass) then
            Begin
              If IgnoreNamed Then
              Begin
                POleVariant(VarResult)^ := GetAsDispatch( Format('%s[method %s]',[Self.Name,DispatchInfo.DelphiName]),
                                                          TObject(NativeInt(ObjectInvoke(DispatchInfo.Instance,
                                                          DispatchInfo.MethodInfo,
                                                          [],
                                                          Slice(PVariantArray(Parms.rgvarg)^, Parms.cArgs))))
                                                        ) as IDispatch;
              End
              Else
              Begin
                POleVariant(VarResult)^ := GetAsDispatch( Format('%s[method %s]',[Self.Name,DispatchInfo.DelphiName]),
                                                          TObject(NativeInt(ObjectInvoke(DispatchInfo.Instance,
                                                          DispatchInfo.MethodInfo,
                                                          Slice(PIntegerArray(Parms.rgDispIDNamedArgs)^, Parms.cNamedArgs),
                                                          Slice(PVariantArray(Parms.rgvarg)^, Parms.cArgs))))
                                                        ) as IDispatch;
              End;
            End
            else
            Begin
              If IgnoreNamed Then
              Begin
                POleVariant(VarResult)^ := ObjectInvoke(DispatchInfo.Instance,
                                                        DispatchInfo.MethodInfo,
                                                        [],
                                                        Slice(PVariantArray(Parms.rgvarg)^, Parms.cArgs)
                                                       );
              End
              Else
              Begin
                POleVariant(VarResult)^ := ObjectInvoke(DispatchInfo.Instance,
                                                        DispatchInfo.MethodInfo,
                                                        Slice(PIntegerArray(Parms.rgDispIDNamedArgs)^, Parms.cNamedArgs),
                                                        Slice(PVariantArray(Parms.rgvarg)^, Parms.cArgs)
                                                       );
              End;
            End;
          end;
        dkSubComponent:
          POleVariant(VarResult)^ := GetAsDispatch( Format('%s[subcomponent %s]',[Self.Name,DispatchInfo.DelphiName]),
                                                    TComponent(DispatchInfo.Instance).Components[DispatchInfo.Index]
                                                  ) as IDispatch;
      end;
    end else
    Begin
      Result := DISP_E_MEMBERNOTFOUND;
      DDU.DispatchEx.Debug('%s.Invoke %s(%d) [unhandled]'#13#10'%s',[Self.Name,DispIDName,DispID,dump_PDP(Parms)]);
    End;
  except
    On E:Exception Do
    Begin
      DDU.DispatchEx.Debug('%s.Invoke %s(%d) Exception[%s] %s',[Self.Name,DispIDName,DispID,E.ClassName,E.Message]);
      if ExcepInfo <> nil then
      begin
        FillChar(ExcepInfo^, SizeOf(TExcepInfo), 0);
        with TExcepInfo(ExcepInfo^) do
        begin
          bstrSource := StringToOleStr(ClassName);
          if ExceptObject is Exception then
            bstrDescription := StringToOleStr(Exception(ExceptObject).Message);
          scode := E_FAIL;
        end;
      End;
      Result := DISP_E_EXCEPTION;
//Result := DISP_E_MEMBERNOTFOUND;
    end;
  end;
end;

procedure TDDUObjectDispatch.NameDispID(DispID : TDispID; Const DispatchName : String);
begin
  If Assigned(fNames) Then
  Begin
    fNames.AddOrSetValue(DispID,DispatchName);
  End;
end;

function TDDUObjectDispatchEx.IDispatchEx_InvokeEx(const DispID: TDispID; const lcid: LCID; const wflags: WORD; const Pdp: PDispParams; Const varResult: POleVariant; Const pei: PExcepInfo; const pspCaller: PServiceProvider): HResult; stdcall;

Var
  TempResult              : OleVariant;
  pResult                 : POleVariant;
  DispatchInfo            : TDispatchInfo;
  ReturnInfo              : PReturnInfo;
  DispIDName              : String;

begin
  Result := S_OK;

//  If Not GetDispIDName(DispID,DispIDName) Then
  Begin
    DispIDName := 'unknown';
  End;

  try
    if GetDispatchInfo(DispID,DispatchInfo) Then
    begin
      DDU.DispatchEx.Debug('%s.InvokeEx %s(%d)'#13#10'%s',[Self.Name,DispIDName,DispID,dump_PDP(pdp)]);
      If Assigned(varResult) Then
      Begin
        pResult := varResult
      End
      Else
      Begin
        pResult := @TempResult;
      End;

      case DispatchInfo.Kind of
        dkProperty:
          begin
            // The high bit set means the DispID is a property not a method.
            // See GetIDsOfNames
            if wFlags and (DISPATCH_PROPERTYPUTREF or DISPATCH_PROPERTYPUT) <> 0 then
            Begin
              if (Pdp.cNamedArgs <> 1) or (PIntegerArray(Pdp.rgDispIDNamedArgs)^[0] <> DispID_PROPERTYPUT) then
              Begin
                Result := DISP_E_MEMBERNOTFOUND;
              End
              else
              Begin
                SetPropValue(DispatchInfo.Instance, DispatchInfo.PropInfo, PVariantArray(Pdp.rgvarg)^[0]);
              End;
            End
            else
            Begin
              if Pdp.cArgs <> 0 then
              Begin
                Result := DISP_E_BADPARAMCOUNT;
              End Else if DispatchInfo.PropInfo^.PropType^.Kind = tkClass then
              Begin
                pResult^ := GetAsDispatch( Format('%s[property %s]',[Self.Name,DispatchInfo.DelphiName ]),
                                           TObject(GetOrdProp(DispatchInfo.Instance, DispatchInfo.PropInfo))
                                         ) as IDispatch;
              End else
              Begin
                pResult^ := GetPropValue(DispatchInfo.Instance,DispatchInfo.PropInfo, False);
              End;
            End;
          end;
        dkMethod:
          begin
            ReturnInfo := PReturnInfo(DispatchInfo.MethodInfo.NameFld.Tail);
            if (ReturnInfo.ReturnType <> nil) and (ReturnInfo.ReturnType^.Kind = tkClass) then
            Begin
              If IgnoreNamed Then
              Begin
                pResult^ := GetAsDispatch( Format('%s[method %s]',[Self.Name,DispatchInfo.DelphiName]),
                                           TObject(NativeInt(ObjectInvoke(DispatchInfo.Instance,
                                           DispatchInfo.MethodInfo,
                                           [],
                                           Slice(PVariantArray(Pdp.rgvarg)^, Pdp.cArgs))))
                                         ) as IDispatch;
              End
              Else
              Begin
                pResult^ := GetAsDispatch( Format('%s[method %s]',[Self.Name ,DispatchInfo.DelphiName]),
                                           TObject(NativeInt(ObjectInvoke(DispatchInfo.Instance,
                                           DispatchInfo.MethodInfo,
                                           Slice(PIntegerArray(Pdp.rgDispIDNamedArgs)^, Pdp.cNamedArgs),
                                           Slice(PVariantArray(Pdp.rgvarg)^, Pdp.cArgs))))
                                         ) as IDispatch;
              End;
            End
            else
            Begin
              If IgnoreNamed Then
              Begin
                pResult^ := ObjectInvoke(DispatchInfo.Instance,
                                         DispatchInfo.MethodInfo,
                                         [],
                                         Slice(PVariantArray(Pdp.rgvarg)^, Pdp.cArgs)
                                        );
              End
              Else
              Begin
                pResult^ := ObjectInvoke(DispatchInfo.Instance,
                                         DispatchInfo.MethodInfo,
                                         Slice(PIntegerArray(Pdp.rgDispIDNamedArgs)^, Pdp.cNamedArgs),
                                         Slice(PVariantArray(Pdp.rgvarg)^, Pdp.cArgs)
                                        );
              End;
            End;
          end;
        dkSubComponent:
          pResult^ := GetAsDispatch( Format('%s[subcomponent %s]',[Self.Name,DispatchInfo.DelphiName]),
                                     TComponent(DispatchInfo.Instance).Components[DispatchInfo.Index]
                                   ) as IDispatch;
      end;
    end else
    Begin
      Result := DISP_E_MEMBERNOTFOUND;
      DDU.DispatchEx.Debug('%s.InvokeEx! %s(%d)'#13#10'%s',[Self.Name,DispIDName,DispID,dump_PDP(pdp)]);
    End;
  except
    On E:Exception Do
    Begin
      DDU.DispatchEx.Debug('%s.InvokeEx %s(%d) Exception[%s] %s',[Self.Name,DispIDName,DispID,E.ClassName,E.Message]);
      if Assigned(pei) Then
      begin
        FillChar(pei^, SizeOf(TExcepInfo), 0);
        pei^.bstrSource := StringToOleStr(ClassName);
        if ExceptObject is Exception then
          PEI^.bstrDescription := StringToOleStr(Exception(ExceptObject).Message);
        PEI^.scode := E_FAIL;
      end;
      Result := DISP_E_EXCEPTION;
    End;
  end;
End;

function TDDUObjectDispatchEx.jsHasOwnProperty(const Name: String): Boolean;

begin
  Result := (GetDispIDByDispatchName(Name)<>DispID_UNKNOWN);
end;

function TDDUObjectDispatch.QueryInterface(const IID: TGUID; out Obj): HRESULT;


begin
  Result := Inherited QueryInterface(IID,Obj);
  if (Result=E_NOINTERFACE) And IsEqualIID(IID, fCustomIID) then
  begin
    GetInterface(IDispatch, Obj);
    Result := S_OK;
    Exit;
  end;
end;

function TDDUObjectDispatch._AddRef: Integer;
begin
  Result := Inherited _AddRef;
end;

function TDDUObjectDispatch._Release: Integer;

begin
  Result := Inherited _Release;
end;

{$ENDREGION}

{$REGION 'TParameters/IParameters'}

{ TParameters }

procedure   TParameters.BuildPositionalDispatchIDs;

Var
  Loop                       : integer;
  Index                      : Integer;
  lLow                        : Integer;
  lHigh                       : Integer;

begin
  If Assigned(fDispatchParameters) Then
  Begin
    lLow  := 0;
    lHigh := fDispatchParameters.cArgs - 1;

    For Loop := lLow To lHigh Do
    Begin
      fDispatchParameterIDs[Loop] := fDispatchParameters.cArgs-1-Loop;
    End;

    if (fDispatchParameters.cNamedArgs>0) Then
    Begin
      for Loop := 0 to fDispatchParameters.cNamedArgs - 1 do  // Hmmm... Review this logic, it seems.... wrong.
      Begin
        Index := fDispatchParameters.rgdispidNamedArgs[Loop];

        If (Index>=lLow) And (Index<=lHigh) Then
        Begin
          fDispatchParameterIDs[Index] := Loop;
        End;

      End;
    End;
  End;
end;
constructor TParameters.Create(const aDPS: PDispParams);
begin
  Inherited Create;
  fDispatchParameters := aDPS;
  fNull               := Null;

  if Assigned(fDispatchParameters) And (fDispatchParameters.cArgs > 0) Then
  begin
    fDispatchParameterIDsSize := fDispatchParameters.cArgs * SizeOf(TDispID);
    GetMem(fDispatchParameterIDs, fDispatchParameterIDsSize);

    BuildPositionalDispatchIDs;
  end
  Else
  Begin
    fDispatchParameterIDs     := Nil;
    fDispatchParameterIDsSize := 0;
  End;
end;
destructor  TParameters.Destroy;
begin
  If Assigned(fDispatchParameterIDs) Then
  Begin
    FreeMem(fDispatchParameterIDs, fDispatchParameterIDsSize);
  End;
  inherited;
end;
function    TParameters.GetParameter(aIndex: Integer): TVariantArg;
begin
  If Assigned(fDispatchParameters) Then
  Begin
    Result := fDispatchParameters.rgvarg[fDispatchParameterIDs^[aIndex]];
  End
  Else
  Begin
    Result := TVariantArg(fNull);
  End;
end;

{$ENDREGION}

{$REGION 'TDispatchHandler/IDispatch.}

{ TDispatch }

function TDispatchHandler.IDispatch_GetIDsOfNames(const IID: TGUID; Names: Pointer; NameCount, LocaleID: Integer; DispIDs: Pointer): HResult;
begin
  Result := GetIDsOfNames(IID,Names,NameCount,LocaleID,DispIDs);
end;
function TDispatchHandler.IDispatch_GetTypeInfo(Index, LocaleID: Integer; out TypeInfo): HResult;
begin
  Result := GetTypeInfo(Index,LocaleID,TypeInfo);
end;
function TDispatchHandler.IDispatch_GetTypeInfoCount(out Count: Integer): HResult;
begin
  Result := GetTypeInfoCount(Count);
end;
function TDispatchHandler.IDispatch_Invoke(DispID: Integer; const IID: TGUID; LocaleID: Integer; Flags: Word; var Params; VarResult, ExcepInfo, ArgErr: Pointer): HResult;

Var
  Parameters              : IParameters;

begin
  If @Params<>Nil Then
  Begin
    Parameters := TParameters.Create(@TDispParams(Params));
  End
  Else
  Begin
    Parameters := TParameters.Create(Nil);
  end;

  Result := Invoke(DispID,IID,LocaleID,Flags,Parameters,VarResult,ExcepInfo,ArgErr);
end;

constructor TDispatchHandler.Create;
begin
  Inherited Create;
  DDU.DispatchEx.Debug('CREATE: %s',[Self.ClassName]);
  Init;
end;

constructor TDispatchHandler.Create(const aName: String; aCustomIID: TGUID);
begin
  Inherited Create;
  Init;
  DDU.DispatchEx.Debug('CREATE: %s//%s',[Self.ClassName,aName]);
  fName := aName;
  fCustomIID  := aCustomIID;
end;

destructor TDispatchHandler.Destroy;
begin
  DDU.DispatchEx.Debug('DESTROY: %s//%s',[Self.ClassName,fName]);
  FreeAndNil(fNameDispIDs);
  inherited;
end;

function TDispatchHandler.GetIDsOfNames(const IID: TGUID; Names: Pointer; NameCount, LocaleID: Integer; DispIDs: Pointer): HResult;

type
  PNames = ^TNames;
  TNames = array[0..100] of POleStr;
  PDispIDs = ^TDispIDs;
  TDispIDs = array[0..100] of Cardinal;

Var
  Loop                    : Integer;
  aName                   : String;

begin
  If Assigned(fGetIDsOfNames) Then
  Begin
    Result := fGetIDsOfNames(IID,Names,NameCount,LocaleID,DispIDs);
  End
  Else
  Begin
    For Loop := 0 To NameCount-1 Do
    Begin
      aName := PChar(PNames(Names)[Loop]);
//      OutputDebugString(PChar(Format('************* GetIdsOfName(%d) %s',[aName])));
    End;

    Result := E_NOTIMPL;
  End;
end;
function TDispatchHandler.GetTypeInfo(Index, LocaleID: Integer; out TypeInfo): HResult;
begin
  If Assigned(fGetTypeInfo) Then
  Begin
    Result := fGetTypeInfo(Index,LocaleID,TypeInfo);
  End
  Else
  Begin
    Result := E_NOTIMPL;
    Pointer(TypeInfo) := nil;
  End;
end;
function TDispatchHandler.GetTypeInfoCount(out Count: Integer): HResult;
begin
  If Assigned(fGetTypeInfoCount) Then
  Begin
    Result := GetTypeInfoCount(Count);
  End
  Else
  Begin
    Result := E_NOTIMPL;
    Count  := 0;
  End;
end;
procedure TDispatchHandler.Init;
begin
  fNameDispIDs   := TDictionary<String,Integer>.Create;
  fCustomIID     := IDispatch;
end;

function TDispatchHandler.Invoke(DispID: Integer; const IID: TGUID; LocaleID: Integer; Flags: Word; Parameters : IParameters; VarResult, ExcepInfo, ArgErr: Pointer): HResult;
begin
  If Assigned(fInvoke) then
  Begin
    Result := fInvoke(DispID,IID,LocaleID,Flags,Parameters,VarResult,ExcepInfo,ArgErr);
  End
  Else
  Begin
    Result := DISP_E_MEMBERNOTFOUND;
  End;
end;

procedure TDispatchHandler.MapNameDispID(const aName: String; aDispID: Integer);
begin
  fNameDispIDs.AddOrSetValue(aName.ToLower,aDispID);
end;

function TDispatchHandler.QueryInterface(const IID: TGUID; out Obj): HRESULT;
begin
  Result := Inherited QueryInterface(IID,Obj);
  if (Result=E_NOINTERFACE) And IsEqualIID(IID, fCustomIID) then
  begin
    GetInterface(IDispatch, Obj);
    Result := S_OK;
    Exit;
  end;
end;

function TDispatchHandler._AddRef: Integer;
begin
  Result := Inherited _AddRef;
end;

function TDispatchHandler._Release: Integer;
begin
  Result := Inherited _Release;
end;

{$ENDREGION}

{$REGION 'TDispatchExHandler/IDispatchEx'}

function TDispatchExHandler.IDispatchEx_DeleteMemberByDispID(const id: TDispID): HResult;
begin
  Result := DeleteMemberByDispID(ID);
end;
function TDispatchExHandler.IDispatchEx_DeleteMemberByName(const bstr: TBSTR; const grfdex: DWORD): HResult;
begin
  Result := DeleteMemberByName(bstr,grfdex);
end;
function TDispatchExHandler.IDispatchEx_GetDispID(const bstrName: TBSTR; const grfdex: DWORD; out id: TDispID): HResult;
begin
  Result := GetDispID(bstrName,grfdex,ID);
end;
function TDispatchExHandler.IDispatchEx_GetMemberName(const id: TDispID; out bstrName: TBSTR): HResult;

Var
  aName                   : String;

begin
  Result := GetMemberName(ID,aName);
  If aName='' Then
  Begin
    bstrName := Nil;
  End
  Else
  Begin
    bstrName := SysAllocString(PChar(aName));
  End;
end;
function TDispatchExHandler.IDispatchEx_GetMemberProperties(const id: TDispID; const grfdexFetch: DWORD; out grfdex: DWORD): HResult;
begin
  Result := GetMemberProperties(ID,grfdexFetch,grfdex);
end;
function TDispatchExHandler.IDispatchEx_GetNameSpaceParent(out unk: IInterface): HResult;
begin
  Result := GetNameSpaceParent(unk);
end;
function TDispatchExHandler.IDispatchEx_GetNextDispID(const grfdex: DWORD; const id: TDispID; out nid: TDispID): HResult;
begin
  Result := GetNextDispID(grfdex,id,nid);
end;
function TDispatchExHandler.IDispatchEx_InvokeEx(const id: TDispID; const lcid: LCID; const wflags: WORD; const pdp: PDispParams; const varRes: POleVariant; const pei: PExcepInfo; const pspCaller: PServiceProvider): HResult;

Var
  Parameters              : IParameters;

begin
  If Assigned(PDP) Then
  Begin
    Parameters := TParameters.Create( pdp );
  End;
  Result := InvokeEx(ID,lcid,wflags,Parameters,varRes,pei,pspCaller);
end;

function TDispatchExHandler.DeleteMemberByDispID(const id: TDispID): HResult;
begin
  If Assigned(fOnDeleteMemberByDispID) Then
  Begin
    Result := fOnDeleteMemberByDispID(ID);
  End
  Else
  Begin
    Result := E_NOTIMPL;
  End;
end;
function TDispatchExHandler.DeleteMemberByName(const Name: String; const grfdex: DWORD): HResult;
begin
  If Assigned(fOnDeleteMemberByName) Then
  Begin
    Result := fOnDeleteMemberByName(Name,grfdex);
  End
  Else
  Begin
    Result := E_NOTIMPL;
  End;
end;
function TDispatchExHandler.GetDispID(const Name: String; const grfdex: DWORD; out id: TDispID): HResult;
begin
  If Assigned(fOnGetDispID) Then
  Begin
    Result := fOnGetDispID(Name,grfdex,ID);
  End
  Else
  Begin
    Result := E_NOTIMPL;
  End;
end;
function TDispatchExHandler.GetMemberName(const id: TDispID; out Name: String): HResult;
begin
  If Assigned(fOnGetMemberName) Then
  Begin
    Result := fOnGetMemberName(id,Name);
  End
  Else
  Begin
    Result   := E_NOTIMPL;
  End;
  Name := '';
end;
function TDispatchExHandler.GetMemberProperties(const id: TDispID; const grfdexFetch: DWORD; out grfdex: DWORD): HResult;
begin
  If Assigned(fOnGetMemberProperties) Then
  Begin
    Result := fOnGetMemberProperties(ID,grfdexFetch,grfDex);
  End
  Else
  Begin
    Result := E_NOTIMPL;
  End;
end;
function TDispatchExHandler.GetNameSpaceParent(out unk: IInterface): HResult;
begin
  If Assigned(fOnGetNameSpaceParent) Then
  Begin
    Result := fOnGetNameSpaceParent(unk);
  End
  Else
  Begin
    Result := E_NOTIMPL;
  End;
end;
function TDispatchExHandler.GetNextDispID(const grfdex: DWORD; const id: TDispID; out nid: TDispID): HResult;
begin
  If Assigned(fOnGetNextDispID) then
  Begin
    Result := fOnGetNextDispID(grfdex,ID,nid);
  End
  Else
  Begin
    Result := E_NOTIMPL;
  End;
end;
function TDispatchExHandler.InvokeEx(const id: TDispID; const lcid: LCID; const wflags: WORD; const Parameters : IParameters; Const varRes: POleVariant; Const pei: PExcepInfo; const pspCaller: PServiceProvider): HResult;
begin
  If Assigned(fOnInvokeEx) Then
  Begin
    Result := fOnInvokeEx(ID,lcid,wFlags,Parameters,varRes,pei,pspCaller);
  End
  Else
  Begin
    Result := DISP_E_MEMBERNOTFOUND;
  End;
end;

{$ENDREGION}


end.
