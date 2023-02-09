library COM_object;

{$IFDEF FPC}
  {$MODE Delphi}
{$ENDIF}

{$IFNDEF FPC }
{$IFDEF RELEASE}
{$WEAKLINKRTTI ON}
{$RTTI EXPLICIT METHODS([]) PROPERTIES([]) FIELDS([])}
{$ENDIF}
{$ENDIF}


//Very simple COM library


uses
  windows;

var // some static variables
// A count of how many objects our DLL has created (by some
// app calling our IClassFactory object's CreateInstance())
// which have not yet been Release()'d by the app
   OutstandingObjects :longint;

// A count of how many apps have locked our DLL via calling our
// IClassFactory object's LockServer()
   LockCount :longint;


//GUID toools ------------------------------------------------------------------
function _CompareTGUID(a,b:pbyte; sz:longword):boolean;  // simple tool
var i,c:longword;
begin
   Result := false;
   c := 0;
   if sz > 0 then for i := 0 to sz-1 do if a[i] = b[i] then inc(c);
   if sz = c then Result := true;
end;

const hex : array [0..15] of ansichar = ('0','1','2','3','4','5','6','7','8','9','A','B','C','D','E','F');
function _StrTGUID(a:pbyte):ansistring;
var i :longword;
begin
   Result := '{';
   for i := 3 downto 0 do Result := Result + hex[ (a[i + 0] shr 4) and $F] +  hex[ a[i + 0]  and $F];
   Result := Result + '-';
   for i := 1 downto 0 do Result := Result + hex[ (a[i + 4] shr 4) and $F] +  hex[ a[i + 4]  and $F];
   Result := Result + '-';
   for i := 1 downto 0 do Result := Result + hex[ (a[i + 6] shr 4) and $F] +  hex[ a[i + 6]  and $F];
   Result := Result + '-';
   for i := 0 to 1 do Result := Result + hex[ (a[i + 8] shr 4) and $F] +  hex[ a[i + 8]  and $F];
   Result := Result + '-';
   for i := 0 to 5 do Result := Result + hex[ (a[i + 10] shr 4) and $F] +  hex[ a[i + 10]  and $F];
   Result := Result + '}'
end;

//--- GUIDs -------------------------------------------------------------------- // use ctrl-shift-G
const
   IMY_FACTORY    : TGUID = '{99BB4CB8-EFF9-4017-B4A8-B8B3B6C8BA63}';  // main key this used to be registered in registry CLSID
   IMY_OBJECT     : TGUID = '{FDB0CC23-EA12-4B8B-9455-DE5C9011B596}';
   IUnknown      : TGUID = '{00000000-0000-0000-C000-000000000046}'; //windows const
   IClassFactory : TGUID = '{00000001-0000-0000-C000-000000000046}';


//--- My Object definition  class and  methods ---------------------------------

type
   MY_OBJECT_INTERFACE = record
      // mandatory class members
      QueryInterface:function(this :pointer; const iid: TGUID; var obj:pointer): HResult;  stdcall;
      AddRef:function(this :pointer): Longint; stdcall;
      Release:function(this :pointer): Longint;  stdcall;
      //user provided members
      ping:function(this :pointer; a:longword): Longint;  stdcall;
      //...
      //here you can add your methods
   end;
   PMY_OBJECT_INTERFACE = ^MY_OBJECT_INTERFACE;

   MY_OBJECT = record // local data segment
      Vtbl :PMY_OBJECT_INTERFACE;  // list with methods
      count :longword;  // count the instances
      aa:longword;  // local my data
   end;
   PMY_OBJECT = ^MY_OBJECT;


// IMPORTANT  you have to put 'this' like firs parameter in real live compiler hide this
// this is pointer to MY_OBJECT to local data segment !!!!

function MY_OBJECT_INTERFACE_QueryInterface(this :pointer; const iid: TGUID; var obj:pointer): HResult;  stdcall;
begin
   obj := nil;
   result := E_NOINTERFACE;
   if _CompareTGUID(@iid, @IMY_OBJECT, SizeOf(TGUID)) or _CompareTGUID(@iid, @IUnknown, SizeOf(TGUID)) then
   begin
      PMY_OBJECT(this).Vtbl.AddRef(this);
      obj := this;
      Result := NOERROR;
   end;
end;

function MY_OBJECT_INTERFACE_AddRef(this :pointer): Longint; stdcall;
begin
   inc(PMY_OBJECT(this).count);
   Result := PMY_OBJECT(this).count;
end;

function MY_OBJECT_INTERFACE_Release(this :pointer): Longint;  stdcall;
begin
   dec(PMY_OBJECT(this).count);
   Result := PMY_OBJECT(this).count;
   if result = 0 then
   begin
      GlobalFree(NativeUInt(PMY_OBJECT(this).Vtbl));
      GlobalFree(NativeUInt(this));
   end;
end;

function MY_OBJECT_INTERFACE_PING(this :pointer; a:longword): Longint; stdcall;
begin
   Result := a + PMY_OBJECT(this).aa;
end;





//------ Factory ---------------------------------------------------------------
// you need to have a factory it is mandatory
// using factory.CreateInstance you will receive the my object !!!

type
   MY_FACTORY_INTERFACE = record
      // mandatory class members for
      QueryInterface:function(this :pointer; const iid: TGUID; var obj:pointer): HResult;  stdcall;
      AddRef:function(this :pointer): Longint; stdcall;
      Release:function(this :pointer): Longint;  stdcall;
      // mandatory for factory
      CreateInstance:function(this :pointer; punkOuter :pointer; const iid: TGUID; obj :pointer): HResult;  stdcall;
      LockServer:function(this :pointer; fLock: BOOL): HResult; stdcall;
      //you can put you slass members or properties down here
   end;
   PMY_FACTORY_INTERFACE = ^MY_FACTORY_INTERFACE;

   MY_FACTORY = record
      Vtbl:PMY_FACTORY_INTERFACE; //Methodes statis
      //data
      ObjCount:longword;
      aa,bb:longword; // just for example

   end;
   PMY_FACTORY = ^MY_FACTORY;



function MY_FACTORY_QueryInterface(this :pointer; const iid: TGUID; var obj:pointer): HResult;  stdcall;
begin
   obj := nil;
   result := E_NOINTERFACE;
   if _CompareTGUID(@iid, @IClassFactory, SizeOf(TGUID)) or _CompareTGUID(@iid, @IUnknown, SizeOf(TGUID)) then
   begin
      PMY_FACTORY(this).Vtbl.AddRef(this);
      obj := this;
      Result := S_OK;
   end;
end;

function MY_FACTORY_AddRef(this :pointer): Longint; stdcall;
begin
   Result := InterlockedIncrement(OutstandingObjects); //winapi secure increment
end;

function MY_FACTORY_Release(this :pointer): Longint;  stdcall;
begin
   Result := InterlockedDecrement(OutstandingObjects); //winapi secure increment
   if Result = 0 then
   begin
      GlobalFree(NativeUInt(PMY_FACTORY(this).Vtbl));
      GlobalFree(NativeUInt(this));
   end;
end;

function MY_FACTORY_CreateInstance(this :pointer; punkOuter :pointer; const iid: TGUID; var obj :pointer): HResult;  stdcall;
var pf:PMY_OBJECT_INTERFACE;
    p:PMY_OBJECT;
begin
   obj := nil;
   Result := CLASS_E_CLASSNOTAVAILABLE;
   if punkOuter <> nil then begin Result := CLASS_E_NOAGGREGATION; Exit; end;
   if _CompareTGUID(@iid, @IMY_OBJECT, SizeOf(TGUID)) then
   begin
      // create my object interface
      pf := pointer(GlobalAlloc(GMEM_FIXED,sizeof(MY_OBJECT_INTERFACE)));

      pf.QueryInterface := @MY_OBJECT_INTERFACE_QueryInterface;
      pf.AddRef         := @MY_OBJECT_INTERFACE_AddRef;
      pf.Release        := @MY_OBJECT_INTERFACE_Release;
      pf.Ping           := @MY_OBJECT_INTERFACE_Ping;

      p := pointer(GlobalAlloc(GMEM_FIXED,sizeof(MY_OBJECT)));
      p.count := 0;
      p.aa:= 10;
      p.Vtbl := pf;

      Result := pf.QueryInterface(p,iid,obj);

   end;
end;

function MY_FACTORY_LockServer(this :pointer; fLock: BOOL): HResult; stdcall;
begin
   if flock then InterlockedIncrement(LockCount)	else InterlockedDecrement(LockCount);
   Result := NOERROR;
end;





/// OLE2 functions /////////////////////////////////////////////////////////////

//------------------------------------------------------------------------------
// DllGetClassObject()
// This is called by the OLE functions CoGetClassObject() or
// CoCreateInstance() in order to get our DLL's IClassFactory object
// (and return it to someone who wants to use it to get ahold of one
// of our IExample objects). Our IClassFactory's CreateInstance() can
// be used to allocate/retrieve our IExample object.
//
// NOTE: After we return the pointer to our IClassFactory, the caller
// will typically call its CreateInstance() function.
// CoCreateInstance do:
// 1.CoGetClassObject(objGuid, context, nil, IID_IClassFactory, &pCF);
// 2.hResult = pCF->CreateInstance(pUnkOuter, riid, ppvOBJ);
// 3.pCF->Release()
//
// OLE -> Factory -> myObject


function DllGetClassObject(const objGuid, factoryGuid: TGUID; var factoryHandle :pointer{pointer to pinter}): HResult; stdcall; export;
var pf :PMY_FACTORY_INTERFACE;
    p:PMy_FACTORY;
begin
   Result := CLASS_E_CLASSNOTAVAILABLE;
   factoryhandle := nil;
   if _CompareTGUID(@objGuid, @IMY_FACTORY, SizeOf(TGUID)) then  // Check that the caller is passing our That's the only object our DLL implements
   begin
      pf := pointer(GlobalAlloc(GMEM_FIXED,sizeof(MY_FACTORY_INTERFACE)));

      pf.QueryInterface := @MY_FACTORY_QueryInterface;
      pf.AddRef         := @MY_FACTORY_AddRef;
      pf.Release        := @MY_FACTORY_Release;
      pf.CreateInstance := @MY_FACTORY_CreateInstance;
      pf.LockServer     := @MY_FACTORY_LockServer;

      p := pointer(GlobalAlloc(GMEM_FIXED,sizeof(MY_FACTORY)));
      p.aa:= 11;
      p.bb:= 22;
      p.Vtbl := pf;
      Result :=  pf.QueryInterface(p,factoryGuid,factoryhandle);   //return ClassFactory
   end;
end;


//------------------------------------------------------------------------------
// DllCanUnloadNow()
// This is called by some OLE function in order to determine
// whether it is safe to unload our DLL from memory.
//
// RETURNS: S_OK if safe to unload, or S_FALSE if not.
function DllCanUnloadNow: HResult; stdcall; export;
begin
   Result := S_OK;
   if (OutstandingObjects or LockCount) <> 0 then Result := S_FALSE;
end;



function DllRegisterServer: HResult; stdcall; export;
var RootKey,Key,Key2,Extra:HKEY;
    sid:ansistring;
    disposition :longword;
    Buf:array[0..254]of char;
begin
   FillChar(Buf,Sizeof(Buf),#0);
   Result := S_FALSE;
   if RegOpenKeyEx(HKEY_LOCAL_MACHINE, 'Software\Classes', 0, KEY_WRITE, rootKey) = 0 then
   begin
      if RegOpenKeyEx(rootKey, 'CLSID', 0, KEY_ALL_ACCESS, Key) = 0 then
      begin
         sid := _StrTGUID(@IMY_FACTORY);
         if RegCreateKeyExA( Key, @sid[1], 0, nil, REG_OPTION_NON_VOLATILE, KEY_WRITE, nil, Key2, @disposition) = 0 then
         begin
            sid := 'Sdex32 interface';
            RegSetValueExA(Key2, nil, 0, REG_SZ, @sid[1], length(sid));
            // Create an "InprocServer32" key whose default value is the path of this DLL
            if RegCreateKeyEx(Key2, 'InProcServer32', 0, nil, REG_OPTION_NON_VOLATILE, KEY_WRITE, nil, Extra, @disposition) = 0 then
            begin
               GetModuleFileName(hInstance,Buf,255);
               sid := ansistring(Buf);
               if RegSetValueExA(Extra, nil, 0, REG_SZ, @sid[1], length(sid)) = 0 then
               begin
                  sid := 'Both';
							    if RegSetValueExA(Extra, 'ThreadingModel', 0, REG_SZ, @sid[1], length(sid))= 0 then
                  begin
                     Result := S_OK;
                  end;
               end;
               RegCloseKey(Extra);
            end;
            RegCloseKey(Key2);
         end;
         RegCloseKey(Key);
      end;
      RegCloseKey(rootKey);
   end;
end;


function DllUnregisterServer: HResult; stdcall;  export;
var RootKey,Key,Key2:HKEY;
    sid:ansistring;
begin
   Result := S_FALSE;
 	 if RegOpenKeyEx(HKEY_LOCAL_MACHINE, 'Software\Classes', 0, KEY_WRITE, rootKey) = 0 then
   begin
			// Delete our CLSID key and everything under it
			if RegOpenKeyEx(rootKey, 'CLSID', 0, KEY_ALL_ACCESS, Key) = 0 then
      begin
         sid := _StrTGUID(@IMY_FACTORY);
         if RegOpenKeyExA(Key, @sid[1], 0, KEY_ALL_ACCESS, Key2) = 0 then
         begin
        		RegDeleteKey(Key2, 'InprocServer32');
  					RegDeleteKey(Key2, 'ThreadingModel');
            RegCloseKey(Key2);
            Result := S_OK;
         end;
 				 RegDeleteKeyA(Key, @sid[1]);
         RegCloseKey(Key);
      end;
      RegCloseKey(rootKey);
   end;
end;

(*
function DllInstall(bInstall: WordBool; pszCmdLine: LPCWSTR): HResult; stdcall;
begin

end;
function LoadTypeLibrary(const ModuleName: string): ITypeLib;
begin

end;
*)


//procedure DllMain(reason:integer);  //i did not need this
//begin
//   case reason of
//      DLL_PROCESS_DETACH : begin end;
//      DLL_PROCESS_ATTACH : begin
//      end;
//      DLL_THREAD_ATTACH : begin end;
//      DLL_THREAD_DETACH : begin end;
//   end;
//end;

exports
   DllRegisterServer,
   DllUnregisterServer,
   DllGetClassObject,
   DllCanUnloadNow;


begin
   DisableThreadLibraryCalls(hinstance);
   OutstandingObjects := 0; //this is done on  DLL_PROCESS_ATTACH
   LockCount := 0;

//   DllProc := DllMain;
//   DllMain(DLL_PROCESS_ATTACH); // begin is executed after OS call dllmain
end.
