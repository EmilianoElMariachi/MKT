
#ifndef _NO_CRT_STDIO_INLINE
#define _NO_CRT_STDIO_INLINE 1
#endif

#ifndef NOMINMAX
#define NOMINMAX
#endif

#include <ntstatus.h>
#define WIN32_NO_STATUS
#include <windows.h>
#ifdef __MINGW64_VERSION_MAJOR
#include "winternl.h"
#else
#include <winternl.h>
#endif
#include <winnt.h>
#include <rpc.h>

#define LOCALHOST_IP L"127.0.0.1"
#define PROTO_SEQ_TCP L"ncacn_ip_tcp"

#ifdef __cplusplus
extern "C" {
#endif

#ifdef _MSC_VER
#define DLL_PROCESS_VERIFIER 4

using RTL_VERIFIER_DLL_LOAD_CALLBACK = VOID(NTAPI *) (PWSTR DllName, PVOID DllBase, SIZE_T DllSize, PVOID Reserved);
using RTL_VERIFIER_DLL_UNLOAD_CALLBACK = VOID(NTAPI *) (PWSTR DllName, PVOID DllBase, SIZE_T DllSize, PVOID Reserved);
using RTL_VERIFIER_NTDLLHEAPFREE_CALLBACK = VOID(NTAPI *) (PVOID AllocationBase, SIZE_T AllocationSize);

typedef struct _RTL_VERIFIER_THUNK_DESCRIPTOR {
  PCSTR ThunkName;
  PVOID ThunkOldAddress;
  PVOID ThunkNewAddress;
} RTL_VERIFIER_THUNK_DESCRIPTOR, *PRTL_VERIFIER_THUNK_DESCRIPTOR;

typedef struct _RTL_VERIFIER_DLL_DESCRIPTOR {
  PCWSTR  DllName;
  DWORD   DllFlags;
  PVOID   DllAddress;
  PRTL_VERIFIER_THUNK_DESCRIPTOR  DllThunks;
} RTL_VERIFIER_DLL_DESCRIPTOR, *PRTL_VERIFIER_DLL_DESCRIPTOR;

typedef struct _RTL_VERIFIER_PROVIDER_DESCRIPTOR {
  DWORD   Length;
  PRTL_VERIFIER_DLL_DESCRIPTOR      ProviderDlls;
  RTL_VERIFIER_DLL_LOAD_CALLBACK    ProviderDllLoadCallback;
  RTL_VERIFIER_DLL_UNLOAD_CALLBACK  ProviderDllUnloadCallback;
  PCWSTR  VerifierImage;
  DWORD   VerifierFlags;
  DWORD   VerifierDebug;
  PVOID   RtlpGetStackTraceAddress;
  PVOID   RtlpDebugPageHeapCreate;
  PVOID   RtlpDebugPageHeapDestroy;
  RTL_VERIFIER_NTDLLHEAPFREE_CALLBACK ProviderNtdllHeapFreeCallback;
} RTL_VERIFIER_PROVIDER_DESCRIPTOR, *PRTL_VERIFIER_PROVIDER_DESCRIPTOR;
#endif // _MSC_VER

NTSYSAPI
NTSTATUS
NTAPI
LdrDisableThreadCalloutsForDll(
  _In_  PVOID DllImageBase
);

NTSYSAPI
NTSTATUS
NTAPI
NtProtectVirtualMemory(
  _In_    HANDLE  ProcessHandle,
  _Inout_ PVOID   *BaseAddress,
  _Inout_ PSIZE_T RegionSize,
  _In_    ULONG   NewProtect,
  _Out_   PULONG  OldProtect
);

NTSYSAPI
ULONG
DbgPrint(
  PCSTR Format,
  ...
);

NTSYSAPI
ULONG
DbgPrintEx(
  ULONG ComponentId,
  ULONG Level,
  PCSTR Format,
  ...
);

NTSYSAPI
NTSTATUS
NTAPI
LdrGetProcedureAddress(
  _In_  PVOID BaseAddress,
  _In_  PANSI_STRING Name,
  _In_  ULONG Ordinal,
  _Out_ PVOID *ProcedureAddress
);

NTSYSAPI
NTSTATUS
NTAPI
LdrGetDllHandle(
  _In_opt_  PWSTR DllPath,
  _In_opt_  PULONG DllCharacteristics,
  _In_      PUNICODE_STRING DllName,
  _Out_     PVOID *DllHandle
);

NTSYSAPI
NTSTATUS
NTAPI
LdrLoadDll(
  _In_opt_  PWSTR SearchPath,
  _In_opt_  PULONG LoadFlags,
  _In_      PUNICODE_STRING Name,
  _Out_opt_ PVOID *BaseAddress
);

#define RTL_CONSTANT_STRING(s) { \
  sizeof(s) - sizeof((s)[0]), \
  sizeof(s), \
  (std::add_pointer_t<std::remove_const_t<std::remove_pointer_t<std::decay_t<decltype(s)>>>>)(s) \
}

#define NtCurrentProcess() (HANDLE(LONG_PTR(-1)))

#ifdef __cplusplus
}
#endif

#include <cstring>
#include <type_traits>
#include <algorithm>
#include <random>

RPC_STATUS RPC_ENTRY RpcStringBindingComposeW_Hook(PWSTR ObjUuid, PWSTR ProtSeq, PWSTR NetworkAddr, PWSTR EndPoint, PWSTR Options, PWSTR* StringBinding);

struct hook_entry
{
  ANSI_STRING     function_name;
  PVOID           old_address;
  PVOID           new_address;
};

struct hook_target_image
{
  UNICODE_STRING  name;
  ULONG           hook_bitmap;
  PVOID           base = nullptr;
};

#define DEFINE_HOOK(name) {RTL_CONSTANT_STRING(#name), nullptr, (PVOID)&name ## _Hook}

static hook_entry s_hooks[] =
{
  DEFINE_HOOK(RpcStringBindingComposeW),
};

static hook_target_image s_target_images[] =
{
  { RTL_CONSTANT_STRING(L"SppExtComObj"), 0b011 },
};

void* get_original_from_hook_address(void* hook_address)
{
  const hook_entry temp_entry{ {}, nullptr, hook_address };

  const auto it = std::find_if(std::begin(s_hooks), std::end(s_hooks), [hook_address](const hook_entry& e)
  {
    return e.new_address == hook_address;
  });

  return it == std::end(s_hooks) ? nullptr : it->old_address;
}

template <typename T>
T* get_original_from_hook_address_wrapper(T* fn)
{
  return (T*)get_original_from_hook_address((void*)fn);
}

#define GET_ORIGINAL_FUNC(name) (*get_original_from_hook_address_wrapper(&name ## _Hook))

static VOID NTAPI DllLoadCallback(PWSTR DllName, PVOID DllBase, SIZE_T DllSize, PVOID Reserved);

static RTL_VERIFIER_DLL_DESCRIPTOR s_dll_descriptors[] = { {} };

static RTL_VERIFIER_PROVIDER_DESCRIPTOR s_provider_descriptor =
{
  sizeof(RTL_VERIFIER_PROVIDER_DESCRIPTOR),
  s_dll_descriptors,
  &DllLoadCallback
};

BOOL WINAPI DllMain(
  PVOID dll_handle,
  DWORD reason,
  PRTL_VERIFIER_PROVIDER_DESCRIPTOR* provider
)
{
  switch (reason)
  {
  case DLL_PROCESS_ATTACH:
    LdrDisableThreadCalloutsForDll(dll_handle);
    break;
  case DLL_PROCESS_VERIFIER:
    *provider = &s_provider_descriptor;
    break;
  default:
    break;
  }
  return TRUE;
}

static void hook_thunks(PVOID base, PIMAGE_THUNK_DATA thunk, PIMAGE_THUNK_DATA original_thunk)
{
  while (original_thunk->u1.AddressOfData)
  {
    if (!(original_thunk->u1.Ordinal & IMAGE_ORDINAL_FLAG))
    {
      ULONG bitmap = 0u;
      for (const auto& image : s_target_images)
        if (image.base == base)
          bitmap = image.hook_bitmap;

      for (auto i = 0u; i < std::size(s_hooks); ++i)
      {
        if (!(bitmap & (1 << i)))
          continue;

        auto& hook = s_hooks[i];

        const auto by_name = PIMAGE_IMPORT_BY_NAME((char*)base + original_thunk->u1.AddressOfData);
        if ((hook.old_address && hook.old_address == PVOID(thunk->u1.Function)) || 0 == strcmp(reinterpret_cast<const char*>(by_name->Name), hook.function_name.Buffer))
        {
          hook.old_address = PVOID(thunk->u1.Function);
          PVOID target = &thunk->u1.Function;
          SIZE_T target_size = sizeof(PVOID);
          ULONG old_protect;
          auto status = NtProtectVirtualMemory(
            NtCurrentProcess(),
            &target,
            &target_size,
            PAGE_EXECUTE_READWRITE,
            &old_protect
          );
          thunk->u1.Function = ULONG_PTR(hook.new_address);
          status = NtProtectVirtualMemory(
            NtCurrentProcess(),
            &target,
            &target_size,
            old_protect,
            &old_protect
          );
        }
      }
    }
    thunk++;
    original_thunk++;
  }
}

static void apply_iat_hooks_on_dll(PVOID dll)
{
  const auto base = PUCHAR(dll);

  const auto dosh = PIMAGE_DOS_HEADER(dll);
  if (dosh->e_magic != IMAGE_DOS_SIGNATURE)
    return;

  const auto nth = PIMAGE_NT_HEADERS(base + dosh->e_lfanew);
  if (nth->Signature != IMAGE_NT_SIGNATURE)
    return;

  const auto import_dir = &nth->OptionalHeader.DataDirectory[IMAGE_DIRECTORY_ENTRY_IMPORT];

  if (import_dir->VirtualAddress == 0 || import_dir->Size == 0)
    return;

  const auto import_begin = PIMAGE_IMPORT_DESCRIPTOR(base + import_dir->VirtualAddress);
  const auto import_end = PIMAGE_IMPORT_DESCRIPTOR(PUCHAR(import_begin) + import_dir->Size);

  for(auto desc = import_begin; desc < import_end; ++desc)
  {
    if (!desc->Name)
      break;

    const auto thunk = PIMAGE_THUNK_DATA(base + desc->FirstThunk);
    const auto original_thunk = PIMAGE_THUNK_DATA(base + desc->OriginalFirstThunk);

    hook_thunks(base, thunk, original_thunk);
  }
}

static VOID NTAPI DllLoadCallback(PWSTR DllName, PVOID DllBase, SIZE_T DllSize, PVOID Reserved)
{
  UNREFERENCED_PARAMETER(DllSize);
  UNREFERENCED_PARAMETER(Reserved);

  for (auto& target : s_target_images)
  {
  if (wcsstr(DllName, L"SppExtComObj"))
    {
      target.base = DllBase;
      apply_iat_hooks_on_dll(DllBase);
    }
  }
}

RPC_STATUS RPC_ENTRY RpcStringBindingComposeW_Hook(PWSTR ObjUuid, PWSTR ProtSeq, PWSTR NetworkAddr, PWSTR EndPoint, PWSTR Options, PWSTR* StringBinding)
{
  if (ProtSeq != nullptr && 0 == _wcsicmp(ProtSeq, PROTO_SEQ_TCP))
  {
    NetworkAddr = const_cast<wchar_t*>(LOCALHOST_IP);
  }

  return GET_ORIGINAL_FUNC(RpcStringBindingComposeW)(ObjUuid, ProtSeq, NetworkAddr, EndPoint, Options, StringBinding);
}
