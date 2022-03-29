
#include <Windows.h>
#include <psapi.h>

#pragma comment (lib, "ntdll.lib")
#pragma comment (linker, "/EXPORT:GetProcessMemoryInfo64")

typedef enum _PROCESSINFOCLASS {
  ProcessVmCounters = 3
} PROCESSINFOCLASS;

typedef struct _VM_COUNTERS {
  SIZE_T PeakVirtualSize;
  SIZE_T VirtualSize;
  ULONG PageFaultCount;
  SIZE_T PeakWorkingSetSize;
  SIZE_T WorkingSetSize;
  SIZE_T QuotaPeakPagedPoolUsage;
  SIZE_T QuotaPagedPoolUsage;
  SIZE_T QuotaPeakNonPagedPoolUsage;
  SIZE_T QuotaNonPagedPoolUsage;
  SIZE_T PagefileUsage;
  SIZE_T PeakPagefileUsage;
  SIZE_T PrivatePageCount;
} VM_COUNTERS;

typedef struct _CLIENT_ID {
  HANDLE  UniqueProcess;
  HANDLE  UniqueThread;
} CLIENT_ID, *PCLIENT_ID;

typedef struct _UNICODE_STRING {
  USHORT Length;
  USHORT MaximumLength;
  PWSTR  Buffer;
} UNICODE_STRING, *PUNICODE_STRING;

typedef struct _OBJECT_ATTRIBUTES {
  ULONG Length;
  HANDLE RootDirectory;
  PUNICODE_STRING ObjectName;
  ULONG Attributes;
  PVOID SecurityDescriptor;
  PVOID SecurityQualityOfService;
} OBJECT_ATTRIBUTES;
typedef OBJECT_ATTRIBUTES *POBJECT_ATTRIBUTES;

NTSTATUS NTAPI
NtOpenProcess (
    __out PHANDLE  ProcessHandle,
    __in ACCESS_MASK  DesiredAccess,
    __in POBJECT_ATTRIBUTES  ObjectAttributes,
    __in_opt PCLIENT_ID  ClientId
    );

NTSTATUS NTAPI
NtQueryInformationProcess(
  IN HANDLE  ProcessHandle,
  IN PROCESSINFOCLASS  ProcessInformationClass,
  OUT PVOID  ProcessInformation,
  IN ULONG  ProcessInformationLength,
  OUT PULONG  ReturnLength  OPTIONAL);

NTSTATUS NTAPI
NtClose(
    IN HANDLE  Handle
    );

#ifndef NT_SUCCESS
#define NT_SUCCESS(x) ((x)>=0)
#define STATUS_SUCCESS ((NTSTATUS)0)
#endif

BOOL WINAPI DllMain(HINSTANCE hinstDLL, DWORD fdwReason, LPVOID lpReserved) {
    return TRUE;
}

BOOL WINAPI GetProcessMemoryInfo64 (DWORD dwPID, PROCESS_MEMORY_COUNTERS *ppsmemCounters) {
	VM_COUNTERS tVmCounters;
	CLIENT_ID tClientId;
	OBJECT_ATTRIBUTES tObjAttr;
	HANDLE hProcess;
	ULONG uRet;
	BOOL bRet = FALSE;

	if (!ppsmemCounters || ppsmemCounters->cb != sizeof(PROCESS_MEMORY_COUNTERS))
		return FALSE;

	tClientId.UniqueThread = 0;
	tClientId.UniqueProcess = (HANDLE)dwPID;

	tObjAttr.Length = sizeof(OBJECT_ATTRIBUTES);
	tObjAttr.Attributes = 0;
	tObjAttr.RootDirectory = 0;
	tObjAttr.ObjectName = 0;
	tObjAttr.SecurityDescriptor = 0;
	tObjAttr.SecurityQualityOfService = 0;

	if (!NT_SUCCESS(NtOpenProcess(&hProcess, PROCESS_QUERY_LIMITED_INFORMATION, &tObjAttr, &tClientId)))
		return FALSE;

	if (NT_SUCCESS(NtQueryInformationProcess(hProcess, ProcessVmCounters, &tVmCounters, sizeof(tVmCounters), &uRet))) {

            ppsmemCounters->PageFaultCount = tVmCounters.PageFaultCount;
            ppsmemCounters->PagefileUsage = tVmCounters.PagefileUsage;
            ppsmemCounters->PeakPagefileUsage = tVmCounters.PeakPagefileUsage;
            ppsmemCounters->PeakWorkingSetSize = tVmCounters.PeakWorkingSetSize;
            ppsmemCounters->QuotaNonPagedPoolUsage = tVmCounters.QuotaNonPagedPoolUsage;
            ppsmemCounters->QuotaPagedPoolUsage = tVmCounters.QuotaPagedPoolUsage;
            ppsmemCounters->QuotaPeakNonPagedPoolUsage = tVmCounters.QuotaPeakNonPagedPoolUsage;
            ppsmemCounters->QuotaPeakPagedPoolUsage = tVmCounters.QuotaPeakPagedPoolUsage;
            ppsmemCounters->WorkingSetSize = tVmCounters.WorkingSetSize;

			bRet = TRUE;

	}

	NtClose(hProcess);

	return bRet;


}