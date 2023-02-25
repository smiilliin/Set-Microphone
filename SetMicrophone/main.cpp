#include <Windows.h>
#include <iostream>
#include <fstream>
#include <string>
#include <functiondiscoverykeys.h>
#include <sys/stat.h>

#include "PolicyConfig.h"

#pragma comment(lib, "winmm.lib")

using namespace std;


HRESULT SetDefaultAudioPlaybackDevice(LPCWSTR devID)
{
	IPolicyConfigVista* pPolicyConfig;
	ERole reserved = eConsole;

	HRESULT hr = CoCreateInstance(__uuidof(CPolicyConfigVistaClient), nullptr, CLSCTX_ALL, __uuidof(IPolicyConfigVista), (LPVOID*)&pPolicyConfig);
	if (FAILED(hr)) {
		pPolicyConfig->Release();
		return hr;
	}
	hr = pPolicyConfig->SetDefaultEndpoint(devID, reserved);
	if (FAILED(hr)) {
		pPolicyConfig->Release();
		return hr;
	}

	pPolicyConfig->Release();
	return hr;
}

wofstream logFile;

int errorCount = 0;

void Log(wstring log) {
	logFile << log << endl;
}
#define ThrowError(err) \
						{ \
							Log(L"[Error line" + to_wstring(__LINE__) + + L" " + err + L"] "); \
							Sleep(1000); \
							errorCount++; \
							if (errorCount > 50) { \
								logFile.close(); \
								exit(1); \
							} \
							else { \
								goto main; \
							} \
						}


int WINAPI WinMain(
	HINSTANCE hInstance,
	HINSTANCE hPrevInstance,
	LPSTR lpCmdLine,
	int nShowCmd
) {
	struct stat info;

	if (stat("SetMicrophone", &info) != 0) {
		if (_wmkdir(L"SetMicrophone") != 0) {
			return MessageBoxW(nullptr, L"[Error] Cannot create a folder setMicrophone", L"SetMicrophone Error", MB_ICONERROR);
		}
	}
	else {
		if (!(info.st_mode & S_IFDIR)) {
			return MessageBoxW(nullptr, L"[Error] SetMicrophone is not a folder", L"SetMicrophone Error", MB_ICONERROR);
		}
	}

	logFile.open("SetMicrophone\\log.txt");
main:
	wifstream toFindDeviceNameFile("SetMicrophone\\device.txt");

	if (!toFindDeviceNameFile.is_open()) ThrowError(L"The device file cannot be opened");

	wstring toFindDeviceName;

	toFindDeviceNameFile >> toFindDeviceName;

	if (toFindDeviceName == L"") ThrowError(L"Device name is not defined");

	wcout.imbue(locale(""));

	HRESULT hr;

	hr = CoInitialize(nullptr);

	if (FAILED(hr)) {
		ThrowError(L"CoInitialize");
	}

	IMMDeviceEnumerator* enumerator;
	hr = CoCreateInstance(__uuidof(MMDeviceEnumerator), nullptr, CLSCTX_ALL, __uuidof(IMMDeviceEnumerator), (void**)&enumerator);

	if (FAILED(hr)) {
		ThrowError(L"CoInitialize");
	}

	IMMDeviceCollection* devices;
	hr = enumerator->EnumAudioEndpoints(eCapture, DEVICE_STATE_ACTIVE, &devices);

	if (FAILED(hr)) {
		ThrowError(L"EnumAudioEndpoints");
	}

	UINT devicesCount;
	hr = devices->GetCount(&devicesCount);

	if (FAILED(hr)) {
		ThrowError(L"GetCount");
	}
	for (UINT i = 0; i < devicesCount; i++) {
		IMMDevice* device;
		hr = devices->Item(i, &device);
		if (FAILED(hr)) {
			ThrowError(L"Item");
		}

		IPropertyStore* propertyStore;
		hr = device->OpenPropertyStore(STGM_READ, &propertyStore);
		if (FAILED(hr)) {
			ThrowError(L"OpenPropertyStore");
		}

		PROPVARIANT deviceProp;

		PropVariantInit(&deviceProp);

		hr = propertyStore->GetValue(PKEY_Device_FriendlyName, &deviceProp);

		wstring deviceName(deviceProp.pwszVal);
		
		if (deviceName.find(toFindDeviceName) != string::npos) {
			LPWSTR deviceID;
			hr = device->GetId(&deviceID);
			if (FAILED(hr)) {
				ThrowError(L"GetId");
			}

			hr = SetDefaultAudioPlaybackDevice(deviceID);

			if (FAILED(hr)) {
				ThrowError(L"SetDefaultAudioPlaybackDevice");
			}

			Log(L"OK " + deviceName);
			break;
		}
	}

	return 0;
}