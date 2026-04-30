#ifndef UNICODE
#define UNICODE
#endif
#ifndef _UNICODE
#define _UNICODE
#endif
#define WIN32_LEAN_AND_MEAN

#include <windows.h>
#include <commctrl.h>
#include <mmdeviceapi.h>
#include <endpointvolume.h>
#include <audioclient.h>
#include <mmreg.h>
#include <ksmedia.h>
#include <functiondiscoverykeys_devpkey.h>
#include <propvarutil.h>
#include <powrprof.h>
#include <shellapi.h>
#include <strsafe.h>
#include <stdlib.h>
#include <stdint.h>
#include <vector>
#include <string>
#include <sstream>
#include <algorithm>

#pragma comment(lib, "comctl32.lib")
#pragma comment(lib, "ole32.lib")
#pragma comment(lib, "propsys.lib")
#pragma comment(lib, "shell32.lib")
#pragma comment(lib, "advapi32.lib")
#pragma comment(lib, "powrprof.lib")

// Undocumented but widely used Windows endpoint policy interface. It is needed
// because the public Core Audio API can enumerate devices but cannot change the
// system default endpoint.
struct IPolicyConfig : public IUnknown {
    virtual HRESULT STDMETHODCALLTYPE GetMixFormat(LPCWSTR, WAVEFORMATEX**) = 0;
    virtual HRESULT STDMETHODCALLTYPE GetDeviceFormat(LPCWSTR, INT, WAVEFORMATEX**) = 0;
    virtual HRESULT STDMETHODCALLTYPE ResetDeviceFormat(LPCWSTR) = 0;
    virtual HRESULT STDMETHODCALLTYPE SetDeviceFormat(LPCWSTR, WAVEFORMATEX*, WAVEFORMATEX*) = 0;
    virtual HRESULT STDMETHODCALLTYPE GetProcessingPeriod(LPCWSTR, INT, INT64*, INT64*) = 0;
    virtual HRESULT STDMETHODCALLTYPE SetProcessingPeriod(LPCWSTR, INT64*) = 0;
    virtual HRESULT STDMETHODCALLTYPE GetShareMode(LPCWSTR, void*) = 0;
    virtual HRESULT STDMETHODCALLTYPE SetShareMode(LPCWSTR, void*) = 0;
    virtual HRESULT STDMETHODCALLTYPE GetPropertyValue(LPCWSTR, const PROPERTYKEY&, PROPVARIANT*) = 0;
    virtual HRESULT STDMETHODCALLTYPE SetPropertyValue(LPCWSTR, const PROPERTYKEY&, PROPVARIANT*) = 0;
    virtual HRESULT STDMETHODCALLTYPE SetDefaultEndpoint(LPCWSTR, ERole) = 0;
    virtual HRESULT STDMETHODCALLTYPE SetEndpointVisibility(LPCWSTR, INT) = 0;
};

static const CLSID CLSID_PolicyConfigClient =
{ 0x870af99c, 0x171d, 0x4f9e, { 0xaf, 0x0d, 0xe6, 0x3d, 0xf4, 0x0c, 0x2b, 0xc9 } };
static const IID IID_IPolicyConfig =
{ 0xf8679f50, 0x850a, 0x41cf, { 0x9c, 0x72, 0x43, 0x0f, 0x29, 0x02, 0x90, 0xc8 } };

static const PROPERTYKEY PKEY_AudioEndpoint_Disable_SysFx_Local =
{ { 0x1da5d803, 0xd492, 0x4edd, { 0x8c, 0x23, 0xe0, 0xc0, 0xff, 0xee, 0x7f, 0x0e } }, 5 };
static const PROPERTYKEY PKEY_AudioEngine_DeviceFormat_Local =
{ { 0xf19f064d, 0x82c, 0x4e27, { 0xbc, 0x73, 0x68, 0x82, 0xa1, 0xbb, 0x8e, 0x4c } }, 0 };

enum ControlId {
    IDC_OUTPUTS = 1001, IDC_INPUTS, IDC_LOG,
    IDC_REFRESH, IDC_DEFAULT_OUT, IDC_DEFAULT_IN,
    IDC_FREQ, IDC_BITS, IDC_CHANNELS,
    IDC_APPLY_OUT, IDC_APPLY_IN, IDC_APPLY_ALL,
    IDC_PRESET_48, IDC_PRESET_44,
    IDC_VOLUME, IDC_MUTE, IDC_APPLY_VOLUME,
    IDC_RESTART_AUDIO, IDC_FIX_CRACKLE, IDC_ADMIN,
    IDC_PRO_TUNE, IDC_RESTORE_TUNE, IDC_SOUND_PANEL,
    IDC_APPLY_SELECTED, IDC_RESET_SELECTED,
    IDC_OPTIMIZE_SERVICES, IDC_LOW_LATENCY_POWER, IDC_POWER_REPORT
    , IDC_DISABLE_EFFECTS, IDC_ENABLE_EFFECTS, IDC_SOUND_SETTINGS,
    IDC_EXCLUSIVE_ON_SELECTED, IDC_EXCLUSIVE_OFF_SELECTED, IDC_EXCLUSIVE_RESTORE_SELECTED,
    IDC_EXCLUSIVE_ON_ALL, IDC_EXCLUSIVE_OFF_ALL,
    IDC_DUCKING_OFF, IDC_DUCKING_RESTORE,
    IDC_MMCSS_APPLY, IDC_MMCSS_RESTORE,
    IDC_UI_DELAY_APPLY, IDC_UI_DELAY_RESTORE,
    IDC_RESTORE_ALL_TWEAKS, IDC_REFRESH_CONFIG,
    IDC_HIDE_SELECTED_DEVICES, IDC_SHOW_SELECTED_DEVICES
};

struct DeviceInfo {
    std::wstring id;
    std::wstring name;
    EDataFlow flow;
    DWORD state;
};

struct AppState {
    HWND hwnd = nullptr;
    HWND outputs = nullptr;
    HWND inputs = nullptr;
    HWND log = nullptr;
    HWND freq = nullptr;
    HWND bits = nullptr;
    HWND channels = nullptr;
    HWND volume = nullptr;
    HWND mute = nullptr;
    HWND config = nullptr;
    HWND tabs = nullptr;
    std::vector<DeviceInfo> outputDevices;
    std::vector<DeviceInfo> inputDevices;
    std::vector<HWND> deviceTabControls;
    std::vector<HWND> tweakTabControls;
    std::vector<HWND> commonControls;
    int activeTab = 0;
    int createTab = 0;
};

static AppState g;

static const wchar_t* kAudioServices[] = {
    L"Audiosrv",
    L"AudioEndpointBuilder",
    L"MMCSS"
};

#ifndef ENDPOINT_SYSFX_ENABLED
#define ENDPOINT_SYSFX_ENABLED 0x00000000
#endif
#ifndef ENDPOINT_SYSFX_DISABLED
#define ENDPOINT_SYSFX_DISABLED 0x00000001
#endif

static bool MmDevicePropertiesPath(const DeviceInfo& device, std::wstring* path);
static HRESULT GetEndpointVolumeState(const DeviceInfo& device, float* volume, BOOL* mute);
static std::wstring GetVolumeText(const DeviceInfo& device);

template <class T>
static void SafeRelease(T** value) {
    if (*value) {
        (*value)->Release();
        *value = nullptr;
    }
}

static std::wstring HrText(HRESULT hr) {
    wchar_t* buffer = nullptr;
    FormatMessageW(FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS,
        nullptr, hr, 0, reinterpret_cast<LPWSTR>(&buffer), 0, nullptr);
    std::wstringstream ss;
    ss << L"0x" << std::hex << static_cast<unsigned long>(hr);
    if (buffer) {
        ss << L" (" << buffer << L")";
        LocalFree(buffer);
    }
    return ss.str();
}

static void Log(const std::wstring& text) {
    int len = GetWindowTextLengthW(g.log);
    SendMessageW(g.log, EM_SETSEL, len, len);
    std::wstring line = text + L"\r\n";
    SendMessageW(g.log, EM_REPLACESEL, FALSE, reinterpret_cast<LPARAM>(line.c_str()));
}

static void TrackControl(HWND hwnd) {
    if (!hwnd) return;
    if (g.createTab == 1) {
        g.tweakTabControls.push_back(hwnd);
    } else if (g.createTab == 2) {
        g.commonControls.push_back(hwnd);
    } else {
        g.deviceTabControls.push_back(hwnd);
    }
}

static void ShowTab(int index) {
    g.activeTab = index;
    for (HWND hwnd : g.deviceTabControls) ShowWindow(hwnd, index == 0 ? SW_SHOW : SW_HIDE);
    for (HWND hwnd : g.tweakTabControls) ShowWindow(hwnd, index == 1 ? SW_SHOW : SW_HIDE);
    for (HWND hwnd : g.commonControls) ShowWindow(hwnd, SW_SHOW);
    if (g.tabs) TabCtrl_SetCurSel(g.tabs, index);
}

static std::wstring LastErrorText(DWORD err = GetLastError()) {
    return HrText(HRESULT_FROM_WIN32(err));
}

static std::wstring ServiceStateText(DWORD state) {
    switch (state) {
    case SERVICE_STOPPED: return L"stopped";
    case SERVICE_START_PENDING: return L"start pending";
    case SERVICE_STOP_PENDING: return L"stop pending";
    case SERVICE_RUNNING: return L"running";
    case SERVICE_CONTINUE_PENDING: return L"continue pending";
    case SERVICE_PAUSE_PENDING: return L"pause pending";
    case SERVICE_PAUSED: return L"paused";
    default: return L"unknown";
    }
}

static bool Confirm(const wchar_t* title, const wchar_t* message) {
    return MessageBoxW(g.hwnd, message, title, MB_ICONWARNING | MB_YESNO | MB_DEFBUTTON2) == IDYES;
}

static bool IsElevated() {
    HANDLE token = nullptr;
    if (!OpenProcessToken(GetCurrentProcess(), TOKEN_QUERY, &token)) return false;
    TOKEN_ELEVATION elevation{};
    DWORD returned = 0;
    bool elevated = GetTokenInformation(token, TokenElevation, &elevation, sizeof(elevation), &returned)
        && elevation.TokenIsElevated;
    CloseHandle(token);
    return elevated;
}

static void RelaunchElevated() {
    wchar_t path[MAX_PATH]{};
    GetModuleFileNameW(nullptr, path, MAX_PATH);
    ShellExecuteW(g.hwnd, L"runas", path, nullptr, nullptr, SW_SHOWNORMAL);
}

static std::wstring GetDeviceName(IMMDevice* device) {
    IPropertyStore* store = nullptr;
    PROPVARIANT value;
    PropVariantInit(&value);
    std::wstring name = L"(unknown device)";
    if (SUCCEEDED(device->OpenPropertyStore(STGM_READ, &store)) &&
        SUCCEEDED(store->GetValue(PKEY_Device_FriendlyName, &value)) &&
        value.vt == VT_LPWSTR && value.pwszVal) {
        name = value.pwszVal;
    }
    PropVariantClear(&value);
    SafeRelease(&store);
    return name;
}

static void AddDeviceToList(HWND list, const DeviceInfo& info) {
    std::wstring prefix = info.state == DEVICE_STATE_ACTIVE ? L"" : L"[disabled/unplugged] ";
    int index = ListView_GetItemCount(list);
    std::wstring text = prefix + info.name;
    LVITEMW item{};
    item.mask = LVIF_TEXT;
    item.iItem = index;
    item.pszText = const_cast<LPWSTR>(text.c_str());
    ListView_InsertItem(list, &item);
    ListView_SetCheckState(list, index, FALSE);
}

static void InitDeviceList(HWND list) {
    ListView_SetExtendedListViewStyle(list, LVS_EX_CHECKBOXES | LVS_EX_FULLROWSELECT | LVS_EX_DOUBLEBUFFER);
    LVCOLUMNW col{};
    col.mask = LVCF_TEXT | LVCF_WIDTH;
    col.cx = 480;
    col.pszText = const_cast<LPWSTR>(L"Device");
    ListView_InsertColumn(list, 0, &col);
}

static void UpdateSelectedConfig();

static void RefreshDevices() {
    ListView_DeleteAllItems(g.outputs);
    ListView_DeleteAllItems(g.inputs);
    g.outputDevices.clear();
    g.inputDevices.clear();

    IMMDeviceEnumerator* enumerator = nullptr;
    HRESULT hr = CoCreateInstance(__uuidof(MMDeviceEnumerator), nullptr, CLSCTX_ALL,
        __uuidof(IMMDeviceEnumerator), reinterpret_cast<void**>(&enumerator));
    if (FAILED(hr)) {
        Log(L"Device enumeration failed: " + HrText(hr));
        return;
    }

    for (EDataFlow flow : { eRender, eCapture }) {
        IMMDeviceCollection* collection = nullptr;
        hr = enumerator->EnumAudioEndpoints(flow, DEVICE_STATE_ACTIVE | DEVICE_STATE_DISABLED | DEVICE_STATE_UNPLUGGED, &collection);
        if (FAILED(hr)) {
            Log(std::wstring(flow == eRender ? L"Output" : L"Input") + L" enumeration failed: " + HrText(hr));
            continue;
        }

        UINT count = 0;
        collection->GetCount(&count);
        for (UINT i = 0; i < count; ++i) {
            IMMDevice* device = nullptr;
            LPWSTR id = nullptr;
            DWORD state = 0;
            if (SUCCEEDED(collection->Item(i, &device)) && SUCCEEDED(device->GetId(&id))) {
                device->GetState(&state);
                DeviceInfo info{ id, GetDeviceName(device), flow, state };
                if (flow == eRender) {
                    g.outputDevices.push_back(info);
                    AddDeviceToList(g.outputs, info);
                } else {
                    g.inputDevices.push_back(info);
                    AddDeviceToList(g.inputs, info);
                }
            }
            if (id) CoTaskMemFree(id);
            SafeRelease(&device);
        }
        SafeRelease(&collection);
    }
    SafeRelease(&enumerator);
    Log(L"Devices refreshed.");
    UpdateSelectedConfig();
}

static int ComboNumber(HWND combo, int fallback) {
    wchar_t text[64]{};
    GetWindowTextW(combo, text, 64);
    int value = _wtoi(text);
    return value > 0 ? value : fallback;
}

static WAVEFORMATEXTENSIBLE MakeFormat(int sampleRate, int bits, int channels) {
    WAVEFORMATEXTENSIBLE format{};
    format.Format.wFormatTag = WAVE_FORMAT_EXTENSIBLE;
    format.Format.nChannels = static_cast<WORD>(channels);
    format.Format.nSamplesPerSec = sampleRate;
    format.Format.wBitsPerSample = static_cast<WORD>(bits);
    format.Format.nBlockAlign = static_cast<WORD>((channels * bits) / 8);
    format.Format.nAvgBytesPerSec = sampleRate * format.Format.nBlockAlign;
    format.Format.cbSize = sizeof(WAVEFORMATEXTENSIBLE) - sizeof(WAVEFORMATEX);
    format.Samples.wValidBitsPerSample = static_cast<WORD>(bits);
    format.dwChannelMask = channels == 1 ? SPEAKER_FRONT_CENTER :
        channels == 2 ? (SPEAKER_FRONT_LEFT | SPEAKER_FRONT_RIGHT) : 0;
    format.SubFormat = bits == 32 ? KSDATAFORMAT_SUBTYPE_IEEE_FLOAT : KSDATAFORMAT_SUBTYPE_PCM;
    return format;
}

static HRESULT GetPolicyConfig(IPolicyConfig** policy) {
    return CoCreateInstance(CLSID_PolicyConfigClient, nullptr, CLSCTX_ALL, IID_IPolicyConfig,
        reinterpret_cast<void**>(policy));
}

static HRESULT SetDeviceFormat(const DeviceInfo& device, int sampleRate, int bits, int channels) {
    WAVEFORMATEXTENSIBLE format = MakeFormat(sampleRate, bits, channels);
    IPolicyConfig* policy = nullptr;
    HRESULT hr = GetPolicyConfig(&policy);
    if (SUCCEEDED(hr)) {
        hr = policy->SetDeviceFormat(device.id.c_str(), &format.Format, &format.Format);
    }
    SafeRelease(&policy);
    return hr;
}

static HRESULT ResetDeviceFormat(const DeviceInfo& device) {
    IPolicyConfig* policy = nullptr;
    HRESULT hr = GetPolicyConfig(&policy);
    if (SUCCEEDED(hr)) {
        hr = policy->ResetDeviceFormat(device.id.c_str());
    }
    SafeRelease(&policy);
    return hr;
}

static HRESULT IsDeviceFormatSupported(const DeviceInfo& device, const WAVEFORMATEX* format, WAVEFORMATEX** closestOut) {
    if (closestOut) *closestOut = nullptr;
    IMMDeviceEnumerator* enumerator = nullptr;
    IMMDevice* endpoint = nullptr;
    IAudioClient* client = nullptr;
    WAVEFORMATEX* closest = nullptr;
    HRESULT hr = CoCreateInstance(__uuidof(MMDeviceEnumerator), nullptr, CLSCTX_ALL,
        __uuidof(IMMDeviceEnumerator), reinterpret_cast<void**>(&enumerator));
    if (SUCCEEDED(hr)) hr = enumerator->GetDevice(device.id.c_str(), &endpoint);
    if (SUCCEEDED(hr)) hr = endpoint->Activate(__uuidof(IAudioClient), CLSCTX_ALL, nullptr, reinterpret_cast<void**>(&client));
    if (SUCCEEDED(hr)) hr = client->IsFormatSupported(AUDCLNT_SHAREMODE_SHARED, format, &closest);
    if (closestOut) {
        *closestOut = closest;
    } else if (closest) {
        CoTaskMemFree(closest);
    }
    SafeRelease(&client);
    SafeRelease(&endpoint);
    SafeRelease(&enumerator);
    return hr;
}

static std::wstring FormatText(const WAVEFORMATEX* format) {
    if (!format) return L"(none)";
    std::wstringstream ss;
    ss << format->nSamplesPerSec << L" Hz / " << format->wBitsPerSample
       << L"-bit / " << format->nChannels << L" channel(s)";
    return ss.str();
}

static std::wstring GetPropertyStoreDeviceFormatText(const DeviceInfo& device) {
    IMMDeviceEnumerator* enumerator = nullptr;
    IMMDevice* endpoint = nullptr;
    IPropertyStore* store = nullptr;
    PROPVARIANT value;
    PropVariantInit(&value);
    HRESULT hr = CoCreateInstance(__uuidof(MMDeviceEnumerator), nullptr, CLSCTX_ALL,
        __uuidof(IMMDeviceEnumerator), reinterpret_cast<void**>(&enumerator));
    if (SUCCEEDED(hr)) hr = enumerator->GetDevice(device.id.c_str(), &endpoint);
    if (SUCCEEDED(hr)) hr = endpoint->OpenPropertyStore(STGM_READ, &store);
    if (SUCCEEDED(hr)) hr = store->GetValue(PKEY_AudioEngine_DeviceFormat_Local, &value);

    std::wstring text;
    if (SUCCEEDED(hr) && value.vt == VT_BLOB && value.blob.cbSize >= sizeof(WAVEFORMATEX)) {
        text = FormatText(reinterpret_cast<const WAVEFORMATEX*>(value.blob.pBlobData));
    } else {
        text = L"unavailable: " + HrText(hr);
    }

    PropVariantClear(&value);
    SafeRelease(&store);
    SafeRelease(&endpoint);
    SafeRelease(&enumerator);
    return text;
}

static std::wstring GetCurrentDeviceFormatText(const DeviceInfo& device) {
    IPolicyConfig* policy = nullptr;
    WAVEFORMATEX* format = nullptr;
    HRESULT hr = GetPolicyConfig(&policy);
    if (SUCCEEDED(hr)) hr = policy->GetDeviceFormat(device.id.c_str(), FALSE, &format);
    std::wstring text = SUCCEEDED(hr) ? FormatText(format) : GetPropertyStoreDeviceFormatText(device);
    if (format) CoTaskMemFree(format);
    SafeRelease(&policy);
    return text;
}

static std::wstring GetMixFormatText(const DeviceInfo& device) {
    IMMDeviceEnumerator* enumerator = nullptr;
    IMMDevice* endpoint = nullptr;
    IAudioClient* client = nullptr;
    WAVEFORMATEX* format = nullptr;
    HRESULT hr = CoCreateInstance(__uuidof(MMDeviceEnumerator), nullptr, CLSCTX_ALL,
        __uuidof(IMMDeviceEnumerator), reinterpret_cast<void**>(&enumerator));
    if (SUCCEEDED(hr)) hr = enumerator->GetDevice(device.id.c_str(), &endpoint);
    if (SUCCEEDED(hr)) hr = endpoint->Activate(__uuidof(IAudioClient), CLSCTX_ALL, nullptr, reinterpret_cast<void**>(&client));
    if (SUCCEEDED(hr)) hr = client->GetMixFormat(&format);
    std::wstring text = SUCCEEDED(hr) ? FormatText(format) : L"unavailable: " + HrText(hr);
    if (format) CoTaskMemFree(format);
    SafeRelease(&client);
    SafeRelease(&endpoint);
    SafeRelease(&enumerator);
    return text;
}

static std::wstring GetSysFxText(const DeviceInfo& device) {
    IMMDeviceEnumerator* enumerator = nullptr;
    IMMDevice* endpoint = nullptr;
    IPropertyStore* store = nullptr;
    PROPVARIANT value;
    PropVariantInit(&value);
    HRESULT hr = CoCreateInstance(__uuidof(MMDeviceEnumerator), nullptr, CLSCTX_ALL,
        __uuidof(IMMDeviceEnumerator), reinterpret_cast<void**>(&enumerator));
    if (SUCCEEDED(hr)) hr = enumerator->GetDevice(device.id.c_str(), &endpoint);
    if (SUCCEEDED(hr)) hr = endpoint->OpenPropertyStore(STGM_READ, &store);
    if (SUCCEEDED(hr)) hr = store->GetValue(PKEY_AudioEndpoint_Disable_SysFx_Local, &value);

    std::wstring text = L"unknown";
    if (SUCCEEDED(hr) && value.vt == VT_UI4) {
        text = value.ulVal == ENDPOINT_SYSFX_DISABLED ? L"disabled" : L"enabled";
    } else if (HRESULT_CODE(hr) == ERROR_NOT_FOUND) {
        text = L"driver default";
    } else if (FAILED(hr)) {
        text = L"unavailable: " + HrText(hr);
    }
    PropVariantClear(&value);
    SafeRelease(&store);
    SafeRelease(&endpoint);
    SafeRelease(&enumerator);

    std::wstring path;
    if (MmDevicePropertiesPath(device, &path)) {
        HKEY key = nullptr;
        if (RegOpenKeyExW(HKEY_LOCAL_MACHINE, path.c_str(), 0, KEY_QUERY_VALUE | KEY_WOW64_64KEY, &key) == ERROR_SUCCESS) {
            DWORD regValue = 0, type = 0, bytes = sizeof(DWORD);
            if (RegQueryValueExW(key, L"{1da5d803-d492-4edd-8c23-e0c0ffee7f0e},5", nullptr, &type,
                reinterpret_cast<BYTE*>(&regValue), &bytes) == ERROR_SUCCESS && type == REG_DWORD) {
                std::wstring regText = regValue == ENDPOINT_SYSFX_DISABLED ? L"disabled" : L"enabled";
                if (text == L"unknown" || text == L"driver default" || text.rfind(L"unavailable:", 0) == 0) {
                    text = regText + L" (registry fallback/pending restart)";
                } else if (text != regText) {
                    text += L"; registry says " + regText + L" pending restart";
                }
            }
            RegCloseKey(key);
        }
    }
    return text;
}

static HRESULT SetEndpointSysFx(const DeviceInfo& device, bool disabled) {
    PROPVARIANT value;
    PropVariantInit(&value);
    value.vt = VT_UI4;
    value.ulVal = disabled ? ENDPOINT_SYSFX_DISABLED : ENDPOINT_SYSFX_ENABLED;
    IPolicyConfig* policy = nullptr;
    HRESULT hr = GetPolicyConfig(&policy);
    if (SUCCEEDED(hr)) {
        hr = policy->SetPropertyValue(device.id.c_str(), PKEY_AudioEndpoint_Disable_SysFx_Local, &value);
    }
    SafeRelease(&policy);
    PropVariantClear(&value);
    return hr;
}

static std::wstring DeviceStateText(DWORD state) {
    switch (state) {
    case DEVICE_STATE_ACTIVE: return L"active";
    case DEVICE_STATE_DISABLED: return L"disabled";
    case DEVICE_STATE_NOTPRESENT: return L"not present";
    case DEVICE_STATE_UNPLUGGED: return L"unplugged";
    default: return L"unknown";
    }
}

static void ApplyFormatToList(const std::vector<DeviceInfo>& devices, const wchar_t* label) {
    int sampleRate = ComboNumber(g.freq, 48000);
    int bits = ComboNumber(g.bits, 24);
    int channels = ComboNumber(g.channels, 2);
    std::wstringstream start;
    start << L"Applying " << sampleRate << L" Hz / " << bits << L"-bit / " << channels << L" channel(s) to " << label << L"...";
    Log(start.str());

    WAVEFORMATEXTENSIBLE format = MakeFormat(sampleRate, bits, channels);
    for (const auto& device : devices) {
        if (device.state != DEVICE_STATE_ACTIVE) {
            Log(L"Skipped inactive device: " + device.name);
            continue;
        }
        WAVEFORMATEX* closest = nullptr;
        HRESULT supported = IsDeviceFormatSupported(device, &format.Format, &closest);
        if (supported == S_OK) {
            Log(L"Driver reports shared-mode support: " + device.name);
        } else if (supported == S_FALSE) {
            Log(L"Driver suggested closest shared format for " + device.name + L": " + FormatText(closest) + L". Trying selected Windows endpoint format anyway.");
        } else {
            Log(L"Shared-mode pre-check could not confirm support for " + device.name + L": " + HrText(supported) + L". Trying selected Windows endpoint format anyway.");
        }
        if (closest) CoTaskMemFree(closest);

        HRESULT hr = SetDeviceFormat(device, sampleRate, bits, channels);
        if (SUCCEEDED(hr)) {
            Log(L"Updated: " + device.name);
        } else {
            Log(L"Windows rejected this endpoint format for " + device.name + L": " + HrText(hr));
        }
    }
    UpdateSelectedConfig();
}

static void ApplyFormatToDevice(const DeviceInfo& device) {
    int sampleRate = ComboNumber(g.freq, 48000);
    int bits = ComboNumber(g.bits, 24);
    int channels = ComboNumber(g.channels, 2);
    WAVEFORMATEXTENSIBLE format = MakeFormat(sampleRate, bits, channels);

    std::wstringstream start;
    start << L"Applying selected format to " << device.name << L": "
          << sampleRate << L" Hz / " << bits << L"-bit / " << channels << L" channel(s)";
    Log(start.str());

    WAVEFORMATEX* closest = nullptr;
    HRESULT supported = IsDeviceFormatSupported(device, &format.Format, &closest);
    if (supported == S_OK) {
        Log(L"Driver reports shared-mode support for selected device.");
    } else if (supported == S_FALSE) {
        Log(L"Driver suggested closest shared format: " + FormatText(closest) + L". Trying selected Windows endpoint format anyway.");
    } else {
        Log(L"Shared-mode pre-check could not confirm support: " + HrText(supported) + L". Trying selected Windows endpoint format anyway.");
    }
    if (closest) CoTaskMemFree(closest);

    HRESULT hr = SetDeviceFormat(device, sampleRate, bits, channels);
    Log(SUCCEEDED(hr) ? L"Selected device format updated." : L"Windows rejected selected endpoint format: " + HrText(hr));
    UpdateSelectedConfig();
}

static void SelectCombo(HWND combo, const wchar_t* value) {
    int index = static_cast<int>(SendMessageW(combo, CB_FINDSTRINGEXACT, static_cast<WPARAM>(-1), reinterpret_cast<LPARAM>(value)));
    if (index >= 0) SendMessageW(combo, CB_SETCURSEL, index, 0);
}

static bool GetCurrentDeviceFormat(const DeviceInfo& device, WAVEFORMATEX* out) {
    if (!out) return false;
    ZeroMemory(out, sizeof(*out));
    IPolicyConfig* policy = nullptr;
    WAVEFORMATEX* format = nullptr;
    HRESULT hr = GetPolicyConfig(&policy);
    if (SUCCEEDED(hr)) hr = policy->GetDeviceFormat(device.id.c_str(), FALSE, &format);
    if (SUCCEEDED(hr) && format) {
        *out = *format;
        CoTaskMemFree(format);
        SafeRelease(&policy);
        return true;
    }
    if (format) CoTaskMemFree(format);
    SafeRelease(&policy);

    IMMDeviceEnumerator* enumerator = nullptr;
    IMMDevice* endpoint = nullptr;
    IPropertyStore* store = nullptr;
    PROPVARIANT value;
    PropVariantInit(&value);
    hr = CoCreateInstance(__uuidof(MMDeviceEnumerator), nullptr, CLSCTX_ALL,
        __uuidof(IMMDeviceEnumerator), reinterpret_cast<void**>(&enumerator));
    if (SUCCEEDED(hr)) hr = enumerator->GetDevice(device.id.c_str(), &endpoint);
    if (SUCCEEDED(hr)) hr = endpoint->OpenPropertyStore(STGM_READ, &store);
    if (SUCCEEDED(hr)) hr = store->GetValue(PKEY_AudioEngine_DeviceFormat_Local, &value);
    bool ok = false;
    if (SUCCEEDED(hr) && value.vt == VT_BLOB && value.blob.cbSize >= sizeof(WAVEFORMATEX)) {
        *out = *reinterpret_cast<const WAVEFORMATEX*>(value.blob.pBlobData);
        ok = true;
    }
    PropVariantClear(&value);
    SafeRelease(&store);
    SafeRelease(&endpoint);
    SafeRelease(&enumerator);
    return ok;
}

static HRESULT SetDefaultDevice(const DeviceInfo& device) {
    IPolicyConfig* policy = nullptr;
    HRESULT hr = GetPolicyConfig(&policy);
    if (SUCCEEDED(hr)) {
        for (ERole role : { eConsole, eMultimedia, eCommunications }) {
            HRESULT roleHr = policy->SetDefaultEndpoint(device.id.c_str(), role);
            if (FAILED(roleHr)) hr = roleHr;
        }
    }
    SafeRelease(&policy);
    return hr;
}

static HRESULT SetDeviceVisibility(const DeviceInfo& device, bool visible) {
    IPolicyConfig* policy = nullptr;
    HRESULT hr = GetPolicyConfig(&policy);
    if (SUCCEEDED(hr)) {
        hr = policy->SetEndpointVisibility(device.id.c_str(), visible ? TRUE : FALSE);
    }
    SafeRelease(&policy);
    return hr;
}

static DeviceInfo* SelectedDevice(HWND list, std::vector<DeviceInfo>& devices) {
    int index = ListView_GetNextItem(list, -1, LVNI_SELECTED);
    if (index < 0) {
        for (int i = 0; i < ListView_GetItemCount(list); ++i) {
            if (ListView_GetCheckState(list, i)) {
                index = i;
                break;
            }
        }
    }
    if (index < 0 || index >= static_cast<int>(devices.size())) return nullptr;
    return &devices[index];
}

static void AddCheckedFromList(HWND list, std::vector<DeviceInfo>& source, std::vector<DeviceInfo*>& selected) {
    for (int i = 0; i < ListView_GetItemCount(list); ++i) {
        if (ListView_GetCheckState(list, i) && i < static_cast<int>(source.size())) selected.push_back(&source[i]);
    }
}

static void AddHighlightedFromList(HWND list, std::vector<DeviceInfo>& source, std::vector<DeviceInfo*>& selected) {
    int index = -1;
    while ((index = ListView_GetNextItem(list, index, LVNI_SELECTED)) >= 0) {
        if (index < static_cast<int>(source.size())) selected.push_back(&source[index]);
    }
}

static void AddFirstHighlightedFromList(HWND list, std::vector<DeviceInfo>& source, std::vector<DeviceInfo*>& selected) {
    int index = ListView_GetNextItem(list, -1, LVNI_SELECTED);
    if (index >= 0 && index < static_cast<int>(source.size())) selected.push_back(&source[index]);
}

static std::vector<DeviceInfo*> SelectedDevices() {
    std::vector<DeviceInfo*> selected;
    AddCheckedFromList(g.outputs, g.outputDevices, selected);
    AddCheckedFromList(g.inputs, g.inputDevices, selected);
    if (selected.empty()) {
        AddHighlightedFromList(g.outputs, g.outputDevices, selected);
        AddHighlightedFromList(g.inputs, g.inputDevices, selected);
    }
    return selected;
}

static std::vector<DeviceInfo*> DisplayDevices() {
    std::vector<DeviceInfo*> display;
    AddFirstHighlightedFromList(g.outputs, g.outputDevices, display);
    AddFirstHighlightedFromList(g.inputs, g.inputDevices, display);
    if (!display.empty()) return display;
    AddCheckedFromList(g.outputs, g.outputDevices, display);
    AddCheckedFromList(g.inputs, g.inputDevices, display);
    return display;
}

static std::vector<DeviceInfo*> AllDevices() {
    std::vector<DeviceInfo*> devices;
    for (auto& d : g.outputDevices) devices.push_back(&d);
    for (auto& d : g.inputDevices) devices.push_back(&d);
    return devices;
}

static bool MmDevicePropertiesPath(const DeviceInfo& device, std::wstring* path);

static std::wstring GetExclusiveModeText(const DeviceInfo& device) {
    std::wstring path;
    if (!MmDevicePropertiesPath(device, &path)) return L"unavailable";
    HKEY key = nullptr;
    LONG result = RegOpenKeyExW(HKEY_LOCAL_MACHINE, path.c_str(), 0, KEY_QUERY_VALUE | KEY_WOW64_64KEY, &key);
    if (result != ERROR_SUCCESS) return L"unavailable";
    DWORD allow = 0, priority = 0, type = 0, bytes = sizeof(DWORD);
    bool hasAllow = RegQueryValueExW(key, L"{b3f8fa53-0004-438e-9003-51a46e139bfc},3", nullptr, &type, reinterpret_cast<BYTE*>(&allow), &bytes) == ERROR_SUCCESS;
    bytes = sizeof(DWORD);
    bool hasPriority = RegQueryValueExW(key, L"{b3f8fa53-0004-438e-9003-51a46e139bfc},4", nullptr, &type, reinterpret_cast<BYTE*>(&priority), &bytes) == ERROR_SUCCESS;
    RegCloseKey(key);
    if (!hasAllow && !hasPriority) return L"driver/default";
    return std::wstring(allow ? L"allowed" : L"blocked") + L", priority " + (priority ? L"on" : L"off");
}

static std::wstring GetLatencyText(const DeviceInfo& device) {
    IMMDeviceEnumerator* enumerator = nullptr;
    IMMDevice* endpoint = nullptr;
    IAudioClient* client = nullptr;
    REFERENCE_TIME defaultPeriod = 0;
    REFERENCE_TIME minimumPeriod = 0;
    HRESULT hr = CoCreateInstance(__uuidof(MMDeviceEnumerator), nullptr, CLSCTX_ALL,
        __uuidof(IMMDeviceEnumerator), reinterpret_cast<void**>(&enumerator));
    if (SUCCEEDED(hr)) hr = enumerator->GetDevice(device.id.c_str(), &endpoint);
    if (SUCCEEDED(hr)) hr = endpoint->Activate(__uuidof(IAudioClient), CLSCTX_ALL, nullptr, reinterpret_cast<void**>(&client));
    if (SUCCEEDED(hr)) hr = client->GetDevicePeriod(&defaultPeriod, &minimumPeriod);
    SafeRelease(&client);
    SafeRelease(&endpoint);
    SafeRelease(&enumerator);
    if (FAILED(hr)) return L"unavailable: " + HrText(hr);
    std::wstringstream ss;
    ss.setf(std::ios::fixed);
    ss.precision(2);
    ss << L"default " << (static_cast<double>(defaultPeriod) / 10000.0)
       << L" ms, minimum " << (static_cast<double>(minimumPeriod) / 10000.0) << L" ms";
    return ss.str();
}

static void SyncFormatControlsFromDevice(const DeviceInfo& device) {
    WAVEFORMATEX format{};
    if (!GetCurrentDeviceFormat(device, &format)) return;
    wchar_t buffer[32]{};
    StringCchPrintfW(buffer, 32, L"%u", format.nSamplesPerSec);
    SelectCombo(g.freq, buffer);
    StringCchPrintfW(buffer, 32, L"%u", format.wBitsPerSample);
    SelectCombo(g.bits, buffer);
    StringCchPrintfW(buffer, 32, L"%u", format.nChannels);
    SelectCombo(g.channels, buffer);
}

static void SyncVolumeControlsFromDevice(const DeviceInfo& device) {
    float volume = 0.0f;
    BOOL mute = FALSE;
    if (SUCCEEDED(GetEndpointVolumeState(device, &volume, &mute))) {
        SendMessageW(g.volume, TBM_SETPOS, TRUE, static_cast<LPARAM>(volume * 100.0f + 0.5f));
        SendMessageW(g.mute, BM_SETCHECK, mute ? BST_CHECKED : BST_UNCHECKED, 0);
    }
}

static void UpdateSelectedConfig() {
    if (!g.config) return;
    std::vector<DeviceInfo*> selected = DisplayDevices();
    std::vector<DeviceInfo*> targets = SelectedDevices();
    std::wstringstream ss;
    if (selected.empty()) {
        ss << L"Check one or more devices to target actions, or highlight a row to inspect current configuration.";
    } else {
        ss << selected.size() << L" displayed device(s); " << targets.size() << L" checked action target(s)\r\n";
        int shown = 0;
        for (DeviceInfo* device : selected) {
            if (shown >= 4) {
                ss << L"...more devices selected\r\n";
                break;
            }
            ss << (device->flow == eRender ? L"Output: " : L"Input: ") << device->name << L"\r\n";
            ss << L"  State: " << DeviceStateText(device->state) << L"\r\n";
            ss << L"  Device format: " << GetCurrentDeviceFormatText(*device) << L"\r\n";
            ss << L"  Mix format: " << GetMixFormatText(*device) << L"\r\n";
            ss << L"  Current latency: " << GetLatencyText(*device) << L"\r\n";
            ss << L"  Volume: " << GetVolumeText(*device) << L"\r\n";
            ss << L"  System effects/APO: " << GetSysFxText(*device) << L"\r\n";
            ss << L"  Exclusive mode: " << GetExclusiveModeText(*device) << L"\r\n";
            ++shown;
        }
        SyncFormatControlsFromDevice(*selected.front());
        SyncVolumeControlsFromDevice(*selected.front());
    }
    SetWindowTextW(g.config, ss.str().c_str());
}

static HRESULT SetEndpointVolume(const DeviceInfo& device, float volume, bool mute) {
    IMMDeviceEnumerator* enumerator = nullptr;
    IMMDevice* endpoint = nullptr;
    IAudioEndpointVolume* endpointVolume = nullptr;
    HRESULT hr = CoCreateInstance(__uuidof(MMDeviceEnumerator), nullptr, CLSCTX_ALL,
        __uuidof(IMMDeviceEnumerator), reinterpret_cast<void**>(&enumerator));
    if (SUCCEEDED(hr)) hr = enumerator->GetDevice(device.id.c_str(), &endpoint);
    if (SUCCEEDED(hr)) hr = endpoint->Activate(__uuidof(IAudioEndpointVolume), CLSCTX_ALL, nullptr, reinterpret_cast<void**>(&endpointVolume));
    if (SUCCEEDED(hr)) hr = endpointVolume->SetMasterVolumeLevelScalar(volume, nullptr);
    if (SUCCEEDED(hr)) hr = endpointVolume->SetMute(mute, nullptr);
    SafeRelease(&endpointVolume);
    SafeRelease(&endpoint);
    SafeRelease(&enumerator);
    return hr;
}

static HRESULT GetEndpointVolumeState(const DeviceInfo& device, float* volume, BOOL* mute) {
    IMMDeviceEnumerator* enumerator = nullptr;
    IMMDevice* endpoint = nullptr;
    IAudioEndpointVolume* endpointVolume = nullptr;
    HRESULT hr = CoCreateInstance(__uuidof(MMDeviceEnumerator), nullptr, CLSCTX_ALL,
        __uuidof(IMMDeviceEnumerator), reinterpret_cast<void**>(&enumerator));
    if (SUCCEEDED(hr)) hr = enumerator->GetDevice(device.id.c_str(), &endpoint);
    if (SUCCEEDED(hr)) hr = endpoint->Activate(__uuidof(IAudioEndpointVolume), CLSCTX_ALL, nullptr, reinterpret_cast<void**>(&endpointVolume));
    if (SUCCEEDED(hr) && volume) hr = endpointVolume->GetMasterVolumeLevelScalar(volume);
    if (SUCCEEDED(hr) && mute) hr = endpointVolume->GetMute(mute);
    SafeRelease(&endpointVolume);
    SafeRelease(&endpoint);
    SafeRelease(&enumerator);
    return hr;
}

static std::wstring GetVolumeText(const DeviceInfo& device) {
    float volume = 0.0f;
    BOOL mute = FALSE;
    HRESULT hr = GetEndpointVolumeState(device, &volume, &mute);
    if (FAILED(hr)) return L"unavailable: " + HrText(hr);
    std::wstringstream ss;
    ss << static_cast<int>(volume * 100.0f + 0.5f) << L"%";
    if (mute) ss << L" (muted)";
    return ss.str();
}

static void ApplyVolumeToSelected() {
    std::vector<DeviceInfo*> devices = SelectedDevices();
    if (devices.empty()) {
        Log(L"Select an input or output device before applying volume.");
        return;
    }
    int pos = static_cast<int>(SendMessageW(g.volume, TBM_GETPOS, 0, 0));
    bool mute = SendMessageW(g.mute, BM_GETCHECK, 0, 0) == BST_CHECKED;
    for (DeviceInfo* device : devices) {
        HRESULT hr = SetEndpointVolume(*device, pos / 100.0f, mute);
        Log(SUCCEEDED(hr) ? L"Volume/mute updated: " + device->name : L"Volume update failed for " + device->name + L": " + HrText(hr));
    }
    UpdateSelectedConfig();
}

static bool QueryService(SC_HANDLE service, SERVICE_STATUS_PROCESS* status) {
    DWORD needed = 0;
    return QueryServiceStatusEx(service, SC_STATUS_PROCESS_INFO, reinterpret_cast<BYTE*>(status), sizeof(*status), &needed) != FALSE;
}

static void LogServiceStatus(SC_HANDLE scm, const wchar_t* name) {
    SC_HANDLE service = OpenServiceW(scm, name, SERVICE_QUERY_STATUS);
    if (!service) {
        Log(std::wstring(L"Related service not accessible: ") + name + L" - " + LastErrorText());
        return;
    }
    SERVICE_STATUS_PROCESS status{};
    if (QueryService(service, &status)) {
        Log(std::wstring(L"Related service: ") + name + L" is " + ServiceStateText(status.dwCurrentState));
    }
    CloseServiceHandle(service);
}

static void LogDependentServices(SC_HANDLE service, const wchar_t* parentName) {
    DWORD bytesNeeded = 0;
    DWORD count = 0;
    EnumDependentServicesW(service, SERVICE_ACTIVE, nullptr, 0, &bytesNeeded, &count);
    if (GetLastError() == ERROR_MORE_DATA && bytesNeeded > 0) {
        std::vector<BYTE> buffer(bytesNeeded);
        auto deps = reinterpret_cast<ENUM_SERVICE_STATUSW*>(buffer.data());
        if (EnumDependentServicesW(service, SERVICE_ACTIVE, deps, bytesNeeded, &bytesNeeded, &count)) {
            for (DWORD i = 0; i < count; ++i) {
                Log(std::wstring(L"Active dependent of ") + parentName + L": " + deps[i].lpServiceName +
                    L" (" + deps[i].lpDisplayName + L")");
            }
            return;
        }
    }
    Log(std::wstring(L"Active dependent of ") + parentName + L": none");
}

static bool StopServiceWithWait(SC_HANDLE service, const wchar_t* name) {
    SERVICE_STATUS_PROCESS status{};
    if (!QueryService(service, &status)) {
        Log(std::wstring(L"Could not query service: ") + name + L" - " + LastErrorText());
        return false;
    }
    Log(std::wstring(L"Stopping service: ") + name + L" (current state: " + ServiceStateText(status.dwCurrentState) + L")");
    if (status.dwCurrentState == SERVICE_STOPPED) {
        Log(std::wstring(L"Already stopped: ") + name);
        return true;
    }

    SERVICE_STATUS ss{};
    if (!ControlService(service, SERVICE_CONTROL_STOP, &ss) && GetLastError() != ERROR_SERVICE_NOT_ACTIVE) {
        Log(std::wstring(L"Stop request failed for ") + name + L": " + LastErrorText());
        return false;
    }
    for (int i = 0; i < 40; ++i) {
        Sleep(250);
        if (!QueryService(service, &status)) break;
        if (status.dwCurrentState == SERVICE_STOPPED) {
            Log(std::wstring(L"Stopped service: ") + name);
            return true;
        }
    }
    Log(std::wstring(L"Service did not stop cleanly: ") + name);
    return false;
}

static bool StartServiceWithWait(SC_HANDLE service, const wchar_t* name) {
    SERVICE_STATUS_PROCESS status{};
    if (!QueryService(service, &status)) {
        Log(std::wstring(L"Could not query service: ") + name + L" - " + LastErrorText());
        return false;
    }
    Log(std::wstring(L"Starting service: ") + name + L" (current state: " + ServiceStateText(status.dwCurrentState) + L")");
    if (status.dwCurrentState == SERVICE_RUNNING) {
        Log(std::wstring(L"Already running: ") + name);
        return true;
    }

    if (!StartServiceW(service, 0, nullptr) && GetLastError() != ERROR_SERVICE_ALREADY_RUNNING) {
        Log(std::wstring(L"Start request failed for ") + name + L": " + LastErrorText());
        return false;
    }
    for (int i = 0; i < 40; ++i) {
        Sleep(250);
        if (!QueryService(service, &status)) break;
        if (status.dwCurrentState == SERVICE_RUNNING) {
            Log(std::wstring(L"Started service: ") + name);
            return true;
        }
    }
    Log(std::wstring(L"Service did not start cleanly: ") + name);
    return false;
}

static void ApplyHighPerformancePower();

static std::wstring ExeDirectory() {
    wchar_t path[MAX_PATH]{};
    GetModuleFileNameW(nullptr, path, MAX_PATH);
    std::wstring full = path;
    size_t slash = full.find_last_of(L"\\/");
    return slash == std::wstring::npos ? L"." : full.substr(0, slash);
}

static void SetServiceAutomaticAndRunning(SC_HANDLE scm, const wchar_t* name) {
    SC_HANDLE service = OpenServiceW(scm, name, SERVICE_CHANGE_CONFIG | SERVICE_START | SERVICE_QUERY_STATUS);
    if (!service) {
        Log(std::wstring(L"Could not open service for optimization: ") + name + L" - " + LastErrorText());
        return;
    }
    if (ChangeServiceConfigW(service, SERVICE_NO_CHANGE, SERVICE_AUTO_START, SERVICE_NO_CHANGE,
        nullptr, nullptr, nullptr, nullptr, nullptr, nullptr, nullptr)) {
        Log(std::wstring(L"Service startup set to Automatic: ") + name);
    } else {
        Log(std::wstring(L"Could not change startup type for ") + name + L": " + LastErrorText());
    }
    StartServiceWithWait(service, name);
    CloseServiceHandle(service);
}

static void OptimizeAudioServices() {
    if (!IsElevated()) {
        Log(L"Optimizing audio services requires administrator rights. Use Restart as admin.");
        return;
    }
    SC_HANDLE scm = OpenSCManagerW(nullptr, nullptr, SC_MANAGER_CONNECT);
    if (!scm) {
        Log(L"Could not open Service Control Manager: " + LastErrorText());
        return;
    }
    Log(L"Optimizing Windows audio services...");
    SetServiceAutomaticAndRunning(scm, L"AudioEndpointBuilder");
    SetServiceAutomaticAndRunning(scm, L"MMCSS");
    SetServiceAutomaticAndRunning(scm, L"Audiosrv");
    LogServiceStatus(scm, L"RpcSs");
    LogServiceStatus(scm, L"DcomLaunch");
    LogServiceStatus(scm, L"PlugPlay");
    Log(L"Audio service optimization complete.");
    CloseServiceHandle(scm);
}

static DWORD RunPowerCfg(const wchar_t* args) {
    wchar_t systemDir[MAX_PATH]{};
    GetSystemDirectoryW(systemDir, MAX_PATH);
    std::wstring exe = std::wstring(systemDir) + L"\\powercfg.exe";
    std::wstring cmd = L"\"" + exe + L"\" " + args;
    std::vector<wchar_t> mutableCmd(cmd.begin(), cmd.end());
    mutableCmd.push_back(L'\0');

    STARTUPINFOW si{};
    PROCESS_INFORMATION pi{};
    si.cb = sizeof(si);
    si.dwFlags = STARTF_USESHOWWINDOW;
    si.wShowWindow = SW_HIDE;
    if (!CreateProcessW(nullptr, mutableCmd.data(), nullptr, nullptr, FALSE, CREATE_NO_WINDOW, nullptr, nullptr, &si, &pi)) {
        Log(std::wstring(L"powercfg failed to start: ") + args + L" - " + LastErrorText());
        return static_cast<DWORD>(-1);
    }
    WaitForSingleObject(pi.hProcess, 30000);
    DWORD exitCode = 0;
    GetExitCodeProcess(pi.hProcess, &exitCode);
    CloseHandle(pi.hThread);
    CloseHandle(pi.hProcess);
    std::wstringstream ss;
    ss << L"powercfg " << args << L" -> exit " << exitCode;
    Log(ss.str());
    return exitCode;
}

static void ApplyLowLatencyPower() {
    ApplyHighPerformancePower();
    RunPowerCfg(L"/setacvalueindex SCHEME_CURRENT SUB_PROCESSOR PROCTHROTTLEMIN 100");
    RunPowerCfg(L"/setacvalueindex SCHEME_CURRENT SUB_PROCESSOR PROCTHROTTLEMAX 100");
    RunPowerCfg(L"/setdcvalueindex SCHEME_CURRENT SUB_PROCESSOR PROCTHROTTLEMIN 100");
    RunPowerCfg(L"/setdcvalueindex SCHEME_CURRENT SUB_PROCESSOR PROCTHROTTLEMAX 100");
    RunPowerCfg(L"/setacvalueindex SCHEME_CURRENT SUB_USB USBSELECTIVE 0");
    RunPowerCfg(L"/setdcvalueindex SCHEME_CURRENT SUB_USB USBSELECTIVE 0");
    RunPowerCfg(L"/setacvalueindex SCHEME_CURRENT SUB_PCIEXPRESS ASPM 0");
    RunPowerCfg(L"/setdcvalueindex SCHEME_CURRENT SUB_PCIEXPRESS ASPM 0");
    RunPowerCfg(L"/setactive SCHEME_CURRENT");
    Log(L"Low-latency power profile applied to the active plan.");
}

static void GeneratePowerReport() {
    std::wstring path = ExeDirectory() + L"\\audio-power-report.html";
    std::wstring report = L"\"" + path + L"\"";
    RunPowerCfg((std::wstring(L"/energy /duration 30 /output ") + report).c_str());
    Log(L"Power diagnostics requested: " + path);
}

static void RestartAudioServices() {
    if (!IsElevated()) {
        Log(L"Restarting audio services requires administrator rights. Use the Admin button.");
        return;
    }
    SC_HANDLE scm = OpenSCManagerW(nullptr, nullptr, SC_MANAGER_CONNECT);
    if (!scm) {
        Log(L"Could not open Service Control Manager.");
        return;
    }
    Log(L"Audio service repair plan:");
    for (const wchar_t* serviceName : kAudioServices) LogServiceStatus(scm, serviceName);

    SC_HANDLE audiosrv = OpenServiceW(scm, L"Audiosrv", SERVICE_STOP | SERVICE_START | SERVICE_QUERY_STATUS | SERVICE_ENUMERATE_DEPENDENTS);
    SC_HANDLE builder = OpenServiceW(scm, L"AudioEndpointBuilder", SERVICE_STOP | SERVICE_START | SERVICE_QUERY_STATUS | SERVICE_ENUMERATE_DEPENDENTS);
    SC_HANDLE mmcss = OpenServiceW(scm, L"MMCSS", SERVICE_START | SERVICE_QUERY_STATUS);
    if (!audiosrv || !builder) {
        Log(L"Could not open required Windows audio services.");
    } else {
        LogDependentServices(audiosrv, L"Audiosrv");
        LogDependentServices(builder, L"AudioEndpointBuilder");
        StopServiceWithWait(audiosrv, L"Audiosrv");
        StopServiceWithWait(builder, L"AudioEndpointBuilder");
        StartServiceWithWait(builder, L"AudioEndpointBuilder");
        if (mmcss) StartServiceWithWait(mmcss, L"MMCSS");
        StartServiceWithWait(audiosrv, L"Audiosrv");
        Log(L"Audio services restart complete. Related service states after restart:");
        for (const wchar_t* serviceName : kAudioServices) LogServiceStatus(scm, serviceName);
    }
    if (audiosrv) CloseServiceHandle(audiosrv);
    if (builder) CloseServiceHandle(builder);
    if (mmcss) CloseServiceHandle(mmcss);
    CloseServiceHandle(scm);
    RefreshDevices();
}

static bool SetRegDword(HKEY root, const wchar_t* path, const wchar_t* name, DWORD value) {
    HKEY key = nullptr;
    LONG result = RegCreateKeyExW(root, path, 0, nullptr, 0, KEY_SET_VALUE | KEY_WOW64_64KEY, nullptr, &key, nullptr);
    if (result != ERROR_SUCCESS) {
        Log(std::wstring(L"Registry open failed: ") + path + L" - " + HrText(HRESULT_FROM_WIN32(result)));
        return false;
    }
    result = RegSetValueExW(key, name, 0, REG_DWORD, reinterpret_cast<const BYTE*>(&value), sizeof(value));
    RegCloseKey(key);
    if (result != ERROR_SUCCESS) {
        Log(std::wstring(L"Registry write failed: ") + name + L" - " + HrText(HRESULT_FROM_WIN32(result)));
        return false;
    }
    std::wstringstream ss;
    ss << L"Registry DWORD set: " << path << L"\\" << name << L" = " << value;
    Log(ss.str());
    return true;
}

static std::wstring RootName(HKEY root) {
    if (root == HKEY_LOCAL_MACHINE) return L"HKLM";
    if (root == HKEY_CURRENT_USER) return L"HKCU";
    return L"HK";
}

static std::wstring BackupName(HKEY root, const wchar_t* path, const wchar_t* name) {
    std::wstring result = RootName(root) + L"|" + path + L"|" + name;
    std::replace(result.begin(), result.end(), L'\\', L'/');
    return result;
}

static void BackupRegValue(HKEY root, const wchar_t* path, const wchar_t* name) {
    HKEY key = nullptr;
    LONG result = RegOpenKeyExW(root, path, 0, KEY_QUERY_VALUE | KEY_WOW64_64KEY, &key);
    DWORD type = REG_NONE;
    DWORD bytes = 0;
    std::vector<BYTE> data;
    bool exists = false;
    if (result == ERROR_SUCCESS) {
        result = RegQueryValueExW(key, name, nullptr, &type, nullptr, &bytes);
        if (result == ERROR_SUCCESS && bytes > 0) {
            data.resize(bytes);
            result = RegQueryValueExW(key, name, nullptr, &type, data.data(), &bytes);
            exists = result == ERROR_SUCCESS;
        }
        RegCloseKey(key);
    }

    HKEY backup = nullptr;
    if (RegCreateKeyExW(HKEY_CURRENT_USER, L"Software\\AudioOptimizer\\Backup", 0, nullptr, 0, KEY_SET_VALUE, nullptr, &backup, nullptr) != ERROR_SUCCESS) {
        return;
    }
    std::wstring base = BackupName(root, path, name);
    DWORD existsDword = exists ? 1 : 0;
    RegSetValueExW(backup, (base + L"|exists").c_str(), 0, REG_DWORD, reinterpret_cast<const BYTE*>(&existsDword), sizeof(existsDword));
    RegSetValueExW(backup, (base + L"|type").c_str(), 0, REG_DWORD, reinterpret_cast<const BYTE*>(&type), sizeof(type));
    if (exists && !data.empty()) {
        RegSetValueExW(backup, (base + L"|data").c_str(), 0, REG_BINARY, data.data(), static_cast<DWORD>(data.size()));
    }
    RegCloseKey(backup);
}

static bool SetRegDwordBacked(HKEY root, const wchar_t* path, const wchar_t* name, DWORD value) {
    BackupRegValue(root, path, name);
    return SetRegDword(root, path, name, value);
}

static bool SetRegStringBacked(HKEY root, const wchar_t* path, const wchar_t* name, const wchar_t* value);

static bool SetRegString(HKEY root, const wchar_t* path, const wchar_t* name, const wchar_t* value) {
    HKEY key = nullptr;
    LONG result = RegCreateKeyExW(root, path, 0, nullptr, 0, KEY_SET_VALUE | KEY_WOW64_64KEY, nullptr, &key, nullptr);
    if (result != ERROR_SUCCESS) {
        Log(std::wstring(L"Registry open failed: ") + path + L" - " + HrText(HRESULT_FROM_WIN32(result)));
        return false;
    }
    DWORD bytes = static_cast<DWORD>((wcslen(value) + 1) * sizeof(wchar_t));
    result = RegSetValueExW(key, name, 0, REG_SZ, reinterpret_cast<const BYTE*>(value), bytes);
    RegCloseKey(key);
    if (result != ERROR_SUCCESS) {
        Log(std::wstring(L"Registry write failed: ") + name + L" - " + HrText(HRESULT_FROM_WIN32(result)));
        return false;
    }
    Log(std::wstring(L"Registry string set: ") + path + L"\\" + name + L" = " + value);
    return true;
}

static bool SetRegStringBacked(HKEY root, const wchar_t* path, const wchar_t* name, const wchar_t* value) {
    BackupRegValue(root, path, name);
    return SetRegString(root, path, name, value);
}

static bool RestoreRegValue(HKEY root, const wchar_t* path, const wchar_t* name) {
    HKEY backup = nullptr;
    if (RegOpenKeyExW(HKEY_CURRENT_USER, L"Software\\AudioOptimizer\\Backup", 0, KEY_QUERY_VALUE, &backup) != ERROR_SUCCESS) {
        Log(L"No registry backup found.");
        return false;
    }
    std::wstring base = BackupName(root, path, name);
    DWORD exists = 0, type = REG_NONE, size = sizeof(DWORD);
    bool ok = RegQueryValueExW(backup, (base + L"|exists").c_str(), nullptr, nullptr, reinterpret_cast<BYTE*>(&exists), &size) == ERROR_SUCCESS;
    size = sizeof(DWORD);
    ok = ok && RegQueryValueExW(backup, (base + L"|type").c_str(), nullptr, nullptr, reinterpret_cast<BYTE*>(&type), &size) == ERROR_SUCCESS;
    if (!ok) {
        RegCloseKey(backup);
        Log(std::wstring(L"No backup for: ") + path + L"\\" + name);
        return false;
    }

    HKEY key = nullptr;
    LONG result = RegCreateKeyExW(root, path, 0, nullptr, 0, KEY_SET_VALUE | KEY_WOW64_64KEY, nullptr, &key, nullptr);
    if (result != ERROR_SUCCESS) {
        RegCloseKey(backup);
        Log(std::wstring(L"Restore open failed: ") + path + L" - " + HrText(HRESULT_FROM_WIN32(result)));
        return false;
    }

    if (!exists) {
        result = RegDeleteValueW(key, name);
        if (result == ERROR_FILE_NOT_FOUND) result = ERROR_SUCCESS;
    } else {
        DWORD bytes = 0;
        result = RegQueryValueExW(backup, (base + L"|data").c_str(), nullptr, nullptr, nullptr, &bytes);
        std::vector<BYTE> data(bytes);
        if (result == ERROR_SUCCESS) {
            result = RegQueryValueExW(backup, (base + L"|data").c_str(), nullptr, nullptr, data.data(), &bytes);
            if (result == ERROR_SUCCESS) result = RegSetValueExW(key, name, 0, type, data.data(), bytes);
        }
    }
    RegCloseKey(key);
    RegCloseKey(backup);
    Log(result == ERROR_SUCCESS ? std::wstring(L"Restored registry value: ") + path + L"\\" + name
        : std::wstring(L"Restore failed: ") + path + L"\\" + name + L" - " + HrText(HRESULT_FROM_WIN32(result)));
    return result == ERROR_SUCCESS;
}

static void ApplyHighPerformancePower() {
    HRESULT hr = PowerSetActiveScheme(nullptr, &GUID_MIN_POWER_SAVINGS);
    if (SUCCEEDED(hr)) {
        Log(L"Power plan set to High performance for lower CPU latency.");
    } else {
        Log(L"Could not set High performance power plan: " + HrText(hr));
    }
}

static void RestoreBalancedPower() {
    HRESULT hr = PowerSetActiveScheme(nullptr, &GUID_TYPICAL_POWER_SAVINGS);
    if (SUCCEEDED(hr)) {
        Log(L"Power plan restored to Balanced.");
    } else {
        Log(L"Could not restore Balanced power plan: " + HrText(hr));
    }
}

static void ApplyProAudioTuning() {
    if (!Confirm(L"Apply Pro Audio Tuning",
        L"This changes Windows multimedia scheduling registry values, disables communications ducking for this user, and switches to High performance power. Continue?")) {
        Log(L"Pro audio tuning cancelled.");
        return;
    }

    SetRegDwordBacked(HKEY_CURRENT_USER, L"Software\\Microsoft\\Multimedia\\Audio", L"UserDuckingPreference", 3);

    if (!IsElevated()) {
        Log(L"Administrator rights required for MMCSS system registry tuning. Use Restart as admin, then run Pro tune again.");
        ApplyHighPerformancePower();
        return;
    }

    const wchar_t* profile = L"SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Multimedia\\SystemProfile";
    const wchar_t* tasks[] = {
        L"SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Multimedia\\SystemProfile\\Tasks\\Audio",
        L"SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Multimedia\\SystemProfile\\Tasks\\Capture",
        L"SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Multimedia\\SystemProfile\\Tasks\\Playback",
        L"SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Multimedia\\SystemProfile\\Tasks\\Pro Audio"
    };
    SetRegDwordBacked(HKEY_LOCAL_MACHINE, profile, L"SystemResponsiveness", 10);
    SetRegDwordBacked(HKEY_LOCAL_MACHINE, profile, L"NetworkThrottlingIndex", 0xffffffff);
    for (const wchar_t* task : tasks) {
        SetRegDwordBacked(HKEY_LOCAL_MACHINE, task, L"Priority", 6);
        SetRegDwordBacked(HKEY_LOCAL_MACHINE, task, L"GPU Priority", 8);
        SetRegDwordBacked(HKEY_LOCAL_MACHINE, task, L"Clock Rate", 10000);
        SetRegStringBacked(HKEY_LOCAL_MACHINE, task, L"Scheduling Category", L"High");
        SetRegStringBacked(HKEY_LOCAL_MACHINE, task, L"SFIO Priority", L"High");
        SetRegStringBacked(HKEY_LOCAL_MACHINE, task, L"Background Only", L"False");
    }
    ApplyLowLatencyPower();
    OptimizeAudioServices();
    Log(L"Pro audio tuning applied. Restart audio services or reboot for all apps/drivers to pick up changes.");
}

static void RestoreSafeTuning() {
    if (!Confirm(L"Restore Safe Tuning",
        L"This restores the app's main tuning knobs to conservative Windows-style values and switches power back to Balanced. Continue?")) {
        Log(L"Restore tuning cancelled.");
        return;
    }

    RestoreRegValue(HKEY_CURRENT_USER, L"Software\\Microsoft\\Multimedia\\Audio", L"UserDuckingPreference");
    if (IsElevated()) {
        const wchar_t* profile = L"SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Multimedia\\SystemProfile";
        const wchar_t* tasks[] = {
            L"SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Multimedia\\SystemProfile\\Tasks\\Audio",
            L"SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Multimedia\\SystemProfile\\Tasks\\Capture",
            L"SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Multimedia\\SystemProfile\\Tasks\\Playback",
            L"SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Multimedia\\SystemProfile\\Tasks\\Pro Audio"
        };
        RestoreRegValue(HKEY_LOCAL_MACHINE, profile, L"SystemResponsiveness");
        RestoreRegValue(HKEY_LOCAL_MACHINE, profile, L"NetworkThrottlingIndex");
        for (const wchar_t* task : tasks) {
            RestoreRegValue(HKEY_LOCAL_MACHINE, task, L"Priority");
            RestoreRegValue(HKEY_LOCAL_MACHINE, task, L"GPU Priority");
            RestoreRegValue(HKEY_LOCAL_MACHINE, task, L"Clock Rate");
            RestoreRegValue(HKEY_LOCAL_MACHINE, task, L"Scheduling Category");
            RestoreRegValue(HKEY_LOCAL_MACHINE, task, L"SFIO Priority");
            RestoreRegValue(HKEY_LOCAL_MACHINE, task, L"Background Only");
        }
    } else {
        Log(L"Administrator rights required to restore HKLM MMCSS values. User ducking and power plan restore will still run.");
    }
    RestoreBalancedPower();
    Log(L"Safe tuning restore complete.");
}

static void FixCracklingPreset() {
    SelectCombo(g.freq, L"48000");
    SelectCombo(g.bits, L"24");
    SelectCombo(g.channels, L"2");
    ApplyFormatToList(g.outputDevices, L"outputs");
    ApplyFormatToList(g.inputDevices, L"inputs");
    SetRegDwordBacked(HKEY_CURRENT_USER, L"Software\\Microsoft\\Multimedia\\Audio", L"UserDuckingPreference", 3);
    RestartAudioServices();
    Log(L"Crackling repair preset finished. Re-open apps that were playing or recording audio.");
}

static HRESULT SetEndpointSysFxRegistryFallback(const DeviceInfo& device, bool disabled);

static void SetSysFxForSelected(bool disabled) {
    std::vector<DeviceInfo*> devices = SelectedDevices();
    if (devices.empty()) {
        Log(L"Select one or more devices before changing system effects.");
        return;
    }
    for (DeviceInfo* device : devices) {
        HRESULT hr = SetEndpointSysFx(*device, disabled);
        bool usedFallback = false;
        if (FAILED(hr)) {
            Log(L"Driver rejected direct effects change for " + device->name + L": " + HrText(hr) + L". Trying registry fallback.");
            hr = SetEndpointSysFxRegistryFallback(*device, disabled);
            usedFallback = SUCCEEDED(hr);
        }
        if (SUCCEEDED(hr)) {
            Log(std::wstring(disabled ? L"System effects disabled for: " : L"System effects enabled for: ") + device->name +
                (usedFallback ? L" (registry fallback written; restart audio services or reboot if driver does not update immediately)" : L""));
        } else {
            Log(L"System effects are not supported or are driver-locked for " + device->name + L": " + HrText(hr));
        }
    }
    UpdateSelectedConfig();
}

static void SetVisibilityForSelected(bool visible) {
    std::vector<DeviceInfo*> devices = SelectedDevices();
    if (devices.empty()) {
        Log(L"Select one or more devices before changing visibility.");
        return;
    }
    for (DeviceInfo* device : devices) {
        HRESULT hr = SetDeviceVisibility(*device, visible);
        Log(SUCCEEDED(hr)
            ? std::wstring(visible ? L"Device shown/enabled in Sound panel: " : L"Device hidden/disabled in Sound panel: ") + device->name
            : std::wstring(L"Device visibility change failed for ") + device->name + L": " + HrText(hr));
    }
    RefreshDevices();
}

static bool MmDevicePropertiesPath(const DeviceInfo& device, std::wstring* path) {
    size_t pos = device.id.find_last_of(L'{');
    if (pos == std::wstring::npos) return false;
    std::wstring guid = device.id.substr(pos);
    const wchar_t* flow = device.flow == eRender ? L"Render" : L"Capture";
    *path = std::wstring(L"SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\MMDevices\\Audio\\") + flow + L"\\" + guid + L"\\Properties";
    return true;
}

static HRESULT SetEndpointSysFxRegistryFallback(const DeviceInfo& device, bool disabled) {
    std::wstring path;
    if (!MmDevicePropertiesPath(device, &path)) return HRESULT_FROM_WIN32(ERROR_PATH_NOT_FOUND);
    DWORD value = disabled ? ENDPOINT_SYSFX_DISABLED : ENDPOINT_SYSFX_ENABLED;
    return SetRegDwordBacked(HKEY_LOCAL_MACHINE, path.c_str(), L"{1da5d803-d492-4edd-8c23-e0c0ffee7f0e},5", value)
        ? S_OK : HRESULT_FROM_WIN32(GetLastError());
}

static void SetExclusiveModeForDevices(const std::vector<DeviceInfo*>& devices, bool enabled) {
    if (!IsElevated()) {
        Log(L"Exclusive-mode registry changes require administrator rights. Use Restart as admin.");
        return;
    }
    if (devices.empty()) {
        Log(L"Select devices first for per-device exclusive mode.");
        return;
    }
    for (DeviceInfo* device : devices) {
        std::wstring path;
        if (!MmDevicePropertiesPath(*device, &path)) {
            Log(L"Could not locate MMDevice registry path for: " + device->name);
            continue;
        }
        DWORD value = enabled ? 1 : 0;
        SetRegDwordBacked(HKEY_LOCAL_MACHINE, path.c_str(), L"{b3f8fa53-0004-438e-9003-51a46e139bfc},3", value);
        SetRegDwordBacked(HKEY_LOCAL_MACHINE, path.c_str(), L"{b3f8fa53-0004-438e-9003-51a46e139bfc},4", value);
        Log(std::wstring(enabled ? L"Exclusive mode enabled for: " : L"Exclusive mode disabled for: ") + device->name);
    }
    Log(L"Restart apps using audio. Some drivers require restarting audio services or rebooting.");
}

static void RestoreExclusiveModeForDevices(const std::vector<DeviceInfo*>& devices) {
    if (!IsElevated()) {
        Log(L"Exclusive-mode restore requires administrator rights. Use Restart as admin.");
        return;
    }
    for (DeviceInfo* device : devices) {
        std::wstring path;
        if (!MmDevicePropertiesPath(*device, &path)) continue;
        RestoreRegValue(HKEY_LOCAL_MACHINE, path.c_str(), L"{b3f8fa53-0004-438e-9003-51a46e139bfc},3");
        RestoreRegValue(HKEY_LOCAL_MACHINE, path.c_str(), L"{b3f8fa53-0004-438e-9003-51a46e139bfc},4");
    }
}

static void RestoreAppRegistryTweaks() {
    RestoreSafeTuning();
    RestoreRegValue(HKEY_CURRENT_USER, L"Control Panel\\Desktop", L"MenuShowDelay");
    Log(L"Requested restore of app-backed registry tweaks.");
}

static void ApplyUiDelayTweak() {
    SetRegStringBacked(HKEY_CURRENT_USER, L"Control Panel\\Desktop", L"MenuShowDelay", L"0");
    Log(L"Explorer/UI menu delay set to 0 for faster device menu response. Sign out or restart Explorer for full effect.");
}

static void RestoreUiDelayTweak() {
    RestoreRegValue(HKEY_CURRENT_USER, L"Control Panel\\Desktop", L"MenuShowDelay");
}

static HWND MakeButton(HWND parent, const wchar_t* text, int id, int x, int y, int w, int h) {
    HWND hwnd = CreateWindowW(L"BUTTON", text, WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON, x, y, w, h, parent, reinterpret_cast<HMENU>(static_cast<INT_PTR>(id)), nullptr, nullptr);
    TrackControl(hwnd);
    return hwnd;
}

static HWND MakeStatic(HWND parent, const wchar_t* text, int x, int y, int w, int h) {
    HWND hwnd = CreateWindowW(L"STATIC", text, WS_CHILD | WS_VISIBLE, x, y, w, h, parent, nullptr, nullptr, nullptr);
    TrackControl(hwnd);
    return hwnd;
}

static HWND MakeGroup(HWND parent, const wchar_t* text, int x, int y, int w, int h) {
    HWND hwnd = CreateWindowW(L"BUTTON", text, WS_CHILD | WS_VISIBLE | BS_GROUPBOX, x, y, w, h, parent, nullptr, nullptr, nullptr);
    TrackControl(hwnd);
    return hwnd;
}

static HWND MakeCombo(HWND parent, int id, int x, int y, int w, const wchar_t* const* values, int count, int selected) {
    HWND combo = CreateWindowW(WC_COMBOBOXW, L"", WS_CHILD | WS_VISIBLE | CBS_DROPDOWNLIST, x, y, w, 180, parent, reinterpret_cast<HMENU>(static_cast<INT_PTR>(id)), nullptr, nullptr);
    TrackControl(combo);
    for (int i = 0; i < count; ++i) SendMessageW(combo, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(values[i]));
    SendMessageW(combo, CB_SETCURSEL, selected, 0);
    return combo;
}

static HMENU MenuId(int id) {
    return reinterpret_cast<HMENU>(static_cast<INT_PTR>(id));
}

static void Layout(HWND hwnd) {
    RECT rc{};
    GetClientRect(hwnd, &rc);
    int width = rc.right - rc.left;
    int height = rc.bottom - rc.top;
    int margin = 12;
    int top = 68;
    int listW = (width - margin * 3) / 2;
    int listH = 220;

    MoveWindow(g.outputs, margin, top, listW, listH, TRUE);
    MoveWindow(g.inputs, margin * 2 + listW, top, listW, listH, TRUE);
    ListView_SetColumnWidth(g.outputs, 0, listW - 8);
    ListView_SetColumnWidth(g.inputs, 0, listW - 8);
    MoveWindow(g.config, margin, 340, width - margin * 2, 118, TRUE);
    MoveWindow(g.log, margin, 720, width - margin * 2, max(90, height - 732), TRUE);
}

static void InitControls(HWND hwnd) {
    g.tabs = CreateWindowW(WC_TABCONTROLW, L"", WS_CHILD | WS_VISIBLE, 8, 8, 1034, 28, hwnd, nullptr, nullptr, nullptr);
    TCITEMW item{};
    item.mask = TCIF_TEXT;
    item.pszText = const_cast<LPWSTR>(L"Devices");
    TabCtrl_InsertItem(g.tabs, 0, &item);
    item.pszText = const_cast<LPWSTR>(L"Advanced Tweaks");
    TabCtrl_InsertItem(g.tabs, 1, &item);

    g.createTab = 0;
    MakeStatic(hwnd, L"Outputs", 12, 46, 120, 20);
    MakeStatic(hwnd, L"Inputs", 540, 46, 120, 20);
    g.outputs = CreateWindowW(WC_LISTVIEWW, L"", WS_CHILD | WS_VISIBLE | WS_BORDER | LVS_REPORT | LVS_SINGLESEL | LVS_SHOWSELALWAYS | WS_VSCROLL, 12, 68, 500, 220, hwnd, MenuId(IDC_OUTPUTS), nullptr, nullptr);
    TrackControl(g.outputs);
    InitDeviceList(g.outputs);
    g.inputs = CreateWindowW(WC_LISTVIEWW, L"", WS_CHILD | WS_VISIBLE | WS_BORDER | LVS_REPORT | LVS_SINGLESEL | LVS_SHOWSELALWAYS | WS_VSCROLL, 528, 68, 500, 220, hwnd, MenuId(IDC_INPUTS), nullptr, nullptr);
    TrackControl(g.inputs);
    InitDeviceList(g.inputs);

    MakeButton(hwnd, L"Refresh devices", IDC_REFRESH, 12, 298, 120, 28);
    MakeButton(hwnd, L"Set default output", IDC_DEFAULT_OUT, 142, 298, 150, 28);
    MakeButton(hwnd, L"Set default input", IDC_DEFAULT_IN, 302, 298, 140, 28);
    MakeButton(hwnd, IsElevated() ? L"Running as admin" : L"Restart as admin", IDC_ADMIN, 452, 298, 150, 28);
    MakeButton(hwnd, L"Open sound panel", IDC_SOUND_PANEL, 612, 298, 140, 28);
    MakeButton(hwnd, L"Sound settings", IDC_SOUND_SETTINGS, 762, 298, 120, 28);
    MakeButton(hwnd, L"Refresh selected config", IDC_REFRESH_CONFIG, 892, 298, 136, 28);

    g.config = CreateWindowW(L"EDIT", L"Select one or more devices to see current configuration.", WS_CHILD | WS_VISIBLE | WS_BORDER | ES_MULTILINE | ES_AUTOVSCROLL | ES_READONLY | WS_VSCROLL, 12, 340, 1016, 118, hwnd, nullptr, nullptr, nullptr);
    TrackControl(g.config);

    MakeGroup(hwnd, L"Selected device format", 12, 468, 500, 92);
    MakeStatic(hwnd, L"Format", 28, 492, 70, 20);
    const wchar_t* rates[] = { L"44100", L"48000", L"88200", L"96000", L"192000" };
    const wchar_t* bits[] = { L"16", L"24", L"32" };
    const wchar_t* channels[] = { L"1", L"2" };
    g.freq = MakeCombo(hwnd, IDC_FREQ, 92, 488, 100, rates, 5, 1);
    g.bits = MakeCombo(hwnd, IDC_BITS, 202, 488, 80, bits, 3, 1);
    g.channels = MakeCombo(hwnd, IDC_CHANNELS, 292, 488, 80, channels, 2, 1);
    MakeButton(hwnd, L"Apply selected", IDC_APPLY_SELECTED, 28, 524, 120, 28);
    MakeButton(hwnd, L"Reset selected", IDC_RESET_SELECTED, 158, 524, 120, 28);
    MakeButton(hwnd, L"Apply outputs", IDC_APPLY_OUT, 288, 524, 100, 28);
    MakeButton(hwnd, L"Apply all", IDC_APPLY_ALL, 398, 524, 90, 28);

    MakeGroup(hwnd, L"Selected device volume / effects", 528, 468, 500, 92);
    g.volume = CreateWindowW(TRACKBAR_CLASSW, L"", WS_CHILD | WS_VISIBLE | TBS_AUTOTICKS, 548, 492, 230, 34, hwnd, MenuId(IDC_VOLUME), nullptr, nullptr);
    TrackControl(g.volume);
    SendMessageW(g.volume, TBM_SETRANGE, TRUE, MAKELPARAM(0, 100));
    SendMessageW(g.volume, TBM_SETPOS, TRUE, 80);
    g.mute = CreateWindowW(L"BUTTON", L"Mute", WS_CHILD | WS_VISIBLE | BS_AUTOCHECKBOX, 788, 496, 70, 24, hwnd, MenuId(IDC_MUTE), nullptr, nullptr);
    TrackControl(g.mute);
    MakeButton(hwnd, L"Apply volume", IDC_APPLY_VOLUME, 868, 492, 105, 30);
    MakeButton(hwnd, L"Disable effects", IDC_DISABLE_EFFECTS, 548, 524, 120, 28);
    MakeButton(hwnd, L"Enable effects", IDC_ENABLE_EFFECTS, 678, 524, 110, 28);
    MakeButton(hwnd, L"Hide selected", IDC_HIDE_SELECTED_DEVICES, 798, 524, 105, 28);
    MakeButton(hwnd, L"Show selected", IDC_SHOW_SELECTED_DEVICES, 913, 524, 105, 28);

    MakeGroup(hwnd, L"Quick repair", 12, 568, 500, 82);
    MakeButton(hwnd, L"1-click 48 kHz / 24-bit", IDC_PRESET_48, 28, 596, 180, 30);
    MakeButton(hwnd, L"44.1 kHz / 16-bit", IDC_PRESET_44, 218, 596, 145, 30);
    MakeButton(hwnd, L"Fix crackling", IDC_FIX_CRACKLE, 373, 596, 120, 30);

    MakeGroup(hwnd, L"System optimization", 528, 568, 500, 112);
    MakeButton(hwnd, L"Restart audio services", IDC_RESTART_AUDIO, 548, 596, 160, 30);
    MakeButton(hwnd, L"Optimize services", IDC_OPTIMIZE_SERVICES, 718, 596, 140, 30);
    MakeButton(hwnd, L"Low latency power", IDC_LOW_LATENCY_POWER, 868, 596, 140, 30);
    MakeButton(hwnd, L"Pro audio tune", IDC_PRO_TUNE, 548, 638, 140, 30);
    MakeButton(hwnd, L"Restore safe tuning", IDC_RESTORE_TUNE, 698, 638, 150, 30);
    MakeButton(hwnd, L"Power report", IDC_POWER_REPORT, 858, 638, 120, 30);

    g.log = CreateWindowW(L"EDIT", L"", WS_CHILD | WS_VISIBLE | WS_BORDER | ES_MULTILINE | ES_AUTOVSCROLL | ES_READONLY | WS_VSCROLL, 12, 720, 1016, 150, hwnd, MenuId(IDC_LOG), nullptr, nullptr);
    g.commonControls.push_back(g.log);

    g.createTab = 1;
    MakeGroup(hwnd, L"MMCSS / Multimedia Scheduler", 12, 50, 500, 140);
    MakeButton(hwnd, L"Apply MMCSS low latency", IDC_MMCSS_APPLY, 32, 82, 180, 30);
    MakeButton(hwnd, L"Restore MMCSS backup", IDC_MMCSS_RESTORE, 222, 82, 170, 30);
    MakeStatic(hwnd, L"Applies SystemResponsiveness, NetworkThrottlingIndex, Audio/Capture/Playback/Pro Audio priority, Clock Rate, Scheduling Category, SFIO Priority.", 32, 122, 450, 46);

    MakeGroup(hwnd, L"Ducking / UI", 528, 50, 500, 140);
    MakeButton(hwnd, L"Disable ducking", IDC_DUCKING_OFF, 548, 82, 140, 30);
    MakeButton(hwnd, L"Restore ducking", IDC_DUCKING_RESTORE, 698, 82, 130, 30);
    MakeButton(hwnd, L"Menu delay 0", IDC_UI_DELAY_APPLY, 548, 122, 130, 30);
    MakeButton(hwnd, L"Restore menu delay", IDC_UI_DELAY_RESTORE, 688, 122, 150, 30);

    MakeGroup(hwnd, L"Exclusive Mode", 12, 210, 500, 155);
    MakeButton(hwnd, L"Enable selected", IDC_EXCLUSIVE_ON_SELECTED, 32, 242, 140, 30);
    MakeButton(hwnd, L"Disable selected", IDC_EXCLUSIVE_OFF_SELECTED, 182, 242, 140, 30);
    MakeButton(hwnd, L"Restore selected", IDC_EXCLUSIVE_RESTORE_SELECTED, 332, 242, 140, 30);
    MakeButton(hwnd, L"Enable all devices", IDC_EXCLUSIVE_ON_ALL, 32, 286, 150, 30);
    MakeButton(hwnd, L"Disable all devices", IDC_EXCLUSIVE_OFF_ALL, 192, 286, 150, 30);
    MakeStatic(hwnd, L"Exclusive mode can reduce DAW/WASAPI latency, but may block other apps. Disable it for shared stability.", 32, 326, 440, 28);

    MakeGroup(hwnd, L"Power / USB / Services", 528, 210, 500, 155);
    MakeButton(hwnd, L"Low latency power", IDC_LOW_LATENCY_POWER, 548, 242, 150, 30);
    MakeButton(hwnd, L"Optimize services", IDC_OPTIMIZE_SERVICES, 708, 242, 140, 30);
    MakeButton(hwnd, L"Restart audio services", IDC_RESTART_AUDIO, 858, 242, 150, 30);
    MakeButton(hwnd, L"Power report", IDC_POWER_REPORT, 548, 286, 130, 30);
    MakeButton(hwnd, L"Restore all backed tweaks", IDC_RESTORE_ALL_TWEAKS, 688, 286, 190, 30);
    MakeStatic(hwnd, L"USB root hub power boxes and Bluetooth codec/HFP options are driver/device-manager settings. Use report + Sound panels for those.", 548, 326, 440, 28);

    MakeGroup(hwnd, L"Quality helpers", 12, 385, 1016, 100);
    MakeButton(hwnd, L"Open classic Sound panel", IDC_SOUND_PANEL, 32, 422, 170, 30);
    MakeButton(hwnd, L"Open Windows Sound settings", IDC_SOUND_SETTINGS, 212, 422, 190, 30);
    MakeButton(hwnd, L"Pro audio tune", IDC_PRO_TUNE, 412, 422, 140, 30);
    MakeButton(hwnd, L"Restore safe tuning", IDC_RESTORE_TUNE, 562, 422, 160, 30);
    MakeStatic(hwnd, L"Driver-specific items such as Loudness Equalization, Dolby/DTS, Realtek jack detection, Bluetooth AAC/HFP, and speaker layout must be exposed by the installed driver.", 32, 458, 940, 20);

    g.createTab = 0;
    ShowTab(0);
}

static void HandleCommand(HWND, WPARAM wParam) {
    switch (LOWORD(wParam)) {
    case IDC_REFRESH: RefreshDevices(); break;
    case IDC_REFRESH_CONFIG: UpdateSelectedConfig(); break;
    case IDC_ADMIN: if (!IsElevated()) RelaunchElevated(); break;
    case IDC_DEFAULT_OUT:
        if (DeviceInfo* d = SelectedDevice(g.outputs, g.outputDevices)) {
            HRESULT hr = SetDefaultDevice(*d);
            Log(SUCCEEDED(hr) ? L"Default output set: " + d->name : L"Default output failed: " + HrText(hr));
        } else Log(L"Select an output device first.");
        break;
    case IDC_DEFAULT_IN:
        if (DeviceInfo* d = SelectedDevice(g.inputs, g.inputDevices)) {
            HRESULT hr = SetDefaultDevice(*d);
            Log(SUCCEEDED(hr) ? L"Default input set: " + d->name : L"Default input failed: " + HrText(hr));
        } else Log(L"Select an input device first.");
        break;
    case IDC_APPLY_OUT: ApplyFormatToList(g.outputDevices, L"outputs"); break;
    case IDC_APPLY_IN: ApplyFormatToList(g.inputDevices, L"inputs"); break;
    case IDC_APPLY_SELECTED:
    {
        std::vector<DeviceInfo*> devices = SelectedDevices();
        if (devices.empty()) {
            Log(L"Select one or more output/input devices first.");
        } else {
            for (DeviceInfo* d : devices) ApplyFormatToDevice(*d);
            UpdateSelectedConfig();
        }
        break;
    }
    case IDC_RESET_SELECTED:
    {
        std::vector<DeviceInfo*> devices = SelectedDevices();
        if (!devices.empty()) {
            for (DeviceInfo* d : devices) {
                HRESULT hr = ResetDeviceFormat(*d);
                Log(SUCCEEDED(hr) ? L"Device format reset to driver default: " + d->name : L"Reset selected format failed for " + d->name + L": " + HrText(hr));
            }
            UpdateSelectedConfig();
        } else {
            Log(L"Select an output or input device first.");
        }
        break;
    }
    case IDC_APPLY_ALL:
        ApplyFormatToList(g.outputDevices, L"outputs");
        ApplyFormatToList(g.inputDevices, L"inputs");
        break;
    case IDC_PRESET_48:
        SelectCombo(g.freq, L"48000"); SelectCombo(g.bits, L"24"); SelectCombo(g.channels, L"2");
        ApplyFormatToList(g.outputDevices, L"outputs");
        ApplyFormatToList(g.inputDevices, L"inputs");
        break;
    case IDC_PRESET_44:
        SelectCombo(g.freq, L"44100"); SelectCombo(g.bits, L"16"); SelectCombo(g.channels, L"2");
        ApplyFormatToList(g.outputDevices, L"outputs");
        ApplyFormatToList(g.inputDevices, L"inputs");
        break;
    case IDC_APPLY_VOLUME: ApplyVolumeToSelected(); break;
    case IDC_RESTART_AUDIO: RestartAudioServices(); break;
    case IDC_FIX_CRACKLE: FixCracklingPreset(); break;
    case IDC_PRO_TUNE: ApplyProAudioTuning(); break;
    case IDC_RESTORE_TUNE: RestoreSafeTuning(); break;
    case IDC_OPTIMIZE_SERVICES: OptimizeAudioServices(); break;
    case IDC_LOW_LATENCY_POWER: ApplyLowLatencyPower(); break;
    case IDC_POWER_REPORT: GeneratePowerReport(); break;
    case IDC_DISABLE_EFFECTS: SetSysFxForSelected(true); break;
    case IDC_ENABLE_EFFECTS: SetSysFxForSelected(false); break;
    case IDC_HIDE_SELECTED_DEVICES: SetVisibilityForSelected(false); break;
    case IDC_SHOW_SELECTED_DEVICES: SetVisibilityForSelected(true); break;
    case IDC_EXCLUSIVE_ON_SELECTED: SetExclusiveModeForDevices(SelectedDevices(), true); break;
    case IDC_EXCLUSIVE_OFF_SELECTED: SetExclusiveModeForDevices(SelectedDevices(), false); break;
    case IDC_EXCLUSIVE_RESTORE_SELECTED: RestoreExclusiveModeForDevices(SelectedDevices()); break;
    case IDC_EXCLUSIVE_ON_ALL: SetExclusiveModeForDevices(AllDevices(), true); break;
    case IDC_EXCLUSIVE_OFF_ALL: SetExclusiveModeForDevices(AllDevices(), false); break;
    case IDC_DUCKING_OFF: SetRegDwordBacked(HKEY_CURRENT_USER, L"Software\\Microsoft\\Multimedia\\Audio", L"UserDuckingPreference", 3); break;
    case IDC_DUCKING_RESTORE: RestoreRegValue(HKEY_CURRENT_USER, L"Software\\Microsoft\\Multimedia\\Audio", L"UserDuckingPreference"); break;
    case IDC_MMCSS_APPLY: ApplyProAudioTuning(); break;
    case IDC_MMCSS_RESTORE: RestoreSafeTuning(); break;
    case IDC_UI_DELAY_APPLY: ApplyUiDelayTweak(); break;
    case IDC_UI_DELAY_RESTORE: RestoreUiDelayTweak(); break;
    case IDC_RESTORE_ALL_TWEAKS: RestoreAppRegistryTweaks(); break;
    case IDC_SOUND_PANEL:
        ShellExecuteW(g.hwnd, L"open", L"mmsys.cpl", nullptr, nullptr, SW_SHOWNORMAL);
        break;
    case IDC_SOUND_SETTINGS:
        ShellExecuteW(g.hwnd, L"open", L"ms-settings:sound", nullptr, nullptr, SW_SHOWNORMAL);
        break;
    }
}

static LRESULT CALLBACK WindowProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    switch (msg) {
    case WM_CREATE:
        g.hwnd = hwnd;
        InitControls(hwnd);
        RefreshDevices();
        return 0;
    case WM_SIZE:
        Layout(hwnd);
        return 0;
    case WM_GETMINMAXINFO:
        reinterpret_cast<MINMAXINFO*>(lParam)->ptMinTrackSize.x = 1060;
        reinterpret_cast<MINMAXINFO*>(lParam)->ptMinTrackSize.y = 900;
        return 0;
    case WM_COMMAND:
        HandleCommand(hwnd, wParam);
        return 0;
    case WM_NOTIFY:
        if (reinterpret_cast<NMHDR*>(lParam)->hwndFrom == g.tabs &&
            reinterpret_cast<NMHDR*>(lParam)->code == TCN_SELCHANGE) {
            ShowTab(TabCtrl_GetCurSel(g.tabs));
            return 0;
        }
        if ((reinterpret_cast<NMHDR*>(lParam)->hwndFrom == g.outputs ||
            reinterpret_cast<NMHDR*>(lParam)->hwndFrom == g.inputs) &&
            reinterpret_cast<NMHDR*>(lParam)->code == LVN_ITEMCHANGED) {
            auto changed = reinterpret_cast<NMLISTVIEW*>(lParam);
            UINT interesting = LVIS_SELECTED | LVIS_STATEIMAGEMASK;
            if ((changed->uChanged & LVIF_STATE) && ((changed->uOldState ^ changed->uNewState) & interesting)) {
                UpdateSelectedConfig();
            }
            return 0;
        }
        break;
    case WM_DESTROY:
        PostQuitMessage(0);
        return 0;
    }
    return DefWindowProcW(hwnd, msg, wParam, lParam);
}

int APIENTRY wWinMain(HINSTANCE instance, HINSTANCE, LPWSTR, int show) {
    CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);

    INITCOMMONCONTROLSEX icc{ sizeof(icc), ICC_BAR_CLASSES | ICC_STANDARD_CLASSES };
    InitCommonControlsEx(&icc);

    WNDCLASSW wc{};
    wc.lpfnWndProc = WindowProc;
    wc.hInstance = instance;
    wc.hCursor = LoadCursorW(nullptr, IDC_ARROW);
    wc.hIcon = LoadIconW(nullptr, IDI_APPLICATION);
    wc.hbrBackground = reinterpret_cast<HBRUSH>(COLOR_WINDOW + 1);
    wc.lpszClassName = L"AudioOptimizerNativeWindow";
    RegisterClassW(&wc);

    HWND hwnd = CreateWindowExW(0, wc.lpszClassName, L"Audio Optimizer - Native Windows Audio Control",
        WS_OVERLAPPEDWINDOW, CW_USEDEFAULT, CW_USEDEFAULT, 1080, 900,
        nullptr, nullptr, instance, nullptr);
    ShowWindow(hwnd, show);
    UpdateWindow(hwnd);

    MSG msg{};
    while (GetMessageW(&msg, nullptr, 0, 0)) {
        TranslateMessage(&msg);
        DispatchMessageW(&msg);
    }

    CoUninitialize();
    return static_cast<int>(msg.wParam);
}
