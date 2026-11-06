
#include "AudioManager.h"
#include "ConstantString.h"
#include "ASAudio.h"      // 引入底層音訊操作類別
#include <Windows.h>
#include <sstream>      // 用於組合字串
#include <shellapi.h>   // 為了使用 ShellExecute 來打開音效設定面板
#include <comdef.h>

extern short leftAudioData[BUFFER_SIZE];
extern short rightAudioData[BUFFER_SIZE];
extern double leftSpectrumData[BUFFER_SIZE];
extern double rightSpectrumData[BUFFER_SIZE];

extern short leftAudioData_SNR[BUFFER_SIZE];
extern short rightAudioData_SNR[BUFFER_SIZE];
extern double leftSpectrumData_SNR[BUFFER_SIZE];
extern double rightSpectrumData_SNR[BUFFER_SIZE];

extern double thd_n[2];
extern double FundamentalLevel_dBFS[2];
extern double freq[2];
extern double snr_result[2];
const double NOT_TESTED_VALUE = -999.0;

extern "C" int fft_thd_n_exe(short*, short*, double*, double*);
extern "C" void fft_get_thd_n_db(double*, double*, double*);
extern "C" int fft_mute_exe(bool, bool, short*, short*, double*, double*);
extern "C" void fft_get_snr(double*);


class ComInitializer {
public:
    ComInitializer() : hr(CoInitializeEx(NULL, COINIT_APARTMENTTHREADED)) {
    }
    ~ComInitializer() {
        // 只在成功初始化後才進行反初始化
        if (SUCCEEDED(hr)) {
            CoUninitialize();
        }
    }
    // 讓外部可以檢查初始化是否成功
    HRESULT GetHResult() const {
        return hr;
    }
private:
    HRESULT hr; // 保存 CoInitializeEx 的回傳值
};


AudioManager::AudioManager() {
    memset(leftAudioData, 0, sizeof(leftAudioData));
    memset(rightAudioData, 0, sizeof(rightAudioData));
    memset(leftSpectrumData, 0, sizeof(leftSpectrumData));
    memset(rightSpectrumData, 0, sizeof(rightSpectrumData));
    memset(thd_n, 0, sizeof(thd_n));
    memset(FundamentalLevel_dBFS, 0, sizeof(FundamentalLevel_dBFS));
    memset(freq, 0, sizeof(freq));
}
const std::string& AudioManager::GetResultString() const {
    return resultString;
}

bool AudioManager::ExecuteTestsFromConfig(const Config& config) {

    // 2. 依序執行設定檔中啟用的各項測試
    if (config.AudioTestEnable) {
        ComInitializer comInit;
        if (FAILED(comInit.GetHResult())) {
            wprintf(L"COM 初始化失敗，直接中止 \nCOM Fail of initialization, stop directly \n");
            resultString = "LOG:ERROR_AUDIO_OTHER#";
            return false; // 如果 COM 初始化失敗，直接中止
        }

        // 重置所有用於測試的資料緩衝區
        memset(leftAudioData, 0, sizeof(leftAudioData));
        memset(rightAudioData, 0, sizeof(rightAudioData));
        memset(leftSpectrumData, 0, sizeof(leftSpectrumData));
        memset(rightSpectrumData, 0, sizeof(rightSpectrumData));
        memset(thd_n, 0, sizeof(thd_n));
        memset(FundamentalLevel_dBFS, 0, sizeof(FundamentalLevel_dBFS));
        memset(freq, 0, sizeof(freq));
        memset(snr_result, 0, sizeof(snr_result));


        thd_n[0] = NOT_TESTED_VALUE;
        thd_n[1] = NOT_TESTED_VALUE;
        FundamentalLevel_dBFS[0] = NOT_TESTED_VALUE;
        FundamentalLevel_dBFS[1] = NOT_TESTED_VALUE;
        freq[0] = NOT_TESTED_VALUE;
        freq[1] = NOT_TESTED_VALUE;
        snr_result[0] = NOT_TESTED_VALUE;
        snr_result[1] = NOT_TESTED_VALUE;
        resultString.clear(); // 每次執行前都清除上一次的結果

        // 1. 檢查裝置是否存在
        if (!CheckAudioDevices(config)) {
            return false; // 裝置檢查失敗，流程中止
        }
        if (!RunAudioTest(config)) return false;
        if (config.snrTestEnable) {
            // SNR 測試必須在 AudioTest (THD+N 測試) 之後執行，
            // 因為它需要 AudioTest 測得的訊號強度。
            if (!config.AudioTestEnable) {
				wprintf(L"SNR 測試需要先啟用 AudioTest，請在設定檔中將 AudioTestEnable 設為 true。\nSNR test requires AudioTest to be enabled, please set AudioTestEnable to true in the configuration file.\n");
                resultString = "LOG:ERROR_AUDIO_OTHER#";
                return false;
            }
            if (!RunSnrTest(config)) return false;
        }
    }
    if (config.PlayWAVFileEnable || config.CloseWAVFileEnable) {
        if (!RunWavPlayback(config)) return false;
    }

    if (config.AudioLoopBackEnable) {
        if (!RunAudioLoopback(config)) return false;
    }

    if (config.SwitchDefaultAudioEnable) {
        if (!RunSwitchDefaultDevice(config)) return false;
    }
	if (config.SetListenToThisDeviceEnable) {
        if (!RunSetListenToThisDevice(config)) return false;
	}

    // 如果所有流程都跑完了，而且 resultString 還是空的，
    // 表示所有功能都被 disable，給一個預設的成功訊息。
    if (resultString.empty()) {
        resultString = "LOG:PASS#";
    }

    return true; // 所有啟用的測試流程都已執行完畢
}

bool AudioManager::CheckAudioDevices(const Config& config) {
    ASAudio& audio = ASAudio::GetInstance();

    if (!config.outDeviceName.empty()) {
        ULONGLONG timeout = GetTickCount64() + 5000; // 5秒超時
        while (audio.GetWaveOutDevice(audio.charToWstring(config.outDeviceName.c_str())) == -1) {
            if (GetTickCount64() > timeout) {
                resultString = "LOG:ERROR_AUDIO_WAVEOUT_DEVICE_NOT_FIND#";
                return false;
            }
            Sleep(100);
        }
        audio.SetSpeakerSystemVolume();
    }

    if (!config.inDeviceName.empty()) {
        ULONGLONG timeout = GetTickCount64() + 5000;
        while (audio.GetWaveInDevice(audio.charToWstring(config.inDeviceName.c_str())) == -1) {
            if (GetTickCount64() > timeout) {
                resultString = "LOG:ERROR_AUDIO_WAVEIN_DEVICE_NOT_FIND#";
                return false;
            }
            Sleep(100);
        }
        audio.SetMicSystemVolume();
    }
    return true;
}

bool AudioManager::RunAudioTest(const Config& config) {
    // 呼叫 C-API 執行核心 FFT 測試
    fft_thd_n_exe(leftAudioData, rightAudioData, leftSpectrumData, rightSpectrumData);
    fft_get_thd_n_db(thd_n, FundamentalLevel_dBFS, freq);

    double levelDifference = std::abs(FundamentalLevel_dBFS[0] - FundamentalLevel_dBFS[1]);

    // 根據設定檔的標準，判斷測試結果
    if (thd_n[0] > config.thd_n || thd_n[1] > config.thd_n) {
        std::stringstream ss;
        ss << "LOG:ERROR_AUDIO_THD_N#";
        resultString = ss.str();
        return false;
    }
    else if (FundamentalLevel_dBFS[0] < config.FundamentalLevel_dBFS || FundamentalLevel_dBFS[1] < config.FundamentalLevel_dBFS) {
        std::stringstream ss;
        ss << "LOG:ERROR_AUDIO_FUNDAMENTAL_LEVEL#";
        resultString = ss.str();
        return false;
    }
    else if (levelDifference > config.maxLevelDifference_dB) {
        std::stringstream ss;
        ss << "LOG:ERROR_AUDIO_LEVEL_IMBALANCE#";
        resultString = ss.str();
        return false;
    }
    else if (freq[0] != config.frequencyL || freq[1] != config.frequencyR) {
        std::stringstream ss;
        ss << "LOG:ERROR_AUDIO_FREQUENCY#";
        resultString = ss.str();
        return false;
    }
    else {
        std::stringstream ss;
        ss << "LOG:PASS#";
        resultString = ss.str();
        return true;
    }
}

bool AudioManager::RunWavPlayback(const Config& config) {
    ASAudio& audio = ASAudio::GetInstance();
    if (config.PlayWAVFileEnable) {
        if (!audio.PlayWavFile(config.AutoCloseWAVFile)) {
            // 從 ASAudio 實例中獲取詳細的錯誤訊息
            resultString = audio.GetLastResult();
            return false;
        }
    }

    if (config.CloseWAVFileEnable) {
        audio.StopPlayingWavFile();
    }
    return true;
}

bool AudioManager::RunAudioLoopback(const Config& config) {
    ASAudio& audio = ASAudio::GetInstance();
    if (config.AudioLoopBackStart) {
        ShellExecute(0, L"open", L"control mmsys.cpl,,1", 0, 0, SW_SHOWNORMAL);
        if (!audio.StartAudioLoopback(audio.GetWaveParm().WaveInDev.c_str(), audio.GetWaveParm().WaveOutDev.c_str())) {
            resultString = audio.GetLastResult();
            return false;
        }
    }
    else {
        audio.StopAudioLoopback();
    }
    return true;
}

bool AudioManager::RunSwitchDefaultDevice(const Config& config) {
    ASAudio& audio = ASAudio::GetInstance();
    std::wstring deviceId;

    Config nonConstConfig = config;

    if (audio.FindDeviceIdByName(nonConstConfig, deviceId, eRender)) {
        if (!audio.SetDefaultAudioPlaybackDevice(deviceId)) {
            wprintf(L"切換預設播放裝置失敗！錯誤碼: 0x%08X\n Change default playback device failed! Error code: 0x%08X\n", audio.GetLastResult());
            resultString = "LOG:ERROR_AUDIO_CHANGE_DEFAULT_DEVICE#";
            return false;
        }
    }
    else {
        resultString = audio.GetLastResult();
        return false;
    }
    return true;
}
bool AudioManager::RunSetListenToThisDevice(const Config& config) {
    ASAudio& audio = ASAudio::GetInstance();
    std::wstring deviceId;

    Config tempConfig = config;
    tempConfig.AudioName = config.AudioName;

    if (audio.FindDeviceIdByName(tempConfig, deviceId, eCapture)) {
        if (!audio.SetListenToThisDevice(deviceId, config.setListen)) {
            wprintf(L"設定監聽裝置失敗！錯誤碼: 0x%08X\n Set listen device failed! Error code: 0x%08X\n", audio.GetLastResult());
            resultString = "LOG:ERROR_AUDIO_SET_LISTEN_FAILED#";
            return false;
        }
    }
    else {
        resultString = audio.GetLastResult(); // 獲取如「找不到裝置」的錯誤
        return false;
    }
    return true;
}
bool AudioManager::RunSnrTest(const Config& config) {
    // 執行靜音錄製以獲取噪聲數據
    // MuteWaveOut=true, MuteWaveIn=false (靜音播放，正常錄音)
    fft_mute_exe(true, false, leftAudioData_SNR, rightAudioData_SNR, leftSpectrumData_SNR, rightSpectrumData_SNR);
    snr_result[0] = snr_result[1] = 0;
    fft_get_snr(snr_result);

    // 判斷測試結果
    if (snr_result[0] < config.snrThreshold || snr_result[1] < config.snrThreshold) {
        std::stringstream ss;
        // 產生失敗的 Log
        wprintf(L"SNR 測試失敗！左聲道: %.2f dB, 右聲道: %.2f dB, 標準: %.2f dB\nSNR test failed! Left: %.2f dB, Right: %.2f dB, Threshold: %.2f dB\n",
            snr_result[0], snr_result[1], config.snrThreshold,
			snr_result[0], snr_result[1], config.snrThreshold);
        ss << "LOG:ERROR_AUDIO_SNR#";
        resultString = ss.str();
        return false;
    }
    else {
        return true;
    }
}