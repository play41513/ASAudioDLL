#include "pch.h"
#include "AudioManager.h"
#include "ConstantString.h"
#include "ASAudio.h"      // 引入底層音訊操作類別
#include <Windows.h>
#include <sstream>      // 用於組合字串
#include <shellapi.h>   // 為了使用 ShellExecute 來打開音效設定面板

extern short leftAudioData[BUFFER_SIZE];
extern short rightAudioData[BUFFER_SIZE];
extern double leftSpectrumData[BUFFER_SIZE];
extern double rightSpectrumData[BUFFER_SIZE];

extern "C" int fft_thd_n_exe(short*, short*, double*, double*);
extern "C" void fft_get_thd_n_db(double*, double*, double*);

extern double thd_n[2];
extern double dB_ValueMax[2];
extern double freq[2];
extern short leftAudioData[];
extern short rightAudioData[];
extern double leftSpectrumData[];
extern double rightSpectrumData[];

AudioManager::AudioManager() {
    // 建構函式，目前不需要做特別的事情
}

const std::string& AudioManager::GetResultString() const {
    return resultString;
}

bool AudioManager::ExecuteTestsFromConfig(const Config& config) {
    resultString.clear(); // 每次執行前都清除上一次的結果

    // 1. 檢查裝置是否存在
    if (!CheckAudioDevices(config)) {
        return false; // 裝置檢查失敗，流程中止
    }

    // 2. 依序執行設定檔中啟用的各項測試
    if (config.AudioTestEnable) {
        if (!RunAudioTest(config)) return false;
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
    // 重置所有用於測試的資料緩衝區
    memset(leftAudioData, 0, sizeof(leftAudioData));
    memset(rightAudioData, 0, sizeof(rightAudioData));
    memset(leftSpectrumData, 0, sizeof(leftSpectrumData));
    memset(rightSpectrumData, 0, sizeof(rightSpectrumData));
    memset(thd_n, 0, sizeof(thd_n));
    memset(dB_ValueMax, 0, sizeof(dB_ValueMax));
    memset(freq, 0, sizeof(freq));

    // 呼叫 C-API 執行核心 FFT 測試
    fft_thd_n_exe(leftAudioData, rightAudioData, leftSpectrumData, rightSpectrumData);
    fft_get_thd_n_db(thd_n, dB_ValueMax, freq);

    // 根據設定檔的標準，判斷測試結果
    if (thd_n[0] > config.thd_n || thd_n[1] > config.thd_n) {
        std::stringstream ss;
        ss << "LOG:ERROR_AUDIO_THD_N_VALUE_" << "L_" << (int)thd_n[0] << "_R_" << (int)thd_n[1] << "#";
        resultString = ss.str();
        return false;
    }
    else if (dB_ValueMax[0] < config.db_ValueMax || dB_ValueMax[1] < config.db_ValueMax) {
        std::stringstream ss;
        ss << "LOG:ERROR_AUDIO_DB_MAX_VALUE_" << "L_" << (int)dB_ValueMax[0] << "_R_" << (int)dB_ValueMax[1] << "#";
        resultString = ss.str();
        return false;
    }
    else if (freq[0] != config.frequencyL || freq[1] != config.frequencyR) {
        std::stringstream ss;
        ss << "LOG:ERROR_AUDIO_FREQUENCY_" << "L_" << (int)freq[0] << "_R_" << (int)freq[1] << "#";
        resultString = ss.str();
        return false;
    }
    else {
        std::stringstream ss;
        ss << "LOG:SUCCESS_AUDIO_TEST_THD_N_" << "L_" << (int)thd_n[0] << "_R_" << (int)thd_n[1]
            << "_dB_ValueMax_L_" << (int)dB_ValueMax[0] << "_R_" << (int)dB_ValueMax[1]
            << "_Frequency_L_" << (int)freq[0] << "_R_" << (int)freq[1] << "#";
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

    // 建立一個可修改的 config 副本，因為 FindDeviceIdByName 可能會修改它
    Config nonConstConfig = config;

    if (audio.FindDeviceIdByName(nonConstConfig, deviceId)) {
        if (!audio.SetDefaultAudioPlaybackDevice(deviceId)) {
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