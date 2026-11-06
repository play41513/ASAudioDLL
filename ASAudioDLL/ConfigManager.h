#pragma once

#include <string>
#include <vector>
#include <Windows.h>

// 定義了所有測試所需的設定參數
struct Config {
    // [Device]
    std::string outDeviceName;
    int outDeviceVolume;
    std::string inDeviceName;
    int inDeviceVolume;

    // [AudioTest]
    bool AudioTestEnable;
    int frequencyL;
    int frequencyR;
    int waveOutDelay;
    int thd_n;
    int FundamentalLevel_dBFS;
    bool snrTestEnable;      // 是否啟用 SNR 測試
    double snrThreshold;     // SNR 的合格標準 (單位: dB)

    // [PlayWAVFile]
    bool PlayWAVFileEnable;
    bool CloseWAVFileEnable;
    bool AutoCloseWAVFile;
    std::string WAVFilePath;
    std::wstring WAVFilePath_w;
    double fundamentalBandwidthHz;
    double maxLevelDifference_dB; //最大音量差異容許值 (dB)

    // [AudioLoopBack]
    bool AudioLoopBackEnable;
    bool AudioLoopBackStart;

    // [SwitchDefaultAudio]
    bool SwitchDefaultAudioEnable;
    std::vector<std::wstring> monitorNames;
    int AudioIndex;

    // [setListenToThisDevice]
    bool SetListenToThisDeviceEnable;
    std::string AudioName;
    int setListen; // -1: 不改變, 0: 取消勾選, 1: 勾選

};

class ConfigManager {
public:
    // 從指定的 INI 檔案路徑讀取設定
    static bool ReadConfig(const std::string& filePath, Config& config);
};