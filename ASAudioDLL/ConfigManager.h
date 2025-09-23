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
    int db_ValueMax;

    // [PlayWAVFile]
    bool PlayWAVFileEnable;
    bool CloseWAVFileEnable;
    bool AutoCloseWAVFile;
    std::string WAVFilePath;
    std::wstring WAVFilePath_w;

    // [AudioLoopBack]
    bool AudioLoopBackEnable;
    bool AudioLoopBackStart;

    // [SwitchDefaultAudio]
    bool SwitchDefaultAudioEnable;
    std::vector<std::wstring> monitorNames;
    std::string AudioName;
    int AudioIndex;
};

class ConfigManager {
public:
    // 從指定的 INI 檔案路徑讀取設定
    static bool ReadConfig(const std::string& filePath, Config& config);
};