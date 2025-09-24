#pragma once

#include <string>
#include <vector>
#include <Windows.h>

// �w�q�F�Ҧ����թһݪ��]�w�Ѽ�
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
    bool snrTestEnable;      // �O�_�ҥ� SNR ����
    double snrThreshold;     // SNR ���X��з� (���: dB)

    // [PlayWAVFile]
    bool PlayWAVFileEnable;
    bool CloseWAVFileEnable;
    bool AutoCloseWAVFile;
    std::string WAVFilePath;
    std::wstring WAVFilePath_w;
    double fundamentalBandwidthHz;

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
    // �q���w�� INI �ɮ׸��|Ū���]�w
    static bool ReadConfig(const std::string& filePath, Config& config);
};