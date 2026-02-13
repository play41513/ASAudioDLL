#include "ConfigManager.h"
#include <fstream>

static std::wstring charToWstring(const char* szIn)
{
    int length = MultiByteToWideChar(CP_ACP, 0, szIn, -1, NULL, 0);
    std::wstring strRet;
    if (length > 0)
    {
        WCHAR* buf = new WCHAR[length];
        ZeroMemory(buf, length * sizeof(WCHAR));
        MultiByteToWideChar(CP_ACP, 0, szIn, -1, buf, length);
        strRet = buf;
        delete[] buf;
    }
    return strRet;
}

bool ConfigManager::ReadConfig(const std::string& filePath, Config& config)
{
    std::ifstream file(filePath);
    if (!file.good())
    {
        return false;
    }
    file.close();

    char buffer[256];
    int iNumber = 0;

    GetPrivateProfileStringA("Device", "outDeviceName", "", buffer, sizeof(buffer), filePath.c_str());
    config.outDeviceName = buffer;

    iNumber = GetPrivateProfileIntA("Device", "outDeviceVolume", 0, filePath.c_str());
    config.outDeviceVolume = iNumber;

    GetPrivateProfileStringA("Device", "inDeviceName", "", buffer, sizeof(buffer), filePath.c_str());
    config.inDeviceName = buffer;

    iNumber = GetPrivateProfileIntA("Device", "inDeviceVolume", 0, filePath.c_str());
    config.inDeviceVolume = iNumber;

    iNumber = GetPrivateProfileIntA("AudioTest", "AudioTestEnable", 0, filePath.c_str());
    config.AudioTestEnable = iNumber == 1 ? true : false;

    iNumber = GetPrivateProfileIntA("AudioTest", "frequencyL", 0, filePath.c_str());
    config.frequencyL = iNumber;
    iNumber = GetPrivateProfileIntA("AudioTest", "frequencyR", 0, filePath.c_str());
    config.frequencyR = iNumber;

    iNumber = GetPrivateProfileIntA("AudioTest", "waveOutDelay", 0, filePath.c_str());
    config.waveOutDelay = iNumber;

    iNumber = GetPrivateProfileIntA("AudioTest", "thd_n", 0, filePath.c_str());
    config.thd_n = iNumber;

    iNumber = GetPrivateProfileIntA("AudioTest", "FundamentalLevel_dBFS", 0, filePath.c_str());
    config.FundamentalLevel_dBFS = iNumber;

    config.fundamentalBandwidthHz = GetPrivateProfileIntA("AudioTest", "FundamentalBandwidthHz", 100, filePath.c_str());
    config.maxLevelDifference_dB = GetPrivateProfileIntA("AudioTest", "MaxLevelDifference_dB", 2, filePath.c_str());

    config.snrTestEnable = GetPrivateProfileIntA("AudioTest", "SNRTestEnable", 0, filePath.c_str()) == 1;
    config.snrThreshold = GetPrivateProfileIntA("AudioTest", "SNRThreshold", 80, filePath.c_str()); // 預設門檻為 80dB

    iNumber = GetPrivateProfileIntA("PlayWAVFile", "PlayWAVFileEnable", 0, filePath.c_str());
    config.PlayWAVFileEnable = iNumber == 1 ? true : false;
    iNumber = GetPrivateProfileIntA("PlayWAVFile", "CloseWAVFileEnable", 0, filePath.c_str());
    config.CloseWAVFileEnable = iNumber == 1 ? true : false;

    iNumber = GetPrivateProfileIntA("PlayWAVFile", "AutoCloseWAVFile", 0, filePath.c_str());
    config.AutoCloseWAVFile = iNumber == 1 ? true : false;

    GetPrivateProfileStringA("PlayWAVFile", "WAVFilePath", "", buffer, sizeof(buffer), filePath.c_str());
    config.WAVFilePath = buffer;

    config.WAVFilePath_w = charToWstring(config.WAVFilePath.c_str());

    iNumber = GetPrivateProfileIntA("AudioLoopBack", "AudioLoopBackEnable", 0, filePath.c_str());
    config.AudioLoopBackEnable = iNumber == 1 ? true : false;
    iNumber = GetPrivateProfileIntA("AudioLoopBack", "AudioLoopBackStart", 0, filePath.c_str());
    config.AudioLoopBackStart = iNumber == 1 ? true : false;

    iNumber = GetPrivateProfileIntA("SwitchDefaultAudio", "SwitchDefaultAudioEnable", 0, filePath.c_str());
    config.SwitchDefaultAudioEnable = iNumber == 1 ? true : false;
    config.monitorNames.clear();
    for (int i = 0;; i++)
    {
        std::string temp = "AudioName";
        if (i > 0)
            temp += std::to_string(i);
        GetPrivateProfileStringA("SwitchDefaultAudio", temp.c_str(), "", buffer, sizeof(buffer), filePath.c_str());
        if (strlen(buffer) == 0)
            break;

        config.monitorNames.push_back(charToWstring(buffer));
    }
    GetPrivateProfileStringA("SwitchDefaultAudio", "AudioName", "", buffer, sizeof(buffer), filePath.c_str());
    config.AudioName = buffer;
    iNumber = GetPrivateProfileIntA("SwitchDefaultAudio", "AudioIndex", 0, filePath.c_str());
    config.AudioIndex = iNumber;

    config.setListen = GetPrivateProfileIntA("SwitchDefaultAudio", "SetListen", -1, filePath.c_str());

    return true;
}