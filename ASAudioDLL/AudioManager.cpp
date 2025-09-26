
#include "AudioManager.h"
#include "ConstantString.h"
#include "ASAudio.h"      // �ޤJ���h���T�ާ@���O
#include <Windows.h>
#include <sstream>      // �Ω�զX�r��
#include <shellapi.h>   // ���F�ϥ� ShellExecute �ӥ��}���ĳ]�w���O
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
extern "C" int fft_thd_n_exe(short*, short*, double*, double*);
extern "C" void fft_get_thd_n_db(double*, double*, double*);
extern "C" int fft_mute_exe(bool, bool, short*, short*, double*, double*);
extern "C" void fft_get_snr(double*);

class ComInitializer {
public:
    ComInitializer() : hr(CoInitializeEx(NULL, COINIT_APARTMENTTHREADED)) {
    }
    ~ComInitializer() {
        // �u�b���\��l�ƫ�~�i��Ϫ�l��
        if (SUCCEEDED(hr)) {
            CoUninitialize();
        }
    }
    // ���~���i�H�ˬd��l�ƬO�_���\
    HRESULT GetHResult() const {
        return hr;
    }
private:
    HRESULT hr; // �O�s CoInitializeEx ���^�ǭ�
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
    ComInitializer comInit;
    if (FAILED(comInit.GetHResult())) {
        resultString = "LOG:ERROR_COM_INITIALIZE_FAILED#";
        return false; // �p�G COM ��l�ƥ��ѡA��������
    }

    resultString.clear(); // �C������e���M���W�@�������G

    // 1. �ˬd�˸m�O�_�s�b
    if (!CheckAudioDevices(config)) {
        return false; // �˸m�ˬd���ѡA�y�{����
    }

    // 2. �̧ǰ���]�w�ɤ��ҥΪ��U������
    if (config.AudioTestEnable) {
        if (!RunAudioTest(config)) return false;
    }
    if (config.snrTestEnable) {
        // �`�N�GSNR ���ե����b AudioTest (THD+N ����) �������A
        // �]�����ݭn AudioTest ���o���T���j�סC
        if (!config.AudioTestEnable) {
            resultString = "LOG:ERROR_SNR_NEEDS_AUDIOTEST_ENABLED#";
            return false;
        }
        if (!RunSnrTest(config)) return false;
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

    // �p�G�Ҧ��y�{���]���F�A�ӥB resultString �٬O�Ū��A
    // ��ܩҦ��\�ೣ�Q disable�A���@�ӹw�]�����\�T���C
    if (resultString.empty()) {
        resultString = "LOG:PASS#";
    }

    return true; // �Ҧ��ҥΪ����լy�{���w���槹��
}

bool AudioManager::CheckAudioDevices(const Config& config) {
    ASAudio& audio = ASAudio::GetInstance();

    if (!config.outDeviceName.empty()) {
        ULONGLONG timeout = GetTickCount64() + 5000; // 5��W��
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
    // ���m�Ҧ��Ω���ժ���ƽw�İ�
    memset(leftAudioData, 0, sizeof(leftAudioData));
    memset(rightAudioData, 0, sizeof(rightAudioData));
    memset(leftSpectrumData, 0, sizeof(leftSpectrumData));
    memset(rightSpectrumData, 0, sizeof(rightSpectrumData));
    memset(thd_n, 0, sizeof(thd_n));
    memset(FundamentalLevel_dBFS, 0, sizeof(FundamentalLevel_dBFS));
    memset(freq, 0, sizeof(freq));

    // �I�s C-API ����֤� FFT ����
    fft_thd_n_exe(leftAudioData, rightAudioData, leftSpectrumData, rightSpectrumData);
    fft_get_thd_n_db(thd_n, FundamentalLevel_dBFS, freq);

    double levelDifference = std::abs(FundamentalLevel_dBFS[0] - FundamentalLevel_dBFS[1]);

    // �ھڳ]�w�ɪ��зǡA�P�_���յ��G
    if (thd_n[0] > config.thd_n || thd_n[1] > config.thd_n) {
        std::stringstream ss;
        ss << "LOG:ERROR_AUDIO_THD_N_VALUE_" << "L_" << (int)thd_n[0] << "_R_" << (int)thd_n[1] << "#";
        resultString = ss.str();
        return false;
    }
    else if (FundamentalLevel_dBFS[0] < config.FundamentalLevel_dBFS || FundamentalLevel_dBFS[1] < config.FundamentalLevel_dBFS) {
        std::stringstream ss;
        ss << "LOG:ERROR_AUDIO_FUNDAMENTAL_LEVEL_VALUE_" << "L_" << (int)FundamentalLevel_dBFS[0] << "_R_" << (int)FundamentalLevel_dBFS[1] << "#";
        resultString = ss.str();
        return false;
    }
    else if (levelDifference > config.maxLevelDifference_dB) {
        std::stringstream ss;
        ss << "LOG:ERROR_AUDIO_LEVEL_IMBALANCE_" << "L_" << (int)FundamentalLevel_dBFS[0] << "_R_" << (int)FundamentalLevel_dBFS[1] << "_DIFF_" << levelDifference << "#";
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
            << "_FundamentalLevel_dBFS_L_" << (int)FundamentalLevel_dBFS[0] << "_R_" << (int)FundamentalLevel_dBFS[1]
            << "_Frequency_L_" << (int)freq[0] << "_R_" << (int)freq[1] << "#";
        resultString = ss.str();
        return true;
    }
}

bool AudioManager::RunWavPlayback(const Config& config) {
    ASAudio& audio = ASAudio::GetInstance();
    if (config.PlayWAVFileEnable) {
        if (!audio.PlayWavFile(config.AutoCloseWAVFile)) {
            // �q ASAudio ��Ҥ�����ԲӪ����~�T��
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

    // �إߤ@�ӥi�ק諸 config �ƥ��A�]�� FindDeviceIdByName �i��|�ק復
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
bool AudioManager::RunSnrTest(const Config& config) {
    // �����R�����s�H������n�ƾ�
    // MuteWaveOut=true, MuteWaveIn=false (�R������A���`����)
    fft_mute_exe(true, false, leftAudioData_SNR, rightAudioData_SNR, leftSpectrumData_SNR, rightSpectrumData_SNR);

    double snr_result[2] = { 0 };
    fft_get_snr(snr_result);

    // �P�_���յ��G
    if (snr_result[0] < config.snrThreshold || snr_result[1] < config.snrThreshold) {
        std::stringstream ss;
        // ���ͥ��Ѫ� Log
        ss << "LOG:ERROR_SNR_VALUE_TOO_LOW_" << "L_" << static_cast<int>(snr_result[0])
            << "_R_" << static_cast<int>(snr_result[1]) << "#";
        resultString = ss.str();
        return false;
    }
    else {
        std::stringstream ss;
        // �N���\�����G���[��{���� resultString �᭱
        // �`�N�G�ڭ̥� += �Ӫ��[�A�Ӥ��O�������
        ss << "SUCCESS_SNR_" << "L_" << static_cast<int>(snr_result[0])
            << "_R_" << static_cast<int>(snr_result[1]) << "#";

        // �p�G���e�w�� AudioTest �����\�T���A�h�X��
        if (resultString.find("SUCCESS_AUDIO_TEST") != std::string::npos) {
            // ���������� '#'
            resultString.pop_back();
            resultString += "_" + ss.str();
        }
        else {
            resultString = "LOG:" + ss.str();
        }
        return true;
    }
}