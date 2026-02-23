#pragma once

#include "ConfigManager.h"
#include <windows.h>
#include <vector>
#include <string>
#include <mutex>
#include <condition_variable>
#include "fftw3.h"
#include "portaudio.h"
#include "ConstantString.h" // 假設您的錯誤碼定義在這裡
#include <mmsystem.h>
#include <mmdeviceapi.h>
#include <endpointvolume.h>
#include <dshow.h>
#include <propkey.h>
#include <functiondiscoverykeys.h>
#include <atlbase.h> // for CComPtr
#include <initguid.h>
#include <atomic>
#include <comdef.h> // for _bstr_t

#pragma comment(lib, "strmiids.lib")
#pragma comment(lib,"winmm.lib")
#pragma warning(disable:4244)

// FFTW分析回呼函式指標
typedef void (*SpectrumAnalysisCallback)(fftw_complex* leftSpectrum, fftw_complex* rightSpectrum);


struct WAVE_ANALYSIS {
    double thd_N;
    double thd_N_dB;
    double fundamentalEnergy;
    double TotalEnergy;
    int TotalEnergyPoint;

    WAVE_ANALYSIS() :
        thd_N(0.0),
        thd_N_dB(0.0),
        fundamentalEnergy(0.0),
        TotalEnergy(0.0),
        TotalEnergyPoint(0)
    {
    }
};

struct WAVE_DATA {
    short audioBuffer[BUFFER_SIZE * 2];
    short LeftAudioBuffer[BUFFER_SIZE];
    short RightAudioBuffer[BUFFER_SIZE];

    WAVE_DATA() {
        memset(audioBuffer, 0, sizeof(audioBuffer));
        memset(LeftAudioBuffer, 0, sizeof(LeftAudioBuffer));
        memset(RightAudioBuffer, 0, sizeof(RightAudioBuffer));
    }
};

// 全域參數結構
struct WAVE_PARM {
    std::wstring WaveOutDev;
    std::wstring WaveInDev;
    std::wstring ActualWaveOutDev;
    std::wstring ActualWaveInDev;
    int WaveOutVolume;
    int WaveInVolume;
    int frequencyL;
    int frequencyR;
    int WaveOutDelay;
    std::wstring AudioFile;
    double fundamentalBandwidthHz;

    WAVE_DATA WAVE_DATA;
    WAVE_ANALYSIS leftWAVE_ANALYSIS;
    WAVE_ANALYSIS rightWAVE_ANALYSIS;
    WAVE_ANALYSIS leftMuteWAVE_ANALYSIS;
    WAVE_ANALYSIS rightMuteWAVE_ANALYSIS;

    bool isRecordingFinished = false;
    std::mutex recordingMutex;
    std::condition_variable recordingFinishedCV;
    bool bMuteTest = false;
    bool firstBufferDiscarded = false;
    // 使用 atomic<bool> 確保在多執行緒環境下安全地寫入和讀取錯誤狀態
    std::atomic<bool> deviceError;
    // 定義錄音的最終結果
    enum class RecordingResult {
        Success,
        Timeout,
        DeviceError
    };
    RecordingResult lastRecordingResult;

    WAVE_PARM() :
        WaveOutVolume(0),
        WaveInVolume(0),
        frequencyL(0),
        frequencyR(0),
        WaveOutDelay(0),
        isRecordingFinished(false),
        bMuteTest(false),
        firstBufferDiscarded(false),
        fundamentalBandwidthHz(100.0),
        deviceError(false),
        lastRecordingResult(RecordingResult::Success) 
    {

    }
};

// PortAudio 回調函式使用的資料結構
typedef struct {
    float phaseL;
    float phaseR;
    double frequencyL;
    double frequencyR;
    PaStream* stream;
    const char* errorMsg;
    int errorcode;
} paTestData;

class ASAudio {
public:
    static ASAudio& GetInstance();

    // 禁止複製和賦值
    ASAudio(const ASAudio&) = delete;
    void operator=(const ASAudio&) = delete;

    void ClearLog();
    void AppendLog(const std::wstring& message);
    const std::wstring& GetLog() const;
    WAVE_PARM& GetWaveParm();
    WAVE_PARM::RecordingResult WaitForRecording(int timeoutSeconds);

    // --- 成員函式宣告 ---
    bool PlayWavFile(bool AutoClose);
    void StopPlayingWavFile();
    MMRESULT StartRecordingAndDrawSpectrum();
    MMRESULT StartRecordingOnly();
    MMRESULT StopRecording();
    DWORD SetMicSystemVolume();
    DWORD SetSpeakerSystemVolume() const;
    void SetMicMute(bool mute);
    bool StartAudioLoopback(const std::wstring captureKeyword, const std::wstring renderKeyword);
    void StopAudioLoopback();
    bool SetDefaultAudioPlaybackDevice(const std::wstring& deviceId);
    bool SetListenToThisDevice(const std::wstring& deviceId, int enable);
    bool FindDeviceIdByName(struct Config& config, std::wstring& outDeviceId, EDataFlow dataFlow = eRender);
    std::wstring charToWstring(const char* szIn);
    void SpectrumAnalysis(double* leftSpectrumData, double* rightSpectrumData);
    void GetSpectrumData(fftw_complex*& leftSpectrum, fftw_complex*& rightSpectrum);

    // 設備查找
    static int GetWaveOutDevice(std::wstring szOutDevName);
    static int GetWaveInDevice(std::wstring szInDevName);
    static int GetPa_WaveOutDevice(std::wstring szOutDevName);
    static int GetPa_WaveInDevice(std::wstring szInDevName);

    // 播放和停止 PortAudio
    DWORD playPortAudioSound();
    void stopPlayback();

    // 獲取最後的錯誤訊息
    const std::string& GetLastResult() const;

    const fftw_complex* GetLeftSpectrumPtr() const;
    const fftw_complex* GetRightSpectrumPtr() const;
    void ExecuteFft();

private:
    ASAudio();
    ~ASAudio();

    void ProcessAudioBuffer();
    static void CALLBACK WaveInProc(HWAVEIN hwi, UINT uMsg, DWORD_PTR dwInstance, DWORD_PTR dwParam1, DWORD_PTR dwParam2);
    static int patestCallback(const void* inputBuffer, void* outputBuffer,
        unsigned long framesPerBuffer,
        const PaStreamCallbackTimeInfo* timeInfo,
        PaStreamCallbackFlags statusFlags,
        void* userData);

    HWAVEIN hWaveIn;
    HWAVEOUT hWaveOut;
    WAVEHDR waveHdr;
    WAVEHDR WaveOutHeader;
    std::vector<char> WaveOutData;

    paTestData data;

    // DirectShow 相關的 COM 介面指標
    CComPtr<IGraphBuilder> pGraph;
    CComPtr<ICaptureGraphBuilder2> pCaptureGraph;
    CComPtr<IMediaControl> pMediaControl;
    CComPtr<IBaseFilter> pAudioCapture;
    CComPtr<IBaseFilter> pAudioRenderer;

    // 其他狀態變數
    bool bFirstWaveInFlag;
    std::string strMacroResult;
    std::wstring allContent; 

    // --- 成員變數 ---
    WAVE_PARM mWAVE_PARM;
    SpectrumAnalysisCallback spectrumCallback;
    size_t bufferIndex;

    fftw_complex* fftInputBufferLeft;
    fftw_complex* fftInputBufferRight;
    fftw_complex* fftOutputBufferLeft;
    fftw_complex* fftOutputBufferRight;

    fftw_plan fftPlanLeft;
    fftw_plan fftPlanRight;
};