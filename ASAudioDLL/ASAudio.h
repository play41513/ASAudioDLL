#pragma once

#include "ConfigManager.h"
#include <windows.h>
#include <vector>
#include <string>
#include <mutex>
#include <condition_variable>
#include "fftw3.h"
#include "portaudio.h"
#include "ConstantString.h" // ���]�z�����~�X�w�q�b�o��
#include <mmsystem.h>
#include <mmdeviceapi.h>
#include <endpointvolume.h>
#include <dshow.h>
#include <propkey.h>
#include <functiondiscoverykeys.h>
#include <atlbase.h> // for CComPtr
#include <initguid.h>
#include <comdef.h> // for _bstr_t

#pragma comment(lib, "strmiids.lib")
#pragma comment(lib,"winmm.lib")
#pragma warning(disable:4244)

// FFTW���R�^�I�禡����
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

// ����ѼƵ��c
struct WAVE_PARM {
    std::wstring WaveOutDev;
    std::wstring WaveInDev;
    int WaveOutVolume;
    int WaveInVolume;
    int frequencyL;
    int frequencyR;
    int WaveOutDelay;
    std::wstring AudioFile;

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

    WAVE_PARM() :
        WaveOutVolume(0),
        WaveInVolume(0),
        frequencyL(0),
        frequencyR(0),
        WaveOutDelay(0),
        isRecordingFinished(false),
        bMuteTest(false),
        firstBufferDiscarded(false)
    {

    }
};

// PortAudio �^�ը禡�ϥΪ���Ƶ��c
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

    // �T��ƻs�M���
    ASAudio(const ASAudio&) = delete;
    void operator=(const ASAudio&) = delete;

    void ClearLog();
    void AppendLog(const std::wstring& message);
    const std::wstring& GetLog() const;
    WAVE_PARM& GetWaveParm();

    // --- �����禡�ŧi ---
    bool PlayWavFile(bool AutoClose);
    void StopPlayingWavFile();
    MMRESULT StartRecordingAndDrawSpectrum();
    MMRESULT StopRecording();
    DWORD SetMicSystemVolume();
    DWORD SetSpeakerSystemVolume() const;
    void SetMicMute(bool mute);
    bool StartAudioLoopback(const std::wstring captureKeyword, const std::wstring renderKeyword);
    void StopAudioLoopback();
    bool SetDefaultAudioPlaybackDevice(const std::wstring& deviceId);
    bool FindDeviceIdByName(struct Config& config, std::wstring& outDeviceId);
    std::wstring charToWstring(const char* szIn);
    void SpectrumAnalysis(double* leftSpectrumData, double* rightSpectrumData);
    void GetSpectrumData(fftw_complex*& leftSpectrum, fftw_complex*& rightSpectrum);

    // �]�Ƭd��
    static int GetWaveOutDevice(std::wstring szOutDevName);
    static int GetWaveInDevice(std::wstring szInDevName);
    static int GetPa_WaveOutDevice(std::wstring szOutDevName);
    static int GetPa_WaveInDevice(std::wstring szInDevName);

    // ����M���� PortAudio
    DWORD playPortAudioSound();
    void stopPlayback();

    // ����̫᪺���~�T��
    const std::string& GetLastResult() const;

    const fftw_complex* GetLeftSpectrumPtr() const;
    const fftw_complex* GetRightSpectrumPtr() const;

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

    // DirectShow ������ COM ��������
    CComPtr<IGraphBuilder> pGraph;
    CComPtr<ICaptureGraphBuilder2> pCaptureGraph;
    CComPtr<IMediaControl> pMediaControl;
    CComPtr<IBaseFilter> pAudioCapture;
    CComPtr<IBaseFilter> pAudioRenderer;

    // ��L���A�ܼ�
    bool bFirstWaveInFlag;
    std::string strMacroResult;
    std::wstring allContent; 

    // --- �����ܼ� ---
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