#pragma once
#ifdef ASAudio_EXPORTS
#define ASAudio_API extern "C" __declspec(dllexport)
#else
#define ASAudio_API extern "C" __declspec(dllimport)
#endif

#include <windows.h>
#include <string>
#include <vector>
#include <mutex>
#include <condition_variable>
#include "fftw3.h"
#include "portaudio.h"
#include "ConstantString.h"



#pragma comment(lib, "strmiids.lib")
#pragma comment(lib,"winmm.lib")
#pragma warning(disable:4244)


void fft_set_test_parm(const char* szOutDevName, int WaveOutVolume, const char* szInDevName, int WaveInVolume, int frequencyL, int frequencyR, int WaveOutdelay);
void fft_set_audio_volume(const char* szOutDevName, int WaveOutVolume, const char* szInDevName, int WaveInVolume);
void fft_get_thd_n_db(double* thd_n, double* dB_ValueMax, double* freq);
void fft_set_system_sound_mute(bool Mute);
int fft_thd_n_exe(short* leftAudioData, short* rightAudioData, double* leftSpectrumData, double* rightSpectrumData);

extern "C" __declspec(dllexport) BSTR MacroTest(HWND ownerHwnd,const char* filePath);

typedef void (*SpectrumAnalysisCallback)(fftw_complex* leftSpectrum, fftw_complex* rightSpectrum);
static const int BUFFER_SIZE = 2205; //控制到頻譜每一個點為20Hz，每五十個點為1K Hz

#define M_PI 3.14159235358979323
#define SAMPLE_RATE 44100       // 定義採樣率為44100Hz
#define AMPLITUDE 0.5           // 定義信號振幅為0.5
#define FRAMES_PER_BUFFER 256   // 定義每個緩衝區的幀數

// 結構體定義：儲存每個區塊的參數
struct Config {
	std::string outDeviceName = ""; // 輸出裝置名稱
	int outDeviceVolume = 0; // 輸出裝置音量

	std::string inDeviceName = ""; // 輸入裝置名稱
	int inDeviceVolume = 0; // 輸入裝置音量

	bool AudioTestEnable = false;
	int frequencyL = 0; // 左聲道頻率
	int frequencyR = 0; // 右聲道頻率
	int waveOutDelay = 0; // 輸出延遲
	int thd_n = 0;
	int db_ValueMax = 0;

	bool PlayWAVFileEnable = false;
	bool CloseWAVFileEnable = false;
	bool AutoCloseWAVFile = false;
	std::string WAVFilePath = "";
	std::wstring WAVFilePath_w = L"";

	bool AudioLoopBackEnable = false;
	bool AudioLoopBackStart = false;


	bool SwitchDefaultAudioEnable = false;
	std::string AudioName = "";
	std::vector<std::wstring> monitorNames;
	int AudioIndex= 0;
};

// 用於儲存音頻信號相位的資料結構
typedef struct {
	float phaseL;   // 左聲道相位
	float phaseR;   // 右聲道相位
	float frequencyL; // 左聲道頻率
	float frequencyR; // 右聲道頻率
	PaStream* stream; // 保存音頻流指針

	int errorcode;
	std::string errorMsg;

} paTestData;

typedef struct _WAVE_DATA
{
	short LeftAudioBuffer[BUFFER_SIZE];  // 左聲道音訊緩衝
	short RightAudioBuffer[BUFFER_SIZE];  // 右聲道音訊緩衝
	short audioBuffer[BUFFER_SIZE * 2];  // 複合左右聲道音訊緩衝
}mWAVE_DATA;
typedef struct _WAVE_ANALYSIS
{
	int TotalEnergyHz;//頻譜最高能量的頻率
	int TotalEnergyPoint;//頻譜最高能量的頻率點
	double TotalEnergy;//頻譜總能量
	double thd_N;
	double thd_N_dB;
	double fundamentalEnergy;//頻譜最高能量點的能量
}WAVE_ANALYSIS;

typedef struct _WAVE_PARM
{
	std::wstring WaveInDev;
	std::wstring WaveOutDev; //設置播放、錄音裝置序列
	int WaveInVolume = 0, WaveOutVolume = 0; //設置播放、錄音裝置音量
	int frequencyL = 0, frequencyR = 0; //設置播放頻率
	int WaveOutDelay = 0;//如播放裝置太慢啟動，可調整此參數延遲錄音


	bool isRecordingFinished = false;// 是否錄音完成的標誌
	bool bMuteTest = false;//靜音測試的標誌
	std::wstring AudioFile;
	mWAVE_DATA WAVE_DATA = { 0 };
	WAVE_ANALYSIS leftWAVE_ANALYSIS = { 0 };
	WAVE_ANALYSIS rightWAVE_ANALYSIS = { 0 };

	WAVE_ANALYSIS leftMuteWAVE_ANALYSIS = { 0 };
	WAVE_ANALYSIS rightMuteWAVE_ANALYSIS = { 0 };
	std::condition_variable recordingFinishedCV; // 錄音完成條件變數
	bool firstBufferDiscarded = false;

	std::mutex recordingMutex; // 錄音互斥鎖
}WAVE_PARM;


class ASAudio
{
public:
	~ASAudio();
	static ASAudio& GetInstance();
	ASAudio(const ASAudio&) = delete;
	ASAudio& operator=(const ASAudio&) = delete;

	MMRESULT StartRecordingAndDrawSpectrum();
	MMRESULT StopRecording();


	DWORD SetMicSystemVolume();
	DWORD SetSpeakerSystemVolume();
	void SetMicMute(bool mute);

	void stopPlayback();

	int GetPa_WaveOutDevice(std::wstring szOutDevName);
	int GetPa_WaveInDevice(std::wstring szInDevName);
	std::wstring charToWstring(const char* szIn);
	void SpectrumAnalysis(double* leftSpectrumData, double* rightSpectrumData);

	void GetSpectrumData(fftw_complex*& leftSpectrum, fftw_complex*& rightSpectrum) {
		leftSpectrum = fftOutputBufferLeft;
		rightSpectrum = fftOutputBufferRight;
	}

	static bool PlayWavFile(bool AutoClose);
	static void StopPlayingWavFile();
	static int GetWaveOutDevice(std::wstring szOutDevName);
	static int GetWaveInDevice(std::wstring szInDevName);
	static bool ReadConfig(const std::string& filePath, Config& config);

	static bool StartAudioLoopback(const std::wstring captureKeyword, const std::wstring renderKeyword);
	static void StopAudioLoopback();
	static bool FindDeviceIdByName(Config& config, std::wstring& outDeviceId);
	static bool SetDefaultAudioPlaybackDevice(const std::wstring& deviceId);

	bool ExecuteTestsFromConfig(const Config& config);


private:

	ASAudio(SpectrumAnalysisCallback callback);
	fftw_complex* fftInputBufferLeft;
	fftw_complex* fftInputBufferRight;
	fftw_complex* fftOutputBufferLeft;
	fftw_complex* fftOutputBufferRight;
	fftw_plan fftPlanLeft;
	fftw_plan fftPlanRight;

	void (*waveInCallback)(WAVEHDR*);
	int bufferIndex;

	static void CALLBACK WaveInProc(HWAVEIN hwi, UINT uMsg, DWORD_PTR dwInstance, DWORD_PTR dwParam1, DWORD_PTR dwParam2);
	SpectrumAnalysisCallback spectrumCallback;
	void ProcessAudioBuffer();
	DWORD playPortAudioSound();
	static int patestCallback(const void* inputBuffer, void* outputBuffer,
		unsigned long framesPerBuffer,
		const PaStreamCallbackTimeInfo* timeInfo,
		PaStreamCallbackFlags statusFlags,
		void* userData);

	bool RunAudioTest(const Config& config);
	bool RunWavPlayback(const Config& config);
	bool RunAudioLoopback(const Config& config);
	bool RunSwitchDefaultDevice(const Config& config);


};