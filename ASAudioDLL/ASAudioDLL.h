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
static const int BUFFER_SIZE = 2205; //������W�ШC�@���I��20Hz�A�C���Q���I��1K Hz

#define M_PI 3.14159235358979323
#define SAMPLE_RATE 44100       // �w�q�ļ˲v��44100Hz
#define AMPLITUDE 0.5           // �w�q�H�����T��0.5
#define FRAMES_PER_BUFFER 256   // �w�q�C�ӽw�İϪ��V��

// ���c��w�q�G�x�s�C�Ӱ϶����Ѽ�
struct Config {
	std::string outDeviceName = ""; // ��X�˸m�W��
	int outDeviceVolume = 0; // ��X�˸m���q

	std::string inDeviceName = ""; // ��J�˸m�W��
	int inDeviceVolume = 0; // ��J�˸m���q

	bool AudioTestEnable = false;
	int frequencyL = 0; // ���n�D�W�v
	int frequencyR = 0; // �k�n�D�W�v
	int waveOutDelay = 0; // ��X����
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

// �Ω��x�s���W�H���ۦ쪺��Ƶ��c
typedef struct {
	float phaseL;   // ���n�D�ۦ�
	float phaseR;   // �k�n�D�ۦ�
	float frequencyL; // ���n�D�W�v
	float frequencyR; // �k�n�D�W�v
	PaStream* stream; // �O�s���W�y���w

	int errorcode;
	std::string errorMsg;

} paTestData;

typedef struct _WAVE_DATA
{
	short LeftAudioBuffer[BUFFER_SIZE];  // ���n�D���T�w��
	short RightAudioBuffer[BUFFER_SIZE];  // �k�n�D���T�w��
	short audioBuffer[BUFFER_SIZE * 2];  // �ƦX���k�n�D���T�w��
}mWAVE_DATA;
typedef struct _WAVE_ANALYSIS
{
	int TotalEnergyHz;//�W�г̰���q���W�v
	int TotalEnergyPoint;//�W�г̰���q���W�v�I
	double TotalEnergy;//�W���`��q
	double thd_N;
	double thd_N_dB;
	double fundamentalEnergy;//�W�г̰���q�I����q
}WAVE_ANALYSIS;

typedef struct _WAVE_PARM
{
	std::wstring WaveInDev;
	std::wstring WaveOutDev; //�]�m����B�����˸m�ǦC
	int WaveInVolume = 0, WaveOutVolume = 0; //�]�m����B�����˸m���q
	int frequencyL = 0, frequencyR = 0; //�]�m�����W�v
	int WaveOutDelay = 0;//�p����˸m�ӺC�ҰʡA�i�վ㦹�ѼƩ������


	bool isRecordingFinished = false;// �O�_�����������лx
	bool bMuteTest = false;//�R�����ժ��лx
	std::wstring AudioFile;
	mWAVE_DATA WAVE_DATA = { 0 };
	WAVE_ANALYSIS leftWAVE_ANALYSIS = { 0 };
	WAVE_ANALYSIS rightWAVE_ANALYSIS = { 0 };

	WAVE_ANALYSIS leftMuteWAVE_ANALYSIS = { 0 };
	WAVE_ANALYSIS rightMuteWAVE_ANALYSIS = { 0 };
	std::condition_variable recordingFinishedCV; // �������������ܼ�
	bool firstBufferDiscarded = false;

	std::mutex recordingMutex; // ����������
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