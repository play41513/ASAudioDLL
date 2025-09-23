#include "ConfigManager.h"
#include "AudioManager.h"
#include "ASAudioDLL.h"
#include "ASAudio.h"
#include <sstream>
#include <fstream>
#include <algorithm>
#include <thread>
#include <chrono>
#include <locale>
#include <codecvt>

#include <mmsystem.h>
#include <mmdeviceapi.h>
#include <endpointvolume.h>
#include <dshow.h>
#include <propkey.h>
#include <functiondiscoverykeys.h>
#include <atlbase.h> // for CComPtr
#include <initguid.h>
#include <comdef.h> // for _bstr_t

short leftAudioData[BUFFER_SIZE] = { 0 };
short rightAudioData[BUFFER_SIZE] = { 0 };
double leftSpectrumData[BUFFER_SIZE] = { 0 };
double rightSpectrumData[BUFFER_SIZE] = { 0 };
double thd_n[2] = { 0 };
double dB_ValueMax[2] = { 0 };
double freq[2] = { 0 };

extern "C" __declspec(dllexport) BSTR MacroTest(HWND ownerHwnd, const char* filePath)
{
	bool shouldRetry;
	std::string finalResult;
	ASAudio& audioInstance = ASAudio::GetInstance(); // 獲取單例

	do {
		shouldRetry = false;
		audioInstance.ClearLog();

		Config configs;
		if (!ConfigManager::ReadConfig(filePath, configs)) {
			finalResult = "LOG:ERROR_CONFIG_NOT_FIND#";
		}
		else {
			// 初始化測試參數，透過 ASAudio 實例設定
			fft_set_test_parm(
				configs.outDeviceName.c_str(), configs.outDeviceVolume,
				configs.inDeviceName.c_str(), configs.inDeviceVolume,
				configs.frequencyL, configs.frequencyR,
				configs.waveOutDelay
			);
			audioInstance.GetWaveParm().AudioFile = configs.WAVFilePath_w;

			// 使用新的 AudioManager 執行測試流程
			AudioManager audioManager;
			audioManager.ExecuteTestsFromConfig(configs);
			finalResult = audioManager.GetResultString(); // 總是獲取結果，無論成功或失敗
		}

		bool hasError = (finalResult.find("ERROR") != std::string::npos);
		if (hasError && ownerHwnd != NULL) {
			std::string strTemp = "Failed: " + finalResult;
			std::wstring wErrorMsg = audioInstance.charToWstring(strTemp.c_str()) + L"\n\n" + audioInstance.GetLog() + L"\n\n是否要重測？(Retry ?)";
			if (MessageBoxW(ownerHwnd, wErrorMsg.c_str(), L"測試失敗", MB_YESNO | MB_ICONERROR) == IDYES) {
				shouldRetry = true;
			}
		}

	} while (shouldRetry);

	if (finalResult.empty()) {
		finalResult = "LOG:PASS#";
	}

	// 將 string 轉為 BSTR
	std::wstring wstr = audioInstance.charToWstring(finalResult.c_str());
	return SysAllocString(wstr.c_str());
}



void __cdecl fft_set_test_parm(
	const char* szOutDevName, int WaveOutVolumeValue
	, const char* szInDevName, int WaveInVolumeValue
	, int frequencyL, int frequencyR
	, int WaveOutdelay)
{
	auto& parm = ASAudio::GetInstance().GetWaveParm();
	parm.WaveOutDev = ASAudio::GetInstance().charToWstring(szOutDevName);
	parm.WaveInDev = ASAudio::GetInstance().charToWstring(szInDevName);
	parm.WaveOutVolume = WaveOutVolumeValue;
	parm.WaveInVolume = WaveInVolumeValue;
	parm.frequencyL = frequencyL;
	parm.frequencyR = frequencyR;
	parm.WaveOutDelay = WaveOutdelay;
}

int __cdecl fft_get_buffer_size()
{
	return BUFFER_SIZE;
}

void __cdecl fft_set_audio_volume(const char* szOutDevName, int WaveOutVolumeValue
	, const char* szInDevName, int WaveInVolumeValue)
{
	// 獲取 ASAudio 的單例 (singleton)
	ASAudio& audioInstance = ASAudio::GetInstance();

	// 透過單例存取其 public 成員 mWAVE_PARM
	auto& parm = audioInstance.GetWaveParm();

	// 將傳入的 C-style 字串轉換為 wstring 並設定參數
	parm.WaveOutDev = audioInstance.charToWstring(szOutDevName);
	parm.WaveInDev = audioInstance.charToWstring(szInDevName);
	parm.WaveOutVolume = WaveOutVolumeValue;
	parm.WaveInVolume = WaveInVolumeValue;

	// 呼叫 ASAudio 的成員函式來實際套用麥克風和喇叭的音量
	audioInstance.SetMicSystemVolume();
	audioInstance.SetSpeakerSystemVolume();
}

BSTR __cdecl fft_get_error_msg(int error_code)
{
	ErrorCode code = static_cast<ErrorCode>(error_code);
	auto it = error_map.find(code);
	if (it != error_map.end()) {
		return SysAllocString(it->second);
	}
	else {
		return SysAllocString(L"Unknown error code");
	}
}

void __cdecl fft_set_system_sound_mute(bool Mute)
{ // 系統參數設置，SPIF_SENDCHANGE 表示立即更改，SPIF_UPDATEINIFILE 表示更新配置文件，第二個參數1、0切換windosw預設、無音效
	int flag = Mute ? 0 : 1;
	SystemParametersInfo(SPI_SETSOUNDSENTRY, flag, NULL, SPIF_SENDCHANGE | SPIF_UPDATEINIFILE);
}

void __cdecl fft_get_thd_n_db(double* thd_n, double* dB_ValueMax, double* freq)
{
	auto& parm = ASAudio::GetInstance().GetWaveParm();
	thd_n[0] = parm.leftWAVE_ANALYSIS.thd_N_dB;
	thd_n[1] = parm.rightWAVE_ANALYSIS.thd_N_dB;

	dB_ValueMax[0] = 10 * log10(parm.leftWAVE_ANALYSIS.fundamentalEnergy);
	dB_ValueMax[1] = 10 * log10(parm.rightWAVE_ANALYSIS.fundamentalEnergy);

	freq[0] = parm.leftWAVE_ANALYSIS.TotalEnergyPoint * 20;
	freq[1] = parm.rightWAVE_ANALYSIS.TotalEnergyPoint * 20;
	/*
	THD+N的dB值範圍
	專業音響設備：
	高品質的專業音響設備通常要求THD+N低於 -80 dB，這意味著失真和噪聲非常低。
	消費電子產品：
	家用音響和消費電子產品的THD+N一般在 -60 dB 到 -80 dB 之間。
	中等品質音響設備：
	THD+N值在 -60 dB 到 -40 dB 之間被認為是中等品質。
	低品質音響設備：
	THD+N值高於 -40 dB 可能表示失真和噪聲較高，音質較差。
	最大dB值範圍
	專業音響設備：
	專業設備的最大dB值（表示信號強度）通常在 80 dB 到 120 dB 之間。
	消費電子產品：
	家用音響和消費電子產品的最大dB值通常在 60 dB 到 100 dB 之間。
	中等品質音響設備：
	最大dB值在 70 dB 到 90 dB 之間。
	低品質音響設備：
	最大dB值低於 70 dB。
	*/
}

void __cdecl fft_get_mute_db(double* dB_ValueMax)
{
	auto& parm = ASAudio::GetInstance().GetWaveParm();
	dB_ValueMax[0] = 10 * log10(parm.leftMuteWAVE_ANALYSIS.TotalEnergy);
	dB_ValueMax[1] = 10 * log10(parm.rightMuteWAVE_ANALYSIS.TotalEnergy);
	/*
	靜音測試的最大dB值範圍
	專業音響設備：
	高品質的專業音響設備在靜音狀態下，最大dB值（背景噪聲電平）通常低於 -90 dB。
	消費電子產品：
	家用音響和消費電子產品的靜音最大dB值一般在 -70 dB 到 -90 dB 之間。
	中等品質音響設備：
	靜音最大dB值在 -60 dB 到 -80 dB 之間。
	低品質音響設備：
	靜音最大dB值高於 -60 dB 表示在靜音狀態下噪聲水平較高。
	*/
}

void __cdecl fft_get_snr(double* snr)
{
	// 計算SNR
	auto& parm = ASAudio::GetInstance().GetWaveParm();
	snr[0] = (parm.leftMuteWAVE_ANALYSIS.TotalEnergy > 0) ? (parm.leftWAVE_ANALYSIS.fundamentalEnergy / parm.leftMuteWAVE_ANALYSIS.TotalEnergy) : 0;
	snr[1] = (parm.rightMuteWAVE_ANALYSIS.TotalEnergy > 0) ? (parm.rightWAVE_ANALYSIS.fundamentalEnergy / parm.rightMuteWAVE_ANALYSIS.TotalEnergy) : 0;

	// 轉換為分貝，這裡的SNR是能量或功率的比值，所以用10 * log10
	double leftSNR_dB = (snr[0] > 0) ? (10 * log10(snr[0])) : -INFINITY;
	double rightSNR_dB = (snr[1] > 0) ? (10 * log10(snr[1])) : -INFINITY;

	// 更新 SNR 值
	snr[0] = leftSNR_dB;
	snr[1] = rightSNR_dB;
	/*
	SNR 測試的範圍
	專業音響設備：
	高品質的專業音響設備通常具有非常高的 SNR，範圍在 90 dB 到 120 dB 之間，有些高端設備甚至可以達到 120 dB 以上。
	消費電子產品：
	家用音響和消費電子產品的 SNR 範圍通常在 70 dB 到 100 dB 之間。
	中等品質音響設備：
	中等品質音響設備的 SNR 通常在 60 dB 到 80 dB 之間。
	低品質音響設備：
	低品質音響設備的 SNR 可能低於 60 dB，顯示信號與噪聲之間的差距較小，音質可能較差。
	SNR 測試的解釋
	高 SNR：
	高 SNR 值表示背景噪聲低，音頻信號清晰且強大。這在專業音響系統中尤為重要，因為它們要求高度清晰的音質和低噪聲。
	低 SNR：
	低 SNR 值表示背景噪聲相對較高，信號不夠清晰。在一些噪聲容忍度較高的應用中，這可能不會造成顯著影響，但在高要求的音頻應用中，低 SNR 可能會顯著降低音質。
	*/
}

// 外部C函數，用於執行FFT轉換
int __cdecl fft_thd_n_exe(short* leftAudioData, short* rightAudioData, double* leftSpectrumData, double* rightSpectrumData)
{
	auto& parm = ASAudio::GetInstance().GetWaveParm();
	parm.bMuteTest = false;
	MMRESULT result;

	result = ASAudio::GetInstance().SetMicSystemVolume();
	result = ASAudio::GetInstance().SetSpeakerSystemVolume();
	if (result != MMSYSERR_NOERROR)
		return result;

	result = ASAudio::GetInstance().StartRecordingAndDrawSpectrum();
	if (result != MMSYSERR_NOERROR)
		return result;

	std::unique_lock<std::mutex> lock(parm.recordingMutex);
	// 等待直到錄音完成
	//while簡寫 當isRecordingFinished == true 跳出
	parm.recordingFinishedCV.wait(lock, [&] { return parm.isRecordingFinished;  });
	parm.isRecordingFinished = false;

	ASAudio::GetInstance().StopRecording(); // 停止錄音
	ASAudio::GetInstance().stopPlayback(); // 停止播放檔

	const fftw_complex* leftSpectrum = ASAudio::GetInstance().GetLeftSpectrumPtr();
	const fftw_complex* rightSpectrum = ASAudio::GetInstance().GetRightSpectrumPtr();

	memcpy(leftAudioData, parm.WAVE_DATA.LeftAudioBuffer, BUFFER_SIZE * sizeof(short));
	memcpy(rightAudioData, parm.WAVE_DATA.RightAudioBuffer, BUFFER_SIZE * sizeof(short));

	double maxAmplitude = 0;
	for (int i = 0; i < BUFFER_SIZE; ++i) {
		double sample = static_cast<double>(leftAudioData[i]);
		if (fabs(sample) > maxAmplitude) {
			maxAmplitude = fabs(sample);
		}
	}
	maxAmplitude = max(maxAmplitude, 1.0); // 避免 log10(0) 的情況
	double maxAmplitude_dB = 20 * log10(maxAmplitude / 32767);

	for (int i = 0; i < BUFFER_SIZE; ++i) {
		leftSpectrumData[i] = sqrt(leftSpectrum[i][0] * leftSpectrum[i][0] + leftSpectrum[i][1] * leftSpectrum[i][1]) / BUFFER_SIZE;
		rightSpectrumData[i] = sqrt(rightSpectrum[i][0] * rightSpectrum[i][0] + rightSpectrum[i][1] * rightSpectrum[i][1]) / BUFFER_SIZE;
	}

	// 分析頻譜THD+N
	ASAudio::GetInstance().SpectrumAnalysis(leftSpectrumData, rightSpectrumData);
	return result;
}

// 外部C函數，用於執行FFT轉換
int __cdecl fft_mute_exe(bool MuteWaveOut, bool MuteWaveIn,
	short* leftAudioData, short* rightAudioData, double* leftSpectrumData, double* rightSpectrumData)
{
	auto& parm = ASAudio::GetInstance().GetWaveParm();
	parm.bMuteTest = true;
	MMRESULT result = MMSYSERR_NOERROR;
	if (MuteWaveIn)
		ASAudio::GetInstance().SetMicMute(true);
	if (MuteWaveOut)
		result = ASAudio::GetInstance().SetSpeakerSystemVolume();

	if (result != MMSYSERR_NOERROR)
		return result;
	result = ASAudio::GetInstance().StartRecordingAndDrawSpectrum();
	if (result != MMSYSERR_NOERROR)
		return result;

	std::unique_lock<std::mutex> lock(parm.recordingMutex);
	// 等待直到錄音完成
	//while簡寫 當isRecordingFinished == true 跳出
	parm.recordingFinishedCV.wait(lock, [&] { return parm.isRecordingFinished;  });
	parm.isRecordingFinished = false;

	ASAudio::GetInstance().StopRecording();// 停止錄音
	ASAudio::GetInstance().stopPlayback(); // 停止播放檔

	fftw_complex* leftSpectrum;
	fftw_complex* rightSpectrum;
	ASAudio::GetInstance().GetSpectrumData(leftSpectrum, rightSpectrum); // 獲取頻譜數據
	memcpy(leftAudioData, parm.WAVE_DATA.LeftAudioBuffer, BUFFER_SIZE * sizeof(short));
	memcpy(rightAudioData, parm.WAVE_DATA.RightAudioBuffer, BUFFER_SIZE * sizeof(short));
	double maxAmplitude = 0;
	for (int i = 0; i < BUFFER_SIZE; ++i) {
		double sample = static_cast<double>(leftAudioData[i]);
		if (fabs(sample) > maxAmplitude) {
			maxAmplitude = fabs(sample);
		}
	}
	maxAmplitude = maxAmplitude >= 32768 ? 32767 : maxAmplitude;
	double maxAmplitude_dB = 20 * log10(maxAmplitude / 32767);

	for (int i = 0; i < BUFFER_SIZE; i++) {
		leftSpectrumData[i] = sqrt((leftSpectrum[i][0] * leftSpectrum[i][0]) + (leftSpectrum[i][1] * leftSpectrum[i][1])) / BUFFER_SIZE;;
		rightSpectrumData[i] = sqrt((rightSpectrum[i][0] * rightSpectrum[i][0]) + (rightSpectrum[i][1] * rightSpectrum[i][1])) / BUFFER_SIZE;;
	}
	//分析頻譜
	ASAudio::GetInstance().SpectrumAnalysis(leftSpectrumData, rightSpectrumData);
	//復原音量
	parm.bMuteTest = false;
	if (MuteWaveIn)
		ASAudio::GetInstance().SetMicMute(false);
	if (MuteWaveOut)
		ASAudio::GetInstance().SetSpeakerSystemVolume();

	return result;
}
