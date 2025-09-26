#include "ConfigManager.h"
#include "AudioManager.h"
#include "ASAudioDLL.h"
#include "ASAudio.h"
#include "FailureDialog.h" 
#include "resource.h"     
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

short leftAudioData_SNR[BUFFER_SIZE] = { 0 };
short rightAudioData_SNR[BUFFER_SIZE] = { 0 };
double leftSpectrumData_SNR[BUFFER_SIZE] = { 0 };
double rightSpectrumData_SNR[BUFFER_SIZE] = { 0 };

double thd_n[2] = { 0 };
double FundamentalLevel_dBFS[2] = { 0 };
double freq[2] = { 0 };

extern HINSTANCE g_hInst;

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
			audioInstance.GetWaveParm().fundamentalBandwidthHz = configs.fundamentalBandwidthHz;

			AudioManager audioManager;
			audioManager.ExecuteTestsFromConfig(configs);
			finalResult = audioManager.GetResultString(); // 總是獲取結果，無論成功或失敗
		}

		bool hasError = (finalResult.find("ERROR") != std::string::npos);
		if (hasError && ownerHwnd != NULL) {
			/*
			std::string strTemp = "Failed: " + finalResult;
			std::wstring wErrorMsg = audioInstance.charToWstring(strTemp.c_str()) + L"\n\n" + audioInstance.GetLog() + L"\n\n是否要重測？(Retry ?)";
			if (MessageBoxW(ownerHwnd, wErrorMsg.c_str(), L"測試失敗", MB_YESNO | MB_ICONERROR) == IDYES) {
				shouldRetry = true;
			}*/
			std::string strTemp = "失敗 (Failed): " + finalResult;
			std::wstring wErrorMsg = audioInstance.charToWstring(strTemp.c_str()) + L"\n" + audioInstance.GetLog();

			AudioData data;
			data.bufferSize = BUFFER_SIZE;
			data.errorMessage = wErrorMsg.c_str();
			// 傳入設定和結果的指標
			data.config = &configs;
			data.thd_n_result = thd_n;
			data.FundamentalLevel_dBFS_result = FundamentalLevel_dBFS;
			data.freq_result = freq;

			// 填充 THD+N 的圖表資料
			data.leftAudioData = leftAudioData;
			data.rightAudioData = rightAudioData;
			data.leftSpectrumData = leftSpectrumData;
			data.rightSpectrumData = rightSpectrumData;

			// <<< 新增：填充 SNR 的圖表資料 >>>
			data.leftAudioData_SNR = leftAudioData_SNR;
			data.rightAudioData_SNR = rightAudioData_SNR;
			data.leftSpectrumData_SNR = leftSpectrumData_SNR;
			data.rightSpectrumData_SNR = rightSpectrumData_SNR;

			// 呼叫 ShowFailureDialog，現在它擁有了兩組資料
			if (ShowFailureDialog(g_hInst, ownerHwnd, &data)) {
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

void __cdecl fft_get_thd_n_db(double* thd_n, double* FundamentalLevel_dBFS, double* freq)
{
	auto& parm = ASAudio::GetInstance().GetWaveParm();
	thd_n[0] = parm.leftWAVE_ANALYSIS.thd_N_dB;
	thd_n[1] = parm.rightWAVE_ANALYSIS.thd_N_dB;

	// 1. 從能量（振幅的平方）開根號，還原回振幅
	double leftAmplitude = sqrt(parm.leftWAVE_ANALYSIS.fundamentalEnergy);
	double rightAmplitude = sqrt(parm.rightWAVE_ANALYSIS.fundamentalEnergy);

	// 2. 以 16-bit 的最大值 32767.0 作為參考，計算 dBFS
	//    使用 20 * log10() 來轉換振幅
	double fullScaleReference = 32767.0;

	if (leftAmplitude > 0) {
		FundamentalLevel_dBFS[0] = 20 * log10(leftAmplitude / fullScaleReference);
	}
	else {
		FundamentalLevel_dBFS[0] = -INFINITY; // 靜音
	}

	if (rightAmplitude > 0) {
		FundamentalLevel_dBFS[1] = 20 * log10(rightAmplitude / fullScaleReference);
	}
	else {
		FundamentalLevel_dBFS[1] = -INFINITY; // 靜音
	}

	// 頻率
	const double sampleRate = 44100.0;
	const int N = BUFFER_SIZE;
	freq[0] = parm.leftWAVE_ANALYSIS.TotalEnergyPoint * (sampleRate / N);
	freq[1] = parm.rightWAVE_ANALYSIS.TotalEnergyPoint * (sampleRate / N);
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

void __cdecl fft_get_mute_db(double* FundamentalLevel_dBFS)
{
	auto& parm = ASAudio::GetInstance().GetWaveParm();
	FundamentalLevel_dBFS[0] = 10 * log10(parm.leftMuteWAVE_ANALYSIS.TotalEnergy);
	FundamentalLevel_dBFS[1] = 10 * log10(parm.rightMuteWAVE_ANALYSIS.TotalEnergy);
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
	auto& parm = ASAudio::GetInstance().GetWaveParm();

	// 計算能量比
	double left_snr_ratio = (parm.leftMuteWAVE_ANALYSIS.TotalEnergy > 0) ? (parm.leftWAVE_ANALYSIS.fundamentalEnergy / parm.leftMuteWAVE_ANALYSIS.TotalEnergy) : 0;
	double right_snr_ratio = (parm.rightMuteWAVE_ANALYSIS.TotalEnergy > 0) ? (parm.rightWAVE_ANALYSIS.fundamentalEnergy / parm.rightMuteWAVE_ANALYSIS.TotalEnergy) : 0;

	// 處理噪聲為零的極端情況 ---
	const double MAX_SNR_CAP = 144.0; // 設定一個非常高的 SNR 上限 (dB)，代表極好的結果

	if (left_snr_ratio > 0) {
		snr[0] = 10 * log10(left_snr_ratio);
	}
	else {
		// 如果能量比為 0 (因為噪聲為 0)，直接賦予上限值
		snr[0] = MAX_SNR_CAP;
	}

	if (right_snr_ratio > 0) {
		snr[1] = 10 * log10(right_snr_ratio);
	}
	else {
		snr[1] = MAX_SNR_CAP;
	}

	// 可以選擇性地對正常計算出的值也設定上限，避免出現過大的數字
	if (snr[0] > MAX_SNR_CAP) snr[0] = MAX_SNR_CAP;
	if (snr[1] > MAX_SNR_CAP) snr[1] = MAX_SNR_CAP;
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

	WAVE_PARM::RecordingResult waitResult = ASAudio::GetInstance().WaitForRecording(10);

	if (waitResult != WAVE_PARM::RecordingResult::Success) {
		ASAudio::GetInstance().StopRecording(); 
		ASAudio::GetInstance().stopPlayback();

		if (waitResult == WAVE_PARM::RecordingResult::Timeout) {
			return ERROR_CODE_UNSET_PARAMETER; 
		}
		else if (waitResult == WAVE_PARM::RecordingResult::DeviceError) {
			return ERROR_CODE_UNSET_PARAMETER; 
		}
	}

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

	if (MuteWaveOut) {
		// 確保任何可能在播放的聲音都停止
		ASAudio::GetInstance().stopPlayback();
	}

	if (MuteWaveIn)
		ASAudio::GetInstance().SetMicMute(true);
	result = ASAudio::GetInstance().StartRecordingOnly();

	if (result != MMSYSERR_NOERROR)
	{
		parm.bMuteTest = false;
		if (MuteWaveIn) 
			ASAudio::GetInstance().SetMicMute(false);
		return result;
	}

	WAVE_PARM::RecordingResult waitResult = ASAudio::GetInstance().WaitForRecording(10);
	if (waitResult != WAVE_PARM::RecordingResult::Success) {
		ASAudio::GetInstance().StopRecording(); 
		ASAudio::GetInstance().stopPlayback();

		// Restore state even on error
		parm.bMuteTest = false;
		if (MuteWaveIn) ASAudio::GetInstance().SetMicMute(false);
		if (MuteWaveOut) ASAudio::GetInstance().SetSpeakerSystemVolume();

		if (waitResult == WAVE_PARM::RecordingResult::Timeout) {
			return ERROR_CODE_UNSET_PARAMETER; 
		}
		else if (waitResult == WAVE_PARM::RecordingResult::DeviceError) {
			return ERROR_CODE_UNSET_PARAMETER; 
		}
	}

	ASAudio::GetInstance().StopRecording();// 停止錄音
	ASAudio::GetInstance().stopPlayback(); // 停止播放檔

	ASAudio::GetInstance().ExecuteFft();

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
