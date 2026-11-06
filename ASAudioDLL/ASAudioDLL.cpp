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
double snr_result[2] = { 0 };

extern HINSTANCE g_hInst;

std::wstring GenerateOperatorDiagnosis(double* thd_n, double* fund_dbfs, double* freq, double* snr) {
	std::wstring msg = L"";
	// 1. 削峰破音 (音量過載)
	if ((fund_dbfs[0] >= -2.0 && thd_n[0] > -30.0) || (fund_dbfs[1] >= -2.0 && thd_n[1] > -30.0)) {
		msg = L"錯誤：音量過大導致破音！\n 動作：請調低系統的播放/錄音音量，再重測一次。";
	}
	// 2. 無聲或極小聲
	else if (fund_dbfs[0] < -20.0 || fund_dbfs[1] < -20.0) {
		msg = L"錯誤：麥克風沒收到聲音！\n 動作：請檢查 1.音源線是否脫落 2.治具接觸不良 3.孔被堵住。";
	}
	// 3. 頻率錯誤 或 接反
	else if (fabs(freq[0] - 2000.0) > 100.0 || fabs(freq[1] - 1000.0) > 100.0) {
		msg = L"錯誤：左右聲道接反，或收到雜音！\n動作：請檢查治具音源線是否插反。";
	}
	// 4. 一般雜音失真 (硬體或線材受損)
	else if (thd_n[0] > -55.0 || thd_n[1] > -55.0) {
		msg = L"錯誤：測出雜音或音質失真！\n動作：請更換良品線材交叉測試，確認是否為板子損壞。";
	}
	// 5. 底噪過高 (電流聲/接地不良)
	// 加入 != -999.0 判斷，確保只有在有測 SNR 時才報錯
	else if (snr != nullptr && snr[0] != -999.0 && (snr[0] < 75.0 || snr[1] < 75.0)) {
		msg = L"錯誤：背景底噪過高 (電流聲)！\n動作：請確認 USB 是否接觸不良，或機台有嚴重的漏電與接地干擾。";
	}
	else {
		msg = L"錯誤：未知異常！\n動作：請拿 Golden Sample 測試確認是機台問題還是板子問題。";
	}
	return msg;
}
extern "C" __declspec(dllexport) BSTR MacroTest(HWND ownerHwnd, const char* filePath)
{
	// 只要進入 MacroTest 就初始化 COM，離開這個函式 (包含 return) 就會自動釋放
	struct AutoComInit {
		bool bInit;
		AutoComInit() {
			HRESULT hr = CoInitializeEx(NULL, COINIT_MULTITHREADED);
			bInit = SUCCEEDED(hr) || hr == RPC_E_CHANGED_MODE;
		}
		~AutoComInit() {
			if (bInit) CoUninitialize();
		}
	} autoCom;

	bool shouldRetry;
	std::string finalResult;
	ASAudio& audioInstance = ASAudio::GetInstance(); // 獲取單例

	do {
		shouldRetry = false;
		audioInstance.StopRecording();
		audioInstance.stopPlayback();
		audioInstance.StopAudioLoopback();
		audioInstance.ClearLog();

		thd_n[0] = thd_n[1] = 0.0;
		FundamentalLevel_dBFS[0] = FundamentalLevel_dBFS[1] = 0.0;
		freq[0] = freq[1] = 0.0;
		snr_result[0] = snr_result[1] = 0.0;

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
		// 1. 判斷是否包含錯誤字串
		bool hasError = (finalResult.find("ERROR") != std::string::npos);

		// 2. 定義哪些功能屬於「品質測試」(需要顯示圖表 UI)
		// 只有 AudioTest 或 SNRTest 啟動時發生的錯誤才需要彈出失敗視窗
		bool isQualityTest = (configs.AudioTestEnable || configs.snrTestEnable);

		// 3. 只有在「是品質測試」且「有錯誤」且「有視窗句柄」時才顯示 Dialog
		if (hasError && ownerHwnd != NULL && isQualityTest) {

			std::string strTemp = "失敗 (Failed): " + finalResult;
			// 準備傳給 UI 的資料
			std::wstring operatorDiagnosis = GenerateOperatorDiagnosis(thd_n, FundamentalLevel_dBFS, freq, snr_result);

			AudioData data;
			data.bufferSize = BUFFER_SIZE;
			data.errorMessage = operatorDiagnosis.c_str();
			// 傳入設定和結果的指標
			data.config = &configs;
			data.thd_n_result = thd_n;
			data.FundamentalLevel_dBFS_result = FundamentalLevel_dBFS;
			data.freq_result = freq;
			data.snr_result = snr_result;

			// 填充 THD+N 的圖表資料
			data.leftAudioData = leftAudioData;
			data.rightAudioData = rightAudioData;
			data.leftSpectrumData = leftSpectrumData;
			data.rightSpectrumData = rightSpectrumData;

			// 填充 SNR 的圖表資料
			data.leftAudioData_SNR = leftAudioData_SNR;
			data.rightAudioData_SNR = rightAudioData_SNR;
			data.leftSpectrumData_SNR = leftSpectrumData_SNR;
			data.rightSpectrumData_SNR = rightSpectrumData_SNR;

			// 從 audioInstance 獲取並填充實際裝置名稱
			data.actualOutDeviceName = audioInstance.GetWaveParm().ActualWaveOutDev;
			data.actualInDeviceName = audioInstance.GetWaveParm().ActualWaveInDev;

			// 呼叫 ShowFailureDialog
			if (ShowFailureDialog(g_hInst, ownerHwnd, &data)) {
				shouldRetry = true;
			}
		}

	} while (shouldRetry);

	audioInstance.StopRecording();
	audioInstance.stopPlayback();
	audioInstance.StopAudioLoopback();

	if (finalResult.empty()) {
		finalResult = "LOG:PASS#";
	}
	//格式化輸出結果
	size_t hashPos = finalResult.find('#');
	std::string baseResult = finalResult;
	if (hashPos != std::string::npos) {
		baseResult = finalResult.substr(0, hashPos);
	}

	bool isPass = (baseResult.find("LOG:PASS") != std::string::npos);

	// 使用一個獨立的 stringstream 來收集 JSON 資料
	std::stringstream jsonData;
	bool hasData = false;
	bool first = true;

	// 輔助函式：用來組合 JSON 鍵值對
	auto appendData = [&](const std::string& key, double value) {
		if (!first) jsonData << ",";
		jsonData << "\"" << key << "\":\"" << value << "\"";
		first = false;
		hasData = true; // 標記有成功加入至少一筆資料
		};

	// isValid：判斷數值是否有效 (過濾掉預設值 0.0 與失敗值 -999.0)
	auto isValid = [](double v) { return v != 0.0 && v != -999.0; };

	// 如果是 PASS，全部帶出；如果是 ERROR，只帶出有成功取到 (有效) 的數值
	if (isPass || isValid(freq[0])) appendData("Freq_L", freq[0]);
	if (isPass || isValid(freq[1])) appendData("Freq_R", freq[1]);

	if (isPass || isValid(FundamentalLevel_dBFS[0])) appendData("Level_L", FundamentalLevel_dBFS[0]);
	if (isPass || isValid(FundamentalLevel_dBFS[1])) appendData("Level_R", FundamentalLevel_dBFS[1]);

	if (isPass || isValid(thd_n[0])) appendData("THD+N_L", thd_n[0]);
	if (isPass || isValid(thd_n[1])) appendData("THD+N_R", thd_n[1]);

	if (isPass || isValid(snr_result[0])) appendData("SNR_L", snr_result[0]);
	if (isPass || isValid(snr_result[1])) appendData("SNR_R", snr_result[1]);

	// 如果有成功加入資料，才將 #DATA:{...} 串接回 finalResult
	if (hasData) {
		finalResult = baseResult + "#DATA:{" + jsonData.str() + "}";
	}
	else {
		// 如果沒有任何資料，就保持原本的 baseResult 並加上 '#' 作為結尾
		finalResult = baseResult + "#";
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
	double fullScaleEnergy = 32767.0 * 32767.0; // 加入滿刻度基準----------------------------------

	FundamentalLevel_dBFS[0] = (parm.leftMuteWAVE_ANALYSIS.TotalEnergy > 0) ?
		10 * log10(parm.leftMuteWAVE_ANALYSIS.TotalEnergy / fullScaleEnergy) : -INFINITY;

	FundamentalLevel_dBFS[1] = (parm.rightMuteWAVE_ANALYSIS.TotalEnergy > 0) ?
		10 * log10(parm.rightMuteWAVE_ANALYSIS.TotalEnergy / fullScaleEnergy) : -INFINITY;
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
		if (snr[0] > MAX_SNR_CAP) snr[0] = MAX_SNR_CAP;
	}
	else {
		// 如果能量比為 0 (因為噪聲為 0)，直接賦予上限值
		snr[0] = MAX_SNR_CAP;
	}

	if (right_snr_ratio > 0) {
		snr[1] = 10 * log10(right_snr_ratio);
		if (snr[1] > MAX_SNR_CAP) snr[1] = MAX_SNR_CAP;
	}
	else {
		snr[1] = MAX_SNR_CAP;
	}
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

	WAVE_PARM::RecordingResult waitResult = ASAudio::GetInstance().WaitForRecording(3);

	if (waitResult != WAVE_PARM::RecordingResult::Success) {
		ASAudio::GetInstance().StopRecording(); 
		ASAudio::GetInstance().stopPlayback();

		if (waitResult == WAVE_PARM::RecordingResult::Timeout) {
			return MMSYSERR_INVALHANDLE;
		}
		else if (waitResult == WAVE_PARM::RecordingResult::DeviceError) {
			return MMSYSERR_INVALHANDLE;
		}
	}

	//ASAudio::GetInstance().ExecuteFft();
	ASAudio::GetInstance().StopRecording(); // 停止錄音
	ASAudio::GetInstance().stopPlayback(); // 停止播放檔


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
			return MMSYSERR_INVALHANDLE;
		}
		else if (waitResult == WAVE_PARM::RecordingResult::DeviceError) {
			return MMSYSERR_INVALHANDLE;
		}
	}

	ASAudio::GetInstance().StopRecording();// 停止錄音
	ASAudio::GetInstance().stopPlayback(); // 停止播放檔


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
