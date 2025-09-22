#include "ASAudioDLL.h"
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

// Global Variables
WAVE_PARM mWAVE_PARM;
WAVEHDR waveHdr; // WAVEHDR結構，用於錄音緩衝
HWAVEIN hWaveIn; // 音訊輸入設備的處理
HWAVEOUT hWaveOut; // 音訊輸出設備的處理
WAVEHDR WaveOutHeader;
std::vector<char> WaveOutData;

paTestData data = { 0 };

short leftAudioData[BUFFER_SIZE] = { 0 };
short rightAudioData[BUFFER_SIZE] = { 0 };
double leftSpectrumData[BUFFER_SIZE] = { 0 };
double rightSpectrumData[BUFFER_SIZE] = { 0 };
double thd_n[2] = { 0 };
double dB_ValueMax[2] = { 0 };
double freq[2] = { 0 };

IGraphBuilder* pGraph = nullptr;
ICaptureGraphBuilder2* pCaptureGraph = nullptr;
IMediaControl* pMediaControl = nullptr;
IBaseFilter* pAudioCapture = nullptr;
IBaseFilter* pAudioRenderer = nullptr;

bool bFirstWaveInFlag;
std::string strMacroResult;
std::wstring allContent;

// Forward Declarations
int GetIMMWaveOutDevice(IMMDeviceCollection* pDevices, std::wstring szOutDevName);
std::wstring Utf8ToWstring(const std::string& str);

extern "C" __declspec(dllexport) BSTR MacroTest(HWND ownerHwnd, const char* filePath)
{
	bool shouldRetry, hasError;
	do {
		shouldRetry = false; // 預設不重試
		strMacroResult.clear(); // 每次重試前清除結果
		allContent.clear();
		hasError = false;

		Config configs;
		if (!ASAudio::ReadConfig(filePath, configs)) {
			hasError = true;
		}

		// --- Check WaveOut Device ---
		if (!hasError && mWAVE_PARM.WaveOutDev != L"") {
			ULONGLONG timeout = GetTickCount64() + 5000;
			while (ASAudio::GetWaveOutDevice(mWAVE_PARM.WaveOutDev.c_str()) == -1) {
				if (GetTickCount64() > timeout) {
					strMacroResult = "LOG:ERROR_AUDIO_WAVEOUT_DEVICE_NOT_FIND#";
					hasError = true;
					break;
				}
				Sleep(100);
			}
			if (!hasError) {
				ASAudio::GetInstance().SetSpeakerSystemVolume();
			}
		}

		// --- Check WaveIn Device ---
		if (!hasError && mWAVE_PARM.WaveInDev != L"") {
			ULONGLONG timeout = GetTickCount64() + 5000;
			while (ASAudio::GetWaveInDevice(mWAVE_PARM.WaveInDev.c_str()) == -1) {
				if (GetTickCount64() > timeout) {
					strMacroResult = "LOG:ERROR_AUDIO_WAVEIN_DEVICE_NOT_FIND#";
					hasError = true;
					break;
				}
				Sleep(100);
			}
			if (!hasError) {
				ASAudio::GetInstance().SetMicSystemVolume();
			}
		}

		// --- Audio Test Logic ---
		if (!hasError && configs.AudioTestEnable) {
			memset(leftAudioData, 0, sizeof(leftAudioData));
			memset(rightAudioData, 0, sizeof(rightAudioData));
			memset(leftSpectrumData, 0, sizeof(leftSpectrumData));
			memset(rightSpectrumData, 0, sizeof(rightSpectrumData));
			memset(thd_n, 0, sizeof(thd_n));
			memset(dB_ValueMax, 0, sizeof(dB_ValueMax));
			memset(freq, 0, sizeof(freq));

			fft_thd_n_exe(leftAudioData, rightAudioData, leftSpectrumData, rightSpectrumData);
			fft_get_thd_n_db(thd_n, dB_ValueMax, freq);

			if (thd_n[0] > configs.thd_n || thd_n[1] > configs.thd_n) {
				std::stringstream ss;
				ss << "LOG:ERROR_AUDIO_THD_N_VALUE_" << "L_" << (int)thd_n[0] << "_R_" << (int)thd_n[1] << "#";
				strMacroResult = ss.str();
				hasError = true;
			}
			else if (dB_ValueMax[0] < configs.db_ValueMax || dB_ValueMax[1] < configs.db_ValueMax) {
				std::stringstream ss;
				ss << "LOG:ERROR_AUDIO_DB_MAX_VALUE_" << "L_" << (int)dB_ValueMax[0] << "_R_" << (int)dB_ValueMax[1] << "#";
				strMacroResult = ss.str();
				hasError = true;
			}
			else if (freq[0] != configs.frequencyL || freq[1] != configs.frequencyR) {
				std::stringstream ss;
				ss << "LOG:ERROR_AUDIO_FREQUENCY_" << "L_" << (int)freq[0] << "_R_" << (int)freq[1] << "#";
				strMacroResult = ss.str();
				hasError = true;
			}
			else {
				std::stringstream ss;
				ss << "LOG:SUCCESS_AUDIO_TEST_THD_N_" << "L_" << (int)thd_n[0] << "_R_" << (int)thd_n[1]
					<< "_dB_ValueMax_L_" << (int)dB_ValueMax[0] << "_R_" << (int)dB_ValueMax[1]
					<< "_Frequency_L_" << (int)freq[0] << "_R_" << (int)freq[1] << "#";
				strMacroResult = ss.str();
			}
		}

		// --- Play WAV File ---
		if (!hasError && configs.PlayWAVFileEnable) {
			if (!ASAudio::PlayWavFile(configs.AutoCloseWAVFile)) {
				hasError = true; // PlayWavFile 內部會設定 strMacroResult
			}
		}

		// --- Close WAV File ---
		if (!hasError && configs.CloseWAVFileEnable) {
			ASAudio::StopPlayingWavFile();
		}

		// --- Audio Loopback ---
		if (!hasError && configs.AudioLoopBackEnable) {
			if (configs.AudioLoopBackStart) {
				ShellExecute(0, L"open", L"control mmsys.cpl,,1", 0, 0, SW_SHOWNORMAL);
				if (!ASAudio::StartAudioLoopback(mWAVE_PARM.WaveInDev.c_str(), mWAVE_PARM.WaveOutDev.c_str())) {
					hasError = true; // StartAudioLoopback 內部會設定 strMacroResult
				}
			}
			else {
				ASAudio::StopAudioLoopback();
			}
		}

		// --- Switch Default Audio Device ---
		if (!hasError && configs.SwitchDefaultAudioEnable) {
			std::wstring deviceId;
			if (ASAudio::FindDeviceIdByName(configs, deviceId)) {
				if (!ASAudio::SetDefaultAudioPlaybackDevice(deviceId)) {
					strMacroResult = "LOG:ERROR_AUDIO_CHANGE_DEFAULT_DEVICE#";
					hasError = true;
				}
			}
			else { // FindDeviceIdByName 內部會設定 strMacroResult
				hasError = true;
			}
		}

		// 如果有任何測試失敗，且是在互動模式下，則跳出提示
		if (hasError) {
			if (ownerHwnd != NULL) { // 如果 ownerHwnd 不為 NULL，才顯示錯誤訊息
				std::string strTemp = "Failed: " + strMacroResult;
				std::wstring wErrorMsg = ASAudio::GetInstance().charToWstring(strTemp.c_str()) + L"\n\n" + allContent + L"\n\n是否要重測？(Retry ?)";
				if (MessageBoxW(ownerHwnd, wErrorMsg.c_str(), L"測試失敗", MB_YESNO | MB_ICONERROR) == IDYES) {
					shouldRetry = true; // 設定旗標以重新執行
				}
			}
		}
	} while (shouldRetry);

	// 如果 strMacroResult 為空 (例如所有功能都 disable)，給一個預設成功訊息
	if (strMacroResult.empty() || !hasError) {
		strMacroResult = "LOG:PASS#";
	}

	// 將 string 轉為 BSTR 回傳
	int len = MultiByteToWideChar(CP_UTF8, 0, strMacroResult.c_str(), -1, nullptr, 0);
	if (len == 0) {
		return SysAllocString(L"LOG:ERROR_AUDIO_OTHER#"); // 避免回傳 nullptr
	}

	std::wstring wstr(len, L'\0');
	MultiByteToWideChar(CP_UTF8, 0, strMacroResult.c_str(), -1, &wstr[0], len);
	return SysAllocString(wstr.c_str());
}

bool ASAudio::ReadConfig(const std::string& filePath, Config& config)
{
	std::ifstream file(filePath);
	if (!file.good())
	{
		strMacroResult = "LOG:ERROR_AUDIO_CONFIG_NOT_FIND#";
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

	iNumber = GetPrivateProfileIntA("AudioTest", "db_ValueMax", 0, filePath.c_str());
	config.db_ValueMax = iNumber;

	iNumber = GetPrivateProfileIntA("PlayWAVFile", "PlayWAVFileEnable", 0, filePath.c_str());
	config.PlayWAVFileEnable = iNumber == 1 ? true : false;
	iNumber = GetPrivateProfileIntA("PlayWAVFile", "CloseWAVFileEnable", 0, filePath.c_str());
	config.CloseWAVFileEnable = iNumber == 1 ? true : false;

	iNumber = GetPrivateProfileIntA("PlayWAVFile", "AutoCloseWAVFile", 0, filePath.c_str());
	config.AutoCloseWAVFile = iNumber == 1 ? true : false;

	GetPrivateProfileStringA("PlayWAVFile", "WAVFilePath", "", buffer, sizeof(buffer), filePath.c_str());
	config.WAVFilePath = buffer;

	iNumber = GetPrivateProfileIntA("AudioLoopBack", "AudioLoopBackEnable", 0, filePath.c_str());
	config.AudioLoopBackEnable = iNumber == 1 ? true : false;
	iNumber = GetPrivateProfileIntA("AudioLoopBack", "AudioLoopBackStart", 0, filePath.c_str());
	config.AudioLoopBackStart = iNumber == 1 ? true : false;

	fft_set_test_parm(
		config.outDeviceName.c_str()
		, config.outDeviceVolume
		, config.inDeviceName.c_str()
		, config.inDeviceVolume
		, config.frequencyL
		, config.frequencyR
		, config.waveOutDelay);
	mWAVE_PARM.AudioFile = ASAudio::GetInstance().charToWstring(config.WAVFilePath.c_str());

	iNumber = GetPrivateProfileIntA("SwitchDefaultAudio", "SwitchDefaultAudioEnable", 0, filePath.c_str());
	config.SwitchDefaultAudioEnable = iNumber == 1 ? true : false;
	config.monitorNames.clear();
	for (int i = 0;;i++)
	{
		std::string temp = "AudioName";
		if(i > 0)
			temp += std::to_string(i);
		GetPrivateProfileStringA("SwitchDefaultAudio", temp.c_str(), "", buffer, sizeof(buffer), filePath.c_str());
		if (strlen(buffer) == 0)
			break;
		config.monitorNames.push_back(ASAudio::GetInstance().charToWstring(buffer));
	}
	GetPrivateProfileStringA("SwitchDefaultAudio", "AudioName", "", buffer, sizeof(buffer), filePath.c_str());
	config.AudioName = buffer;
	iNumber = GetPrivateProfileIntA("SwitchDefaultAudio", "AudioIndex", 0, filePath.c_str());
	config.AudioIndex = iNumber;

	return true;
}

void __cdecl fft_set_test_parm(
	const char* szOutDevName, int WaveOutVolumeValue
	, const char* szInDevName, int WaveInVolumeValue
	, int frequencyL, int frequencyR
	, int WaveOutdelay)
{
	mWAVE_PARM.WaveOutDev = ASAudio::GetInstance().charToWstring(szOutDevName);
	mWAVE_PARM.WaveInDev = ASAudio::GetInstance().charToWstring(szInDevName);
	mWAVE_PARM.WaveOutVolume = WaveOutVolumeValue;
	mWAVE_PARM.WaveInVolume = WaveInVolumeValue;
	mWAVE_PARM.frequencyL = frequencyL;
	mWAVE_PARM.frequencyR = frequencyR;
	mWAVE_PARM.WaveOutDelay = WaveOutdelay;
}

int __cdecl fft_get_buffer_size()
{
	return BUFFER_SIZE;
}

void __cdecl fft_set_audio_volume(const char* szOutDevName, int WaveOutVolumeValue
	, const char* szInDevName, int WaveInVolumeValue)
{
	mWAVE_PARM.WaveOutDev = ASAudio::GetInstance().charToWstring(szOutDevName);
	mWAVE_PARM.WaveInDev = ASAudio::GetInstance().charToWstring(szInDevName);
	mWAVE_PARM.WaveOutVolume = WaveOutVolumeValue;
	mWAVE_PARM.WaveInVolume = WaveInVolumeValue;
	ASAudio::GetInstance().SetMicSystemVolume();
	ASAudio::GetInstance().SetSpeakerSystemVolume();
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
	thd_n[0] = mWAVE_PARM.leftWAVE_ANALYSIS.thd_N_dB;
	thd_n[1] = mWAVE_PARM.rightWAVE_ANALYSIS.thd_N_dB;

	dB_ValueMax[0] = 10 * log10(mWAVE_PARM.leftWAVE_ANALYSIS.fundamentalEnergy);
	dB_ValueMax[1] = 10 * log10(mWAVE_PARM.rightWAVE_ANALYSIS.fundamentalEnergy);

	freq[0] = mWAVE_PARM.leftWAVE_ANALYSIS.TotalEnergyPoint * 20;
	freq[1] = mWAVE_PARM.rightWAVE_ANALYSIS.TotalEnergyPoint * 20;
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
	dB_ValueMax[0] = 10 * log10(mWAVE_PARM.leftMuteWAVE_ANALYSIS.TotalEnergy);
	dB_ValueMax[1] = 10 * log10(mWAVE_PARM.rightMuteWAVE_ANALYSIS.TotalEnergy);
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
	snr[0] = (mWAVE_PARM.leftMuteWAVE_ANALYSIS.TotalEnergy > 0) ? (mWAVE_PARM.leftWAVE_ANALYSIS.fundamentalEnergy / mWAVE_PARM.leftMuteWAVE_ANALYSIS.TotalEnergy) : 0;
	snr[1] = (mWAVE_PARM.rightMuteWAVE_ANALYSIS.TotalEnergy > 0) ? (mWAVE_PARM.rightWAVE_ANALYSIS.fundamentalEnergy / mWAVE_PARM.rightMuteWAVE_ANALYSIS.TotalEnergy) : 0;

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
	mWAVE_PARM.bMuteTest = false;
	MMRESULT result;

	result = ASAudio::GetInstance().SetMicSystemVolume();
	result = ASAudio::GetInstance().SetSpeakerSystemVolume();
	if (result != MMSYSERR_NOERROR)
		return result;

	result = ASAudio::GetInstance().StartRecordingAndDrawSpectrum();
	if (result != MMSYSERR_NOERROR)
		return result;

	std::unique_lock<std::mutex> lock(mWAVE_PARM.recordingMutex);
	// 等待直到錄音完成
	//while簡寫 當isRecordingFinished == true 跳出
	mWAVE_PARM.recordingFinishedCV.wait(lock, [] { return mWAVE_PARM.isRecordingFinished;  });
	mWAVE_PARM.isRecordingFinished = false;

	ASAudio::GetInstance().StopRecording(); // 停止錄音
	ASAudio::GetInstance().stopPlayback(); // 停止播放檔

	// 確保 FFTW 結果指標被分配
	fftw_complex* leftSpectrum = (fftw_complex*)fftw_malloc(sizeof(fftw_complex) * BUFFER_SIZE);
	fftw_complex* rightSpectrum = (fftw_complex*)fftw_malloc(sizeof(fftw_complex) * BUFFER_SIZE);

	ASAudio::GetInstance().GetSpectrumData(leftSpectrum, rightSpectrum);// 獲取頻譜數據

	memcpy(leftAudioData, mWAVE_PARM.WAVE_DATA.LeftAudioBuffer, BUFFER_SIZE * sizeof(short));
	memcpy(rightAudioData, mWAVE_PARM.WAVE_DATA.RightAudioBuffer, BUFFER_SIZE * sizeof(short));

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
	mWAVE_PARM.bMuteTest = true;
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

	std::unique_lock<std::mutex> lock(mWAVE_PARM.recordingMutex);
	// 等待直到錄音完成
	//while簡寫 當isRecordingFinished == true 跳出
	mWAVE_PARM.recordingFinishedCV.wait(lock, [] { return mWAVE_PARM.isRecordingFinished;  });
	mWAVE_PARM.isRecordingFinished = false;

	ASAudio::GetInstance().StopRecording();// 停止錄音
	ASAudio::GetInstance().stopPlayback(); // 停止播放檔

	fftw_complex* leftSpectrum;
	fftw_complex* rightSpectrum;
	ASAudio::GetInstance().GetSpectrumData(leftSpectrum, rightSpectrum); // 獲取頻譜數據
	memcpy(leftAudioData, mWAVE_PARM.WAVE_DATA.LeftAudioBuffer, BUFFER_SIZE * sizeof(short));
	memcpy(rightAudioData, mWAVE_PARM.WAVE_DATA.RightAudioBuffer, BUFFER_SIZE * sizeof(short));
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
	mWAVE_PARM.bMuteTest = false;
	if (MuteWaveIn)
		ASAudio::GetInstance().SetMicMute(false);
	if (MuteWaveOut)
		ASAudio::GetInstance().SetSpeakerSystemVolume();

	return result;
}

void ASAudio::SpectrumAnalysis(double* leftSpectrumData, double* rightSpectrumData)
{
	const double Fs = 44100; // 取樣頻率
	const int freqBins = BUFFER_SIZE / 2 + 1; // 頻率bin數量
	double leftMagnitude, rightMagnitude; // 當前頻率bin的幅度

	auto& leftTempWAVE_ANALYSIS = mWAVE_PARM.bMuteTest ? mWAVE_PARM.leftMuteWAVE_ANALYSIS : mWAVE_PARM.leftWAVE_ANALYSIS;
	auto& RightTempWAVE_ANALYSIS = mWAVE_PARM.bMuteTest ? mWAVE_PARM.rightMuteWAVE_ANALYSIS : mWAVE_PARM.rightWAVE_ANALYSIS;
	leftTempWAVE_ANALYSIS.TotalEnergyPoint = 0; // 左聲道最大能量點初始化
	RightTempWAVE_ANALYSIS.TotalEnergyPoint = 0; // 右聲道最大能量點初始化
	leftTempWAVE_ANALYSIS.TotalEnergy = 0; // 左聲道總能量初始化
	RightTempWAVE_ANALYSIS.TotalEnergy = 0; // 右聲道總能量初始化

	double maxLeftMagnitude = 0;
	double maxRightMagnitude = 0;

	// 計算總能量和最大能量點
	for (int i = 2; i < freqBins; ++i) {//頻率0忽略不計
		leftMagnitude = leftSpectrumData[i]; // 當前左聲道頻率bin的幅度
		rightMagnitude = rightSpectrumData[i]; // 當前右聲道頻率bin的幅度

		// 累加所有頻率成分的能量（幅度的平方）
		leftTempWAVE_ANALYSIS.TotalEnergy += leftMagnitude * leftMagnitude;
		RightTempWAVE_ANALYSIS.TotalEnergy += rightMagnitude * rightMagnitude;

		// 更新最大能量點和最大值
		if (leftMagnitude > maxLeftMagnitude) {
			maxLeftMagnitude = leftMagnitude;
			leftTempWAVE_ANALYSIS.TotalEnergyPoint = i;
		}
		if (rightMagnitude > maxRightMagnitude) {
			maxRightMagnitude = rightMagnitude;
			RightTempWAVE_ANALYSIS.TotalEnergyPoint = i;
		}
	}

	double Harmonic1, Harmonic2; // 諧波幅度
	double totalHarmonicEnergy; // 諧波總能量

	// 計算THD+N 左聲道
	int energyPoint = leftTempWAVE_ANALYSIS.TotalEnergyPoint;
	if (energyPoint * 2 < freqBins) {
		Harmonic1 = leftSpectrumData[energyPoint * 2];
		Harmonic2 = (energyPoint * 3 < freqBins) ? leftSpectrumData[energyPoint * 3] : 0;

		totalHarmonicEnergy = Harmonic1 * Harmonic1 + Harmonic2 * Harmonic2;
		// 取得基波能量
		double fundamentalEnergy = leftSpectrumData[energyPoint] * leftSpectrumData[energyPoint];
		leftTempWAVE_ANALYSIS.fundamentalEnergy = fundamentalEnergy;
		// 計算THD+N
		leftTempWAVE_ANALYSIS.thd_N = sqrt(totalHarmonicEnergy / fundamentalEnergy);

		// 轉換為分貝
		leftTempWAVE_ANALYSIS.thd_N_dB = (leftTempWAVE_ANALYSIS.thd_N > 0)
			? 20 * log10(leftTempWAVE_ANALYSIS.thd_N)
			: -INFINITY; // 處理0或負值情況
	}
	else {
		leftTempWAVE_ANALYSIS.thd_N_dB = 0;
	}

	// 計算THD+N 右聲道
	energyPoint = RightTempWAVE_ANALYSIS.TotalEnergyPoint;
	if (energyPoint * 2 < freqBins) {
		Harmonic1 = rightSpectrumData[energyPoint * 2];
		Harmonic2 = (energyPoint * 3 < freqBins) ? rightSpectrumData[energyPoint * 3] : 0;

		totalHarmonicEnergy = Harmonic1 * Harmonic1 + Harmonic2 * Harmonic2;
		// 取得基波能量
		double fundamentalEnergy = rightSpectrumData[energyPoint] * rightSpectrumData[energyPoint];
		RightTempWAVE_ANALYSIS.fundamentalEnergy = fundamentalEnergy;
		// 計算THD+N
		RightTempWAVE_ANALYSIS.thd_N = sqrt(totalHarmonicEnergy / fundamentalEnergy);

		// 轉換為分貝
		RightTempWAVE_ANALYSIS.thd_N_dB = (RightTempWAVE_ANALYSIS.thd_N > 0)
			? 20 * log10(RightTempWAVE_ANALYSIS.thd_N)
			: -INFINITY; // 處理0或負值情況
	}
	else {
		RightTempWAVE_ANALYSIS.thd_N_dB = 0;
	}

	double totalEnergy = 0;
	for (int i = 0; i < freqBins; ++i) {
		double magnitude = leftSpectrumData[i];
		totalEnergy += magnitude * magnitude; // 能量是幅度的平方
	}
	double volume_dB = 10 * log10(totalEnergy);
}

ASAudio & ASAudio::GetInstance()
{
	static ASAudio instance([](fftw_complex* leftSpectrum, fftw_complex* rightSpectrum) {
				// This is a dummy callback, as the original global instance had one.
			});
	return instance;
}

// ASAudio類的構造函數
ASAudio::ASAudio(SpectrumAnalysisCallback callback) : spectrumCallback(callback)
{
	// 初始化音訊輸入和輸出相關變數
	hWaveIn = NULL; // 音訊輸入設備句柄
	hWaveOut = NULL; // 音訊輸出設備句柄
	waveInCallback = nullptr; // 音訊輸入回呼函式
	memset(&waveHdr, 0, sizeof(WAVEHDR)); // 初始化音訊輸入回呼用的WAVEHDR結構
	memset(mWAVE_PARM.WAVE_DATA.audioBuffer, 0, BUFFER_SIZE * 2 * sizeof(short)); // 初始化音訊緩衝，兩聲道
	bufferIndex = 0; // 音訊緩衝索引

	// 初始化 FFT 所需的變數和計劃
	fftInputBufferLeft = (fftw_complex*)fftw_malloc(sizeof(fftw_complex) * BUFFER_SIZE);
	fftInputBufferRight = (fftw_complex*)fftw_malloc(sizeof(fftw_complex) * BUFFER_SIZE);
	fftOutputBufferLeft = (fftw_complex*)fftw_malloc(sizeof(fftw_complex) * BUFFER_SIZE);
	fftOutputBufferRight = (fftw_complex*)fftw_malloc(sizeof(fftw_complex) * BUFFER_SIZE);

	// 創建 FFT 計劃，用於對輸入進行傅立葉變換
	fftPlanLeft = fftw_plan_dft_1d(BUFFER_SIZE, fftInputBufferLeft, fftOutputBufferLeft, FFTW_FORWARD, FFTW_ESTIMATE);
	fftPlanRight = fftw_plan_dft_1d(BUFFER_SIZE, fftInputBufferRight, fftOutputBufferRight, FFTW_FORWARD, FFTW_ESTIMATE);
}

// ASAudio類的析構函數
ASAudio::~ASAudio() {
	fftw_destroy_plan(fftPlanLeft);
	fftw_destroy_plan(fftPlanRight);
	fftw_free(fftInputBufferLeft);
	fftw_free(fftInputBufferRight);
	fftw_free(fftOutputBufferLeft);
	fftw_free(fftOutputBufferRight);
}

// 播放WAV檔案
bool ASAudio::PlayWavFile(bool AutoClose) {
	MMCKINFO ChunkInfo;
	MMCKINFO FormatChunkInfo = { 0 };
	MMCKINFO DataChunkInfo = { 0 };
	WAVEFORMATEX wfx;
	int DataSize;

	// Zero out the ChunkInfo structure.
	memset(&ChunkInfo, 0, sizeof(MMCKINFO));

	// 打開WAVE文件，返回一個HMMIO句柄
	HMMIO handle = mmioOpen((LPWSTR)mWAVE_PARM.AudioFile.c_str(), 0, MMIO_READ);

	// 進入RIFF區塊(RIFF Chunk)
	mmioDescend(handle, &ChunkInfo, 0, MMIO_FINDRIFF);

	// 進入fmt區塊(RIFF子塊，包含音訊結構的信息)
	FormatChunkInfo.ckid = mmioStringToFOURCCA("fmt", 0); // 尋找fmt子塊
	mmioDescend(handle, &FormatChunkInfo, &ChunkInfo, MMIO_FINDCHUNK);

	// 讀取wav結構信息
	memset(&wfx, 0, sizeof(WAVEFORMATEX));
	mmioRead(handle, (char*)&wfx, FormatChunkInfo.cksize);

	// 檢查聲道數量
	if (wfx.nChannels != 2) {
		// 錯誤：不是雙聲道
		mmioClose(handle, 0);
		strMacroResult = "LOG:ERROR_AUDIO_WAVE_FILE_FORMAT_NOT_SUPPORT#";
		return false;
	}

	// 跳出fmt區塊
	mmioAscend(handle, &FormatChunkInfo, 0);

	// 進入data區塊(包含所有的數據波形)
	DataChunkInfo.ckid = mmioStringToFOURCCA("data", 0);
	mmioDescend(handle, &DataChunkInfo, &ChunkInfo, MMIO_FINDCHUNK);

	// 獲得data區塊的大小
	DataSize = DataChunkInfo.cksize;

	// 分配緩衝區
	WaveOutData.resize(DataSize);
	mmioRead(handle, WaveOutData.data(), DataSize);

	// 打開輸出設備
	int iDevIndex = GetWaveOutDevice(mWAVE_PARM.WaveOutDev);
	MMRESULT result = waveOutOpen(&hWaveOut, iDevIndex, &wfx, 0, 0, WAVE_FORMAT_QUERY);
	if (result != MMSYSERR_NOERROR) {
		// 錯誤
		strMacroResult = "LOG:ERROR_AUDIO_WAVEOUT_DEVICE_NOT_FIND_2#";
		return false;
	}

	result = waveOutOpen(&hWaveOut, iDevIndex, &wfx, NULL, NULL, CALLBACK_NULL);
	if (result != MMSYSERR_NOERROR) {
		// 錯誤
		waveOutClose(hWaveOut);
		strMacroResult = "LOG:ERROR_AUDIO_WAVEOUT_DEVICE_NOT_FIND_3#";
		return false;
	}

	// 設置wave header.
	memset(&WaveOutHeader, 0, sizeof(WaveOutHeader));
	WaveOutHeader.lpData = WaveOutData.data();
	WaveOutHeader.dwBufferLength = DataSize;

	// 準備wave header.
	waveOutPrepareHeader(hWaveOut, &WaveOutHeader, sizeof(WAVEHDR));

	// 將緩衝區資料寫入撥放設備(開始撥放).
	waveOutWrite(hWaveOut, &WaveOutHeader, sizeof(WAVEHDR));
	if (AutoClose)
	{
		// 等待播放完成
		while (!(WaveOutHeader.dwFlags & WHDR_DONE)) {
			Sleep(100); // 避免 CPU 過度佔用
		}

		// 停止播放並重置管理器
		ASAudio::StopPlayingWavFile();
	}
	return true;
}

// 停止播放WAV檔案
void ASAudio::StopPlayingWavFile() {
	// 停止播放並重置管理器
	waveOutReset(hWaveOut);
	// 關閉撥放設備
	waveOutClose(hWaveOut);
	// 清理用WaveOutPrepareHeader準備的Wave。
	waveOutUnprepareHeader(hWaveOut, &WaveOutHeader, sizeof(WAVEHDR));
	WaveOutData.clear();
}

// 開始錄音
MMRESULT ASAudio::StartRecordingAndDrawSpectrum() {
	// 設置音訊格式
	WAVEFORMATEX format{};
	format.wFormatTag = WAVE_FORMAT_PCM;
	format.nChannels = 2;
	format.nSamplesPerSec = 44100;
	format.wBitsPerSample = 16;
	format.nBlockAlign = format.nChannels * format.wBitsPerSample / 8;
	format.nAvgBytesPerSec = format.nSamplesPerSec * format.nBlockAlign;
	format.cbSize = 0;

	MMRESULT result;
	// 清理音訊緩衝區，避免遺留數據
	memset(mWAVE_PARM.WAVE_DATA.audioBuffer, 0, BUFFER_SIZE * sizeof(short) * 2);
	// 打開音訊輸入設備
	int iDevIndex = GetWaveInDevice(mWAVE_PARM.WaveInDev);
	if (iDevIndex == -1)
	{
		data.errorcode = ERROR_CODE_NOT_FIND_WAVEIN_DEV;
		result = ERROR_CODE_NOT_FIND_WAVEIN_DEV;
	}
	else
	{
		mWAVE_PARM.firstBufferDiscarded = false;
		result = waveInOpen(&hWaveIn, iDevIndex, &format, (DWORD_PTR)WaveInProc, (DWORD_PTR)this, CALLBACK_FUNCTION | WAVE_FORMAT_DIRECT);
		if (result == MMSYSERR_NOERROR) {
			result = playPortAudioSound();
			if (result == MMSYSERR_NOERROR)
			{
				Sleep(100);
				Sleep(mWAVE_PARM.WaveOutDelay);
				bFirstWaveInFlag = true;
				waveHdr.lpData = (LPSTR)mWAVE_PARM.WAVE_DATA.audioBuffer;
				waveHdr.dwBufferLength = BUFFER_SIZE * sizeof(short) * 2;
				waveHdr.dwBytesRecorded = 0;
				waveHdr.dwUser = 0;
				waveHdr.dwFlags = 0;
				waveHdr.dwLoops = 0;
				waveInPrepareHeader(hWaveIn, &waveHdr, sizeof(WAVEHDR));
				waveInAddBuffer(hWaveIn, &waveHdr, sizeof(WAVEHDR));
				waveInStart(hWaveIn);
			}
		}
		else
		{
			return result;
		}
	}
	return data.errorcode;
}

// 停止錄音
MMRESULT ASAudio::StopRecording() {
	MMRESULT result = waveInStop(hWaveIn); // Stop recording
	if (result == MMSYSERR_NOERROR) {
		result = waveInReset(hWaveIn);
		result = waveInClose(hWaveIn); // Close the recording device
	}
	waveInPrepareHeader(hWaveIn, &waveHdr, sizeof(WAVEHDR));
	VirtualFree(&waveHdr.lpData, 0, MEM_RELEASE);
	return result;
}

// 錄音回調函數
void CALLBACK ASAudio::WaveInProc(HWAVEIN hwi, UINT uMsg, DWORD_PTR dwInstance, DWORD_PTR dwParam1, DWORD_PTR dwParam2) {
	if (uMsg == MM_WIM_DATA) {
		ASAudio* asAudio = reinterpret_cast<ASAudio*>(dwInstance);
		WAVEHDR* waveHdr = reinterpret_cast<WAVEHDR*>(dwParam1);
		// 處理緩衝區中的音訊
		if (waveHdr->dwBytesRecorded > 0) {
			if (bFirstWaveInFlag)
			{
				bFirstWaveInFlag = false;
				// 為求錄製音源乾淨，捨棄第一筆資料，重新準備並添加新的緩衝區
				waveInUnprepareHeader(hwi, waveHdr, sizeof(WAVEHDR));
				waveInPrepareHeader(hwi, waveHdr, sizeof(WAVEHDR));
				waveInAddBuffer(hwi, waveHdr, sizeof(WAVEHDR));
			}
			else
			{
				asAudio->ProcessAudioBuffer();
			}
		}
	}
}

// 處理音訊緩衝區
void ASAudio::ProcessAudioBuffer() {
	short* audioSamples = reinterpret_cast<short*>(waveHdr.lpData);
	for (int i = 0; i < BUFFER_SIZE; i++) {
		mWAVE_PARM.WAVE_DATA.audioBuffer[bufferIndex + i * 2] = reinterpret_cast<short*>(waveHdr.lpData)[i * 2];
		mWAVE_PARM.WAVE_DATA.audioBuffer[bufferIndex + i * 2 + 1] = reinterpret_cast<short*>(waveHdr.lpData)[i * 2 + 1];
	}
	bufferIndex += BUFFER_SIZE * 2;

	if (bufferIndex >= BUFFER_SIZE * 2) {
		// 錄音存到雙聲道資訊滿(BUFFER_SIZE*2)
		// 填充 FFT 緩衝區
		for (int i = 0; i < BUFFER_SIZE; i++) {
			fftInputBufferLeft[i][0] = mWAVE_PARM.WAVE_DATA.audioBuffer[i * 2];
			fftInputBufferLeft[i][1] = 0.0;
			fftInputBufferRight[i][0] = mWAVE_PARM.WAVE_DATA.audioBuffer[i * 2 + 1];
			fftInputBufferRight[i][1] = 0.0;

			mWAVE_PARM.WAVE_DATA.LeftAudioBuffer[i] = mWAVE_PARM.WAVE_DATA.audioBuffer[i * 2];
			mWAVE_PARM.WAVE_DATA.RightAudioBuffer[i] = mWAVE_PARM.WAVE_DATA.audioBuffer[i * 2 + 1];
		}

		// 執行FFT轉換
		fftw_execute(fftPlanLeft);
		fftw_execute(fftPlanRight);

		// 通知已完成錄音
		std::lock_guard<std::mutex> lock(mWAVE_PARM.recordingMutex);
		mWAVE_PARM.isRecordingFinished = true;
		mWAVE_PARM.recordingFinishedCV.notify_one();

		bufferIndex = 0;
	}
}

DWORD ASAudio::SetMicSystemVolume()
{
	//初始化相關結構
	MMRESULT rc;
	HMIXER hMixer;
	MIXERLINE mxl;
	MIXERLINECONTROLS mxlc;
	MIXERCONTROL mxc;

	//遍歷系統的混音器，直到找到麥克風的混音器，紀錄該設備ID
	for (int deviceID = 0; true; deviceID++)
	{
		//打開當前的混音器，deviceID為混音器的ID
		rc = mixerOpen(&hMixer, deviceID, 0, 0, MIXER_OBJECTF_MIXER);
		if (rc != MMSYSERR_NOERROR)
			break;
		ZeroMemory(&mxl, sizeof(MIXERLINE));
		mxl.cbStruct = sizeof(MIXERLINE);
		//指出需要獲取的通道
		//聲道的音頻輸出用MIXERLINE_COMPONENTTYPE_DST_SPEAKERS
		//聲道的音頻輸入用MIXERLINE_COMPONENTTYPE_DST_WAVEIN
		mxl.dwComponentType = MIXERLINE_COMPONENTTYPE_DST_WAVEIN;
		rc = mixerGetLineInfo((HMIXEROBJ)hMixer, &mxl, MIXER_GETLINEINFOF_COMPONENTTYPE);
		if (rc != MMSYSERR_NOERROR)
			continue;
		//取得混音器設備的指定線路信息成功的話，則將連結數保存
		DWORD dwConnections = mxl.cConnections;
		DWORD dwLineID = -1;
		for (DWORD i = 0; i < dwConnections; i++)
		{
			mxl.dwSource = i;
			//根據SourceID獲得連結的信息
			rc = mixerGetLineInfo((HMIXEROBJ)hMixer, &mxl, MIXER_OBJECTF_HMIXER | MIXER_GETLINEINFOF_SOURCE);
			if (rc == MMSYSERR_NOERROR)
			{
				//如果當前設備類型為麥克風，跳出循環
				//MIXERLINE_COMPONENTTYPE_DST_SPEAKERS
				//MIXERLINE_COMPONENTTYPE_SRC_MICROPHONE
				if (mxl.dwComponentType == MIXERLINE_COMPONENTTYPE_SRC_MICROPHONE)
				{
					dwLineID = mxl.dwLineID;
					break;
				}
			}
		}
		if (dwLineID == -1)
			continue;
		ZeroMemory(&mxc, sizeof(MIXERCONTROL));
		mxc.cbStruct = sizeof(mxc);
		ZeroMemory(&mxlc, sizeof(MIXERLINECONTROLS));

		//控制類型為控制音量
		mxlc.cbStruct = sizeof(mxlc);
		mxlc.dwLineID = dwLineID;
		mxlc.dwControlType = MIXERCONTROL_CONTROLTYPE_VOLUME;
		mxlc.cControls = 1;
		mxlc.pamxctrl = &mxc;
		mxlc.cbmxctrl = sizeof(mxc);

		//取得控制器信息
		rc = mixerGetLineControls((HMIXEROBJ)hMixer, &mxlc, MIXER_GETLINECONTROLSF_ONEBYTYPE);
		if (MMSYSERR_NOERROR == rc)
		{
			MIXERCONTROLDETAILS mxcd;
			MIXERCONTROLDETAILS_SIGNED volStruct;

			ZeroMemory(&mxcd, sizeof(mxcd));
			mxcd.cbStruct = sizeof(mxcd);
			mxcd.dwControlID = mxc.dwControlID;
			mxcd.paDetails = &volStruct;
			mxcd.cbDetails = sizeof(volStruct);
			mxcd.cChannels = 1;

			//獲得音量值，取得的信息放在mxcd中
			rc = mixerGetControlDetails((HMIXEROBJ)hMixer, &mxcd, MIXER_GETCONTROLDETAILSF_VALUE);

			//初始化錄音大小的信息
			MIXERCONTROLDETAILS_UNSIGNED mxcdVolume_Set = { mxc.Bounds.dwMaximum * mWAVE_PARM.WaveInVolume / 100 };
			MIXERCONTROLDETAILS mxcd_Set = { 0 };
			mxcd_Set.cbStruct = sizeof(MIXERCONTROLDETAILS);
			mxcd_Set.dwControlID = mxc.dwControlID;
			mxcd_Set.cChannels = 1;
			mxcd_Set.cMultipleItems = 0;
			mxcd_Set.cbDetails = sizeof(MIXERCONTROLDETAILS_UNSIGNED);
			mxcd_Set.paDetails = &mxcdVolume_Set;

			//設置錄音大小
			mixerSetControlDetails((HMIXEROBJ)(hMixer), &mxcd_Set, MIXER_OBJECTF_HMIXER | MIXER_SETCONTROLDETAILSF_VALUE);
			mixerClose(hMixer);
		}
	}
	return 0;
}

DWORD ASAudio::SetSpeakerSystemVolume()
{
	HRESULT hrInit = CoInitialize(NULL);
	if (FAILED(hrInit)) {
	}

	HRESULT hr;
	IMMDeviceEnumerator* pEnumerator = NULL;
	IMMDeviceCollection* pDevices = NULL;
	IMMDevice* pDevice = NULL;
	IAudioEndpointVolume* pAudioEndpointVolume = NULL;
	DWORD dwResult = ERROR_CODE_SET_SPEAKER_SYSTEM_VOLUME;

	hr = CoCreateInstance(__uuidof(MMDeviceEnumerator), NULL, CLSCTX_ALL, __uuidof(IMMDeviceEnumerator), (void**)&pEnumerator);
	if (SUCCEEDED(hr)) {
		hr = pEnumerator->EnumAudioEndpoints(eRender, DEVICE_STATE_ACTIVE, &pDevices);
		if (SUCCEEDED(hr)) {
			UINT deviceCount;
			pDevices->GetCount(&deviceCount);

			if (deviceCount > 0) {
				int iDevIndex = GetIMMWaveOutDevice(pDevices, mWAVE_PARM.WaveOutDev);
				if (iDevIndex == -1)
				{
					dwResult = ERROR_CODE_NOT_FIND_WAVEOUT_DEV;
				}
				else
				{
					pDevices->Item(iDevIndex, &pDevice);

					hr = pDevice->Activate(__uuidof(IAudioEndpointVolume), CLSCTX_ALL, NULL, (void**)&pAudioEndpointVolume);
					if (SUCCEEDED(hr)) {
						// 目標音量
						float targetTemp = 0;
						// 當前音量
						float currentVolume = 0;
						pAudioEndpointVolume->GetMasterVolumeLevelScalar(&currentVolume);

						if (mWAVE_PARM.bMuteTest)
							targetTemp = 0;
						else
						{
							targetTemp = (float)mWAVE_PARM.WaveOutVolume / 100;
							// 先檢查設備是否支持音量步進
							UINT stepCount;
							UINT currentStep;
							HRESULT hr = pAudioEndpointVolume->GetVolumeStepInfo(&currentStep, &stepCount);

							if (SUCCEEDED(hr) && stepCount > 1) {
								// 如果支持步進音量控制，則使用 VolumeStepUp 或 VolumeStepDown
								if (targetTemp > currentVolume) {
									while (currentVolume < targetTemp) {
										pAudioEndpointVolume->VolumeStepUp(NULL);
										pAudioEndpointVolume->GetMasterVolumeLevelScalar(&currentVolume);
									}
								}
								else if (targetTemp < currentVolume) {
									while (currentVolume > targetTemp) {
										pAudioEndpointVolume->VolumeStepDown(NULL);
										pAudioEndpointVolume->GetMasterVolumeLevelScalar(&currentVolume);
									}
								}
							}
							else {
								// 如果不支持步進，則使用 SetMasterVolumeLevelScalar
								pAudioEndpointVolume->SetMasterVolumeLevelScalar(targetTemp, NULL);
							}
						}
						dwResult = 0;

						pAudioEndpointVolume->Release();
					}
					pDevice->Release();
				}
			}
			pDevices->Release();
		}
		pEnumerator->Release();
	}

	CoUninitialize();
	return dwResult;
}

void ASAudio::SetMicMute(bool mute)
{
	MMRESULT rc;
	HMIXER hMixer;
	MIXERLINE mxl;
	MIXERLINECONTROLS mxlc;
	MIXERCONTROL mxc;

	// 遍歷系統的混音器
	for (int deviceID = 0; true; deviceID++)
	{
		// 打開當前的混音器，deviceID為混音器的ID
		rc = mixerOpen(&hMixer, deviceID, 0, 0, MIXER_OBJECTF_MIXER);
		if (rc != MMSYSERR_NOERROR)
			break;

		ZeroMemory(&mxl, sizeof(MIXERLINE));
		mxl.cbStruct = sizeof(MIXERLINE);
		mxl.dwComponentType = MIXERLINE_COMPONENTTYPE_DST_WAVEIN;
		rc = mixerGetLineInfo((HMIXEROBJ)hMixer, &mxl, MIXER_GETLINEINFOF_COMPONENTTYPE);
		if (rc != MMSYSERR_NOERROR)
		{
			mixerClose(hMixer);
			continue;
		}

		DWORD dwConnections = mxl.cConnections;
		DWORD dwLineID = -1;
		for (DWORD i = 0; i < dwConnections; i++)
		{
			mxl.dwSource = i;
			rc = mixerGetLineInfo((HMIXEROBJ)hMixer, &mxl, MIXER_OBJECTF_HMIXER | MIXER_GETLINEINFOF_SOURCE);
			if (rc == MMSYSERR_NOERROR)
			{
				if (mxl.dwComponentType == MIXERLINE_COMPONENTTYPE_SRC_MICROPHONE)
				{
					dwLineID = mxl.dwLineID;
					break;
				}
			}
		}
		if (dwLineID == -1)
		{
			mixerClose(hMixer);
			continue;
		}

		ZeroMemory(&mxc, sizeof(MIXERCONTROL));
		ZeroMemory(&mxlc, sizeof(MIXERLINECONTROLS));
		mxlc.cbStruct = sizeof(mxlc);
		mxlc.dwLineID = dwLineID;
		mxlc.dwControlType = MIXERCONTROL_CONTROLTYPE_MUTE;
		mxlc.cControls = 1;
		mxlc.pamxctrl = &mxc;
		mxlc.cbmxctrl = sizeof(mxc);

		rc = mixerGetLineControls((HMIXEROBJ)hMixer, &mxlc, MIXER_GETLINECONTROLSF_ONEBYTYPE);
		if (rc == MMSYSERR_NOERROR)
		{
			MIXERCONTROLDETAILS mxcd;
			MIXERCONTROLDETAILS_BOOLEAN muteStruct = { 0 };

			ZeroMemory(&mxcd, sizeof(mxcd));
			mxcd.cbStruct = sizeof(mxcd);
			mxcd.dwControlID = mxc.dwControlID;
			mxcd.paDetails = &muteStruct;
			mxcd.cbDetails = sizeof(muteStruct);
			mxcd.cChannels = 1;

			// 設置靜音
			muteStruct.fValue = mute ? 1 : 0;

			rc = mixerSetControlDetails((HMIXEROBJ)hMixer, &mxcd, MIXER_OBJECTF_HMIXER | MIXER_SETCONTROLDETAILSF_VALUE);
			mixerClose(hMixer);
			if (rc != MMSYSERR_NOERROR)
			{
				// 處理設置失敗的情況
				break;
			}
			return; // 成功設置後，退出函數
		}
		mixerClose(hMixer);
	}
}

std::wstring ASAudio::charToWstring(const char* szIn)
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

DWORD ASAudio::playPortAudioSound() {
	PaStream* stream;
	PaError err;

	// 初始化 PortAudio
	err = Pa_Initialize();
	if (err != paNoError) {
		data.errorcode = ERROR_CODE_OPEN_INIT_WAVEOUT_DEV;
		return data.errorcode;
	}

	data = { 0.0, 0.0, 0.0, 0.0, nullptr };
	data.phaseL = 0.0;
	data.phaseR = 0.0; // 初始化右聲道相位
	data.frequencyL = mWAVE_PARM.frequencyL; // 設置左聲道頻率
	data.frequencyR = mWAVE_PARM.frequencyR; // 設置右聲道頻率

	// 配置輸出流參數
	PaStreamParameters outputParameters = { 0 };
	outputParameters.device = GetPa_WaveOutDevice(mWAVE_PARM.WaveOutDev); // 設定輸出裝置
	if (outputParameters.device == paNoDevice) {
		data.errorcode = ERROR_CODE_NOT_FIND_WAVEOUT_DEV;
		Pa_Terminate();
		return data.errorcode;
	}
	outputParameters.channelCount = 2; // 雙聲道輸出
	outputParameters.sampleFormat = paFloat32; // 32位浮點數格式
	outputParameters.suggestedLatency = Pa_GetDeviceInfo(outputParameters.device)->defaultLowOutputLatency;
	outputParameters.hostApiSpecificStreamInfo = NULL;

	// 打開音頻流
	err = Pa_OpenStream(&stream,
		NULL, // no input
		&outputParameters,
		SAMPLE_RATE,
		FRAMES_PER_BUFFER,
		paClipOff, // 不進行裁剪
		patestCallback, // 設定回調函數
		&data); // 傳遞用戶資料

	if (err != paNoError) {
		data.errorMsg = Pa_GetErrorText(err);
		data.errorcode = ERROR_CODE_OPEN_WAVEOUT_DEV;
		Pa_Terminate();
		return data.errorcode;
	}

	data.stream = stream; // 保存流指針到資料結構中

	// 開始音頻流
	err = Pa_StartStream(stream);
	if (err != paNoError) {
		data.errorcode = ERROR_CODE_START_WAVEOUT_DEV;
		Pa_CloseStream(stream);
		Pa_Terminate();
		return data.errorcode;
	}
	return err;
}

// 停止撥放的函式
void ASAudio::stopPlayback() {
	if (data.stream != nullptr) {
		Pa_StopStream(data.stream);
		Pa_CloseStream(data.stream);
		data.stream = nullptr;
	}
	Pa_Terminate();
}

int ASAudio::GetWaveOutDevice(std::wstring szOutDevName)
{
	// 獲取音頻輸入設備的數量
	UINT deviceCount = waveOutGetNumDevs();
	for (UINT i = 0; i < deviceCount; ++i) {
		WAVEOUTCAPSA waveOutCaps;
		MMRESULT result = waveOutGetDevCapsA(i, &waveOutCaps, sizeof(WAVEOUTCAPSA));
		if (result == MMSYSERR_NOERROR) {
			std::string deviceName(waveOutCaps.szPname); // 將設備名稱轉換為 std::string
			// 將szOutDevName轉換為std::string
			int bufferSize = WideCharToMultiByte(CP_UTF8, 0, szOutDevName.c_str(), -1, NULL, 0, NULL, NULL) - 1;//這邊需要-1，讓後面的find比對正確的字元數量
			std::string targetName(bufferSize, '\0');
			WideCharToMultiByte(CP_UTF8, 0, szOutDevName.c_str(), -1, &targetName[0], bufferSize, NULL, NULL);

			if (deviceName.find(targetName) != std::string::npos) {
				return i; // 找到包含指定子字串的設備，返回設備 ID
			}
		}
	}
	return -1; // 未找到設備，返回 -1
}

int ASAudio::GetWaveInDevice(std::wstring szInDevName)
{
	// 獲取音頻輸入設備的數量
	UINT deviceCount = waveInGetNumDevs();
	for (UINT i = 0; i < deviceCount; ++i) {
		WAVEINCAPSA waveInCaps;
		MMRESULT result = waveInGetDevCapsA(i, &waveInCaps, sizeof(WAVEINCAPSA));
		if (result == MMSYSERR_NOERROR) {
			std::string deviceName(waveInCaps.szPname); // 將設備名稱轉換為 std::string
			// 將szInDevName轉換為std::string
			int bufferSize = WideCharToMultiByte(CP_UTF8, 0, szInDevName.c_str(), -1, NULL, 0, NULL, NULL) - 1;//這邊需要-1，讓後面的find比對正確的字元數量
			std::string targetName(bufferSize, '\0');
			WideCharToMultiByte(CP_UTF8, 0, szInDevName.c_str(), -1, &targetName[0], bufferSize, NULL, NULL);

			if (deviceName.find(targetName) != std::string::npos) {
				return i; // 找到包含指定子字串的設備，返回設備 ID
			}
		}
	}
	return -1; // 未找到設備，返回 -1
}

int ASAudio::GetPa_WaveOutDevice(std::wstring szOutDevName)
{
	PaError err = Pa_Initialize();
	int deviceCount = Pa_GetDeviceCount();
	for (int i = 0; i < deviceCount; ++i) {
		const PaDeviceInfo* deviceInfo = Pa_GetDeviceInfo(i);
		if (deviceInfo) {
			if (deviceInfo->maxOutputChannels > 0) { // 輸出裝置
				std::string deviceName(deviceInfo->name); // 將設備名稱轉換為 std::string
				// 將szOutDevName轉換為std::string
				int bufferSize = WideCharToMultiByte(CP_UTF8, 0, szOutDevName.c_str(), -1, NULL, 0, NULL, NULL) - 1;//這邊需要-1，讓後面的find比對正確的字元數量
				std::string targetName(bufferSize, '\0');
				WideCharToMultiByte(CP_UTF8, 0, szOutDevName.c_str(), -1, &targetName[0], bufferSize, NULL, NULL);

				if (deviceName.find(targetName) != std::string::npos) {
					Pa_Terminate();
					return i; // 找到包含指定子字串的設備，返回設備 ID
				}
			}
		}
	}
	Pa_Terminate();
	return -1;
}

int ASAudio::GetPa_WaveInDevice(std::wstring szInDevName)
{
	int deviceCount = Pa_GetDeviceCount();
	for (int i = 0; i < deviceCount; ++i) {
		const PaDeviceInfo* deviceInfo = Pa_GetDeviceInfo(i);
		if (deviceInfo) {
			if (deviceInfo->maxInputChannels > 0) { // 僅考慮輸入設備
				std::string deviceName(deviceInfo->name); // 將設備名稱轉換為 std::string
				// 將szInDevName轉換為std::string
				int bufferSize = WideCharToMultiByte(CP_UTF8, 0, szInDevName.c_str(), -1, NULL, 0, NULL, NULL) - 1;//這邊需要-1，讓後面的find比對正確的字元數量
				std::string targetName(bufferSize, '\0');
				WideCharToMultiByte(CP_UTF8, 0, szInDevName.c_str(), -1, &targetName[0], bufferSize, NULL, NULL);

				if (deviceName.find(targetName) != std::string::npos) {
					return i; // 找到包含指定子字串的設備，返回設備 ID
				}
			}
		}
	}
	return -1; // 未找到設備，返回 -1
}

int ASAudio::patestCallback(const void* inputBuffer, void* outputBuffer,
	unsigned long framesPerBuffer,
	const PaStreamCallbackTimeInfo* timeInfo,
	PaStreamCallbackFlags statusFlags,
	void* userData)
{
	paTestData* data = (paTestData*)userData;
	float* out = (float*)outputBuffer;
	unsigned long i;
	(void)inputBuffer; // 防止未使用變量警告

	for (i = 0; i < framesPerBuffer; i++) {
		*out++ = sinf(data->phaseL);
		data->phaseL += 2 * M_PI * data->frequencyL / SAMPLE_RATE;
		*out++ = sinf(data->phaseR);
		data->phaseR += 2 * M_PI * data->frequencyR / SAMPLE_RATE;
	}
	return paContinue; // 繼續播放
}

bool ASAudio::StartAudioLoopback(const std::wstring captureKeyword, const std::wstring renderKeyword) {
	HRESULT hr;
	// 初始化 COM
	HRESULT hrInit = CoInitialize(nullptr);
	if (FAILED(hrInit)) {
		//std::wcout << L"COM 初始化失敗！" << std::endl;
		strMacroResult = "LOG:ERROR_AUDIO_LOOPBACK_COM_INITIALIZE#";
		return false;
	}

	// 建立 DirectShow Graph
	hr = CoCreateInstance(CLSID_FilterGraph, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&pGraph));
	if (FAILED(hr)) {
		strMacroResult = "LOG:ERROR_AUDIO_LOOPBACK_FILTER_GRAPH#";
		//std::wcout << L"無法建立 Filter Graph！" << std::endl;
		return false;
	}

	// 建立 Capture Graph
	hr = CoCreateInstance(CLSID_CaptureGraphBuilder2, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&pCaptureGraph));
	if (FAILED(hr)) {
		//std::wcout << L"無法建立 Capture Graph！" << std::endl;
		strMacroResult = "LOG:ERROR_AUDIO_LOOPBACK_CAPTURE_GRAPH#";
		pGraph->Release();
		CoUninitialize();
		return false;
	}

	pCaptureGraph->SetFiltergraph(pGraph);

	// 找到匹配的錄音裝置
	ICreateDevEnum* pDevEnum = nullptr;
	IEnumMoniker* pEnum = nullptr;
	IMoniker* pMoniker = nullptr;

	CoCreateInstance(CLSID_SystemDeviceEnum, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&pDevEnum));
	if (FAILED(pDevEnum->CreateClassEnumerator(CLSID_AudioInputDeviceCategory, &pEnum, 0)) || !pEnum) {
		//std::wcout << L"未找到錄音裝置！" << std::endl;
		strMacroResult = "LOG:ERROR_AUDIO_WAVEIN_DEVICE_NOT_FIND#";
		goto fail;
	}

	while (pEnum->Next(1, &pMoniker, nullptr) == S_OK) {
		IPropertyBag* pPropBag;
		pMoniker->BindToStorage(nullptr, nullptr, IID_PPV_ARGS(&pPropBag));

		VARIANT varName;
		VariantInit(&varName);
		pPropBag->Read(L"FriendlyName", &varName, nullptr);

		// **如果裝置名稱包含關鍵字，就使用這個裝置**
		if (wcsstr(varName.bstrVal, captureKeyword.c_str())) {
			hr = pMoniker->BindToObject(nullptr, nullptr, IID_IBaseFilter, (void**)&pAudioCapture);
			if (SUCCEEDED(hr)) {
				VariantClear(&varName);
				pPropBag->Release();
				pMoniker->Release();
				break; // 找到符合的裝置就跳出迴圈
			}
		}
		VariantClear(&varName);
		pPropBag->Release();
		pMoniker->Release();
	}
	pEnum->Release();

	if (!pAudioCapture) {
		//std::wcout << L"未找到符合的錄音裝置！" << std::endl;
		strMacroResult = "LOG:ERROR_AUDIO_WAVEIN_DEVICE_NOT_FIND#";
		goto fail;
	}

	// 找到匹配的播放裝置
	if (FAILED(pDevEnum->CreateClassEnumerator(CLSID_AudioRendererCategory, &pEnum, 0)) || !pEnum) {
		//std::wcout << L"未找到播放裝置！" << std::endl;
		strMacroResult = "LOG:ERROR_AUDIO_WAVEOUT_DEVICE_NOT_FIND_4#";
		goto fail;
	}
	while (pEnum->Next(1, &pMoniker, nullptr) == S_OK) {
		IPropertyBag* pPropBag;
		pMoniker->BindToStorage(nullptr, nullptr, IID_PPV_ARGS(&pPropBag));

		VARIANT varName;
		VariantInit(&varName);
		pPropBag->Read(L"FriendlyName", &varName, nullptr);

		// **如果裝置名稱包含關鍵字，就使用這個裝置**
		if (wcsstr(varName.bstrVal, renderKeyword.c_str())) {
			hr = pMoniker->BindToObject(nullptr, nullptr, IID_IBaseFilter, (void**)&pAudioRenderer);
			if (SUCCEEDED(hr)) {
				VariantClear(&varName);
				pPropBag->Release();
				pMoniker->Release();
				break; // 找到符合的裝置就跳出迴圈
			}
		}
		VariantClear(&varName);
		pPropBag->Release();
		pMoniker->Release();
	}
	pEnum->Release();
	pDevEnum->Release();

	if (!pAudioRenderer) {
		//std::wcout << L"未找到符合的播放裝置！" << std::endl;
		strMacroResult = "LOG:ERROR_AUDIO_LOOPBACK_NO_MATCH_RENDER_DEVICE#";
		goto fail;
	}

	// 加入裝置到 Filter Graph
	pGraph->AddFilter(pAudioCapture, L"Audio Capture");
	pGraph->AddFilter(pAudioRenderer, L"Audio Renderer");

	// 連接錄音裝置到播放裝置
	hr = pCaptureGraph->RenderStream(&PIN_CATEGORY_CAPTURE, &MEDIATYPE_Audio, pAudioCapture, nullptr, pAudioRenderer);
	if (FAILED(hr)) {
		//std::wcout << L"無法建立錄音到播放的串流！" << std::endl;
		strMacroResult = "LOG:ERROR_AUDIO_LOOPBACK_RENDER_STREAM#";
		goto fail;
	}

	// 控制播放
	pGraph->QueryInterface(IID_PPV_ARGS(&pMediaControl));
	pMediaControl->Run();

	//std::wcout << L"正在轉發音訊... 按 Enter 結束" << std::endl;
	//std::wcin.get();

	return true;

fail:
	if (pAudioCapture) pAudioCapture->Release();
	if (pAudioRenderer) pAudioRenderer->Release();
	if (pCaptureGraph) pCaptureGraph->Release();
	if (pGraph) pGraph->Release();
	CoUninitialize();
	return false;
}

void ASAudio::StopAudioLoopback()
{
	pMediaControl->Stop();
	if (pAudioCapture) pAudioCapture->Release();
	if (pAudioRenderer) pAudioRenderer->Release();
	if (pCaptureGraph) pCaptureGraph->Release();
	if (pGraph) pGraph->Release();
	CoUninitialize();
}

int GetIMMWaveOutDevice(IMMDeviceCollection* pDevices, std::wstring szOutDevName)
{
	UINT deviceCount = 0;
	HRESULT hr = pDevices->GetCount(&deviceCount);
	if (FAILED(hr)) {
		return -1;
	}

	for (UINT i = 0; i < deviceCount; ++i) {
		IMMDevice* pDevice = nullptr;
		LPWSTR deviceId = nullptr;
		IPropertyStore* pProps = nullptr;

		// 取得裝置
		if (SUCCEEDED(pDevices->Item(i, &pDevice))) {
			// 取得裝置屬性
			if (SUCCEEDED(pDevice->OpenPropertyStore(STGM_READ, &pProps))) {
				PROPVARIANT varName;
				PropVariantInit(&varName);

				// 取得裝置名稱
				if (SUCCEEDED(pProps->GetValue(PKEY_Device_FriendlyName, &varName))) {
					std::wstring deviceName(varName.pwszVal);

					// 檢查名稱是否匹配
					if (deviceName.find(szOutDevName) != std::wstring::npos) {
						PropVariantClear(&varName);
						pProps->Release();
						pDevice->Release();
						return i; // 找到匹配裝置
					}
				}
				PropVariantClear(&varName);
			}
			if (pProps) pProps->Release();
		}
		if (pDevice) pDevice->Release();
	}
	return -1; // 找不到裝置
}

// 將 std::string 轉成 std::wstring（UTF-8 → UTF-16）
std::wstring Utf8ToWstring(const std::string& str) {
	if (str.empty()) return std::wstring();
	int size_needed = MultiByteToWideChar(CP_UTF8, 0, &str[0], (int)str.size(), NULL, 0);
	std::wstring wstrTo(size_needed, 0);
	MultiByteToWideChar(CP_UTF8, 0, &str[0], (int)str.size(), &wstrTo[0], size_needed);
	return wstrTo;
}

bool ASAudio::SetDefaultAudioPlaybackDevice(const std::wstring& deviceId)
{
	// 組合完整命令，例如：
	// SoundVolumeView.exe /SetDefault "{deviceId}" 0
	// 0: All, 1: Console, 2: Multimedia
	std::wstring command = L"SoundVolumeView.exe /SetDefault \"" + deviceId + L"\" 0";

	// 使用 _wsystem 以支援 Unicode 字串
	int result = _wsystem(command.c_str());

	// 回傳成功與否（0 表示成功）
	return result == 0;
}

bool ASAudio::FindDeviceIdByName(Config& config, std::wstring& outDeviceId)
{
	if (config.monitorNames.empty()) {
		strMacroResult = "LOG:ERROR_AUDIO_NO_TARGET_NAMES_PROVIDED#";
		return false;
	}

	CComPtr<IMMDeviceEnumerator> enumerator;
	CComPtr<IMMDeviceCollection> collection;
	CComPtr<IMMDevice> device;

	CoInitializeEx(NULL, COINIT_MULTITHREADED);
	HRESULT hr = CoCreateInstance(__uuidof(MMDeviceEnumerator), NULL, CLSCTX_ALL,
		__uuidof(IMMDeviceEnumerator), (void**)&enumerator);

	if (FAILED(hr)) {
		wprintf(L"初始化音訊裝置枚舉器失敗！錯誤碼: 0x%08X\n", hr);
		strMacroResult = "LOG:ERROR_AUDIO_ENUM_INIT_FAIL#";
		return false;
	}

	hr = enumerator->EnumAudioEndpoints(eRender, DEVICE_STATE_ACTIVE, &collection);
	if (FAILED(hr)) {
		wprintf(L"獲取音訊裝置失敗！錯誤碼: 0x%08X\n", hr);
		strMacroResult = "LOG:ERROR_AUDIO_ENUM_GET_FAIL#";
		return false;
	}

	UINT count;
	collection->GetCount(&count);
	wprintf(L"找到 %u 個音訊裝置:\n\n", count);

	if(count > 0)
		allContent += L"[系統中的裝置 System Audio Device]\n";
	int indexCountdown = config.AudioIndex;
	for (UINT i = 0; i < count; i++)
	{
		IMMDevice* rawDevice = nullptr;
		hr = collection->Item(i, &rawDevice);
		if (SUCCEEDED(hr) && rawDevice != nullptr)
		{
			device.Attach(rawDevice); // 手動附加
		}
		else
		{
			wprintf(L"Skip device %u: hr = 0x%08X\n", i, hr);
			continue;
		}

		LPWSTR deviceId = nullptr;
		device->GetId(&deviceId);

		CComPtr<IPropertyStore> store;
		hr = device->OpenPropertyStore(STGM_READ, &store);
		if (FAILED(hr) || !store) {
			CoTaskMemFree(deviceId);
			continue;
		}
		PROPVARIANT friendlyName;
		PropVariantInit(&friendlyName);
		hr = store->GetValue(PKEY_Device_FriendlyName, &friendlyName);
		if (SUCCEEDED(hr)) {
			wprintf(L"裝置 %d: %s\n", i, friendlyName.pwszVal);
			allContent += friendlyName.pwszVal;
			allContent += L"\n";
			wprintf(L" ID: %s\n\n", deviceId);

			// 檢查當前裝置名稱是否包含在我們的目標列表中
			bool isMatch = false;
			for (const auto& targetName : config.monitorNames) {
				if (wcsstr(friendlyName.pwszVal, targetName.c_str())) {
					isMatch = true;
					break; // 找到一個匹配就夠了，跳出內層迴圈
				}
			}

			if (isMatch) {
				// 這是一個符合條件的裝置
				if (indexCountdown <= 0)
				{
					// 這就是我們要找的第 N 個裝置
					outDeviceId = deviceId;
					CoTaskMemFree(deviceId);
					PropVariantClear(&friendlyName);
					CoUninitialize();
					allContent = L""; // 清空日誌
					return true;
				}
				else
				{
					// 還沒輪到，計數器減一
					indexCountdown--;
				}
			}
		}
		CoTaskMemFree(deviceId);
		PropVariantClear(&friendlyName);
	}
	CoUninitialize();

	// 如果迴圈跑完還沒找到
	allContent += L"\n\n[要找的裝置 Target Audio Devices]\n";
	for (const auto& name : config.monitorNames) {
		allContent += name + L", ";
	}
	strMacroResult = "LOG:ERROR_AUDIO_ENUM_NOT_FIND#";
	return false;
}