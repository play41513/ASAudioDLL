#include "ASAudio.h"
#include "ConstantString.h"
#include <unordered_map>

const std::unordered_map<ErrorCode, const wchar_t*> error_map = {
	{ErrorCode::NOT_FIND_WAVEIN_DEV, L"ERROR_CODE_NOT_FIND_WAVEIN_DEV"},
	{ErrorCode::OPEN_WAVEIN_DEV, L"ERROR_CODE_OPEN_WAVEIN_DEV"},
	{ErrorCode::NOT_FIND_WAVEOUT_DEV, L"ERROR_CODE_NOT_FIND_WAVEOUT_DEV"},
	{ErrorCode::OPEN_INIT_WAVEOUT_DEV, L"ERROR_CODE_OPEN_INIT_WAVEOUT_DEV"},
	{ErrorCode::OPEN_WAVEOUT_DEV, L"ERROR_CODE_OPEN_WAVEOUT_DEV"},
	{ErrorCode::START_WAVEOUT_DEV, L"ERROR_CODE_START_WAVEOUT_DEV"},
	{ErrorCode::SET_SPEAKER_SYSTEM_VOLUME, L"ERROR_CODE_SET_SPEAKER_SYSTEM_VOLUME"},
	{ErrorCode::UNSET_PARAMETER, L"ERROR_CODE_UNSET_PARAMETER"}
};
const fftw_complex* ASAudio::GetLeftSpectrumPtr() const {
	return fftOutputBufferLeft;
}
const fftw_complex* ASAudio::GetRightSpectrumPtr() const {
	return fftOutputBufferRight;
}

static int GetIMMWaveOutDevice(IMMDeviceCollection* pDevices, std::wstring szOutDevName)
{
	UINT deviceCount = 0;
	HRESULT hr = pDevices->GetCount(&deviceCount);
	if (FAILED(hr)) {
		return -1;
	}

	for (UINT i = 0; i < deviceCount; ++i) {
		CComPtr<IMMDevice> pDevice;
		CComPtr<IPropertyStore> pProps;

		// 取得裝置
		if (SUCCEEDED(pDevices->Item(i, &pDevice))) {
			// 取得裝置屬性
			if (SUCCEEDED(pDevice->OpenPropertyStore(STGM_READ, &pProps))) {
				PROPVARIANT varName;
				PropVariantInit(&varName);

				// 取得裝置名稱
				if (SUCCEEDED(pProps->GetValue(PKEY_Device_FriendlyName, &varName))) {
					std::wstring deviceName(varName.pwszVal);

					// 檢查名稱是否相符
					if (deviceName.find(szOutDevName) != std::wstring::npos) {
						PropVariantClear(&varName);
						return i; // 找到相符裝置
					}
				}
				PropVariantClear(&varName);
			}
		}
	}
	return -1; // 找不到裝置
}

// 將 std::string 轉為 std::wstring（UTF-8 到 UTF-16）
static std::wstring Utf8ToWstring(const std::string& str) {
	if (str.empty()) return std::wstring();
	int size_needed = MultiByteToWideChar(CP_UTF8, 0, &str[0], (int)str.size(), NULL, 0);
	std::wstring wstrTo(size_needed, 0);
	MultiByteToWideChar(CP_UTF8, 0, &str[0], (int)str.size(), &wstrTo[0], size_needed);
	return wstrTo;
}
void ASAudio::GetSpectrumData(fftw_complex*& leftSpectrum, fftw_complex*& rightSpectrum)
{
	leftSpectrum = fftOutputBufferLeft;
	rightSpectrum = fftOutputBufferRight;
}
ASAudio& ASAudio::GetInstance()
{
	static ASAudio instance; 
	return instance;
}
const std::string& ASAudio::GetLastResult() const
{
	return strMacroResult;
}

ASAudio::ASAudio() :
	spectrumCallback(nullptr), 
	bufferIndex(0),
	hWaveIn(NULL),
	hWaveOut(NULL),
	pGraph(nullptr),
	pCaptureGraph(nullptr),
	pMediaControl(nullptr),
	pAudioCapture(nullptr),
	pAudioRenderer(nullptr),
	bFirstWaveInFlag(false),
	waveHdr{},
	WaveOutHeader{},
	data{}
{

	memset(&waveHdr, 0, sizeof(WAVEHDR)); // 初始化錄音的WAVEHDR結構
	memset(mWAVE_PARM.WAVE_DATA.audioBuffer, 0, sizeof(mWAVE_PARM.WAVE_DATA.audioBuffer)); // 初始化音訊緩衝區，歸零

	// 初始化 FFT 相關的記憶體和計畫
	fftInputBufferLeft = (fftw_complex*)fftw_malloc(sizeof(fftw_complex) * BUFFER_SIZE);
	fftInputBufferRight = (fftw_complex*)fftw_malloc(sizeof(fftw_complex) * BUFFER_SIZE);
	fftOutputBufferLeft = (fftw_complex*)fftw_malloc(sizeof(fftw_complex) * BUFFER_SIZE);
	fftOutputBufferRight = (fftw_complex*)fftw_malloc(sizeof(fftw_complex) * BUFFER_SIZE);

	// 創建 FFT 計畫，以便將來進行快速傅立葉變換
	fftPlanLeft = fftw_plan_dft_1d(BUFFER_SIZE, fftInputBufferLeft, fftOutputBufferLeft, FFTW_FORWARD, FFTW_ESTIMATE);
	fftPlanRight = fftw_plan_dft_1d(BUFFER_SIZE, fftInputBufferRight, fftOutputBufferRight, FFTW_FORWARD, FFTW_ESTIMATE);

	PaError err = Pa_Initialize();
	if (err != paNoError) {
		//
	}
}
ASAudio::~ASAudio() {
	fftw_destroy_plan(fftPlanLeft);
	fftw_destroy_plan(fftPlanRight);
	fftw_free(fftInputBufferLeft);
	fftw_free(fftInputBufferRight);
	fftw_free(fftOutputBufferLeft);
	fftw_free(fftOutputBufferRight);

	Pa_Terminate();
}

WAVE_PARM& ASAudio::GetWaveParm() {
	return mWAVE_PARM;
}

void ASAudio::ClearLog() {
	allContent.clear();
}

void ASAudio::AppendLog(const std::wstring& message) {
	allContent += message;
}

const std::wstring& ASAudio::GetLog() const {
	return allContent;
}
WAVE_PARM::RecordingResult ASAudio::WaitForRecording(int timeoutSeconds) {
	std::unique_lock<std::mutex> lock(mWAVE_PARM.recordingMutex);

	// 在每次等待之前，重置狀態旗標
	mWAVE_PARM.isRecordingFinished = false;
	mWAVE_PARM.deviceError = false;

	// 使用 wait_for 等待，它會在指定時間後自動返回
	// 第二個參數是一個 lambda 函式，用於防止 "偽喚醒" (spurious wakeups)
	if (!mWAVE_PARM.recordingFinishedCV.wait_for(lock, std::chrono::seconds(timeoutSeconds), [&] {
		return mWAVE_PARM.isRecordingFinished;
		}))
	{
		// 如果 wait_for 返回 false，代表是「逾時」而不是被 notify 喚醒
		mWAVE_PARM.lastRecordingResult = WAVE_PARM::RecordingResult::Timeout;
		return WAVE_PARM::RecordingResult::Timeout;
	}

	// 程式執行到這裡，代表是被 notify 喚醒。
	// 檢查是「正常完成」還是「裝置錯誤」造成的喚醒。
	if (mWAVE_PARM.deviceError) {
		mWAVE_PARM.lastRecordingResult = WAVE_PARM::RecordingResult::DeviceError;
		return WAVE_PARM::RecordingResult::DeviceError;
	}

	// 如果沒有逾時，也沒有裝置錯誤，那代表錄音成功完成。
	mWAVE_PARM.lastRecordingResult = WAVE_PARM::RecordingResult::Success;
	return WAVE_PARM::RecordingResult::Success;
}

void ASAudio::SpectrumAnalysis(double* leftSpectrumData, double* rightSpectrumData)
{
	const int N = BUFFER_SIZE;
	const int freqBins = N / 2;

	// Hann Window ---
	fftw_complex* windowedInputLeft = (fftw_complex*)fftw_malloc(sizeof(fftw_complex) * N);
	fftw_complex* windowedInputRight = (fftw_complex*)fftw_malloc(sizeof(fftw_complex) * N);

	double windowSum = 0.0;
	for (int i = 0; i < N; i++) {
		double window = 0.5 * (1 - cos(2 * M_PI * i / (N - 1)));
		windowSum += window;
		windowedInputLeft[i][0] = fftInputBufferLeft[i][0] * window;
		windowedInputLeft[i][1] = 0.0;
		windowedInputRight[i][0] = fftInputBufferRight[i][0] * window;
		windowedInputRight[i][1] = 0.0;
	}

	// --- 2. 執行 FFT ---
	fftw_execute_dft(fftPlanLeft, windowedInputLeft, fftOutputBufferLeft);
	fftw_execute_dft(fftPlanRight, windowedInputRight, fftOutputBufferRight);

	fftw_free(windowedInputLeft);
	fftw_free(windowedInputRight);

	// --- 3. 計算正確縮放的頻譜振幅 (Amplitude) ---
	// (頻譜縮放) & #2 (窗函數補償)
	// Coherent Gain of Hann window is sum of window samples divided by N
	double coherentGain = windowSum / N;
	double scalingFactor = 2.0 / (N * coherentGain);

	for (int i = 0; i < freqBins; ++i) {
		double rawMagnitudeL = sqrt(fftOutputBufferLeft[i][0] * fftOutputBufferLeft[i][0] + fftOutputBufferLeft[i][1] * fftOutputBufferLeft[i][1]);
		double rawMagnitudeR = sqrt(fftOutputBufferRight[i][0] * fftOutputBufferRight[i][0] + fftOutputBufferRight[i][1] * fftOutputBufferRight[i][1]);

		// 對於 DC (i=0) 和 Nyquist (如果 N 是偶數，在 i=N/2) 分量，不乘以 2
		if (i == 0) {
			leftSpectrumData[i] = rawMagnitudeL / (N * coherentGain);
			rightSpectrumData[i] = rawMagnitudeR / (N * coherentGain);
		}
		else {
			leftSpectrumData[i] = rawMagnitudeL * scalingFactor;
			rightSpectrumData[i] = rawMagnitudeR * scalingFactor;
		}
	}

	// --- 左聲道 ---
	auto& leftAnalysis = mWAVE_PARM.bMuteTest ? mWAVE_PARM.leftMuteWAVE_ANALYSIS : mWAVE_PARM.leftWAVE_ANALYSIS;
	{
		// 4a. 尋找基頻 (Fundamental) 的最高點
		double maxAmplitude = 0.0;
		int fundamentalBin = 0;
		for (int i = 2; i < freqBins; ++i) { 
			if (leftSpectrumData[i] > maxAmplitude) {
				maxAmplitude = leftSpectrumData[i];
				fundamentalBin = i;
			}
		}
		leftAnalysis.TotalEnergyPoint = fundamentalBin;

		// -60 dBFS 是一個合理的閾值，代表訊號強度低於最大值的 0.1%
		double fullScale = 32767.0;
		if (!mWAVE_PARM.bMuteTest && (maxAmplitude < fullScale * 0.001)) { // ~ -60 dBFS
			leftAnalysis.thd_N_dB = 0; // 使用 0 dB 表示錯誤/無效訊號
			leftAnalysis.TotalEnergyPoint = 0;
			leftAnalysis.TotalEnergy = 0; 
			leftAnalysis.fundamentalEnergy = 0;
		}
		else {
			// 4c. 使用正確的能量計算 THD+N
			double totalEnergy = 0;
			double fundamentalEnergy = 0;
			const double sampleRate = 44100.0;

			double bandwidthHz = GetWaveParm().fundamentalBandwidthHz; 
			if (bandwidthHz <= 0) bandwidthHz = 100.0;
			double freqPerBin = sampleRate / N;
			int peakWidth = static_cast<int>(round(bandwidthHz / freqPerBin));

			// 計算總能量
			for (int i = 2; i < freqBins; ++i) {
				totalEnergy += leftSpectrumData[i] * leftSpectrumData[i];
			}

			// 整合基頻能量
			for (int i = -peakWidth; i <= peakWidth; ++i) {
				int bin = fundamentalBin + i;
				if (bin > 1 && bin < freqBins) {
					fundamentalEnergy += leftSpectrumData[bin] * leftSpectrumData[bin];
				}
			}

			leftAnalysis.TotalEnergy = totalEnergy;
			leftAnalysis.fundamentalEnergy = fundamentalEnergy;

			// 計算 THD+N
			double noiseAndDistortionEnergy = totalEnergy - fundamentalEnergy;
			if (fundamentalEnergy > 0 && noiseAndDistortionEnergy > 0) {
				double thd_n_ratio = sqrt(noiseAndDistortionEnergy / fundamentalEnergy);
				leftAnalysis.thd_N_dB = 20 * log10(thd_n_ratio);
			}
			else {
				leftAnalysis.thd_N_dB = -INFINITY;
			}
		}
	}

	// --- 5. 針對右聲道進行分析---
	auto& rightAnalysis = mWAVE_PARM.bMuteTest ? mWAVE_PARM.rightMuteWAVE_ANALYSIS : mWAVE_PARM.rightWAVE_ANALYSIS;
	{
		double maxAmplitude = 0.0;
		int fundamentalBin = 0;
		for (int i = 2; i < freqBins; ++i) {
			if (rightSpectrumData[i] > maxAmplitude) {
				maxAmplitude = rightSpectrumData[i];
				fundamentalBin = i;
			}
		}
		rightAnalysis.TotalEnergyPoint = fundamentalBin;

		double fullScale = 32767.0;
		if (!mWAVE_PARM.bMuteTest && (maxAmplitude < fullScale * 0.001)) { // ~ -60 dBFS
			rightAnalysis.thd_N_dB = 0;
			rightAnalysis.TotalEnergyPoint = 0;
			rightAnalysis.TotalEnergy = 0;
			rightAnalysis.fundamentalEnergy = 0;
		}
		else {
			double totalEnergy = 0;
			double fundamentalEnergy = 0;

			const double sampleRate = 44100.0;
			double bandwidthHz = GetWaveParm().fundamentalBandwidthHz; // 假設 config 被存在 WAVE_PARM
			if (bandwidthHz <= 0) bandwidthHz = 100.0;
			double freqPerBin = sampleRate / N;
			int peakWidth = static_cast<int>(round(bandwidthHz / freqPerBin));

			for (int i = 2; i < freqBins; ++i) {
				totalEnergy += rightSpectrumData[i] * rightSpectrumData[i];
			}

			for (int i = -peakWidth; i <= peakWidth; ++i) {
				int bin = fundamentalBin + i;
				if (bin > 1 && bin < freqBins) {
					fundamentalEnergy += rightSpectrumData[bin] * rightSpectrumData[bin];
				}
			}

			rightAnalysis.TotalEnergy = totalEnergy;
			rightAnalysis.fundamentalEnergy = fundamentalEnergy;

			double noiseAndDistortionEnergy = totalEnergy - fundamentalEnergy;
			if (fundamentalEnergy > 0 && noiseAndDistortionEnergy > 0) {
				double thd_n_ratio = sqrt(noiseAndDistortionEnergy / fundamentalEnergy);
				rightAnalysis.thd_N_dB = 20 * log10(thd_n_ratio);
			}
			else {
				rightAnalysis.thd_N_dB = -INFINITY;
			}
		}
	}
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

	// 開啟WAVE檔，回傳一個HMMIO控制代碼
	HMMIO handle = mmioOpen((LPWSTR)mWAVE_PARM.AudioFile.c_str(), 0, MMIO_READ);

	// 進入RIFF區塊(RIFF Chunk)
	mmioDescend(handle, &ChunkInfo, 0, MMIO_FINDRIFF);

	// 進入fmt區塊(RIFF的子區塊，內含音訊格式資訊)
	FormatChunkInfo.ckid = mmioStringToFOURCCA("fmt", 0); // 尋找fmt子區塊
	mmioDescend(handle, &FormatChunkInfo, &ChunkInfo, MMIO_FINDCHUNK);

	// 讀取wav格式資訊
	memset(&wfx, 0, sizeof(WAVEFORMATEX));
	mmioRead(handle, (char*)&wfx, FormatChunkInfo.cksize);

	// 檢查聲道數量
	if (wfx.nChannels != 2) {
		// 如果不是雙聲道
		mmioClose(handle, 0);
		strMacroResult = "LOG:ERROR_AUDIO_WAVE_FILE_FORMAT_NOT_SUPPORT#";
		return false;
	}

	// 離開fmt區塊
	mmioAscend(handle, &FormatChunkInfo, 0);

	// 進入data區塊(內含實際的音訊樣本)
	DataChunkInfo.ckid = mmioStringToFOURCCA("data", 0);
	mmioDescend(handle, &DataChunkInfo, &ChunkInfo, MMIO_FINDCHUNK);

	// 取得data區塊的大小
	DataSize = DataChunkInfo.cksize;

	// 準備緩衝區
	WaveOutData.resize(DataSize);
	mmioRead(handle, WaveOutData.data(), DataSize);

	// 開啟音效輸出裝置
	int iDevIndex = GetWaveOutDevice(mWAVE_PARM.WaveOutDev);
	MMRESULT result = waveOutOpen(&hWaveOut, iDevIndex, &wfx, 0, 0, WAVE_FORMAT_QUERY);
	if (result != MMSYSERR_NOERROR) {
		// 失敗
		strMacroResult = "LOG:ERROR_AUDIO_WAVEOUT_DEVICE_NOT_FIND_2#";
		return false;
	}

	result = waveOutOpen(&hWaveOut, iDevIndex, &wfx, NULL, NULL, CALLBACK_NULL);
	if (result != MMSYSERR_NOERROR) {
		// 失敗
		waveOutClose(hWaveOut);
		strMacroResult = "LOG:ERROR_AUDIO_WAVEOUT_DEVICE_NOT_FIND_3#";
		return false;
	}

	// 設定wave header.
	memset(&WaveOutHeader, 0, sizeof(WaveOutHeader));
	WaveOutHeader.lpData = WaveOutData.data();
	WaveOutHeader.dwBufferLength = DataSize;

	// 準備wave header.
	waveOutPrepareHeader(hWaveOut, &WaveOutHeader, sizeof(WAVEHDR));

	// 將緩衝區資料寫入音效裝置(開始播放).
	waveOutWrite(hWaveOut, &WaveOutHeader, sizeof(WAVEHDR));
	if (AutoClose)
	{
		// 等待播放完畢
		while (!(WaveOutHeader.dwFlags & WHDR_DONE)) {
			Sleep(100); // 避免 CPU 過度使用
		}

		// 停止播放並釋放資源
		ASAudio::StopPlayingWavFile();
	}
	return true;
}

// 停止播放WAV檔案
void ASAudio::StopPlayingWavFile() {
	// 停止播放並釋放資源
	waveOutReset(hWaveOut);
	// 關閉音效裝置
	waveOutClose(hWaveOut);
	// 清除與WaveOutPrepareHeader準備的Wave。
	waveOutUnprepareHeader(hWaveOut, &WaveOutHeader, sizeof(WAVEHDR));
	WaveOutData.clear();
}

// 開始錄音
MMRESULT ASAudio::StartRecordingAndDrawSpectrum() {
	// 設定音訊格式
	WAVEFORMATEX format{};
	format.wFormatTag = WAVE_FORMAT_PCM;
	format.nChannels = 2;
	format.nSamplesPerSec = 44100;
	format.wBitsPerSample = 16;
	format.nBlockAlign = format.nChannels * format.wBitsPerSample / 8;
	format.nAvgBytesPerSec = format.nSamplesPerSec * format.nBlockAlign;
	format.cbSize = 0;

	MMRESULT result;
	// 清除音訊緩衝區，避免舊資料
	memset(mWAVE_PARM.WAVE_DATA.audioBuffer, 0, BUFFER_SIZE * sizeof(short) * 2);
	// 開啟音訊輸入裝置
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
	return result;
}

// 錄音回呼函式
void CALLBACK ASAudio::WaveInProc(HWAVEIN hwi, UINT uMsg, DWORD_PTR dwInstance, DWORD_PTR dwParam1, DWORD_PTR dwParam2) {

	ASAudio* asAudio = reinterpret_cast<ASAudio*>(dwInstance);
	if (!asAudio) return;
	switch (uMsg) {
		case MM_WIM_DATA: 
		{
			WAVEHDR* waveHdr = reinterpret_cast<WAVEHDR*>(dwParam1);
			if (waveHdr->dwBytesRecorded > 0) {
				// 使用 asAudio 來存取成員變數 bFirstWaveInFlag
				if (asAudio->bFirstWaveInFlag)
				{
					// 使用 asAudio 來存取並修改成員變數 bFirstWaveInFlag
					asAudio->bFirstWaveInFlag = false;

					// 重新提交緩衝區，以丟棄第一個緩衝區
					waveInUnprepareHeader(hwi, waveHdr, sizeof(WAVEHDR));
					waveInPrepareHeader(hwi, waveHdr, sizeof(WAVEHDR));
					waveInAddBuffer(hwi, waveHdr, sizeof(WAVEHDR));
				}
				else
				{
					// 呼叫成員函式來處理
					asAudio->ProcessAudioBuffer();
				}
			}
			break;
		}
		case MM_WIM_CLOSE:
		{
			// 裝置已被關閉或拔除
			// 設定錯誤旗標並立即通知等待中的主執行緒
			std::lock_guard<std::mutex> lock(asAudio->mWAVE_PARM.recordingMutex);
			asAudio->mWAVE_PARM.deviceError = true;
			asAudio->mWAVE_PARM.isRecordingFinished = true; // 將其標記為 "完成" 以便喚醒 wait
			asAudio->mWAVE_PARM.recordingFinishedCV.notify_one(); // 立即喚醒
			break;
		}
	}
}

// 處理音訊緩衝區
void ASAudio::ProcessAudioBuffer() {
	short* recordedSamples = reinterpret_cast<short*>(waveHdr.lpData);

	for (int i = 0; i < BUFFER_SIZE; i++) {
		// 左聲道樣本
		short leftSample = recordedSamples[i * 2];
		mWAVE_PARM.WAVE_DATA.LeftAudioBuffer[i] = leftSample;
		fftInputBufferLeft[i][0] = static_cast<double>(leftSample);
		fftInputBufferLeft[i][1] = 0.0;

		// 右聲道樣本
		short rightSample = recordedSamples[i * 2 + 1];
		mWAVE_PARM.WAVE_DATA.RightAudioBuffer[i] = rightSample;
		fftInputBufferRight[i][0] = static_cast<double>(rightSample);
		fftInputBufferRight[i][1] = 0.0;
	}

	// 執行FFT轉換
	fftw_execute(fftPlanLeft);
	fftw_execute(fftPlanRight);

	// 通知錄音完成
	std::lock_guard<std::mutex> lock(mWAVE_PARM.recordingMutex);
	mWAVE_PARM.isRecordingFinished = true;
	mWAVE_PARM.recordingFinishedCV.notify_one();
}

DWORD ASAudio::SetMicSystemVolume()
{
	//宣告混音器相關結構
	MMRESULT rc;
	HMIXER hMixer;
	MIXERLINE mxl;
	MIXERLINECONTROLS mxlc;
	MIXERCONTROL mxc;

	//遍歷所有混音器裝置
	for (int deviceID = 0; true; deviceID++)
	{
		// 1. 開啟指定混音器
		rc = mixerOpen(&hMixer, deviceID, 0, 0, MIXER_OBJECTF_MIXER);
		if (rc != MMSYSERR_NOERROR) {
			// 如果連裝置都打不開了，代表已經找完所有裝置，結束迴圈
			break;
		}

		// 2. 指定要找的是錄音線路
		ZeroMemory(&mxl, sizeof(MIXERLINE));
		mxl.cbStruct = sizeof(MIXERLINE);
		mxl.dwComponentType = MIXERLINE_COMPONENTTYPE_DST_WAVEIN;
		rc = mixerGetLineInfo((HMIXEROBJ)hMixer, &mxl, MIXER_GETLINEINFOF_COMPONENTTYPE);
		if (rc != MMSYSERR_NOERROR) {
			// 如果這個裝置沒有錄音線路，就關閉它，然後檢查下一個
			mixerClose(hMixer);
			continue;
		}

		// 3. 在這個錄音線路中，尋找麥克風來源
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
					break; // 找到了麥克風，跳出這個小迴圈
				}
			}
		}

		// 如果在這個裝置上沒找到麥克風線路，就關閉它，然後檢查下一個
		if (dwLineID == -1) {
			mixerClose(hMixer);
			continue;
		}

		// 4. 尋找音量控制器
		ZeroMemory(&mxc, sizeof(MIXERCONTROL));
		mxc.cbStruct = sizeof(mxc);
		ZeroMemory(&mxlc, sizeof(MIXERLINECONTROLS));
		mxlc.cbStruct = sizeof(mxlc);
		mxlc.dwLineID = dwLineID;
		mxlc.dwControlType = MIXERCONTROL_CONTROLTYPE_VOLUME;
		mxlc.cControls = 1;
		mxlc.pamxctrl = &mxc;
		mxlc.cbmxctrl = sizeof(mxc);
		rc = mixerGetLineControls((HMIXEROBJ)hMixer, &mxlc, MIXER_GETLINECONTROLSF_ONEBYTYPE);

		// 5. 如果找到了音量控制器，就設定音量
		if (MMSYSERR_NOERROR == rc)
		{
			MIXERCONTROLDETAILS mxcd;
			MIXERCONTROLDETAILS_SIGNED volStruct{};

			ZeroMemory(&mxcd, sizeof(mxcd));
			mxcd.cbStruct = sizeof(mxcd);
			mxcd.dwControlID = mxc.dwControlID;
			mxcd.paDetails = &volStruct;
			mxcd.cbDetails = sizeof(volStruct);
			mxcd.cChannels = 1;

			rc = mixerGetControlDetails((HMIXEROBJ)hMixer, &mxcd, MIXER_GETCONTROLDETAILSF_VALUE);

			MIXERCONTROLDETAILS_UNSIGNED mxcdVolume_Set = { mxc.Bounds.dwMaximum * mWAVE_PARM.WaveInVolume / 100 };
			MIXERCONTROLDETAILS mxcd_Set = { 0 };
			mxcd_Set.cbStruct = sizeof(MIXERCONTROLDETAILS);
			mxcd_Set.dwControlID = mxc.dwControlID;
			mxcd_Set.cChannels = 1;
			mxcd_Set.cMultipleItems = 0;
			mxcd_Set.cbDetails = sizeof(MIXERCONTROLDETAILS_UNSIGNED);
			mxcd_Set.paDetails = &mxcdVolume_Set;

			mixerSetControlDetails((HMIXEROBJ)(hMixer), &mxcd_Set, MIXER_OBJECTF_HMIXER | MIXER_SETCONTROLDETAILSF_VALUE);

			// 6. 完成，關閉裝置並直接跳出最外層的大迴圈
			mixerClose(hMixer);
			break;
		}
		mixerClose(hMixer);
	}
	return 0;
}

DWORD ASAudio::SetSpeakerSystemVolume() const
{
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
						// 目前音量
						float currentVolume = 0;
						pAudioEndpointVolume->GetMasterVolumeLevelScalar(&currentVolume);

						if (mWAVE_PARM.bMuteTest)
							targetTemp = 0;
						else
						{
							targetTemp = (float)mWAVE_PARM.WaveOutVolume / 100;
							// 先檢查裝置是否支援步進音量
							UINT stepCount;
							UINT currentStep;
							HRESULT hr = pAudioEndpointVolume->GetVolumeStepInfo(&currentStep, &stepCount);

							if (SUCCEEDED(hr) && stepCount > 1) {
								// 如果支援步進音量，則使用 VolumeStepUp 和 VolumeStepDown
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
								// 如果不支援步進，則使用 SetMasterVolumeLevelScalar
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
	return dwResult;
}

void ASAudio::SetMicMute(bool mute)
{
	MMRESULT rc;
	HMIXER hMixer;
	MIXERLINE mxl;
	MIXERLINECONTROLS mxlc;
	MIXERCONTROL mxc;

	// 遍歷所有混音器
	for (int deviceID = 0; true; deviceID++)
	{
		// 開啟指定混音器，deviceID為混音器ID
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

			// 設定靜音
			muteStruct.fValue = mute ? 1 : 0;

			rc = mixerSetControlDetails((HMIXEROBJ)hMixer, &mxcd, MIXER_OBJECTF_HMIXER | MIXER_SETCONTROLDETAILSF_VALUE);
			mixerClose(hMixer);
			if (rc != MMSYSERR_NOERROR)
			{
				// 處理設定失敗的狀況
				break;
			}
			return; // 成功設定，就離開
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

	data = { 0.0, 0.0, 0.0, 0.0, nullptr };
	data.phaseL = 0.0;
	data.phaseR = 0.0; // 初始化右聲道相位
	data.frequencyL = mWAVE_PARM.frequencyL; // 設定左聲道頻率
	data.frequencyR = mWAVE_PARM.frequencyR; // 設定右聲道頻率

	// 準備音訊串流參數
	PaStreamParameters outputParameters = { 0 };
	outputParameters.device = GetPa_WaveOutDevice(mWAVE_PARM.WaveOutDev); // 設定輸出裝置
	if (outputParameters.device == paNoDevice) {
		data.errorcode = ERROR_CODE_NOT_FIND_WAVEOUT_DEV;
		return data.errorcode;
	}
	outputParameters.channelCount = 2; // 雙聲道輸出
	outputParameters.sampleFormat = paFloat32; // 32位元浮點數格式
	outputParameters.suggestedLatency = Pa_GetDeviceInfo(outputParameters.device)->defaultLowOutputLatency;
	outputParameters.hostApiSpecificStreamInfo = NULL;

	// 開啟音訊串流
	err = Pa_OpenStream(&stream,
		NULL, // no input
		&outputParameters,
		SAMPLE_RATE,
		FRAMES_PER_BUFFER,
		paClipOff, // 不裁剪樣本
		patestCallback, // 設定回呼函式
		&data); // 傳遞使用者資料

	if (err != paNoError) {
		data.errorMsg = Pa_GetErrorText(err);
		data.errorcode = ERROR_CODE_OPEN_WAVEOUT_DEV;
		return data.errorcode;
	}

	data.stream = stream; // 保存串流指標到結構中

	// 開始音訊串流
	err = Pa_StartStream(stream);
	if (err != paNoError) {
		data.errorcode = ERROR_CODE_START_WAVEOUT_DEV;
		Pa_CloseStream(stream);
		return data.errorcode;
	}
	return err;
}

// 停止播放的函式
void ASAudio::stopPlayback() {
	if (data.stream != nullptr) {
		Pa_StopStream(data.stream);
		Pa_CloseStream(data.stream);
		data.stream = nullptr;
	}
}

int ASAudio::GetWaveOutDevice(std::wstring szOutDevName)
{
	UINT deviceCount = waveOutGetNumDevs();
	for (UINT i = 0; i < deviceCount; ++i) {
		WAVEOUTCAPSW waveOutCaps; 
		MMRESULT result = waveOutGetDevCapsW(i, &waveOutCaps, sizeof(WAVEOUTCAPSW)); 
		if (result == MMSYSERR_NOERROR) {
			std::wstring deviceName(waveOutCaps.szPname);

			if (deviceName.find(szOutDevName) != std::wstring::npos) {
				ASAudio::GetInstance().mWAVE_PARM.ActualWaveOutDev = deviceName;
				return i;
			}
		}
	}
	ASAudio::GetInstance().GetWaveParm().ActualWaveOutDev.clear();
	return -1;
}

int ASAudio::GetWaveInDevice(std::wstring szInDevName)
{
	// 取得音效輸入裝置的數量
	UINT deviceCount = waveInGetNumDevs();
	for (UINT i = 0; i < deviceCount; ++i) {
		WAVEINCAPSW waveInCaps;
		MMRESULT result = waveInGetDevCapsW(i, &waveInCaps, sizeof(WAVEINCAPSW));
		if (result == MMSYSERR_NOERROR) {
			std::wstring deviceName(waveInCaps.szPname);

			if (deviceName.find(szInDevName) != std::string::npos) {
				ASAudio::GetInstance().mWAVE_PARM.ActualWaveInDev = deviceName;
				return i; 
			}
		}
	}
	ASAudio::GetInstance().GetWaveParm().ActualWaveInDev.clear();
	return -1; // 找不到裝置，回傳 -1
}

int ASAudio::GetPa_WaveOutDevice(std::wstring szOutDevName)
{
	int deviceCount = Pa_GetDeviceCount();
	for (int i = 0; i < deviceCount; ++i) {
		const PaDeviceInfo* deviceInfo = Pa_GetDeviceInfo(i);
		if (deviceInfo) {
			if (deviceInfo->maxOutputChannels > 0) { // 輸出裝置
				std::string deviceName(deviceInfo->name); // 將裝置名稱轉為 std::string
				// 將szOutDevName轉為std::string
				int bufferSize = WideCharToMultiByte(CP_UTF8, 0, szOutDevName.c_str(), -1, NULL, 0, NULL, NULL) - 1;//這裡需要-1，不然會導致find找不到正確的字串數量
				std::string targetName(bufferSize, ' ');
				WideCharToMultiByte(CP_UTF8, 0, szOutDevName.c_str(), -1, &targetName[0], bufferSize, NULL, NULL);

				if (deviceName.find(targetName) != std::string::npos) {
					return i; // 若含有指定關鍵字的裝置，回傳裝置 ID
				}
			}
		}
	}
	return -1;
}

int ASAudio::GetPa_WaveInDevice(std::wstring szInDevName)
{
	int deviceCount = Pa_GetDeviceCount();
	for (int i = 0; i < deviceCount; ++i) {
		const PaDeviceInfo* deviceInfo = Pa_GetDeviceInfo(i);
		if (deviceInfo) {
			if (deviceInfo->maxInputChannels > 0) { // 只找尋有輸入的裝置
				std::string deviceName(deviceInfo->name); // 將裝置名稱轉為 std::string
				// 將szInDevName轉為std::string
				int bufferSize = WideCharToMultiByte(CP_UTF8, 0, szInDevName.c_str(), -1, NULL, 0, NULL, NULL) - 1;//這裡需要-1，不然會導致find找不到正確的字串數量
				std::string targetName(bufferSize, ' ');
				WideCharToMultiByte(CP_UTF8, 0, szInDevName.c_str(), -1, &targetName[0], bufferSize, NULL, NULL);

				if (deviceName.find(targetName) != std::string::npos) {
					return i; // 若含有指定關鍵字的裝置，回傳裝置 ID
				}
			}
		}
	}
	return -1; // 找不到裝置，回傳 -1
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
	(void)inputBuffer; // 避免未使用的變數警告

	for (i = 0; i < framesPerBuffer; i++) {
		*out++ = sinf(data->phaseL);
		data->phaseL += 2 * M_PI * data->frequencyL / SAMPLE_RATE;
		*out++ = sinf(data->phaseR);
		data->phaseR += 2 * M_PI * data->frequencyR / SAMPLE_RATE;
	}
	return paContinue; // 繼續播放
}

/**
 * @brief 開始音訊迴路 (Loopback)。
 * @param captureKeyword 用於識別擷取裝置 (麥克風) 名稱的關鍵字。
 * @param renderKeyword 用於識別渲染裝置 (喇叭) 名稱的關鍵字。
 * @return bool 操作是否成功。
 * @note 此函式依賴於類別成員 pGraph, pCaptureGraph 等被宣告為 CComPtr 智慧指標，
 * 以實現自動資源管理，避免在錯誤發生時產生資源洩漏。
 */
bool ASAudio::StartAudioLoopback(const std::wstring captureKeyword, const std::wstring renderKeyword) {
	HRESULT hr;

	// 建立 Graph (圖形管理器) 和 Capture Graph Builder
	// CComPtr 會在物件建立失敗時保持為 NULL
	hr = pGraph.CoCreateInstance(CLSID_FilterGraph);
	if (FAILED(hr)) {
		strMacroResult = "LOG:ERROR_AUDIO_LOOPBACK_CREATE_GRAPH#";
		return false;
	}

	hr = pCaptureGraph.CoCreateInstance(CLSID_CaptureGraphBuilder2);
	if (FAILED(hr)) {
		strMacroResult = "LOG:ERROR_AUDIO_LOOPBACK_CREATE_BUILDER#";
		return false;
	}

	// 將 filter graph 關聯到 capture graph builder
	hr = pCaptureGraph->SetFiltergraph(pGraph);
	if (FAILED(hr)) {
		strMacroResult = "LOG:ERROR_AUDIO_LOOPBACK_SET_FILTERGRAPH#";
		return false;
	}

	// 建立裝置列舉器
	CComPtr<ICreateDevEnum> pDevEnum;
	hr = pDevEnum.CoCreateInstance(CLSID_SystemDeviceEnum);
	if (FAILED(hr)) {
		strMacroResult = "LOG:ERROR_AUDIO_LOOPBACK_CREATE_DEV_ENUM#";
		return false;
	}

	// --- 尋找指定的擷取裝置篩選器 ---
	CComPtr<IEnumMoniker> pEnumCapture;
	if (SUCCEEDED(pDevEnum->CreateClassEnumerator(CLSID_AudioInputDeviceCategory, &pEnumCapture, 0)) && pEnumCapture) {
		CComPtr<IMoniker> pMoniker;
		while (pEnumCapture->Next(1, &pMoniker, nullptr) == S_OK) {
			CComPtr<IPropertyBag> pPropBag;
			if (SUCCEEDED(pMoniker->BindToStorage(nullptr, nullptr, IID_PPV_ARGS(&pPropBag)))) {
				VARIANT varName;
				VariantInit(&varName);
				if (SUCCEEDED(pPropBag->Read(L"FriendlyName", &varName, nullptr))) {
					if (wcsstr(varName.bstrVal, captureKeyword.c_str())) {
						// 找到相符的裝置，建立 filter
						pMoniker->BindToObject(nullptr, nullptr, IID_IBaseFilter, (void**)&pAudioCapture);
					}
				}
				VariantClear(&varName);
			}
			if (pAudioCapture) break; // 找到後跳出迴圈
			pMoniker.Release(); // 釋放當前 moniker 繼續尋找下一個
		}
	}

	if (!pAudioCapture) {
		strMacroResult = "LOG:ERROR_AUDIO_WAVEIN_DEVICE_NOT_FIND#";
		return false; // CComPtr 會自動釋放已建立的 pGraph, pCaptureGraph 等物件
	}

	// --- 尋找指定的渲染裝置篩選器 ---
	CComPtr<IEnumMoniker> pEnumRender;
	if (SUCCEEDED(pDevEnum->CreateClassEnumerator(CLSID_AudioRendererCategory, &pEnumRender, 0)) && pEnumRender) {
		CComPtr<IMoniker> pMoniker;
		while (pEnumRender->Next(1, &pMoniker, nullptr) == S_OK) {
			CComPtr<IPropertyBag> pPropBag;
			if (SUCCEEDED(pMoniker->BindToStorage(nullptr, nullptr, IID_PPV_ARGS(&pPropBag)))) {
				VARIANT varName;
				VariantInit(&varName);
				if (SUCCEEDED(pPropBag->Read(L"FriendlyName", &varName, nullptr))) {
					if (wcsstr(varName.bstrVal, renderKeyword.c_str())) {
						pMoniker->BindToObject(nullptr, nullptr, IID_IBaseFilter, (void**)&pAudioRenderer);
					}
				}
				VariantClear(&varName);
			}
			if (pAudioRenderer) break; // 找到後跳出迴圈
			pMoniker.Release();
		}
	}

	if (!pAudioRenderer) {
		strMacroResult = "LOG:ERROR_AUDIO_LOOPBACK_NO_MATCH_RENDER_DEVICE#";
		return false; // CComPtr 會自動釋放已建立的物件
	}

	// 將篩選器加入圖形
	hr = pGraph->AddFilter(pAudioCapture, L"Audio Capture");
	if (FAILED(hr)) {
		strMacroResult = "LOG:ERROR_AUDIO_LOOPBACK_ADD_CAPTURE_FILTER#";
		return false;
	}

	hr = pGraph->AddFilter(pAudioRenderer, L"Audio Renderer");
	if (FAILED(hr)) {
		strMacroResult = "LOG:ERROR_AUDIO_LOOPBACK_ADD_RENDER_FILTER#";
		return false;
	}

	// 連接擷取與渲染裝置的資料流
	hr = pCaptureGraph->RenderStream(&PIN_CATEGORY_CAPTURE, &MEDIATYPE_Audio, pAudioCapture, nullptr, pAudioRenderer);
	if (FAILED(hr)) {
		strMacroResult = "LOG:ERROR_AUDIO_LOOPBACK_RENDER_STREAM#";
		return false;
	}

	// 取得媒體控制介面並執行
	hr = pGraph->QueryInterface(IID_PPV_ARGS(&pMediaControl));
	if (FAILED(hr) || !pMediaControl) {
		strMacroResult = "LOG:ERROR_AUDIO_LOOPBACK_GET_MEDIA_CONTROL#";
		return false;
	}

	pMediaControl->Run();
	return true;
}

void ASAudio::StopAudioLoopback()
{
	if (pMediaControl) {
		pMediaControl->Stop();
	}

	pMediaControl = nullptr;
	pAudioCapture = nullptr;
	pAudioRenderer = nullptr;
	pCaptureGraph = nullptr;
	pGraph = nullptr;
}


bool ASAudio::SetDefaultAudioPlaybackDevice(const std::wstring& deviceId)
{
	// 組成外部指令，例如
	// SoundVolumeView.exe /SetDefault "{deviceId}" 0
	// 0: All, 1: Console, 2: Multimedia
	std::wstring command = L"SoundVolumeView.exe /SetDefault \"" + deviceId + L"\" 0";

	// 使用 _wsystem 來執行 Unicode 字串
	int result = _wsystem(command.c_str());

	// 回傳是否成功(0 代表成功)
	return result == 0;
}
bool ASAudio::SetListenToThisDevice(const std::wstring& deviceId, int enable)
{
	// SoundVolumeView.exe /SetListenToThisDevice "{deviceId}" {1 or 0}
	std::wstring command = L"SoundVolumeView.exe /SetListenToThisDevice \"" + deviceId + L"\" " + std::to_wstring(enable);

	int result = _wsystem(command.c_str());
	return result == 0;
}

bool ASAudio::FindDeviceIdByName(Config& config, std::wstring& outDeviceId, EDataFlow dataFlow)
{
	if (config.monitorNames.empty()) {
		strMacroResult = "LOG:ERROR_AUDIO_NO_TARGET_NAMES_PROVIDED#";
		return false;
	}

	CComPtr<IMMDeviceEnumerator> enumerator;
	CComPtr<IMMDeviceCollection> collection;
	CComPtr<IMMDevice> device;

	HRESULT hr = CoCreateInstance(__uuidof(MMDeviceEnumerator), NULL, CLSCTX_ALL,
		__uuidof(IMMDeviceEnumerator), (void**)&enumerator);

	if (FAILED(hr)) {
		wprintf(L"初始化音訊裝置枚舉器失敗！錯誤碼: 0x%08X\n", hr);
		strMacroResult = "LOG:ERROR_AUDIO_ENUM_INIT_FAIL#";
		return false;
	}

	hr = enumerator->EnumAudioEndpoints(dataFlow, DEVICE_STATE_ACTIVE, &collection);
	if (FAILED(hr)) {
		wprintf(L"獲取音訊裝置失敗！錯誤碼: 0x%08X\n", hr);
		strMacroResult = "LOG:ERROR_AUDIO_ENUM_GET_FAIL#";
		return false;
	}

	UINT count;
	collection->GetCount(&count);
	wprintf(L"找到 %u 個音訊裝置:\n\n", count);

	if (count > 0)
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

	// 如果迴圈跑完還沒找到
	allContent += L"\n\n[要找的裝置 Target Audio Devices]\n";
	for (const auto& name : config.monitorNames) {
		allContent += name + L", ";
	}
	strMacroResult = "LOG:ERROR_AUDIO_ENUM_NOT_FIND#";
	return false;
}
void ASAudio::ExecuteFft() {
	fftw_execute(fftPlanLeft);
	fftw_execute(fftPlanRight);
}
MMRESULT ASAudio::StartRecordingOnly() {
	// 設定音訊格式
	WAVEFORMATEX format{};
	format.wFormatTag = WAVE_FORMAT_PCM;
	format.nChannels = 2;
	format.nSamplesPerSec = 44100;
	format.wBitsPerSample = 16;
	format.nBlockAlign = format.nChannels * format.wBitsPerSample / 8;
	format.nAvgBytesPerSec = format.nSamplesPerSec * format.nBlockAlign;
	format.cbSize = 0;

	MMRESULT result;
	// 清除音訊緩衝區
	memset(mWAVE_PARM.WAVE_DATA.audioBuffer, 0, BUFFER_SIZE * sizeof(short) * 2);

	// 開啟音訊輸入裝置
	int iDevIndex = GetWaveInDevice(mWAVE_PARM.WaveInDev);
	if (iDevIndex == -1)
	{
		data.errorcode = ERROR_CODE_NOT_FIND_WAVEIN_DEV;
		return ERROR_CODE_NOT_FIND_WAVEIN_DEV;
	}

	// 直接開啟並開始錄音，完全不呼叫播放函式
	mWAVE_PARM.firstBufferDiscarded = false;
	result = waveInOpen(&hWaveIn, iDevIndex, &format, (DWORD_PTR)WaveInProc, (DWORD_PTR)this, CALLBACK_FUNCTION | WAVE_FORMAT_DIRECT);
	if (result == MMSYSERR_NOERROR) {
		Sleep(100);
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
	return result;
}
