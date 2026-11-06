#include "ASAudio.h"
#include "ConstantString.h"
#include <unordered_map>
#include <cmath>

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

	//預先計算 Blackman-Harris Window
	windowSum = 0.0;
	for (int i = 0; i < BUFFER_SIZE; i++) {
		double ratio = (double)i / (BUFFER_SIZE - 1);
		windowFunction[i] = 0.35875
			- 0.48829 * cos(2.0 * M_PI * ratio)
			+ 0.14128 * cos(4.0 * M_PI * ratio)
			- 0.01168 * cos(6.0 * M_PI * ratio);
		windowSum += windowFunction[i];
	}

	// 初始化 FFT 相關的記憶體和計畫
	fftInputBufferLeft = (fftw_complex*)fftw_malloc(sizeof(fftw_complex) * BUFFER_SIZE);
	fftInputBufferRight = (fftw_complex*)fftw_malloc(sizeof(fftw_complex) * BUFFER_SIZE);
	fftOutputBufferLeft = (fftw_complex*)fftw_malloc(sizeof(fftw_complex) * BUFFER_SIZE);
	fftOutputBufferRight = (fftw_complex*)fftw_malloc(sizeof(fftw_complex) * BUFFER_SIZE);
	windowedInputLeft = (fftw_complex*)fftw_malloc(sizeof(fftw_complex) * BUFFER_SIZE);
	windowedInputRight = (fftw_complex*)fftw_malloc(sizeof(fftw_complex) * BUFFER_SIZE);

	// 創建 FFT 計畫，以便將來進行快速傅立葉變換
	fftPlanLeft = fftw_plan_dft_1d(BUFFER_SIZE, fftInputBufferLeft, fftOutputBufferLeft, FFTW_FORWARD, FFTW_ESTIMATE);
	fftPlanRight = fftw_plan_dft_1d(BUFFER_SIZE, fftInputBufferRight, fftOutputBufferRight, FFTW_FORWARD, FFTW_ESTIMATE);

	Pa_Initialize();
}
ASAudio::~ASAudio() {
	fftw_destroy_plan(fftPlanLeft);
	fftw_destroy_plan(fftPlanRight);
	fftw_free(fftInputBufferLeft);
	fftw_free(fftInputBufferRight);
	fftw_free(fftOutputBufferLeft);
	fftw_free(fftOutputBufferRight);

	fftw_free(windowedInputLeft);
	fftw_free(windowedInputRight);
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

	for (int i = 0; i < N; i++) {
		windowedInputLeft[i][0] = fftInputBufferLeft[i][0] * windowFunction[i];
		windowedInputLeft[i][1] = 0.0;
		windowedInputRight[i][0] = fftInputBufferRight[i][0] * windowFunction[i];
		windowedInputRight[i][1] = 0.0;
	}

	// --- 2. 執行 FFT ---
	fftw_execute_dft(fftPlanLeft, windowedInputLeft, fftOutputBufferLeft);
	fftw_execute_dft(fftPlanRight, windowedInputRight, fftOutputBufferRight);

	// --- 3. 計算振幅與還原 Scaling ---
	double coherentGain = windowSum / N;
	double scalingFactor = 2.0 / (N * coherentGain); // 動態適應 Blackman-Harris 的能量衰減

	for (int i = 0; i < freqBins; ++i) {
		double rawMagnitudeL = sqrt(fftOutputBufferLeft[i][0] * fftOutputBufferLeft[i][0] + fftOutputBufferLeft[i][1] * fftOutputBufferLeft[i][1]);
		double rawMagnitudeR = sqrt(fftOutputBufferRight[i][0] * fftOutputBufferRight[i][0] + fftOutputBufferRight[i][1] * fftOutputBufferRight[i][1]);

		if (i == 0) {
			leftSpectrumData[i] = rawMagnitudeL / (N * coherentGain);
			rightSpectrumData[i] = rawMagnitudeR / (N * coherentGain);
		}
		else {
			leftSpectrumData[i] = rawMagnitudeL * scalingFactor;
			rightSpectrumData[i] = rawMagnitudeR * scalingFactor;
		}
	}

	// --- 4. 計算左聲道 THD+N ---
	auto& leftAnalysis = mWAVE_PARM.bMuteTest ? mWAVE_PARM.leftMuteWAVE_ANALYSIS : mWAVE_PARM.leftWAVE_ANALYSIS;

	const double sampleRate = 44100.0;
	double freqPerBin = sampleRate / N;

	// 高通濾波：排除 20Hz 以下的環境低頻底噪與 DC 漂移 (USB 麥克風常見致命傷)
	int minBin = static_cast<int>(20.0 / freqPerBin);
	if (minBin < 2) minBin = 2;

	double maxAmplitude = 0.0;
	int fundamentalBin = minBin;
	for (int j = minBin; j < freqBins; ++j) {
		if (leftSpectrumData[j] > maxAmplitude) {
			maxAmplitude = leftSpectrumData[j];
			fundamentalBin = j;
		}
	}
	leftAnalysis.TotalEnergyPoint = fundamentalBin;

	double fullScale = 32767.0;
	if (!mWAVE_PARM.bMuteTest && (maxAmplitude < fullScale * 0.001)) {
		leftAnalysis.thd_N_dB = 0;
		leftAnalysis.TotalEnergyPoint = 0;
		leftAnalysis.TotalEnergy = 0;
		leftAnalysis.fundamentalEnergy = 0;
	}
	else {
		double totalEnergy = 0.0;
		double fundamentalEnergy = 0.0;

		// 建議將 Config 中的 bandwidth 設為 200Hz 以上
		double bandwidthHz = GetWaveParm().fundamentalBandwidthHz > 0 ? GetWaveParm().fundamentalBandwidthHz : 200.0;
		int peakWidth = static_cast<int>(round(bandwidthHz / freqPerBin));
		if (peakWidth < 6) peakWidth = 6; // Blackman-Harris 主瓣較寬，強制保留足夠的 Bins 收集主頻能量

		// 獨立迴圈 1：計算總頻譜能量 (從 20Hz 開始)
		for (int j = minBin; j < freqBins; ++j) {
			totalEnergy += leftSpectrumData[j] * leftSpectrumData[j];
		}

		// 獨立迴圈 2：計算基頻能量
		for (int j = -peakWidth; j <= peakWidth; ++j) {
			int bin = fundamentalBin + j;
			if (bin >= minBin && bin < freqBins) {
				fundamentalEnergy += leftSpectrumData[bin] * leftSpectrumData[bin];
			}
		}

		leftAnalysis.TotalEnergy = totalEnergy;
		leftAnalysis.fundamentalEnergy = fundamentalEnergy;

		// 業界標準 THD+N 計算：(總能量 - 基頻能量) / 總能量
		double noiseAndDistortionEnergy = totalEnergy - fundamentalEnergy;
		if (noiseAndDistortionEnergy < 0) noiseAndDistortionEnergy = 0.0; // 浮點數防呆

		if (totalEnergy > 0 && noiseAndDistortionEnergy > 0) {
			double thd_n_ratio = sqrt(noiseAndDistortionEnergy / totalEnergy);
			leftAnalysis.thd_N_dB = 20 * log10(thd_n_ratio);
		}
		else {
			leftAnalysis.thd_N_dB = -INFINITY;
		}
	}

	// --- 5. 計算右聲道 THD+N (邏輯完全同左聲道) ---
	auto& rightAnalysis = mWAVE_PARM.bMuteTest ? mWAVE_PARM.rightMuteWAVE_ANALYSIS : mWAVE_PARM.rightWAVE_ANALYSIS;

	maxAmplitude = 0.0;
	fundamentalBin = minBin;
	for (int j = minBin; j < freqBins; ++j) {
		if (rightSpectrumData[j] > maxAmplitude) {
			maxAmplitude = rightSpectrumData[j];
			fundamentalBin = j;
		}
	}
	rightAnalysis.TotalEnergyPoint = fundamentalBin;

	if (!mWAVE_PARM.bMuteTest && (maxAmplitude < fullScale * 0.001)) {
		rightAnalysis.thd_N_dB = 0;
		rightAnalysis.TotalEnergyPoint = 0;
		rightAnalysis.TotalEnergy = 0;
		rightAnalysis.fundamentalEnergy = 0;
	}
	else {
		double totalEnergy = 0.0;
		double fundamentalEnergy = 0.0;

		double bandwidthHz = GetWaveParm().fundamentalBandwidthHz > 0 ? GetWaveParm().fundamentalBandwidthHz : 200.0;
		int peakWidth = static_cast<int>(round(bandwidthHz / freqPerBin));
		if (peakWidth < 6) peakWidth = 6;

		for (int j = minBin; j < freqBins; ++j) {
			totalEnergy += rightSpectrumData[j] * rightSpectrumData[j];
		}

		for (int j = -peakWidth; j <= peakWidth; ++j) {
			int bin = fundamentalBin + j;
			if (bin >= minBin && bin < freqBins) {
				fundamentalEnergy += rightSpectrumData[bin] * rightSpectrumData[bin];
			}
		}

		rightAnalysis.TotalEnergy = totalEnergy;
		rightAnalysis.fundamentalEnergy = fundamentalEnergy;

		double noiseAndDistortionEnergy = totalEnergy - fundamentalEnergy;
		if (noiseAndDistortionEnergy < 0) noiseAndDistortionEnergy = 0.0;

		if (totalEnergy > 0 && noiseAndDistortionEnergy > 0) {
			double thd_n_ratio = sqrt(noiseAndDistortionEnergy / totalEnergy);
			rightAnalysis.thd_N_dB = 20 * log10(thd_n_ratio);
		}
		else {
			rightAnalysis.thd_N_dB = -INFINITY;
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
	if (!handle) {
		strMacroResult = "LOG:ERROR_AUDIO_WAVE_FILE_OPEN#";
		return false;
	}

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
		wprintf(L"此檔案不是雙聲道！\n This file is not stereo!\n");
		strMacroResult = "LOG:ERROR_AUDIO_WAVE_FILE_OPEN#";
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
	// 資料已讀取完畢，必須立刻關閉檔案釋放 Handle
	mmioClose(handle, 0);
	// 開啟音效輸出裝置
	int iDevIndex = GetWaveOutDevice(mWAVE_PARM.WaveOutDev);
	HANDLE hPlayEvent = NULL;
	MMRESULT result;

	if (AutoClose) {
		hPlayEvent = CreateEvent(NULL, FALSE, FALSE, NULL);
		result = waveOutOpen(&hWaveOut, iDevIndex, &wfx, (DWORD_PTR)hPlayEvent, NULL, CALLBACK_EVENT);
	}
	else {
		result = waveOutOpen(&hWaveOut, iDevIndex, &wfx, NULL, NULL, CALLBACK_NULL);
	}

	if (result != MMSYSERR_NOERROR) {
		if (hPlayEvent) CloseHandle(hPlayEvent); // 開啟失敗要釋放
		strMacroResult = "LOG:ERROR_AUDIO_WAVEOUT_DEVICE_NOT_FIND#";
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
		DWORD expectedPlayTimeMs = (DataSize * 1000) / wfx.nAvgBytesPerSec;
		DWORD maxWaitTimeMs = expectedPlayTimeMs + 2000;

		DWORD waitResult = WaitForSingleObject(hPlayEvent, maxWaitTimeMs);

		ASAudio::StopPlayingWavFile(); 
		CloseHandle(hPlayEvent); 

		if (waitResult != WAIT_OBJECT_0) {
			wprintf(L"播放WAV檔案超時或裝置已斷線！\n" "Playback of WAV file timed out or device disconnected!\n");
			strMacroResult = "LOG:ERROR_AUDIO_WAVEOUT_PLAYBACK_TIMEOUT_OR_DISCONNECTED#";
			return false;
		}
		return true;
	}
	return true;
}

// 停止播放WAV檔案
void ASAudio::StopPlayingWavFile() {
	waveOutReset(hWaveOut);
	waveOutUnprepareHeader(hWaveOut, &WaveOutHeader, sizeof(WAVEHDR));
	waveOutClose(hWaveOut);
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
		wprintf(L"找不到指定的錄音裝置！\n Can't find the specified recording device!\n");
		strMacroResult = "LOG:ERROR_AUDIO_WAVEIN_DEVICE_NOT_FIND#";
		result = MMSYSERR_NODRIVER;
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

				mWAVE_PARM.isRecordingFinished = false;
				mWAVE_PARM.deviceError = false;
				waveInStart(hWaveIn);
			}
		}
		else
		{
			return result;
		}
	}
	return result;
}

// 停止錄音
MMRESULT ASAudio::StopRecording() {
	MMRESULT result = waveInStop(hWaveIn);
	if (result == MMSYSERR_NOERROR) {
		result = waveInReset(hWaveIn);
		waveInUnprepareHeader(hWaveIn, &waveHdr, sizeof(WAVEHDR));
		result = waveInClose(hWaveIn);
		return result;
	}
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
					// 捨棄第一個音訊緩衝區，直接重新AddBuffer
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

	// 1. 計算實際錄音取得的有效樣本數 (雙聲道，每個樣本佔 sizeof(short) * 2 = 4 bytes)
	int validSamples = waveHdr.dwBytesRecorded / (sizeof(short) * 2);

	// 防禦性設計：確保不超過預設的緩衝區大小
	if (validSamples > BUFFER_SIZE) {
		validSamples = BUFFER_SIZE;
	}
	std::lock_guard<std::mutex> lock(mWAVE_PARM.recordingMutex);
	// 2. 處理實際錄製到的有效資料
	for (int i = 0; i < validSamples; i++) {
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

	// 3. 將未填滿的剩餘空間補 0，避免 FFT 運算受到記憶體中的殘留垃圾資料干擾
	for (int i = validSamples; i < BUFFER_SIZE; i++) {
		mWAVE_PARM.WAVE_DATA.LeftAudioBuffer[i] = 0;
		fftInputBufferLeft[i][0] = 0.0;
		fftInputBufferLeft[i][1] = 0.0;

		mWAVE_PARM.WAVE_DATA.RightAudioBuffer[i] = 0;
		fftInputBufferRight[i][0] = 0.0;
		fftInputBufferRight[i][1] = 0.0;
	}

	// 通知錄音完成
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
	DWORD dwResult = 0;
	CComPtr<IMMDeviceEnumerator> pEnumerator; // 改用 CComPtr

	HRESULT hr = pEnumerator.CoCreateInstance(__uuidof(MMDeviceEnumerator));
	if (SUCCEEDED(hr)) {
		CComPtr<IMMDeviceCollection> pDevices; // 改用 CComPtr
		hr = pEnumerator->EnumAudioEndpoints(eRender, DEVICE_STATE_ACTIVE, &pDevices);
		if (SUCCEEDED(hr)) {
			UINT deviceCount = 0;
			pDevices->GetCount(&deviceCount);
			if (deviceCount > 0) {
				int iDevIndex = GetIMMWaveOutDevice(pDevices, mWAVE_PARM.WaveOutDev);
				if (iDevIndex == -1) {
					dwResult = MMSYSERR_INVALHANDLE;
				}
				else {
					CComPtr<IMMDevice> pDevice; // 改用 CComPtr
					pDevices->Item(iDevIndex, &pDevice);
					if (pDevice) {
						CComPtr<IAudioEndpointVolume> pAudioEndpointVolume; // 改用 CComPtr
						hr = pDevice->Activate(__uuidof(IAudioEndpointVolume), CLSCTX_ALL, NULL, (void**)&pAudioEndpointVolume);
						if (SUCCEEDED(hr) && pAudioEndpointVolume) {
							float targetTemp = mWAVE_PARM.bMuteTest ? 0.0f : ((float)mWAVE_PARM.WaveOutVolume / 100.0f);
							hr = pAudioEndpointVolume->SetMasterVolumeLevelScalar(targetTemp, NULL);

							if (SUCCEEDED(hr)) {
								dwResult = 0;
							}
						}
					}
				}
			}
		}
	}
	return MMSYSERR_NOERROR;
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
		if (rc != MMSYSERR_NOERROR) {
			break; 
		}
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
				continue;
			}
			return; // 成功設定，就離開
		}
		mixerClose(hMixer);
	}
}

std::wstring ASAudio::charToWstring(const char* szIn)
{
	if (!szIn) return L"";
	int length = MultiByteToWideChar(CP_ACP, 0, szIn, -1, NULL, 0);
	if (length <= 0) return L"";

	std::wstring strRet(length - 1, L'\0');
	MultiByteToWideChar(CP_ACP, 0, szIn, -1, &strRet[0], length);
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
		wprintf(L"找不到指定的音效輸出裝置！\n Can't find the specified audio output device!\n");
		return 1;
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
		wprintf(L"開啟音訊串流失敗！\nFailed to open audio stream!\n%s\n", Pa_GetErrorText(err));
		return 1;
	}

	data.stream = stream; // 保存串流指標到結構中

	// 開始音訊串流
	err = Pa_StartStream(stream);
	if (err != paNoError) {
		wprintf(L"開始音訊串流失敗！\nFailed to start audio stream!\n%s\n", Pa_GetErrorText(err));
		Pa_CloseStream(stream);
		return 1;
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
				int bufferSize = WideCharToMultiByte(CP_UTF8, 0, szOutDevName.c_str(), (int)szOutDevName.length(), NULL, 0, NULL, NULL);
				std::string targetName(bufferSize, '\0');
				WideCharToMultiByte(CP_UTF8, 0, szOutDevName.c_str(), (int)szOutDevName.length(), &targetName[0], bufferSize, NULL, NULL);

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
				int bufferSize = WideCharToMultiByte(CP_UTF8, 0, szInDevName.c_str(), (int)szInDevName.length(), NULL, 0, NULL, NULL);
				std::string targetName(bufferSize, '\0');
				WideCharToMultiByte(CP_UTF8, 0, szInDevName.c_str(), (int)szInDevName.length(), &targetName[0], bufferSize, NULL, NULL);
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
	(void)inputBuffer;

	for (i = 0; i < framesPerBuffer; i++) {
		*out++ = sinf(data->phaseL);
		data->phaseL += (float)(2 * M_PI * data->frequencyL / SAMPLE_RATE);
		if (data->phaseL >= (float)(2 * M_PI)) data->phaseL -= (float)(2 * M_PI); 

		*out++ = sinf(data->phaseR);
		data->phaseR += (float)(2 * M_PI * data->frequencyR / SAMPLE_RATE);
		if (data->phaseR >= (float)(2 * M_PI)) data->phaseR -= (float)(2 * M_PI); 
	}
	return paContinue;
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
	StopAudioLoopback();
	HRESULT hr;

	// 建立 Graph (圖形管理器) 和 Capture Graph Builder
	// CComPtr 會在物件建立失敗時保持為 NULL
	hr = pGraph.CoCreateInstance(CLSID_FilterGraph);
	if (FAILED(hr)) {
		wprintf(L"建立圖形管理器失敗！錯誤碼: 0x%08X\n Building graph failed! Error code: 0x%08X\n", hr);
		strMacroResult = "LOG:ERROR_AUDIO_LOOPBACK_SETTING#";
		return false;
	}

	hr = pCaptureGraph.CoCreateInstance(CLSID_CaptureGraphBuilder2);
	if (FAILED(hr)) {
		wprintf(L"建立 Capture Graph Builder 失敗！錯誤碼: 0x%08X\n Building Capture Graph Builder failed! Error code: 0x%08X\n", hr);
		strMacroResult = "LOG:ERROR_AUDIO_LOOPBACK_SETTING#";
		return false;
	}

	// 將 filter graph 關聯到 capture graph builder
	hr = pCaptureGraph->SetFiltergraph(pGraph);
	if (FAILED(hr)) {
		wprintf(L"將 Filter Graph 關聯到 Capture Graph Builder 失敗！錯誤碼: 0x%08X\n Setting filter graph failed! Error code: 0x%08X\n", hr);
		strMacroResult = "LOG:ERROR_AUDIO_LOOPBACK_SETTING#";
		return false;
	}

	// 建立裝置列舉器
	CComPtr<ICreateDevEnum> pDevEnum;
	hr = pDevEnum.CoCreateInstance(CLSID_SystemDeviceEnum);
	if (FAILED(hr)) {
		wprintf(L"建立裝置列舉器失敗！錯誤碼: 0x%08X\n Creating device enumerator failed! Error code: 0x%08X\n", hr);
		strMacroResult = "LOG:ERROR_AUDIO_LOOPBACK_SETTING#";
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
		wprintf(L"找不到指定的錄音裝置！\n Can't find the specified capture device!\n");
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
		wprintf(L"找不到指定的播放裝置！\n Can't find the specified render device!\n");
		strMacroResult = "LOG:ERROR_AUDIO_WAVEOUT_DEVICE_NOT_FIND#";
		return false; // CComPtr 會自動釋放已建立的物件
	}

	// 將篩選器加入圖形
	hr = pGraph->AddFilter(pAudioCapture, L"Audio Capture");
	if (FAILED(hr)) {
		wprintf(L"加入擷取篩選器失敗！錯誤碼: 0x%08X\n Adding capture filter failed! Error code: 0x%08X\n", hr);
		strMacroResult = "LOG:ERROR_AUDIO_LOOPBACK_SETTING#";
		return false;
	}

	hr = pGraph->AddFilter(pAudioRenderer, L"Audio Renderer");
	if (FAILED(hr)) {
		wprintf(L"加入渲染篩選器失敗！錯誤碼: 0x%08X\n Adding render filter failed! Error code: 0x%08X\n", hr);
		strMacroResult = "LOG:ERROR_AUDIO_LOOPBACK_SETTING#";
		return false;
	}

	// 連接擷取與渲染裝置的資料流
	hr = pCaptureGraph->RenderStream(&PIN_CATEGORY_CAPTURE, &MEDIATYPE_Audio, pAudioCapture, nullptr, pAudioRenderer);
	if (FAILED(hr)) {
		wprintf(L"連接擷取與渲染裝置失敗！錯誤碼: 0x%08X\n Connecting capture and render devices failed! Error code: 0x%08X\n", hr);
		strMacroResult = "LOG:ERROR_AUDIO_LOOPBACK_SETTING#";
		return false;
	}

	// 取得媒體控制介面並執行
	hr = pGraph->QueryInterface(IID_PPV_ARGS(&pMediaControl));
	if (FAILED(hr) || !pMediaControl) {
		wprintf(L"取得媒體控制介面失敗！錯誤碼: 0x%08X\n Getting media control interface failed! Error code: 0x%08X\n", hr);
		strMacroResult = "LOG:ERROR_AUDIO_LOOPBACK_SETTING#";
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
	// SoundVolumeView.exe /SetDefault "{deviceId}" 0
	// 0: All, 1: Console, 2: Multimedia
	wprintf(L"SoundVolumeView.exe /SetDefault \"%s\" 0\n", deviceId.c_str());
	std::wstring command = L"SoundVolumeView.exe /SetDefault \"" + deviceId + L"\" 0";

	int result = _wsystem(command.c_str());

	// 回傳是否成功(0 代表成功)
	return result == 0;
}
bool ASAudio::SetListenToThisDevice(const std::wstring& deviceId, int enable)
{
	// SoundVolumeView.exe /SetListenToThisDevice "{deviceId}" {1 or 0}
	wprintf(L"SoundVolumeView.exe /SetListenToThisDevice \"%s\" %d\n", deviceId.c_str(), enable);
	std::wstring command = L"SoundVolumeView.exe /SetListenToThisDevice \"" + deviceId + L"\" " + std::to_wstring(enable);

	int result = _wsystem(command.c_str());
	return result == 0;
}

bool ASAudio::FindDeviceIdByName(Config& config, std::wstring& outDeviceId, EDataFlow dataFlow)
{
	if (config.monitorNames.empty()) {
		wprintf(L"未提供目標名稱！\n No target names provided!\n");
		strMacroResult = "LOG:ERROR_CONFIG_NOT_FIND#";
		return false;
	}

	CComPtr<IMMDeviceEnumerator> enumerator;
	CComPtr<IMMDeviceCollection> collection;

	HRESULT hr = CoCreateInstance(__uuidof(MMDeviceEnumerator), NULL, CLSCTX_ALL,
		__uuidof(IMMDeviceEnumerator), (void**)&enumerator);

	if (FAILED(hr)) {
		wprintf(L"初始化音訊裝置枚舉器失敗！錯誤碼: 0x%08X\n", hr);
		strMacroResult = "LOG:ERROR_AUDIO_FIND_DEVICE#";
		return false;
	}

	hr = enumerator->EnumAudioEndpoints(dataFlow, DEVICE_STATE_ACTIVE, &collection);
	if (FAILED(hr)) {
		wprintf(L"獲取音訊裝置失敗！錯誤碼: 0x%08X\n", hr);
		strMacroResult = "LOG:ERROR_AUDIO_FIND_DEVICE#";
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
		CComPtr<IMMDevice> device;
		hr = collection->Item(i, &device);
		if (FAILED(hr) || device == nullptr) {
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
	wprintf(L"找不到符合條件的裝置！\n Can't find a matching device!\n");
	strMacroResult = "LOG:ERROR_AUDIO_FIND_DEVICE#";
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
		wprintf(L"找不到指定的錄音裝置！\n Can't find the specified recording device!\n");
		strMacroResult = "LOG:ERROR_AUDIO_WAVEIN_DEVICE_NOT_FIND#";
		return MMSYSERR_NODRIVER;
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

		mWAVE_PARM.isRecordingFinished = false;
		mWAVE_PARM.deviceError = false;
		waveInStart(hWaveIn);
	}
	return result;
}
