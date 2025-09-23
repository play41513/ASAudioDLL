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

void ASAudio::SpectrumAnalysis(double* leftSpectrumData, double* rightSpectrumData)
{
	const double Fs = 44100; // 取樣頻率
	const int freqBins = BUFFER_SIZE / 2 + 1; // 頻率bin數量
	double leftMagnitude, rightMagnitude; // 當前頻率bin的振幅

	auto& leftTempWAVE_ANALYSIS = mWAVE_PARM.bMuteTest ? mWAVE_PARM.leftMuteWAVE_ANALYSIS : mWAVE_PARM.leftWAVE_ANALYSIS;
	auto& RightTempWAVE_ANALYSIS = mWAVE_PARM.bMuteTest ? mWAVE_PARM.rightMuteWAVE_ANALYSIS : mWAVE_PARM.rightWAVE_ANALYSIS;
	leftTempWAVE_ANALYSIS.TotalEnergyPoint = 0; // 左聲道最大能量點歸零
	RightTempWAVE_ANALYSIS.TotalEnergyPoint = 0; // 右聲道最大能量點歸零
	leftTempWAVE_ANALYSIS.TotalEnergy = 0; // 左聲道總能量歸零
	RightTempWAVE_ANALYSIS.TotalEnergy = 0; // 右聲道總能量歸零

	double maxLeftMagnitude = 0;
	double maxRightMagnitude = 0;

	// 計算總能量和最大能量點
	for (int i = 2; i < freqBins; ++i) {//頻率0和1不計算
		leftMagnitude = leftSpectrumData[i]; // 當前左聲道頻率bin的振幅
		rightMagnitude = rightSpectrumData[i]; // 當前右聲道頻率bin的振幅

		// 累加所有頻率的能量（振幅的平方）
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

	double Harmonic1, Harmonic2; // 諧波振幅
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
			: -INFINITY; // 避免log(0)
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
			: -INFINITY; // 避免log(0)
	}
	else {
		RightTempWAVE_ANALYSIS.thd_N_dB = 0;
	}

	double totalEnergy = 0;
	for (int i = 0; i < freqBins; ++i) {
		double magnitude = leftSpectrumData[i];
		totalEnergy += magnitude * magnitude; // 能量是振幅的平方
	}
	double volume_dB = 10 * log10(totalEnergy);
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
	if (uMsg == MM_WIM_DATA) {
		// 取得目前 ASAudio 物件的指標
		ASAudio* asAudio = reinterpret_cast<ASAudio*>(dwInstance);
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
	}
}

// 處理音訊緩衝區
void ASAudio::ProcessAudioBuffer() {
	short* audioSamples = reinterpret_cast<short*>(waveHdr.lpData);
	for (int i = 0; i < BUFFER_SIZE; i++) {
		mWAVE_PARM.WAVE_DATA.audioBuffer[bufferIndex + (size_t)i * 2] = reinterpret_cast<short*>(waveHdr.lpData)[i * 2];
		mWAVE_PARM.WAVE_DATA.audioBuffer[bufferIndex + (size_t)i * 2 + 1] = reinterpret_cast<short*>(waveHdr.lpData)[i * 2 + 1];
	}
	bufferIndex += (size_t)BUFFER_SIZE * 2;

	if (bufferIndex >= (size_t)BUFFER_SIZE * 2) {
		// 已經錄到足夠的樣本數(BUFFER_SIZE*2)
		// 填滿 FFT 緩衝區
		for (int i = 0; i < (size_t)BUFFER_SIZE; i++) {
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

		// 通知錄音完成
		std::lock_guard<std::mutex> lock(mWAVE_PARM.recordingMutex);
		mWAVE_PARM.isRecordingFinished = true;
		mWAVE_PARM.recordingFinishedCV.notify_one();

		bufferIndex = 0;
	}
}

DWORD ASAudio::SetMicSystemVolume()
{
	//宣告混音器相關結構
	MMRESULT rc;
	HMIXER hMixer;
	MIXERLINE mxl;
	MIXERLINECONTROLS mxlc;
	MIXERCONTROL mxc;

	//遍歷所有混音器，找到錄音相關的混音器，並設定裝置ID
	for (int deviceID = 0; true; deviceID++)
	{
		//開啟指定混音器，deviceID為混音器ID
		rc = mixerOpen(&hMixer, deviceID, 0, 0, MIXER_OBJECTF_MIXER);
		if (rc != MMSYSERR_NOERROR)
			break;
		ZeroMemory(&mxl, sizeof(MIXERLINE));
		mxl.cbStruct = sizeof(MIXERLINE);
		//指定要查詢的線路類型
		//線路類型為目標的為MIXERLINE_COMPONENTTYPE_DST_SPEAKERS
		//線路類型為來源的為MIXERLINE_COMPONENTTYPE_DST_WAVEIN
		mxl.dwComponentType = MIXERLINE_COMPONENTTYPE_DST_WAVEIN;
		rc = mixerGetLineInfo((HMIXEROBJ)hMixer, &mxl, MIXER_GETLINEINFOF_COMPONENTTYPE);
		if (rc != MMSYSERR_NOERROR)
			continue;
		//取得混音器裝置所支援的音訊線路數量，並逐一檢查
		DWORD dwConnections = mxl.cConnections;
		DWORD dwLineID = -1;
		for (DWORD i = 0; i < dwConnections; i++)
		{
			mxl.dwSource = i;
			//依據SourceID取得線路資訊
			rc = mixerGetLineInfo((HMIXEROBJ)hMixer, &mxl, MIXER_OBJECTF_HMIXER | MIXER_GETLINEINFOF_SOURCE);
			if (rc == MMSYSERR_NOERROR)
			{
				//如果該裝置是麥克風，則跳出
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

		//指定要查詢的控制項
		mxlc.cbStruct = sizeof(mxlc);
		mxlc.dwLineID = dwLineID;
		mxlc.dwControlType = MIXERCONTROL_CONTROLTYPE_VOLUME;
		mxlc.cControls = 1;
		mxlc.pamxctrl = &mxc;
		mxlc.cbmxctrl = sizeof(mxc);

		//取得音量控制資訊
		rc = mixerGetLineControls((HMIXEROBJ)hMixer, &mxlc, MIXER_GETLINECONTROLSF_ONEBYTYPE);
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

			//取得音量值，取得的資訊放在mxcd
			rc = mixerGetControlDetails((HMIXEROBJ)hMixer, &mxcd, MIXER_GETCONTROLDETAILSF_VALUE);

			//初始化音量大小資訊
			MIXERCONTROLDETAILS_UNSIGNED mxcdVolume_Set = { mxc.Bounds.dwMaximum * mWAVE_PARM.WaveInVolume / 100 };
			MIXERCONTROLDETAILS mxcd_Set = { 0 };
			mxcd_Set.cbStruct = sizeof(MIXERCONTROLDETAILS);
			mxcd_Set.dwControlID = mxc.dwControlID;
			mxcd_Set.cChannels = 1;
			mxcd_Set.cMultipleItems = 0;
			mxcd_Set.cbDetails = sizeof(MIXERCONTROLDETAILS_UNSIGNED);
			mxcd_Set.paDetails = &mxcdVolume_Set;

			//設定音量大小
			mixerSetControlDetails((HMIXEROBJ)(hMixer), &mxcd_Set, MIXER_OBJECTF_HMIXER | MIXER_SETCONTROLDETAILSF_VALUE);
			mixerClose(hMixer);
		}
	}
	return 0;
}

DWORD ASAudio::SetSpeakerSystemVolume() const
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
	// 取得音效輸出裝置的數量
	UINT deviceCount = waveOutGetNumDevs();
	for (UINT i = 0; i < deviceCount; ++i) {
		WAVEOUTCAPSA waveOutCaps;
		MMRESULT result = waveOutGetDevCapsA(i, &waveOutCaps, sizeof(WAVEOUTCAPSA));
		if (result == MMSYSERR_NOERROR) {
			std::string deviceName(waveOutCaps.szPname); // 將裝置名稱轉為 std::string
			// 將szOutDevName轉為std::string
			int bufferSize = WideCharToMultiByte(CP_UTF8, 0, szOutDevName.c_str(), -1, NULL, 0, NULL, NULL) - 1;//這裡需要-1，不然會導致find找不到正確的字串數量
			std::string targetName(bufferSize, ' ');
			WideCharToMultiByte(CP_UTF8, 0, szOutDevName.c_str(), -1, &targetName[0], bufferSize, NULL, NULL);

			if (deviceName.find(targetName) != std::string::npos) {
				return i; // 若含有指定關鍵字的裝置，回傳裝置 ID
			}
		}
	}
	return -1; // 找不到裝置，回傳 -1
}

int ASAudio::GetWaveInDevice(std::wstring szInDevName)
{
	// 取得音效輸入裝置的數量
	UINT deviceCount = waveInGetNumDevs();
	for (UINT i = 0; i < deviceCount; ++i) {
		WAVEINCAPSA waveInCaps;
		MMRESULT result = waveInGetDevCapsA(i, &waveInCaps, sizeof(WAVEINCAPSA));
		if (result == MMSYSERR_NOERROR) {
			std::string deviceName(waveInCaps.szPname); // 將裝置名稱轉為 std::string
			// 將szInDevName轉為std::string
			int bufferSize = WideCharToMultiByte(CP_UTF8, 0, szInDevName.c_str(), -1, NULL, 0, NULL, NULL) - 1;//這裡需要-1，不然會導致find找不到正確的字串數量
			std::string targetName(bufferSize, ' ');
			WideCharToMultiByte(CP_UTF8, 0, szInDevName.c_str(), -1, &targetName[0], bufferSize, NULL, NULL);

			if (deviceName.find(targetName) != std::string::npos) {
				return i; // 若含有指定關鍵字的裝置，回傳裝置 ID
			}
		}
	}
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

bool ASAudio::StartAudioLoopback(const std::wstring captureKeyword, const std::wstring renderKeyword) {
	HRESULT hr;
	HRESULT hrInit = CoInitialize(nullptr);
	if (FAILED(hrInit)) {
		strMacroResult = "LOG:ERROR_AUDIO_LOOPBACK_COM_INITIALIZE#";
		return false;
	}

	// 建立 Graph (圖形管理器)
	hr = CoCreateInstance(CLSID_FilterGraph, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&pGraph));
	if (FAILED(hr)) { return false; }
	hr = CoCreateInstance(CLSID_CaptureGraphBuilder2, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&pCaptureGraph));
	if (FAILED(hr)) { CoUninitialize(); return false; }
	pCaptureGraph->SetFiltergraph(pGraph);

	// 建立裝置列舉器
	CComPtr<ICreateDevEnum> pDevEnum;
	hr = CoCreateInstance(CLSID_SystemDeviceEnum, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&pDevEnum));

	// --- 尋找指定的擷取與渲染裝置的篩選器 ---
	if (SUCCEEDED(hr)) {
		CComPtr<IEnumMoniker> pEnum;
		if (SUCCEEDED(pDevEnum->CreateClassEnumerator(CLSID_AudioInputDeviceCategory, &pEnum, 0)) && pEnum) {
			// 將 pMoniker 宣告移至迴圈內
			while (true) {
				CComPtr<IMoniker> pMoniker;
				if (pEnum->Next(1, &pMoniker, nullptr) != S_OK) {
					break; // 找不到更多裝置，跳出迴圈
				}

				CComPtr<IPropertyBag> pPropBag;
				if (SUCCEEDED(pMoniker->BindToStorage(nullptr, nullptr, IID_PPV_ARGS(&pPropBag))) && pPropBag) {
					VARIANT varName;
					VariantInit(&varName);
					if (SUCCEEDED(pPropBag->Read(L"FriendlyName", &varName, nullptr))) {
						if (wcsstr(varName.bstrVal, captureKeyword.c_str())) {
							pMoniker->BindToObject(nullptr, nullptr, IID_IBaseFilter, (void**)&pAudioCapture);
						}
					}
					VariantClear(&varName);
				}
				if (pAudioCapture) break; // 找到裝置，跳出迴圈
			}
		}
	}

	if (!pAudioCapture) {
		strMacroResult = "LOG:ERROR_AUDIO_WAVEIN_DEVICE_NOT_FIND#";
		CoUninitialize();
		return false;
	}

	if (SUCCEEDED(hr)) {
		CComPtr<IEnumMoniker> pEnum;
		if (SUCCEEDED(pDevEnum->CreateClassEnumerator(CLSID_AudioRendererCategory, &pEnum, 0)) && pEnum) {
			while (true) {
				CComPtr<IMoniker> pMoniker;
				if (pEnum->Next(1, &pMoniker, nullptr) != S_OK) {
					break;
				}

				CComPtr<IPropertyBag> pPropBag;
				if (SUCCEEDED(pMoniker->BindToStorage(nullptr, nullptr, IID_PPV_ARGS(&pPropBag))) && pPropBag) {
					VARIANT varName;
					VariantInit(&varName);
					if (SUCCEEDED(pPropBag->Read(L"FriendlyName", &varName, nullptr))) {
						if (wcsstr(varName.bstrVal, renderKeyword.c_str())) {
							pMoniker->BindToObject(nullptr, nullptr, IID_IBaseFilter, (void**)&pAudioRenderer);
						}
					}
					VariantClear(&varName);
				}
				if (pAudioRenderer) break;
			}
		}
	}

	if (!pAudioRenderer) {
		strMacroResult = "LOG:ERROR_AUDIO_LOOPBACK_NO_MATCH_RENDER_DEVICE#";
		CoUninitialize();
		return false;
	}

	// 將篩選器加入圖形...
	pGraph->AddFilter(pAudioCapture, L"Audio Capture");
	pGraph->AddFilter(pAudioRenderer, L"Audio Renderer");
	hr = pCaptureGraph->RenderStream(&PIN_CATEGORY_CAPTURE, &MEDIATYPE_Audio, pAudioCapture, nullptr, pAudioRenderer);
	if (FAILED(hr)) {
		strMacroResult = "LOG:ERROR_AUDIO_LOOPBACK_RENDER_STREAM#";
		return false;
	}
	pGraph->QueryInterface(IID_PPV_ARGS(&pMediaControl));
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

	CoUninitialize();
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
