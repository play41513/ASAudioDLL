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

		// ���o�˸m
		if (SUCCEEDED(pDevices->Item(i, &pDevice))) {
			// ���o�˸m�ݩ�
			if (SUCCEEDED(pDevice->OpenPropertyStore(STGM_READ, &pProps))) {
				PROPVARIANT varName;
				PropVariantInit(&varName);

				// ���o�˸m�W��
				if (SUCCEEDED(pProps->GetValue(PKEY_Device_FriendlyName, &varName))) {
					std::wstring deviceName(varName.pwszVal);

					// �ˬd�W�٬O�_�۲�
					if (deviceName.find(szOutDevName) != std::wstring::npos) {
						PropVariantClear(&varName);
						return i; // ���۲Ÿ˸m
					}
				}
				PropVariantClear(&varName);
			}
		}
	}
	return -1; // �䤣��˸m
}

// �N std::string �ର std::wstring�]UTF-8 �� UTF-16�^
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

	memset(&waveHdr, 0, sizeof(WAVEHDR)); // ��l�ƿ�����WAVEHDR���c
	memset(mWAVE_PARM.WAVE_DATA.audioBuffer, 0, sizeof(mWAVE_PARM.WAVE_DATA.audioBuffer)); // ��l�ƭ��T�w�İϡA�k�s

	// ��l�� FFT �������O����M�p�e
	fftInputBufferLeft = (fftw_complex*)fftw_malloc(sizeof(fftw_complex) * BUFFER_SIZE);
	fftInputBufferRight = (fftw_complex*)fftw_malloc(sizeof(fftw_complex) * BUFFER_SIZE);
	fftOutputBufferLeft = (fftw_complex*)fftw_malloc(sizeof(fftw_complex) * BUFFER_SIZE);
	fftOutputBufferRight = (fftw_complex*)fftw_malloc(sizeof(fftw_complex) * BUFFER_SIZE);

	// �Ы� FFT �p�e�A�H�K�N�Ӷi��ֳt�ť߸��ܴ�
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

	// �b�C�����ݤ��e�A���m���A�X��
	mWAVE_PARM.isRecordingFinished = false;
	mWAVE_PARM.deviceError = false;

	// �ϥ� wait_for ���ݡA���|�b���w�ɶ���۰ʪ�^
	// �ĤG�ӰѼƬO�@�� lambda �禡�A�Ω󨾤� "�����" (spurious wakeups)
	if (!mWAVE_PARM.recordingFinishedCV.wait_for(lock, std::chrono::seconds(timeoutSeconds), [&] {
		return mWAVE_PARM.isRecordingFinished;
		}))
	{
		// �p�G wait_for ��^ false�A�N��O�u�O�ɡv�Ӥ��O�Q notify ���
		mWAVE_PARM.lastRecordingResult = WAVE_PARM::RecordingResult::Timeout;
		return WAVE_PARM::RecordingResult::Timeout;
	}

	// �{�������o�̡A�N��O�Q notify ����C
	// �ˬd�O�u���`�����v�٬O�u�˸m���~�v�y��������C
	if (mWAVE_PARM.deviceError) {
		mWAVE_PARM.lastRecordingResult = WAVE_PARM::RecordingResult::DeviceError;
		return WAVE_PARM::RecordingResult::DeviceError;
	}

	// �p�G�S���O�ɡA�]�S���˸m���~�A���N��������\�����C
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

	// --- 2. ���� FFT ---
	fftw_execute_dft(fftPlanLeft, windowedInputLeft, fftOutputBufferLeft);
	fftw_execute_dft(fftPlanRight, windowedInputRight, fftOutputBufferRight);

	fftw_free(windowedInputLeft);
	fftw_free(windowedInputRight);

	// --- 3. �p�⥿�T�Y���W�Ю��T (Amplitude) ---
	// (�W���Y��) & #2 (����Ƹ��v)
	// Coherent Gain of Hann window is sum of window samples divided by N
	double coherentGain = windowSum / N;
	double scalingFactor = 2.0 / (N * coherentGain);

	for (int i = 0; i < freqBins; ++i) {
		double rawMagnitudeL = sqrt(fftOutputBufferLeft[i][0] * fftOutputBufferLeft[i][0] + fftOutputBufferLeft[i][1] * fftOutputBufferLeft[i][1]);
		double rawMagnitudeR = sqrt(fftOutputBufferRight[i][0] * fftOutputBufferRight[i][0] + fftOutputBufferRight[i][1] * fftOutputBufferRight[i][1]);

		// ��� DC (i=0) �M Nyquist (�p�G N �O���ơA�b i=N/2) ���q�A�����H 2
		if (i == 0) {
			leftSpectrumData[i] = rawMagnitudeL / (N * coherentGain);
			rightSpectrumData[i] = rawMagnitudeR / (N * coherentGain);
		}
		else {
			leftSpectrumData[i] = rawMagnitudeL * scalingFactor;
			rightSpectrumData[i] = rawMagnitudeR * scalingFactor;
		}
	}

	// --- ���n�D ---
	auto& leftAnalysis = mWAVE_PARM.bMuteTest ? mWAVE_PARM.leftMuteWAVE_ANALYSIS : mWAVE_PARM.leftWAVE_ANALYSIS;
	{
		// 4a. �M����W (Fundamental) ���̰��I
		double maxAmplitude = 0.0;
		int fundamentalBin = 0;
		for (int i = 2; i < freqBins; ++i) { 
			if (leftSpectrumData[i] > maxAmplitude) {
				maxAmplitude = leftSpectrumData[i];
				fundamentalBin = i;
			}
		}
		leftAnalysis.TotalEnergyPoint = fundamentalBin;

		// -60 dBFS �O�@�ӦX�z���H�ȡA�N��T���j�קC��̤j�Ȫ� 0.1%
		double fullScale = 32767.0;
		if (!mWAVE_PARM.bMuteTest && (maxAmplitude < fullScale * 0.001)) { // ~ -60 dBFS
			leftAnalysis.thd_N_dB = 0; // �ϥ� 0 dB ��ܿ��~/�L�İT��
			leftAnalysis.TotalEnergyPoint = 0;
			leftAnalysis.TotalEnergy = 0; // �̦n�]�N��q�k�s
			leftAnalysis.fundamentalEnergy = 0;
		}
		else {
			// 4c. �ץ��I #3 & #4: �ϥΥ��T����q�p�� THD+N
			double totalEnergy = 0;
			double fundamentalEnergy = 0;
			const double sampleRate = 44100.0;

			double bandwidthHz = GetWaveParm().fundamentalBandwidthHz; // ���] config �Q�s�b WAVE_PARM
			if (bandwidthHz <= 0) bandwidthHz = 100.0;
			double freqPerBin = sampleRate / N;
			int peakWidth = static_cast<int>(round(bandwidthHz / freqPerBin));

			// �p���`��q
			for (int i = 2; i < freqBins; ++i) {
				totalEnergy += leftSpectrumData[i] * leftSpectrumData[i];
			}

			// ��X���W��q
			for (int i = -peakWidth; i <= peakWidth; ++i) {
				int bin = fundamentalBin + i;
				if (bin > 1 && bin < freqBins) {
					fundamentalEnergy += leftSpectrumData[bin] * leftSpectrumData[bin];
				}
			}

			leftAnalysis.TotalEnergy = totalEnergy;
			leftAnalysis.fundamentalEnergy = fundamentalEnergy;

			// �p�� THD+N
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

	// --- 5. �w��k�n�D�i����R---
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
			double bandwidthHz = GetWaveParm().fundamentalBandwidthHz; // ���] config �Q�s�b WAVE_PARM
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
// ����WAV�ɮ�
bool ASAudio::PlayWavFile(bool AutoClose) {
	MMCKINFO ChunkInfo;
	MMCKINFO FormatChunkInfo = { 0 };
	MMCKINFO DataChunkInfo = { 0 };
	WAVEFORMATEX wfx;
	int DataSize;

	// Zero out the ChunkInfo structure.
	memset(&ChunkInfo, 0, sizeof(MMCKINFO));

	// �}��WAVE�ɡA�^�Ǥ@��HMMIO����N�X
	HMMIO handle = mmioOpen((LPWSTR)mWAVE_PARM.AudioFile.c_str(), 0, MMIO_READ);

	// �i�JRIFF�϶�(RIFF Chunk)
	mmioDescend(handle, &ChunkInfo, 0, MMIO_FINDRIFF);

	// �i�Jfmt�϶�(RIFF���l�϶��A���t���T�榡��T)
	FormatChunkInfo.ckid = mmioStringToFOURCCA("fmt", 0); // �M��fmt�l�϶�
	mmioDescend(handle, &FormatChunkInfo, &ChunkInfo, MMIO_FINDCHUNK);

	// Ū��wav�榡��T
	memset(&wfx, 0, sizeof(WAVEFORMATEX));
	mmioRead(handle, (char*)&wfx, FormatChunkInfo.cksize);

	// �ˬd�n�D�ƶq
	if (wfx.nChannels != 2) {
		// �p�G���O���n�D
		mmioClose(handle, 0);
		strMacroResult = "LOG:ERROR_AUDIO_WAVE_FILE_FORMAT_NOT_SUPPORT#";
		return false;
	}

	// ���}fmt�϶�
	mmioAscend(handle, &FormatChunkInfo, 0);

	// �i�Jdata�϶�(���t��ڪ����T�˥�)
	DataChunkInfo.ckid = mmioStringToFOURCCA("data", 0);
	mmioDescend(handle, &DataChunkInfo, &ChunkInfo, MMIO_FINDCHUNK);

	// ���odata�϶����j�p
	DataSize = DataChunkInfo.cksize;

	// �ǳƽw�İ�
	WaveOutData.resize(DataSize);
	mmioRead(handle, WaveOutData.data(), DataSize);

	// �}�ҭ��Ŀ�X�˸m
	int iDevIndex = GetWaveOutDevice(mWAVE_PARM.WaveOutDev);
	MMRESULT result = waveOutOpen(&hWaveOut, iDevIndex, &wfx, 0, 0, WAVE_FORMAT_QUERY);
	if (result != MMSYSERR_NOERROR) {
		// ����
		strMacroResult = "LOG:ERROR_AUDIO_WAVEOUT_DEVICE_NOT_FIND_2#";
		return false;
	}

	result = waveOutOpen(&hWaveOut, iDevIndex, &wfx, NULL, NULL, CALLBACK_NULL);
	if (result != MMSYSERR_NOERROR) {
		// ����
		waveOutClose(hWaveOut);
		strMacroResult = "LOG:ERROR_AUDIO_WAVEOUT_DEVICE_NOT_FIND_3#";
		return false;
	}

	// �]�wwave header.
	memset(&WaveOutHeader, 0, sizeof(WaveOutHeader));
	WaveOutHeader.lpData = WaveOutData.data();
	WaveOutHeader.dwBufferLength = DataSize;

	// �ǳ�wave header.
	waveOutPrepareHeader(hWaveOut, &WaveOutHeader, sizeof(WAVEHDR));

	// �N�w�İϸ�Ƽg�J���ĸ˸m(�}�l����).
	waveOutWrite(hWaveOut, &WaveOutHeader, sizeof(WAVEHDR));
	if (AutoClose)
	{
		// ���ݼ��񧹲�
		while (!(WaveOutHeader.dwFlags & WHDR_DONE)) {
			Sleep(100); // �קK CPU �L�רϥ�
		}

		// ����������귽
		ASAudio::StopPlayingWavFile();
	}
	return true;
}

// �����WAV�ɮ�
void ASAudio::StopPlayingWavFile() {
	// ����������귽
	waveOutReset(hWaveOut);
	// �������ĸ˸m
	waveOutClose(hWaveOut);
	// �M���PWaveOutPrepareHeader�ǳƪ�Wave�C
	waveOutUnprepareHeader(hWaveOut, &WaveOutHeader, sizeof(WAVEHDR));
	WaveOutData.clear();
}

// �}�l����
MMRESULT ASAudio::StartRecordingAndDrawSpectrum() {
	// �]�w���T�榡
	WAVEFORMATEX format{};
	format.wFormatTag = WAVE_FORMAT_PCM;
	format.nChannels = 2;
	format.nSamplesPerSec = 44100;
	format.wBitsPerSample = 16;
	format.nBlockAlign = format.nChannels * format.wBitsPerSample / 8;
	format.nAvgBytesPerSec = format.nSamplesPerSec * format.nBlockAlign;
	format.cbSize = 0;

	MMRESULT result;
	// �M�����T�w�İϡA�קK�¸��
	memset(mWAVE_PARM.WAVE_DATA.audioBuffer, 0, BUFFER_SIZE * sizeof(short) * 2);
	// �}�ҭ��T��J�˸m
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

// �������
MMRESULT ASAudio::StopRecording() {
	MMRESULT result = waveInStop(hWaveIn); // Stop recording
	if (result == MMSYSERR_NOERROR) {
		result = waveInReset(hWaveIn);
		result = waveInClose(hWaveIn); // Close the recording device
	}
	waveInPrepareHeader(hWaveIn, &waveHdr, sizeof(WAVEHDR));
	return result;
}

// �����^�I�禡
void CALLBACK ASAudio::WaveInProc(HWAVEIN hwi, UINT uMsg, DWORD_PTR dwInstance, DWORD_PTR dwParam1, DWORD_PTR dwParam2) {

	ASAudio* asAudio = reinterpret_cast<ASAudio*>(dwInstance);
	if (!asAudio) return;
	switch (uMsg) {
		case MM_WIM_DATA: 
		{
			WAVEHDR* waveHdr = reinterpret_cast<WAVEHDR*>(dwParam1);
			if (waveHdr->dwBytesRecorded > 0) {
				// �ϥ� asAudio �Ӧs�������ܼ� bFirstWaveInFlag
				if (asAudio->bFirstWaveInFlag)
				{
					// �ϥ� asAudio �Ӧs���íק令���ܼ� bFirstWaveInFlag
					asAudio->bFirstWaveInFlag = false;

					// ���s����w�İϡA�H���Ĥ@�ӽw�İ�
					waveInUnprepareHeader(hwi, waveHdr, sizeof(WAVEHDR));
					waveInPrepareHeader(hwi, waveHdr, sizeof(WAVEHDR));
					waveInAddBuffer(hwi, waveHdr, sizeof(WAVEHDR));
				}
				else
				{
					// �I�s�����禡�ӳB�z
					asAudio->ProcessAudioBuffer();
				}
			}
			break;
		}
		case MM_WIM_CLOSE:
		{
			// �˸m�w�Q�����Ωް�
			// �]�w���~�X�ШåߧY�q�����ݤ����D�����
			std::lock_guard<std::mutex> lock(asAudio->mWAVE_PARM.recordingMutex);
			asAudio->mWAVE_PARM.deviceError = true;
			asAudio->mWAVE_PARM.isRecordingFinished = true; // �N��аO�� "����" �H�K��� wait
			asAudio->mWAVE_PARM.recordingFinishedCV.notify_one(); // �ߧY���
			break;
		}
	}
}

// �B�z���T�w�İ�
void ASAudio::ProcessAudioBuffer() {
	short* audioSamples = reinterpret_cast<short*>(waveHdr.lpData);
	for (int i = 0; i < BUFFER_SIZE; i++) {
		mWAVE_PARM.WAVE_DATA.audioBuffer[bufferIndex + (size_t)i * 2] = reinterpret_cast<short*>(waveHdr.lpData)[i * 2];
		mWAVE_PARM.WAVE_DATA.audioBuffer[bufferIndex + (size_t)i * 2 + 1] = reinterpret_cast<short*>(waveHdr.lpData)[i * 2 + 1];
	}
	bufferIndex += (size_t)BUFFER_SIZE * 2;

	if (bufferIndex >= (size_t)BUFFER_SIZE * 2) {
		// �w�g���쨬�����˥���(BUFFER_SIZE*2)
		// �� FFT �w�İ�
		for (int i = 0; i < (size_t)BUFFER_SIZE; i++) {
			fftInputBufferLeft[i][0] = mWAVE_PARM.WAVE_DATA.audioBuffer[i * 2];
			fftInputBufferLeft[i][1] = 0.0;
			fftInputBufferRight[i][0] = mWAVE_PARM.WAVE_DATA.audioBuffer[i * 2 + 1];
			fftInputBufferRight[i][1] = 0.0;

			mWAVE_PARM.WAVE_DATA.LeftAudioBuffer[i] = mWAVE_PARM.WAVE_DATA.audioBuffer[i * 2];
			mWAVE_PARM.WAVE_DATA.RightAudioBuffer[i] = mWAVE_PARM.WAVE_DATA.audioBuffer[i * 2 + 1];
		}

		// ����FFT�ഫ
		fftw_execute(fftPlanLeft);
		fftw_execute(fftPlanRight);

		// �q����������
		std::lock_guard<std::mutex> lock(mWAVE_PARM.recordingMutex);
		mWAVE_PARM.isRecordingFinished = true;
		mWAVE_PARM.recordingFinishedCV.notify_one();

		bufferIndex = 0;
	}
}

DWORD ASAudio::SetMicSystemVolume()
{
	//�ŧi�V�����������c
	MMRESULT rc;
	HMIXER hMixer;
	MIXERLINE mxl;
	MIXERLINECONTROLS mxlc;
	MIXERCONTROL mxc;

	//�M���Ҧ��V�����˸m
	for (int deviceID = 0; true; deviceID++)
	{
		// 1. �}�ҫ��w�V����
		rc = mixerOpen(&hMixer, deviceID, 0, 0, MIXER_OBJECTF_MIXER);
		if (rc != MMSYSERR_NOERROR) {
			// �p�G�s�˸m�������}�F�A�N��w�g�䧹�Ҧ��˸m�A�����j��
			break;
		}

		// 2. ���w�n�䪺�O�����u��
		ZeroMemory(&mxl, sizeof(MIXERLINE));
		mxl.cbStruct = sizeof(MIXERLINE);
		mxl.dwComponentType = MIXERLINE_COMPONENTTYPE_DST_WAVEIN;
		rc = mixerGetLineInfo((HMIXEROBJ)hMixer, &mxl, MIXER_GETLINEINFOF_COMPONENTTYPE);
		if (rc != MMSYSERR_NOERROR) {
			// �p�G�o�Ӹ˸m�S�������u���A�N�������A�M���ˬd�U�@��
			mixerClose(hMixer);
			continue;
		}

		// 3. �b�o�ӿ����u�����A�M����J���ӷ�
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
					break; // ���F���J���A���X�o�Ӥp�j��
				}
			}
		}

		// �p�G�b�o�Ӹ˸m�W�S�����J���u���A�N�������A�M���ˬd�U�@��
		if (dwLineID == -1) {
			mixerClose(hMixer);
			continue;
		}

		// 4. �M�䭵�q���
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

		// 5. �p�G���F���q����A�N�]�w���q
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

			// 6. �����A�����˸m�ê������X�̥~�h���j�j��
			mixerClose(hMixer);
			break;
		}
		mixerClose(hMixer);
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
						// �ؼЭ��q
						float targetTemp = 0;
						// �ثe���q
						float currentVolume = 0;
						pAudioEndpointVolume->GetMasterVolumeLevelScalar(&currentVolume);

						if (mWAVE_PARM.bMuteTest)
							targetTemp = 0;
						else
						{
							targetTemp = (float)mWAVE_PARM.WaveOutVolume / 100;
							// ���ˬd�˸m�O�_�䴩�B�i���q
							UINT stepCount;
							UINT currentStep;
							HRESULT hr = pAudioEndpointVolume->GetVolumeStepInfo(&currentStep, &stepCount);

							if (SUCCEEDED(hr) && stepCount > 1) {
								// �p�G�䴩�B�i���q�A�h�ϥ� VolumeStepUp �M VolumeStepDown
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
								// �p�G���䴩�B�i�A�h�ϥ� SetMasterVolumeLevelScalar
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

	// �M���Ҧ��V����
	for (int deviceID = 0; true; deviceID++)
	{
		// �}�ҫ��w�V�����AdeviceID���V����ID
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

			// �]�w�R��
			muteStruct.fValue = mute ? 1 : 0;

			rc = mixerSetControlDetails((HMIXEROBJ)hMixer, &mxcd, MIXER_OBJECTF_HMIXER | MIXER_SETCONTROLDETAILSF_VALUE);
			mixerClose(hMixer);
			if (rc != MMSYSERR_NOERROR)
			{
				// �B�z�]�w���Ѫ����p
				break;
			}
			return; // ���\�]�w�A�N���}
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
	data.phaseR = 0.0; // ��l�ƥk�n�D�ۦ�
	data.frequencyL = mWAVE_PARM.frequencyL; // �]�w���n�D�W�v
	data.frequencyR = mWAVE_PARM.frequencyR; // �]�w�k�n�D�W�v

	// �ǳƭ��T��y�Ѽ�
	PaStreamParameters outputParameters = { 0 };
	outputParameters.device = GetPa_WaveOutDevice(mWAVE_PARM.WaveOutDev); // �]�w��X�˸m
	if (outputParameters.device == paNoDevice) {
		data.errorcode = ERROR_CODE_NOT_FIND_WAVEOUT_DEV;
		return data.errorcode;
	}
	outputParameters.channelCount = 2; // ���n�D��X
	outputParameters.sampleFormat = paFloat32; // 32�줸�B�I�Ʈ榡
	outputParameters.suggestedLatency = Pa_GetDeviceInfo(outputParameters.device)->defaultLowOutputLatency;
	outputParameters.hostApiSpecificStreamInfo = NULL;

	// �}�ҭ��T��y
	err = Pa_OpenStream(&stream,
		NULL, // no input
		&outputParameters,
		SAMPLE_RATE,
		FRAMES_PER_BUFFER,
		paClipOff, // �����ż˥�
		patestCallback, // �]�w�^�I�禡
		&data); // �ǻ��ϥΪ̸��

	if (err != paNoError) {
		data.errorMsg = Pa_GetErrorText(err);
		data.errorcode = ERROR_CODE_OPEN_WAVEOUT_DEV;
		return data.errorcode;
	}

	data.stream = stream; // �O�s��y���Ш쵲�c��

	// �}�l���T��y
	err = Pa_StartStream(stream);
	if (err != paNoError) {
		data.errorcode = ERROR_CODE_START_WAVEOUT_DEV;
		Pa_CloseStream(stream);
		return data.errorcode;
	}
	return err;
}

// ����񪺨禡
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
				return i;
			}
		}
	}
	return -1;
}

int ASAudio::GetWaveInDevice(std::wstring szInDevName)
{
	// ���o���Ŀ�J�˸m���ƶq
	UINT deviceCount = waveInGetNumDevs();
	for (UINT i = 0; i < deviceCount; ++i) {
		WAVEINCAPSW waveInCaps;
		MMRESULT result = waveInGetDevCapsW(i, &waveInCaps, sizeof(WAVEINCAPSW));
		if (result == MMSYSERR_NOERROR) {
			std::wstring deviceName(waveInCaps.szPname);

			if (deviceName.find(szInDevName) != std::string::npos) {
				return i; 
			}
		}
	}
	return -1; // �䤣��˸m�A�^�� -1
}

int ASAudio::GetPa_WaveOutDevice(std::wstring szOutDevName)
{
	int deviceCount = Pa_GetDeviceCount();
	for (int i = 0; i < deviceCount; ++i) {
		const PaDeviceInfo* deviceInfo = Pa_GetDeviceInfo(i);
		if (deviceInfo) {
			if (deviceInfo->maxOutputChannels > 0) { // ��X�˸m
				std::string deviceName(deviceInfo->name); // �N�˸m�W���ର std::string
				// �NszOutDevName�ରstd::string
				int bufferSize = WideCharToMultiByte(CP_UTF8, 0, szOutDevName.c_str(), -1, NULL, 0, NULL, NULL) - 1;//�o�̻ݭn-1�A���M�|�ɭPfind�䤣�쥿�T���r��ƶq
				std::string targetName(bufferSize, ' ');
				WideCharToMultiByte(CP_UTF8, 0, szOutDevName.c_str(), -1, &targetName[0], bufferSize, NULL, NULL);

				if (deviceName.find(targetName) != std::string::npos) {
					return i; // �Y�t�����w����r���˸m�A�^�Ǹ˸m ID
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
			if (deviceInfo->maxInputChannels > 0) { // �u��M����J���˸m
				std::string deviceName(deviceInfo->name); // �N�˸m�W���ର std::string
				// �NszInDevName�ରstd::string
				int bufferSize = WideCharToMultiByte(CP_UTF8, 0, szInDevName.c_str(), -1, NULL, 0, NULL, NULL) - 1;//�o�̻ݭn-1�A���M�|�ɭPfind�䤣�쥿�T���r��ƶq
				std::string targetName(bufferSize, ' ');
				WideCharToMultiByte(CP_UTF8, 0, szInDevName.c_str(), -1, &targetName[0], bufferSize, NULL, NULL);

				if (deviceName.find(targetName) != std::string::npos) {
					return i; // �Y�t�����w����r���˸m�A�^�Ǹ˸m ID
				}
			}
		}
	}
	return -1; // �䤣��˸m�A�^�� -1
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
	(void)inputBuffer; // �קK���ϥΪ��ܼ�ĵ�i

	for (i = 0; i < framesPerBuffer; i++) {
		*out++ = sinf(data->phaseL);
		data->phaseL += 2 * M_PI * data->frequencyL / SAMPLE_RATE;
		*out++ = sinf(data->phaseR);
		data->phaseR += 2 * M_PI * data->frequencyR / SAMPLE_RATE;
	}
	return paContinue; // �~�򼽩�
}

bool ASAudio::StartAudioLoopback(const std::wstring captureKeyword, const std::wstring renderKeyword) {
	HRESULT hr;
	HRESULT hrInit = CoInitialize(nullptr);
	if (FAILED(hrInit)) {
		strMacroResult = "LOG:ERROR_AUDIO_LOOPBACK_COM_INITIALIZE#";
		return false;
	}

	// �إ� Graph (�ϧκ޲z��)
	hr = CoCreateInstance(CLSID_FilterGraph, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&pGraph));
	if (FAILED(hr)) { return false; }
	hr = CoCreateInstance(CLSID_CaptureGraphBuilder2, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&pCaptureGraph));
	if (FAILED(hr)) { CoUninitialize(); return false; }
	pCaptureGraph->SetFiltergraph(pGraph);

	// �إ߸˸m�C�|��
	CComPtr<ICreateDevEnum> pDevEnum;
	hr = CoCreateInstance(CLSID_SystemDeviceEnum, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&pDevEnum));

	// --- �M����w���^���P��V�˸m���z�ﾹ ---
	if (SUCCEEDED(hr)) {
		CComPtr<IEnumMoniker> pEnum;
		if (SUCCEEDED(pDevEnum->CreateClassEnumerator(CLSID_AudioInputDeviceCategory, &pEnum, 0)) && pEnum) {
			// �N pMoniker �ŧi���ܰj�餺
			while (true) {
				CComPtr<IMoniker> pMoniker;
				if (pEnum->Next(1, &pMoniker, nullptr) != S_OK) {
					break; // �䤣���h�˸m�A���X�j��
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
				if (pAudioCapture) break; // ���˸m�A���X�j��
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

	// �N�z�ﾹ�[�J�ϧ�...
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
	// �զ��~�����O�A�Ҧp
	// SoundVolumeView.exe /SetDefault "{deviceId}" 0
	// 0: All, 1: Console, 2: Multimedia
	std::wstring command = L"SoundVolumeView.exe /SetDefault \"" + deviceId + L"\" 0";

	// �ϥ� _wsystem �Ӱ��� Unicode �r��
	int result = _wsystem(command.c_str());

	// �^�ǬO�_���\(0 �N���\)
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
		wprintf(L"��l�ƭ��T�˸m�T�|�����ѡI���~�X: 0x%08X\n", hr);
		strMacroResult = "LOG:ERROR_AUDIO_ENUM_INIT_FAIL#";
		return false;
	}

	hr = enumerator->EnumAudioEndpoints(eRender, DEVICE_STATE_ACTIVE, &collection);
	if (FAILED(hr)) {
		wprintf(L"������T�˸m���ѡI���~�X: 0x%08X\n", hr);
		strMacroResult = "LOG:ERROR_AUDIO_ENUM_GET_FAIL#";
		return false;
	}

	UINT count;
	collection->GetCount(&count);
	wprintf(L"��� %u �ӭ��T�˸m:\n\n", count);

	if (count > 0)
		allContent += L"[�t�Τ����˸m System Audio Device]\n";
	int indexCountdown = config.AudioIndex;
	for (UINT i = 0; i < count; i++)
	{
		IMMDevice* rawDevice = nullptr;
		hr = collection->Item(i, &rawDevice);
		if (SUCCEEDED(hr) && rawDevice != nullptr)
		{
			device.Attach(rawDevice); // ��ʪ��[
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
			wprintf(L"�˸m %d: %s\n", i, friendlyName.pwszVal);
			allContent += friendlyName.pwszVal;
			allContent += L"\n";
			wprintf(L" ID: %s\n\n", deviceId);

			// �ˬd��e�˸m�W�٬O�_�]�t�b�ڭ̪��ؼЦC��
			bool isMatch = false;
			for (const auto& targetName : config.monitorNames) {
				if (wcsstr(friendlyName.pwszVal, targetName.c_str())) {
					isMatch = true;
					break; // ���@�Ӥǰt�N���F�A���X���h�j��
				}
			}

			if (isMatch) {
				// �o�O�@�ӲŦX���󪺸˸m
				if (indexCountdown <= 0)
				{
					// �o�N�O�ڭ̭n�䪺�� N �Ӹ˸m
					outDeviceId = deviceId;
					CoTaskMemFree(deviceId);
					PropVariantClear(&friendlyName);
					CoUninitialize();
					allContent = L""; // �M�Ť�x
					return true;
				}
				else
				{
					// �٨S����A�p�ƾ���@
					indexCountdown--;
				}
			}
		}
		CoTaskMemFree(deviceId);
		PropVariantClear(&friendlyName);
	}
	CoUninitialize();

	// �p�G�j��]���٨S���
	allContent += L"\n\n[�n�䪺�˸m Target Audio Devices]\n";
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
	// �]�w���T�榡
	WAVEFORMATEX format{};
	format.wFormatTag = WAVE_FORMAT_PCM;
	format.nChannels = 2;
	format.nSamplesPerSec = 44100;
	format.wBitsPerSample = 16;
	format.nBlockAlign = format.nChannels * format.wBitsPerSample / 8;
	format.nAvgBytesPerSec = format.nSamplesPerSec * format.nBlockAlign;
	format.cbSize = 0;

	MMRESULT result;
	// �M�����T�w�İ�
	memset(mWAVE_PARM.WAVE_DATA.audioBuffer, 0, BUFFER_SIZE * sizeof(short) * 2);

	// �}�ҭ��T��J�˸m
	int iDevIndex = GetWaveInDevice(mWAVE_PARM.WaveInDev);
	if (iDevIndex == -1)
	{
		data.errorcode = ERROR_CODE_NOT_FIND_WAVEIN_DEV;
		return ERROR_CODE_NOT_FIND_WAVEIN_DEV;
	}

	// �����}�Ҩö}�l�����A�������I�s����禡
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
