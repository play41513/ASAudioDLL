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
WAVEHDR waveHdr; // WAVEHDR���c�A�Ω�����w��
HWAVEIN hWaveIn; // ���T��J�]�ƪ��B�z
HWAVEOUT hWaveOut; // ���T��X�]�ƪ��B�z
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
		shouldRetry = false; // �w�]������
		strMacroResult.clear(); // �C�����իe�M�����G
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
				hasError = true; // PlayWavFile �����|�]�w strMacroResult
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
					hasError = true; // StartAudioLoopback �����|�]�w strMacroResult
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
			else { // FindDeviceIdByName �����|�]�w strMacroResult
				hasError = true;
			}
		}

		// �p�G��������ե��ѡA�B�O�b���ʼҦ��U�A�h���X����
		if (hasError) {
			if (ownerHwnd != NULL) { // �p�G ownerHwnd ���� NULL�A�~��ܿ��~�T��
				std::string strTemp = "Failed: " + strMacroResult;
				std::wstring wErrorMsg = ASAudio::GetInstance().charToWstring(strTemp.c_str()) + L"\n\n" + allContent + L"\n\n�O�_�n�����H(Retry ?)";
				if (MessageBoxW(ownerHwnd, wErrorMsg.c_str(), L"���ե���", MB_YESNO | MB_ICONERROR) == IDYES) {
					shouldRetry = true; // �]�w�X�ХH���s����
				}
			}
		}
	} while (shouldRetry);

	// �p�G strMacroResult ���� (�Ҧp�Ҧ��\�ೣ disable)�A���@�ӹw�]���\�T��
	if (strMacroResult.empty() || !hasError) {
		strMacroResult = "LOG:PASS#";
	}

	// �N string �ର BSTR �^��
	int len = MultiByteToWideChar(CP_UTF8, 0, strMacroResult.c_str(), -1, nullptr, 0);
	if (len == 0) {
		return SysAllocString(L"LOG:ERROR_AUDIO_OTHER#"); // �קK�^�� nullptr
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
{ // �t�ΰѼƳ]�m�ASPIF_SENDCHANGE ��ܥߧY���ASPIF_UPDATEINIFILE ��ܧ�s�t�m���A�ĤG�ӰѼ�1�B0����windosw�w�]�B�L����
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
	THD+N��dB�Ƚd��
	�M�~���T�]�ơG
	���~�誺�M�~���T�]�Ƴq�`�n�DTHD+N�C�� -80 dB�A�o�N���ۥ��u�M���n�D�`�C�C
	���O�q�l���~�G
	�a�έ��T�M���O�q�l���~��THD+N�@��b -60 dB �� -80 dB �����C
	�����~�譵�T�]�ơG
	THD+N�Ȧb -60 dB �� -40 dB �����Q�{���O�����~��C
	�C�~�譵�T�]�ơG
	THD+N�Ȱ��� -40 dB �i���ܥ��u�M���n�����A������t�C
	�̤jdB�Ƚd��
	�M�~���T�]�ơG
	�M�~�]�ƪ��̤jdB�ȡ]��ܫH���j�ס^�q�`�b 80 dB �� 120 dB �����C
	���O�q�l���~�G
	�a�έ��T�M���O�q�l���~���̤jdB�ȳq�`�b 60 dB �� 100 dB �����C
	�����~�譵�T�]�ơG
	�̤jdB�Ȧb 70 dB �� 90 dB �����C
	�C�~�譵�T�]�ơG
	�̤jdB�ȧC�� 70 dB�C
	*/
}

void __cdecl fft_get_mute_db(double* dB_ValueMax)
{
	dB_ValueMax[0] = 10 * log10(mWAVE_PARM.leftMuteWAVE_ANALYSIS.TotalEnergy);
	dB_ValueMax[1] = 10 * log10(mWAVE_PARM.rightMuteWAVE_ANALYSIS.TotalEnergy);
	/*
	�R�����ժ��̤jdB�Ƚd��
	�M�~���T�]�ơG
	���~�誺�M�~���T�]�Ʀb�R�����A�U�A�̤jdB�ȡ]�I�����n�q���^�q�`�C�� -90 dB�C
	���O�q�l���~�G
	�a�έ��T�M���O�q�l���~���R���̤jdB�Ȥ@��b -70 dB �� -90 dB �����C
	�����~�譵�T�]�ơG
	�R���̤jdB�Ȧb -60 dB �� -80 dB �����C
	�C�~�譵�T�]�ơG
	�R���̤jdB�Ȱ��� -60 dB ��ܦb�R�����A�U���n���������C
	*/
}

void __cdecl fft_get_snr(double* snr)
{
	// �p��SNR
	snr[0] = (mWAVE_PARM.leftMuteWAVE_ANALYSIS.TotalEnergy > 0) ? (mWAVE_PARM.leftWAVE_ANALYSIS.fundamentalEnergy / mWAVE_PARM.leftMuteWAVE_ANALYSIS.TotalEnergy) : 0;
	snr[1] = (mWAVE_PARM.rightMuteWAVE_ANALYSIS.TotalEnergy > 0) ? (mWAVE_PARM.rightWAVE_ANALYSIS.fundamentalEnergy / mWAVE_PARM.rightMuteWAVE_ANALYSIS.TotalEnergy) : 0;

	// �ഫ�������A�o�̪�SNR�O��q�Υ\�v����ȡA�ҥH��10 * log10
	double leftSNR_dB = (snr[0] > 0) ? (10 * log10(snr[0])) : -INFINITY;
	double rightSNR_dB = (snr[1] > 0) ? (10 * log10(snr[1])) : -INFINITY;

	// ��s SNR ��
	snr[0] = leftSNR_dB;
	snr[1] = rightSNR_dB;
	/*
	SNR ���ժ��d��
	�M�~���T�]�ơG
	���~�誺�M�~���T�]�Ƴq�`�㦳�D�`���� SNR�A�d��b 90 dB �� 120 dB �����A���ǰ��ݳ]�ƬƦܥi�H�F�� 120 dB �H�W�C
	���O�q�l���~�G
	�a�έ��T�M���O�q�l���~�� SNR �d��q�`�b 70 dB �� 100 dB �����C
	�����~�譵�T�]�ơG
	�����~�譵�T�]�ƪ� SNR �q�`�b 60 dB �� 80 dB �����C
	�C�~�譵�T�]�ơG
	�C�~�譵�T�]�ƪ� SNR �i��C�� 60 dB�A��ܫH���P���n�������t�Z���p�A����i����t�C
	SNR ���ժ�����
	�� SNR�G
	�� SNR �Ȫ�ܭI�����n�C�A���W�H���M���B�j�j�C�o�b�M�~���T�t�Τ��׬����n�A�]�����̭n�D���ײM��������M�C���n�C
	�C SNR�G
	�C SNR �Ȫ�ܭI�����n�۹�����A�H�������M���C�b�@�Ǿ��n�e�ԫ׸��������Τ��A�o�i�ण�|�y����ۼv�T�A���b���n�D�����W���Τ��A�C SNR �i��|��ۭ��C����C
	*/
}

// �~��C��ơA�Ω����FFT�ഫ
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
	// ���ݪ����������
	//while²�g ��isRecordingFinished == true ���X
	mWAVE_PARM.recordingFinishedCV.wait(lock, [] { return mWAVE_PARM.isRecordingFinished;  });
	mWAVE_PARM.isRecordingFinished = false;

	ASAudio::GetInstance().StopRecording(); // �������
	ASAudio::GetInstance().stopPlayback(); // �������

	// �T�O FFTW ���G���гQ���t
	fftw_complex* leftSpectrum = (fftw_complex*)fftw_malloc(sizeof(fftw_complex) * BUFFER_SIZE);
	fftw_complex* rightSpectrum = (fftw_complex*)fftw_malloc(sizeof(fftw_complex) * BUFFER_SIZE);

	ASAudio::GetInstance().GetSpectrumData(leftSpectrum, rightSpectrum);// ����W�мƾ�

	memcpy(leftAudioData, mWAVE_PARM.WAVE_DATA.LeftAudioBuffer, BUFFER_SIZE * sizeof(short));
	memcpy(rightAudioData, mWAVE_PARM.WAVE_DATA.RightAudioBuffer, BUFFER_SIZE * sizeof(short));

	double maxAmplitude = 0;
	for (int i = 0; i < BUFFER_SIZE; ++i) {
		double sample = static_cast<double>(leftAudioData[i]);
		if (fabs(sample) > maxAmplitude) {
			maxAmplitude = fabs(sample);
		}
	}
	maxAmplitude = max(maxAmplitude, 1.0); // �קK log10(0) �����p
	double maxAmplitude_dB = 20 * log10(maxAmplitude / 32767);

	for (int i = 0; i < BUFFER_SIZE; ++i) {
		leftSpectrumData[i] = sqrt(leftSpectrum[i][0] * leftSpectrum[i][0] + leftSpectrum[i][1] * leftSpectrum[i][1]) / BUFFER_SIZE;
		rightSpectrumData[i] = sqrt(rightSpectrum[i][0] * rightSpectrum[i][0] + rightSpectrum[i][1] * rightSpectrum[i][1]) / BUFFER_SIZE;
	}

	// ���R�W��THD+N
	ASAudio::GetInstance().SpectrumAnalysis(leftSpectrumData, rightSpectrumData);
	return result;
}

// �~��C��ơA�Ω����FFT�ഫ
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
	// ���ݪ����������
	//while²�g ��isRecordingFinished == true ���X
	mWAVE_PARM.recordingFinishedCV.wait(lock, [] { return mWAVE_PARM.isRecordingFinished;  });
	mWAVE_PARM.isRecordingFinished = false;

	ASAudio::GetInstance().StopRecording();// �������
	ASAudio::GetInstance().stopPlayback(); // �������

	fftw_complex* leftSpectrum;
	fftw_complex* rightSpectrum;
	ASAudio::GetInstance().GetSpectrumData(leftSpectrum, rightSpectrum); // ����W�мƾ�
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
	//���R�W��
	ASAudio::GetInstance().SpectrumAnalysis(leftSpectrumData, rightSpectrumData);
	//�_�쭵�q
	mWAVE_PARM.bMuteTest = false;
	if (MuteWaveIn)
		ASAudio::GetInstance().SetMicMute(false);
	if (MuteWaveOut)
		ASAudio::GetInstance().SetSpeakerSystemVolume();

	return result;
}

void ASAudio::SpectrumAnalysis(double* leftSpectrumData, double* rightSpectrumData)
{
	const double Fs = 44100; // �����W�v
	const int freqBins = BUFFER_SIZE / 2 + 1; // �W�vbin�ƶq
	double leftMagnitude, rightMagnitude; // ��e�W�vbin���T��

	auto& leftTempWAVE_ANALYSIS = mWAVE_PARM.bMuteTest ? mWAVE_PARM.leftMuteWAVE_ANALYSIS : mWAVE_PARM.leftWAVE_ANALYSIS;
	auto& RightTempWAVE_ANALYSIS = mWAVE_PARM.bMuteTest ? mWAVE_PARM.rightMuteWAVE_ANALYSIS : mWAVE_PARM.rightWAVE_ANALYSIS;
	leftTempWAVE_ANALYSIS.TotalEnergyPoint = 0; // ���n�D�̤j��q�I��l��
	RightTempWAVE_ANALYSIS.TotalEnergyPoint = 0; // �k�n�D�̤j��q�I��l��
	leftTempWAVE_ANALYSIS.TotalEnergy = 0; // ���n�D�`��q��l��
	RightTempWAVE_ANALYSIS.TotalEnergy = 0; // �k�n�D�`��q��l��

	double maxLeftMagnitude = 0;
	double maxRightMagnitude = 0;

	// �p���`��q�M�̤j��q�I
	for (int i = 2; i < freqBins; ++i) {//�W�v0�������p
		leftMagnitude = leftSpectrumData[i]; // ��e���n�D�W�vbin���T��
		rightMagnitude = rightSpectrumData[i]; // ��e�k�n�D�W�vbin���T��

		// �֥[�Ҧ��W�v��������q�]�T�ת�����^
		leftTempWAVE_ANALYSIS.TotalEnergy += leftMagnitude * leftMagnitude;
		RightTempWAVE_ANALYSIS.TotalEnergy += rightMagnitude * rightMagnitude;

		// ��s�̤j��q�I�M�̤j��
		if (leftMagnitude > maxLeftMagnitude) {
			maxLeftMagnitude = leftMagnitude;
			leftTempWAVE_ANALYSIS.TotalEnergyPoint = i;
		}
		if (rightMagnitude > maxRightMagnitude) {
			maxRightMagnitude = rightMagnitude;
			RightTempWAVE_ANALYSIS.TotalEnergyPoint = i;
		}
	}

	double Harmonic1, Harmonic2; // �Ӫi�T��
	double totalHarmonicEnergy; // �Ӫi�`��q

	// �p��THD+N ���n�D
	int energyPoint = leftTempWAVE_ANALYSIS.TotalEnergyPoint;
	if (energyPoint * 2 < freqBins) {
		Harmonic1 = leftSpectrumData[energyPoint * 2];
		Harmonic2 = (energyPoint * 3 < freqBins) ? leftSpectrumData[energyPoint * 3] : 0;

		totalHarmonicEnergy = Harmonic1 * Harmonic1 + Harmonic2 * Harmonic2;
		// ���o��i��q
		double fundamentalEnergy = leftSpectrumData[energyPoint] * leftSpectrumData[energyPoint];
		leftTempWAVE_ANALYSIS.fundamentalEnergy = fundamentalEnergy;
		// �p��THD+N
		leftTempWAVE_ANALYSIS.thd_N = sqrt(totalHarmonicEnergy / fundamentalEnergy);

		// �ഫ������
		leftTempWAVE_ANALYSIS.thd_N_dB = (leftTempWAVE_ANALYSIS.thd_N > 0)
			? 20 * log10(leftTempWAVE_ANALYSIS.thd_N)
			: -INFINITY; // �B�z0�έt�ȱ��p
	}
	else {
		leftTempWAVE_ANALYSIS.thd_N_dB = 0;
	}

	// �p��THD+N �k�n�D
	energyPoint = RightTempWAVE_ANALYSIS.TotalEnergyPoint;
	if (energyPoint * 2 < freqBins) {
		Harmonic1 = rightSpectrumData[energyPoint * 2];
		Harmonic2 = (energyPoint * 3 < freqBins) ? rightSpectrumData[energyPoint * 3] : 0;

		totalHarmonicEnergy = Harmonic1 * Harmonic1 + Harmonic2 * Harmonic2;
		// ���o��i��q
		double fundamentalEnergy = rightSpectrumData[energyPoint] * rightSpectrumData[energyPoint];
		RightTempWAVE_ANALYSIS.fundamentalEnergy = fundamentalEnergy;
		// �p��THD+N
		RightTempWAVE_ANALYSIS.thd_N = sqrt(totalHarmonicEnergy / fundamentalEnergy);

		// �ഫ������
		RightTempWAVE_ANALYSIS.thd_N_dB = (RightTempWAVE_ANALYSIS.thd_N > 0)
			? 20 * log10(RightTempWAVE_ANALYSIS.thd_N)
			: -INFINITY; // �B�z0�έt�ȱ��p
	}
	else {
		RightTempWAVE_ANALYSIS.thd_N_dB = 0;
	}

	double totalEnergy = 0;
	for (int i = 0; i < freqBins; ++i) {
		double magnitude = leftSpectrumData[i];
		totalEnergy += magnitude * magnitude; // ��q�O�T�ת�����
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

// ASAudio�����c�y���
ASAudio::ASAudio(SpectrumAnalysisCallback callback) : spectrumCallback(callback)
{
	// ��l�ƭ��T��J�M��X�����ܼ�
	hWaveIn = NULL; // ���T��J�]�ƥy�`
	hWaveOut = NULL; // ���T��X�]�ƥy�`
	waveInCallback = nullptr; // ���T��J�^�I�禡
	memset(&waveHdr, 0, sizeof(WAVEHDR)); // ��l�ƭ��T��J�^�I�Ϊ�WAVEHDR���c
	memset(mWAVE_PARM.WAVE_DATA.audioBuffer, 0, BUFFER_SIZE * 2 * sizeof(short)); // ��l�ƭ��T�w�ġA���n�D
	bufferIndex = 0; // ���T�w�į���

	// ��l�� FFT �һݪ��ܼƩM�p��
	fftInputBufferLeft = (fftw_complex*)fftw_malloc(sizeof(fftw_complex) * BUFFER_SIZE);
	fftInputBufferRight = (fftw_complex*)fftw_malloc(sizeof(fftw_complex) * BUFFER_SIZE);
	fftOutputBufferLeft = (fftw_complex*)fftw_malloc(sizeof(fftw_complex) * BUFFER_SIZE);
	fftOutputBufferRight = (fftw_complex*)fftw_malloc(sizeof(fftw_complex) * BUFFER_SIZE);

	// �Ы� FFT �p���A�Ω���J�i��ť߸��ܴ�
	fftPlanLeft = fftw_plan_dft_1d(BUFFER_SIZE, fftInputBufferLeft, fftOutputBufferLeft, FFTW_FORWARD, FFTW_ESTIMATE);
	fftPlanRight = fftw_plan_dft_1d(BUFFER_SIZE, fftInputBufferRight, fftOutputBufferRight, FFTW_FORWARD, FFTW_ESTIMATE);
}

// ASAudio�����R�c���
ASAudio::~ASAudio() {
	fftw_destroy_plan(fftPlanLeft);
	fftw_destroy_plan(fftPlanRight);
	fftw_free(fftInputBufferLeft);
	fftw_free(fftInputBufferRight);
	fftw_free(fftOutputBufferLeft);
	fftw_free(fftOutputBufferRight);
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

	// ���}WAVE���A��^�@��HMMIO�y�`
	HMMIO handle = mmioOpen((LPWSTR)mWAVE_PARM.AudioFile.c_str(), 0, MMIO_READ);

	// �i�JRIFF�϶�(RIFF Chunk)
	mmioDescend(handle, &ChunkInfo, 0, MMIO_FINDRIFF);

	// �i�Jfmt�϶�(RIFF�l���A�]�t���T���c���H��)
	FormatChunkInfo.ckid = mmioStringToFOURCCA("fmt", 0); // �M��fmt�l��
	mmioDescend(handle, &FormatChunkInfo, &ChunkInfo, MMIO_FINDCHUNK);

	// Ū��wav���c�H��
	memset(&wfx, 0, sizeof(WAVEFORMATEX));
	mmioRead(handle, (char*)&wfx, FormatChunkInfo.cksize);

	// �ˬd�n�D�ƶq
	if (wfx.nChannels != 2) {
		// ���~�G���O���n�D
		mmioClose(handle, 0);
		strMacroResult = "LOG:ERROR_AUDIO_WAVE_FILE_FORMAT_NOT_SUPPORT#";
		return false;
	}

	// ���Xfmt�϶�
	mmioAscend(handle, &FormatChunkInfo, 0);

	// �i�Jdata�϶�(�]�t�Ҧ����ƾڪi��)
	DataChunkInfo.ckid = mmioStringToFOURCCA("data", 0);
	mmioDescend(handle, &DataChunkInfo, &ChunkInfo, MMIO_FINDCHUNK);

	// ��odata�϶����j�p
	DataSize = DataChunkInfo.cksize;

	// ���t�w�İ�
	WaveOutData.resize(DataSize);
	mmioRead(handle, WaveOutData.data(), DataSize);

	// ���}��X�]��
	int iDevIndex = GetWaveOutDevice(mWAVE_PARM.WaveOutDev);
	MMRESULT result = waveOutOpen(&hWaveOut, iDevIndex, &wfx, 0, 0, WAVE_FORMAT_QUERY);
	if (result != MMSYSERR_NOERROR) {
		// ���~
		strMacroResult = "LOG:ERROR_AUDIO_WAVEOUT_DEVICE_NOT_FIND_2#";
		return false;
	}

	result = waveOutOpen(&hWaveOut, iDevIndex, &wfx, NULL, NULL, CALLBACK_NULL);
	if (result != MMSYSERR_NOERROR) {
		// ���~
		waveOutClose(hWaveOut);
		strMacroResult = "LOG:ERROR_AUDIO_WAVEOUT_DEVICE_NOT_FIND_3#";
		return false;
	}

	// �]�mwave header.
	memset(&WaveOutHeader, 0, sizeof(WaveOutHeader));
	WaveOutHeader.lpData = WaveOutData.data();
	WaveOutHeader.dwBufferLength = DataSize;

	// �ǳ�wave header.
	waveOutPrepareHeader(hWaveOut, &WaveOutHeader, sizeof(WAVEHDR));

	// �N�w�İϸ�Ƽg�J����]��(�}�l����).
	waveOutWrite(hWaveOut, &WaveOutHeader, sizeof(WAVEHDR));
	if (AutoClose)
	{
		// ���ݼ��񧹦�
		while (!(WaveOutHeader.dwFlags & WHDR_DONE)) {
			Sleep(100); // �קK CPU �L�צ���
		}

		// �����í��m�޲z��
		ASAudio::StopPlayingWavFile();
	}
	return true;
}

// �����WAV�ɮ�
void ASAudio::StopPlayingWavFile() {
	// �����í��m�޲z��
	waveOutReset(hWaveOut);
	// ��������]��
	waveOutClose(hWaveOut);
	// �M�z��WaveOutPrepareHeader�ǳƪ�Wave�C
	waveOutUnprepareHeader(hWaveOut, &WaveOutHeader, sizeof(WAVEHDR));
	WaveOutData.clear();
}

// �}�l����
MMRESULT ASAudio::StartRecordingAndDrawSpectrum() {
	// �]�m���T�榡
	WAVEFORMATEX format{};
	format.wFormatTag = WAVE_FORMAT_PCM;
	format.nChannels = 2;
	format.nSamplesPerSec = 44100;
	format.wBitsPerSample = 16;
	format.nBlockAlign = format.nChannels * format.wBitsPerSample / 8;
	format.nAvgBytesPerSec = format.nSamplesPerSec * format.nBlockAlign;
	format.cbSize = 0;

	MMRESULT result;
	// �M�z���T�w�İϡA�קK��d�ƾ�
	memset(mWAVE_PARM.WAVE_DATA.audioBuffer, 0, BUFFER_SIZE * sizeof(short) * 2);
	// ���}���T��J�]��
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
	VirtualFree(&waveHdr.lpData, 0, MEM_RELEASE);
	return result;
}

// �����^�ը��
void CALLBACK ASAudio::WaveInProc(HWAVEIN hwi, UINT uMsg, DWORD_PTR dwInstance, DWORD_PTR dwParam1, DWORD_PTR dwParam2) {
	if (uMsg == MM_WIM_DATA) {
		ASAudio* asAudio = reinterpret_cast<ASAudio*>(dwInstance);
		WAVEHDR* waveHdr = reinterpret_cast<WAVEHDR*>(dwParam1);
		// �B�z�w�İϤ������T
		if (waveHdr->dwBytesRecorded > 0) {
			if (bFirstWaveInFlag)
			{
				bFirstWaveInFlag = false;
				// ���D���s�������b�A�˱�Ĥ@����ơA���s�ǳƨòK�[�s���w�İ�
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

// �B�z���T�w�İ�
void ASAudio::ProcessAudioBuffer() {
	short* audioSamples = reinterpret_cast<short*>(waveHdr.lpData);
	for (int i = 0; i < BUFFER_SIZE; i++) {
		mWAVE_PARM.WAVE_DATA.audioBuffer[bufferIndex + i * 2] = reinterpret_cast<short*>(waveHdr.lpData)[i * 2];
		mWAVE_PARM.WAVE_DATA.audioBuffer[bufferIndex + i * 2 + 1] = reinterpret_cast<short*>(waveHdr.lpData)[i * 2 + 1];
	}
	bufferIndex += BUFFER_SIZE * 2;

	if (bufferIndex >= BUFFER_SIZE * 2) {
		// �����s�����n�D��T��(BUFFER_SIZE*2)
		// ��R FFT �w�İ�
		for (int i = 0; i < BUFFER_SIZE; i++) {
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

		// �q���w��������
		std::lock_guard<std::mutex> lock(mWAVE_PARM.recordingMutex);
		mWAVE_PARM.isRecordingFinished = true;
		mWAVE_PARM.recordingFinishedCV.notify_one();

		bufferIndex = 0;
	}
}

DWORD ASAudio::SetMicSystemVolume()
{
	//��l�Ƭ������c
	MMRESULT rc;
	HMIXER hMixer;
	MIXERLINE mxl;
	MIXERLINECONTROLS mxlc;
	MIXERCONTROL mxc;

	//�M���t�Ϊ��V�����A��������J�����V�����A�����ӳ]��ID
	for (int deviceID = 0; true; deviceID++)
	{
		//���}��e���V�����AdeviceID���V������ID
		rc = mixerOpen(&hMixer, deviceID, 0, 0, MIXER_OBJECTF_MIXER);
		if (rc != MMSYSERR_NOERROR)
			break;
		ZeroMemory(&mxl, sizeof(MIXERLINE));
		mxl.cbStruct = sizeof(MIXERLINE);
		//���X�ݭn������q�D
		//�n�D�����W��X��MIXERLINE_COMPONENTTYPE_DST_SPEAKERS
		//�n�D�����W��J��MIXERLINE_COMPONENTTYPE_DST_WAVEIN
		mxl.dwComponentType = MIXERLINE_COMPONENTTYPE_DST_WAVEIN;
		rc = mixerGetLineInfo((HMIXEROBJ)hMixer, &mxl, MIXER_GETLINEINFOF_COMPONENTTYPE);
		if (rc != MMSYSERR_NOERROR)
			continue;
		//���o�V�����]�ƪ����w�u���H�����\���ܡA�h�N�s���ƫO�s
		DWORD dwConnections = mxl.cConnections;
		DWORD dwLineID = -1;
		for (DWORD i = 0; i < dwConnections; i++)
		{
			mxl.dwSource = i;
			//�ھ�SourceID��o�s�����H��
			rc = mixerGetLineInfo((HMIXEROBJ)hMixer, &mxl, MIXER_OBJECTF_HMIXER | MIXER_GETLINEINFOF_SOURCE);
			if (rc == MMSYSERR_NOERROR)
			{
				//�p�G��e�]�����������J���A���X�`��
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

		//��������������q
		mxlc.cbStruct = sizeof(mxlc);
		mxlc.dwLineID = dwLineID;
		mxlc.dwControlType = MIXERCONTROL_CONTROLTYPE_VOLUME;
		mxlc.cControls = 1;
		mxlc.pamxctrl = &mxc;
		mxlc.cbmxctrl = sizeof(mxc);

		//���o����H��
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

			//��o���q�ȡA���o���H����bmxcd��
			rc = mixerGetControlDetails((HMIXEROBJ)hMixer, &mxcd, MIXER_GETCONTROLDETAILSF_VALUE);

			//��l�ƿ����j�p���H��
			MIXERCONTROLDETAILS_UNSIGNED mxcdVolume_Set = { mxc.Bounds.dwMaximum * mWAVE_PARM.WaveInVolume / 100 };
			MIXERCONTROLDETAILS mxcd_Set = { 0 };
			mxcd_Set.cbStruct = sizeof(MIXERCONTROLDETAILS);
			mxcd_Set.dwControlID = mxc.dwControlID;
			mxcd_Set.cChannels = 1;
			mxcd_Set.cMultipleItems = 0;
			mxcd_Set.cbDetails = sizeof(MIXERCONTROLDETAILS_UNSIGNED);
			mxcd_Set.paDetails = &mxcdVolume_Set;

			//�]�m�����j�p
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
						// �ؼЭ��q
						float targetTemp = 0;
						// ��e���q
						float currentVolume = 0;
						pAudioEndpointVolume->GetMasterVolumeLevelScalar(&currentVolume);

						if (mWAVE_PARM.bMuteTest)
							targetTemp = 0;
						else
						{
							targetTemp = (float)mWAVE_PARM.WaveOutVolume / 100;
							// ���ˬd�]�ƬO�_������q�B�i
							UINT stepCount;
							UINT currentStep;
							HRESULT hr = pAudioEndpointVolume->GetVolumeStepInfo(&currentStep, &stepCount);

							if (SUCCEEDED(hr) && stepCount > 1) {
								// �p�G����B�i���q����A�h�ϥ� VolumeStepUp �� VolumeStepDown
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
								// �p�G������B�i�A�h�ϥ� SetMasterVolumeLevelScalar
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

	// �M���t�Ϊ��V����
	for (int deviceID = 0; true; deviceID++)
	{
		// ���}��e���V�����AdeviceID���V������ID
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

			// �]�m�R��
			muteStruct.fValue = mute ? 1 : 0;

			rc = mixerSetControlDetails((HMIXEROBJ)hMixer, &mxcd, MIXER_OBJECTF_HMIXER | MIXER_SETCONTROLDETAILSF_VALUE);
			mixerClose(hMixer);
			if (rc != MMSYSERR_NOERROR)
			{
				// �B�z�]�m���Ѫ����p
				break;
			}
			return; // ���\�]�m��A�h�X���
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

	// ��l�� PortAudio
	err = Pa_Initialize();
	if (err != paNoError) {
		data.errorcode = ERROR_CODE_OPEN_INIT_WAVEOUT_DEV;
		return data.errorcode;
	}

	data = { 0.0, 0.0, 0.0, 0.0, nullptr };
	data.phaseL = 0.0;
	data.phaseR = 0.0; // ��l�ƥk�n�D�ۦ�
	data.frequencyL = mWAVE_PARM.frequencyL; // �]�m���n�D�W�v
	data.frequencyR = mWAVE_PARM.frequencyR; // �]�m�k�n�D�W�v

	// �t�m��X�y�Ѽ�
	PaStreamParameters outputParameters = { 0 };
	outputParameters.device = GetPa_WaveOutDevice(mWAVE_PARM.WaveOutDev); // �]�w��X�˸m
	if (outputParameters.device == paNoDevice) {
		data.errorcode = ERROR_CODE_NOT_FIND_WAVEOUT_DEV;
		Pa_Terminate();
		return data.errorcode;
	}
	outputParameters.channelCount = 2; // ���n�D��X
	outputParameters.sampleFormat = paFloat32; // 32��B�I�Ʈ榡
	outputParameters.suggestedLatency = Pa_GetDeviceInfo(outputParameters.device)->defaultLowOutputLatency;
	outputParameters.hostApiSpecificStreamInfo = NULL;

	// ���}���W�y
	err = Pa_OpenStream(&stream,
		NULL, // no input
		&outputParameters,
		SAMPLE_RATE,
		FRAMES_PER_BUFFER,
		paClipOff, // ���i�����
		patestCallback, // �]�w�^�ը��
		&data); // �ǻ��Τ���

	if (err != paNoError) {
		data.errorMsg = Pa_GetErrorText(err);
		data.errorcode = ERROR_CODE_OPEN_WAVEOUT_DEV;
		Pa_Terminate();
		return data.errorcode;
	}

	data.stream = stream; // �O�s�y���w���Ƶ��c��

	// �}�l���W�y
	err = Pa_StartStream(stream);
	if (err != paNoError) {
		data.errorcode = ERROR_CODE_START_WAVEOUT_DEV;
		Pa_CloseStream(stream);
		Pa_Terminate();
		return data.errorcode;
	}
	return err;
}

// ����񪺨禡
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
	// ������W��J�]�ƪ��ƶq
	UINT deviceCount = waveOutGetNumDevs();
	for (UINT i = 0; i < deviceCount; ++i) {
		WAVEOUTCAPSA waveOutCaps;
		MMRESULT result = waveOutGetDevCapsA(i, &waveOutCaps, sizeof(WAVEOUTCAPSA));
		if (result == MMSYSERR_NOERROR) {
			std::string deviceName(waveOutCaps.szPname); // �N�]�ƦW���ഫ�� std::string
			// �NszOutDevName�ഫ��std::string
			int bufferSize = WideCharToMultiByte(CP_UTF8, 0, szOutDevName.c_str(), -1, NULL, 0, NULL, NULL) - 1;//�o��ݭn-1�A���᭱��find��勵�T���r���ƶq
			std::string targetName(bufferSize, '\0');
			WideCharToMultiByte(CP_UTF8, 0, szOutDevName.c_str(), -1, &targetName[0], bufferSize, NULL, NULL);

			if (deviceName.find(targetName) != std::string::npos) {
				return i; // ���]�t���w�l�r�ꪺ�]�ơA��^�]�� ID
			}
		}
	}
	return -1; // �����]�ơA��^ -1
}

int ASAudio::GetWaveInDevice(std::wstring szInDevName)
{
	// ������W��J�]�ƪ��ƶq
	UINT deviceCount = waveInGetNumDevs();
	for (UINT i = 0; i < deviceCount; ++i) {
		WAVEINCAPSA waveInCaps;
		MMRESULT result = waveInGetDevCapsA(i, &waveInCaps, sizeof(WAVEINCAPSA));
		if (result == MMSYSERR_NOERROR) {
			std::string deviceName(waveInCaps.szPname); // �N�]�ƦW���ഫ�� std::string
			// �NszInDevName�ഫ��std::string
			int bufferSize = WideCharToMultiByte(CP_UTF8, 0, szInDevName.c_str(), -1, NULL, 0, NULL, NULL) - 1;//�o��ݭn-1�A���᭱��find��勵�T���r���ƶq
			std::string targetName(bufferSize, '\0');
			WideCharToMultiByte(CP_UTF8, 0, szInDevName.c_str(), -1, &targetName[0], bufferSize, NULL, NULL);

			if (deviceName.find(targetName) != std::string::npos) {
				return i; // ���]�t���w�l�r�ꪺ�]�ơA��^�]�� ID
			}
		}
	}
	return -1; // �����]�ơA��^ -1
}

int ASAudio::GetPa_WaveOutDevice(std::wstring szOutDevName)
{
	PaError err = Pa_Initialize();
	int deviceCount = Pa_GetDeviceCount();
	for (int i = 0; i < deviceCount; ++i) {
		const PaDeviceInfo* deviceInfo = Pa_GetDeviceInfo(i);
		if (deviceInfo) {
			if (deviceInfo->maxOutputChannels > 0) { // ��X�˸m
				std::string deviceName(deviceInfo->name); // �N�]�ƦW���ഫ�� std::string
				// �NszOutDevName�ഫ��std::string
				int bufferSize = WideCharToMultiByte(CP_UTF8, 0, szOutDevName.c_str(), -1, NULL, 0, NULL, NULL) - 1;//�o��ݭn-1�A���᭱��find��勵�T���r���ƶq
				std::string targetName(bufferSize, '\0');
				WideCharToMultiByte(CP_UTF8, 0, szOutDevName.c_str(), -1, &targetName[0], bufferSize, NULL, NULL);

				if (deviceName.find(targetName) != std::string::npos) {
					Pa_Terminate();
					return i; // ���]�t���w�l�r�ꪺ�]�ơA��^�]�� ID
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
			if (deviceInfo->maxInputChannels > 0) { // �ȦҼ{��J�]��
				std::string deviceName(deviceInfo->name); // �N�]�ƦW���ഫ�� std::string
				// �NszInDevName�ഫ��std::string
				int bufferSize = WideCharToMultiByte(CP_UTF8, 0, szInDevName.c_str(), -1, NULL, 0, NULL, NULL) - 1;//�o��ݭn-1�A���᭱��find��勵�T���r���ƶq
				std::string targetName(bufferSize, '\0');
				WideCharToMultiByte(CP_UTF8, 0, szInDevName.c_str(), -1, &targetName[0], bufferSize, NULL, NULL);

				if (deviceName.find(targetName) != std::string::npos) {
					return i; // ���]�t���w�l�r�ꪺ�]�ơA��^�]�� ID
				}
			}
		}
	}
	return -1; // �����]�ơA��^ -1
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
	(void)inputBuffer; // ����ϥ��ܶqĵ�i

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
	// ��l�� COM
	HRESULT hrInit = CoInitialize(nullptr);
	if (FAILED(hrInit)) {
		//std::wcout << L"COM ��l�ƥ��ѡI" << std::endl;
		strMacroResult = "LOG:ERROR_AUDIO_LOOPBACK_COM_INITIALIZE#";
		return false;
	}

	// �إ� DirectShow Graph
	hr = CoCreateInstance(CLSID_FilterGraph, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&pGraph));
	if (FAILED(hr)) {
		strMacroResult = "LOG:ERROR_AUDIO_LOOPBACK_FILTER_GRAPH#";
		//std::wcout << L"�L�k�إ� Filter Graph�I" << std::endl;
		return false;
	}

	// �إ� Capture Graph
	hr = CoCreateInstance(CLSID_CaptureGraphBuilder2, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&pCaptureGraph));
	if (FAILED(hr)) {
		//std::wcout << L"�L�k�إ� Capture Graph�I" << std::endl;
		strMacroResult = "LOG:ERROR_AUDIO_LOOPBACK_CAPTURE_GRAPH#";
		pGraph->Release();
		CoUninitialize();
		return false;
	}

	pCaptureGraph->SetFiltergraph(pGraph);

	// ���ǰt�������˸m
	ICreateDevEnum* pDevEnum = nullptr;
	IEnumMoniker* pEnum = nullptr;
	IMoniker* pMoniker = nullptr;

	CoCreateInstance(CLSID_SystemDeviceEnum, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&pDevEnum));
	if (FAILED(pDevEnum->CreateClassEnumerator(CLSID_AudioInputDeviceCategory, &pEnum, 0)) || !pEnum) {
		//std::wcout << L"���������˸m�I" << std::endl;
		strMacroResult = "LOG:ERROR_AUDIO_WAVEIN_DEVICE_NOT_FIND#";
		goto fail;
	}

	while (pEnum->Next(1, &pMoniker, nullptr) == S_OK) {
		IPropertyBag* pPropBag;
		pMoniker->BindToStorage(nullptr, nullptr, IID_PPV_ARGS(&pPropBag));

		VARIANT varName;
		VariantInit(&varName);
		pPropBag->Read(L"FriendlyName", &varName, nullptr);

		// **�p�G�˸m�W�٥]�t����r�A�N�ϥγo�Ӹ˸m**
		if (wcsstr(varName.bstrVal, captureKeyword.c_str())) {
			hr = pMoniker->BindToObject(nullptr, nullptr, IID_IBaseFilter, (void**)&pAudioCapture);
			if (SUCCEEDED(hr)) {
				VariantClear(&varName);
				pPropBag->Release();
				pMoniker->Release();
				break; // ���ŦX���˸m�N���X�j��
			}
		}
		VariantClear(&varName);
		pPropBag->Release();
		pMoniker->Release();
	}
	pEnum->Release();

	if (!pAudioCapture) {
		//std::wcout << L"�����ŦX�������˸m�I" << std::endl;
		strMacroResult = "LOG:ERROR_AUDIO_WAVEIN_DEVICE_NOT_FIND#";
		goto fail;
	}

	// ���ǰt������˸m
	if (FAILED(pDevEnum->CreateClassEnumerator(CLSID_AudioRendererCategory, &pEnum, 0)) || !pEnum) {
		//std::wcout << L"����켽��˸m�I" << std::endl;
		strMacroResult = "LOG:ERROR_AUDIO_WAVEOUT_DEVICE_NOT_FIND_4#";
		goto fail;
	}
	while (pEnum->Next(1, &pMoniker, nullptr) == S_OK) {
		IPropertyBag* pPropBag;
		pMoniker->BindToStorage(nullptr, nullptr, IID_PPV_ARGS(&pPropBag));

		VARIANT varName;
		VariantInit(&varName);
		pPropBag->Read(L"FriendlyName", &varName, nullptr);

		// **�p�G�˸m�W�٥]�t����r�A�N�ϥγo�Ӹ˸m**
		if (wcsstr(varName.bstrVal, renderKeyword.c_str())) {
			hr = pMoniker->BindToObject(nullptr, nullptr, IID_IBaseFilter, (void**)&pAudioRenderer);
			if (SUCCEEDED(hr)) {
				VariantClear(&varName);
				pPropBag->Release();
				pMoniker->Release();
				break; // ���ŦX���˸m�N���X�j��
			}
		}
		VariantClear(&varName);
		pPropBag->Release();
		pMoniker->Release();
	}
	pEnum->Release();
	pDevEnum->Release();

	if (!pAudioRenderer) {
		//std::wcout << L"�����ŦX������˸m�I" << std::endl;
		strMacroResult = "LOG:ERROR_AUDIO_LOOPBACK_NO_MATCH_RENDER_DEVICE#";
		goto fail;
	}

	// �[�J�˸m�� Filter Graph
	pGraph->AddFilter(pAudioCapture, L"Audio Capture");
	pGraph->AddFilter(pAudioRenderer, L"Audio Renderer");

	// �s�������˸m�켽��˸m
	hr = pCaptureGraph->RenderStream(&PIN_CATEGORY_CAPTURE, &MEDIATYPE_Audio, pAudioCapture, nullptr, pAudioRenderer);
	if (FAILED(hr)) {
		//std::wcout << L"�L�k�إ߿����켽�񪺦�y�I" << std::endl;
		strMacroResult = "LOG:ERROR_AUDIO_LOOPBACK_RENDER_STREAM#";
		goto fail;
	}

	// �����
	pGraph->QueryInterface(IID_PPV_ARGS(&pMediaControl));
	pMediaControl->Run();

	//std::wcout << L"���b��o���T... �� Enter ����" << std::endl;
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

		// ���o�˸m
		if (SUCCEEDED(pDevices->Item(i, &pDevice))) {
			// ���o�˸m�ݩ�
			if (SUCCEEDED(pDevice->OpenPropertyStore(STGM_READ, &pProps))) {
				PROPVARIANT varName;
				PropVariantInit(&varName);

				// ���o�˸m�W��
				if (SUCCEEDED(pProps->GetValue(PKEY_Device_FriendlyName, &varName))) {
					std::wstring deviceName(varName.pwszVal);

					// �ˬd�W�٬O�_�ǰt
					if (deviceName.find(szOutDevName) != std::wstring::npos) {
						PropVariantClear(&varName);
						pProps->Release();
						pDevice->Release();
						return i; // ���ǰt�˸m
					}
				}
				PropVariantClear(&varName);
			}
			if (pProps) pProps->Release();
		}
		if (pDevice) pDevice->Release();
	}
	return -1; // �䤣��˸m
}

// �N std::string �ন std::wstring�]UTF-8 �� UTF-16�^
std::wstring Utf8ToWstring(const std::string& str) {
	if (str.empty()) return std::wstring();
	int size_needed = MultiByteToWideChar(CP_UTF8, 0, &str[0], (int)str.size(), NULL, 0);
	std::wstring wstrTo(size_needed, 0);
	MultiByteToWideChar(CP_UTF8, 0, &str[0], (int)str.size(), &wstrTo[0], size_needed);
	return wstrTo;
}

bool ASAudio::SetDefaultAudioPlaybackDevice(const std::wstring& deviceId)
{
	// �զX����R�O�A�Ҧp�G
	// SoundVolumeView.exe /SetDefault "{deviceId}" 0
	// 0: All, 1: Console, 2: Multimedia
	std::wstring command = L"SoundVolumeView.exe /SetDefault \"" + deviceId + L"\" 0";

	// �ϥ� _wsystem �H�䴩 Unicode �r��
	int result = _wsystem(command.c_str());

	// �^�Ǧ��\�P�_�]0 ��ܦ��\�^
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

	if(count > 0)
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